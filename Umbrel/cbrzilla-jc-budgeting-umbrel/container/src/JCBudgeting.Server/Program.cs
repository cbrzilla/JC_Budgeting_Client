using JCBudgeting.Core;
using Microsoft.AspNetCore.Http.HttpResults;
using Microsoft.Extensions.Options;
using Microsoft.Data.Sqlite;
using System.Text.Json;
using System.Globalization;
using System.Net;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Collections.Concurrent;
using JCBudgeting.Server;

SQLitePCL.Batteries.Init();
ServerAppLogger.Initialize();
ServerAppLogger.LogInfo($"Server startup. BaseDirectory={AppContext.BaseDirectory}");
AppDomain.CurrentDomain.UnhandledException += (_, eventArgs) =>
{
    var exception = eventArgs.ExceptionObject as Exception;
    ServerAppLogger.LogError($"Unhandled AppDomain exception. IsTerminating={eventArgs.IsTerminating}.", exception);
};
TaskScheduler.UnobservedTaskException += (_, eventArgs) =>
{
    ServerAppLogger.LogError("Unobserved task exception.", eventArgs.Exception);
    eventArgs.SetObserved();
};

var builder = WebApplication.CreateBuilder(args);
builder.Configuration.AddJsonFile("budgetserver.local.json", optional: true, reloadOnChange: true);
builder.Logging.ClearProviders();

builder.Services.Configure<BudgetServerOptions>(builder.Configuration.GetSection("BudgetServer"));
builder.Services.AddCors(options =>
{
    options.AddDefaultPolicy(policy =>
        policy.AllowAnyOrigin()
              .AllowAnyHeader()
              .AllowAnyMethod());
});

var app = builder.Build();
var runtimeOptionsMonitor = app.Services.GetRequiredService<IOptionsMonitor<BudgetServerOptions>>();
var runtimeConsole = new ServerRuntimeConsole();
var startupUrls = FirstNonBlank(
    Environment.GetEnvironmentVariable("ASPNETCORE_URLS"),
    builder.Configuration["Urls"],
    "http://0.0.0.0:5099");
var startupOptions = GetEffectiveBudgetServerOptions(runtimeOptionsMonitor.CurrentValue);
WriteServerStartupBanner(startupUrls, startupOptions.DatabasePath);
runtimeConsole.WatchDatabase(startupOptions.DatabasePath);
IDisposable? optionsMonitorSubscription = runtimeOptionsMonitor.OnChange(updatedOptions =>
{
    runtimeConsole.WatchDatabase(GetEffectiveBudgetServerOptions(updatedOptions).DatabasePath);
});
app.Lifetime.ApplicationStopping.Register(() =>
{
    ServerAppLogger.LogInfo("Server is stopping.");
    optionsMonitorSubscription?.Dispose();
    runtimeConsole.Dispose();
});

app.Use(async (context, next) =>
{
    try
    {
        runtimeConsole.LogClientConnection(context);
        await next();

        if (context.Response.StatusCode >= 500)
        {
            ServerAppLogger.LogWarning(
                $"Server returned HTTP {context.Response.StatusCode} for {context.Request.Method} {context.Request.Path} from {NormalizeIpAddress(context.Connection.RemoteIpAddress)}.");
        }
    }
    catch (Exception ex)
    {
        ServerAppLogger.LogError(
            $"Unhandled request exception for {context.Request.Method} {context.Request.Path} from {NormalizeIpAddress(context.Connection.RemoteIpAddress)}.",
            ex);
        throw;
    }
});

app.UseCors();
app.UseDefaultFiles();
app.UseStaticFiles(new StaticFileOptions
{
    OnPrepareResponse = context =>
    {
        context.Context.Response.Headers["Cache-Control"] = "no-store, no-cache, must-revalidate";
        context.Context.Response.Headers["Pragma"] = "no-cache";
        context.Context.Response.Headers["Expires"] = "0";
    }
});

app.MapGet("/setup", () => TypedResults.Redirect("/setup.html"));

app.MapGet("/api/server-config", (IOptionsMonitor<BudgetServerOptions> options, IConfiguration configuration) =>
{
    var configPath = GetStandaloneServerConfigPath();
    var currentOptions = GetEffectiveBudgetServerOptions(options.CurrentValue);
    var resolved = ResolvePaths(currentOptions, null, null);
    var effectiveSettings = BuildWorkspaceSettings(currentOptions, resolved.SettingsPath, resolved.DatabasePath);
    var urls = FirstNonBlank(
        Environment.GetEnvironmentVariable("ASPNETCORE_URLS"),
        configuration["Urls"],
        "http://0.0.0.0:5099");

    return TypedResults.Ok(new ServerConfigResponse
    {
        ConfigPath = configPath,
        DatabasePath = resolved.DatabasePath,
        SettingsPath = string.Empty,
        BudgetPeriod = effectiveSettings.BudgetPeriod,
        BudgetStartDate = effectiveSettings.BudgetStartDate,
        BudgetYears = effectiveSettings.BudgetYears,
        Urls = urls,
        AccessUrls = GetSuggestedAccessUrls(urls),
        ConfigFileExists = File.Exists(configPath),
        DatabaseExists = !string.IsNullOrWhiteSpace(resolved.DatabasePath) && File.Exists(resolved.DatabasePath),
        SettingsExists = !string.IsNullOrWhiteSpace(resolved.SettingsPath) && File.Exists(resolved.SettingsPath),
        Databases = GetAvailableDatabaseFiles(resolved.DatabasePath)
    });
});

app.MapPost("/api/server-config", Results<Ok<ServerConfigSaveResponse>, BadRequest<ApiErrorResponse>> (
    ServerConfigSaveRequest request) =>
{
    try
    {
        var saveResult = SaveStandaloneServerConfig(request);
        runtimeConsole.WatchDatabase(saveResult.DatabasePath);

        return TypedResults.Ok(new ServerConfigSaveResponse
        {
            Message = "Saved standalone server configuration. Database and budget settings are live now. Restart only if you changed access mode or port.",
            ConfigPath = saveResult.ConfigPath,
            DatabasePath = saveResult.DatabasePath,
            SettingsPath = string.Empty,
            Urls = saveResult.Urls,
            AccessUrls = GetSuggestedAccessUrls(saveResult.Urls)
        });
    }
    catch (InvalidOperationException ex)
    {
        ServerAppLogger.LogError("Failed to save standalone server configuration.", ex);
        return TypedResults.BadRequest(new ApiErrorResponse
        {
            Error = ex.Message
        });
    }
});

app.MapPost("/api/server-config/create-database", Results<Ok<ServerConfigCreateDatabaseResponse>, BadRequest<ApiErrorResponse>> (
    ServerConfigSaveRequest request) =>
{
    try
    {
        var databasePath = ResolveCreateDatabaseTargetPath(request.DatabasePath);
        if (string.IsNullOrWhiteSpace(databasePath))
        {
            return TypedResults.BadRequest(new ApiErrorResponse
            {
                Error = "Database path is required to create a new database."
            });
        }

        databasePath = EnsureUniqueDatabasePath(databasePath);

        CreateDatabaseFromPackagedTemplate(databasePath);
        var saveResult = SaveStandaloneServerConfig(new ServerConfigSaveRequest
        {
            DatabasePath = databasePath,
            BudgetPeriod = request.BudgetPeriod,
            BudgetStartDate = request.BudgetStartDate,
            BudgetYears = request.BudgetYears,
            Urls = request.Urls
        });
        runtimeConsole.WatchDatabase(saveResult.DatabasePath);

        return TypedResults.Ok(new ServerConfigCreateDatabaseResponse
        {
            Message = "Created a new database from the blank template and saved the standalone server configuration. Database and budget settings are live now. Restart only if you changed access mode or port.",
            ConfigPath = saveResult.ConfigPath,
            DatabasePath = saveResult.DatabasePath,
            SettingsPath = string.Empty,
            Urls = saveResult.Urls,
            AccessUrls = GetSuggestedAccessUrls(saveResult.Urls)
        });
    }
    catch (InvalidOperationException ex)
    {
        ServerAppLogger.LogError("Failed to create and save a standalone server database.", ex);
        return TypedResults.BadRequest(new ApiErrorResponse
        {
            Error = ex.Message
        });
    }
});

app.MapPost("/api/server-config/database/upload", async Task<Results<Ok<ServerConfigDatabaseTransferResponse>, BadRequest<ApiErrorResponse>>> (
    HttpRequest request,
    IOptionsMonitor<BudgetServerOptions> options) =>
{
    try
    {
        if (!request.HasFormContentType)
        {
            return TypedResults.BadRequest(new ApiErrorResponse
            {
                Error = "Upload must use multipart form data."
            });
        }

        var form = await request.ReadFormAsync();
        var file = form.Files.FirstOrDefault();
        if (file is null || file.Length <= 0)
        {
            return TypedResults.BadRequest(new ApiErrorResponse
            {
                Error = "Select a .jcbdb file to upload."
            });
        }

        var extension = Path.GetExtension(file.FileName);
        if (!string.Equals(extension, ".jcbdb", StringComparison.OrdinalIgnoreCase))
        {
            return TypedResults.BadRequest(new ApiErrorResponse
            {
                Error = "Only .jcbdb database files can be uploaded."
            });
        }

        var currentOptions = GetEffectiveBudgetServerOptions(options.CurrentValue);
        var targetPath = ResolveUploadDatabaseTargetPath(file.FileName);
        if (string.IsNullOrWhiteSpace(targetPath))
        {
            return TypedResults.BadRequest(new ApiErrorResponse
            {
                Error = "The server does not have a valid target database path configured."
            });
        }

        var targetDirectory = Path.GetDirectoryName(targetPath);
        if (!string.IsNullOrWhiteSpace(targetDirectory))
        {
            Directory.CreateDirectory(targetDirectory);
        }

        await using (var sourceStream = file.OpenReadStream())
        await using (var destinationStream = File.Create(targetPath))
        {
            await sourceStream.CopyToAsync(destinationStream);
        }

        var uploadedSettings = new BudgetWorkspaceSettings
        {
            BudgetPeriod = currentOptions.BudgetPeriod,
            BudgetStartDate = currentOptions.BudgetStartDate,
            BudgetYears = currentOptions.BudgetYears
        };
        BudgetWorkspaceService.TryApplyBudgetTimelineSettingsFromDatabase(targetPath, uploadedSettings);

        var saveResult = SaveStandaloneServerConfig(new ServerConfigSaveRequest
        {
            DatabasePath = targetPath,
            BudgetPeriod = uploadedSettings.BudgetPeriod,
            BudgetStartDate = uploadedSettings.BudgetStartDate,
            BudgetYears = uploadedSettings.BudgetYears,
            Urls = FirstNonBlank(
                Environment.GetEnvironmentVariable("ASPNETCORE_URLS"),
                builder.Configuration["Urls"],
                "http://0.0.0.0:5099")
        });
        runtimeConsole.WatchDatabase(saveResult.DatabasePath);

        return TypedResults.Ok(new ServerConfigDatabaseTransferResponse
        {
            Message = $"Uploaded database to {targetPath}.",
            DatabasePath = saveResult.DatabasePath
        });
    }
    catch (InvalidOperationException ex)
    {
        ServerAppLogger.LogError("Failed to upload a standalone server database.", ex);
        return TypedResults.BadRequest(new ApiErrorResponse
        {
            Error = ex.Message
        });
    }
});

app.MapGet("/api/server-config/database/download", async Task<Results<FileContentHttpResult, BadRequest<ApiErrorResponse>>> (
    IOptionsMonitor<BudgetServerOptions> options) =>
{
    try
    {
        var currentOptions = GetEffectiveBudgetServerOptions(options.CurrentValue);
        var resolved = ResolvePaths(currentOptions, null, null);
        if (string.IsNullOrWhiteSpace(resolved.DatabasePath) || !File.Exists(resolved.DatabasePath))
        {
            return TypedResults.BadRequest(new ApiErrorResponse
            {
                Error = "The server does not currently have a database file available to download."
            });
        }

        byte[] bytes;
        await using (var stream = new FileStream(
            resolved.DatabasePath,
            FileMode.Open,
            FileAccess.Read,
            FileShare.ReadWrite | FileShare.Delete))
        {
            bytes = new byte[stream.Length];
            var offset = 0;
            while (offset < bytes.Length)
            {
                var read = await stream.ReadAsync(bytes, offset, bytes.Length - offset);
                if (read <= 0)
                {
                    break;
                }

                offset += read;
            }
        }

        var downloadName = Path.GetFileName(resolved.DatabasePath);
        if (string.IsNullOrWhiteSpace(downloadName))
        {
            downloadName = "server_budget_database.jcbdb";
        }

        return TypedResults.File(bytes, "application/octet-stream", downloadName);
    }
    catch (InvalidOperationException ex)
    {
        ServerAppLogger.LogError("Failed to download the standalone server database.", ex);
        return TypedResults.BadRequest(new ApiErrorResponse
        {
            Error = ex.Message
        });
    }
});

app.MapGet("/api/health", (IOptions<BudgetServerOptions> options) =>
{
    var resolved = ResolvePaths(options.Value, null, null);
    var dbExists = !string.IsNullOrWhiteSpace(resolved.DatabasePath) && File.Exists(resolved.DatabasePath);
    var settingsExists = !string.IsNullOrWhiteSpace(resolved.SettingsPath) && File.Exists(resolved.SettingsPath);
    var usesInlineBudgetSettings = !settingsExists &&
                                  (!string.IsNullOrWhiteSpace(options.Value.BudgetPeriod) ||
                                   !string.IsNullOrWhiteSpace(options.Value.BudgetStartDate) ||
                                   options.Value.BudgetYears > 0);

    return TypedResults.Ok(new
    {
        status = "ok",
        configuredDatabasePath = resolved.DatabasePath,
        configuredSettingsPath = resolved.SettingsPath,
        databaseExists = dbExists,
        settingsExists = settingsExists,
        usesInlineBudgetSettings,
        urls = app.Urls,
        generatedAt = DateTimeOffset.Now
    });
});

app.MapGet("/api/change-token", Results<Ok<ChangeTokenResponse>, BadRequest<ApiErrorResponse>> (
    string? databasePath,
    string? settingsPath,
    IOptions<BudgetServerOptions> options) =>
{
    var resolved = ResolvePaths(options.Value, databasePath, settingsPath);
    if (string.IsNullOrWhiteSpace(resolved.DatabasePath) || !File.Exists(resolved.DatabasePath))
    {
        return TypedResults.BadRequest(new ApiErrorResponse
        {
            Error = "A valid budget database could not be found."
        });
    }

    var fileInfo = new FileInfo(resolved.DatabasePath);
    var lastWriteUtc = fileInfo.LastWriteTimeUtc;
    var length = fileInfo.Exists ? fileInfo.Length : 0L;
    return TypedResults.Ok(new ChangeTokenResponse
    {
        DatabasePath = resolved.DatabasePath,
        LastWriteUtc = lastWriteUtc,
        Length = length,
        Token = $"{lastWriteUtc.Ticks}:{length}"
    });
});

app.MapGet("/api/overview/summary", Results<Ok<OverviewSummaryResponse>, BadRequest<ApiErrorResponse>> (
    string? databasePath,
    string? settingsPath,
    IOptions<BudgetServerOptions> options) =>
{
    var contextResult = TryBuildSnapshotContext(options.Value, databasePath, settingsPath);
    if (contextResult.Error is not null)
    {
        return TypedResults.BadRequest(contextResult.Error);
    }

    var resolved = contextResult.Paths!;
    var snapshot = contextResult.Snapshot!;
    if (snapshot.PeriodSummaries.Count == 0)
    {
        return TypedResults.Ok(new OverviewSummaryResponse
        {
            DatabasePath = resolved.DatabasePath,
            SettingsPath = resolved.SettingsPath,
            GeneratedAt = DateTimeOffset.Now,
            BudgetPeriod = snapshot.BudgetPeriod,
            BudgetYears = snapshot.BudgetYears
        });
    }

    var currentIndex = Math.Clamp(snapshot.CurrentPeriodIndex, 0, snapshot.PeriodSummaries.Count - 1);
    var current = snapshot.PeriodSummaries[currentIndex];
    var currentDebtBalance = snapshot.DebtDisplayBalances.Sum(item =>
        currentIndex < item.Value.Length ? item.Value[currentIndex] : 0m);
    var lowestAccount = snapshot.AccountRunningBalances
        .Select(entry =>
        {
            var values = entry.Value;
            if (values.Length == 0)
            {
                return null;
            }

            var lowest = values.Min();
            var lowestIndex = Array.IndexOf(values, lowest);
            return new
            {
                Label = entry.Key,
                Value = lowest,
                Index = lowestIndex
            };
        })
        .Where(entry => entry is not null)
        .OrderBy(entry => entry!.Value)
        .FirstOrDefault();

    var response = new OverviewSummaryResponse
    {
        DatabasePath = resolved.DatabasePath,
        SettingsPath = resolved.SettingsPath,
        GeneratedAt = DateTimeOffset.Now,
        BudgetPeriod = snapshot.BudgetPeriod,
        BudgetYears = snapshot.BudgetYears,
        CurrentPeriodIndex = currentIndex,
        CurrentPeriodStart = current.PeriodStart,
        TotalPeriodCount = snapshot.TotalPeriodCount,
        IncomeTotal = current.IncomeTotal,
        PlannedOutflow = current.ExpenseTotal + current.DebtTotal,
        ExpenseTotal = current.ExpenseTotal,
        DebtPaymentTotal = current.DebtTotal,
        SavingsContributionTotal = current.SavingsTotal,
        NetFlow = current.NetCashFlow,
        CurrentDebtBalance = currentDebtBalance,
        AccountCount = snapshot.AccountRunningBalances.Count,
        SavingsCount = snapshot.SavingsRunningBalances.Count,
        DebtCount = snapshot.DebtRunningBalances.Count,
        LowestAccount = lowestAccount is null ? null : new OverviewSummaryAccountLow
        {
            Label = lowestAccount.Label,
            Value = lowestAccount.Value,
            PeriodStart = snapshot.PeriodSummaries[Math.Clamp(lowestAccount.Index, 0, snapshot.PeriodSummaries.Count - 1)].PeriodStart
        }
    };

    return TypedResults.Ok(response);
});

app.MapGet("/api/periods", Results<Ok<IReadOnlyList<BudgetPeriodResponse>>, BadRequest<ApiErrorResponse>> (
    string? databasePath,
    string? settingsPath,
    IOptions<BudgetServerOptions> options) =>
{
    var contextResult = TryBuildSnapshotContext(options.Value, databasePath, settingsPath);
    if (contextResult.Error is not null)
    {
        return TypedResults.BadRequest(contextResult.Error);
    }

    var snapshot = contextResult.Snapshot!;
    var currentIndex = snapshot.PeriodSummaries.Count == 0
        ? 0
        : Math.Clamp(snapshot.CurrentPeriodIndex, 0, snapshot.PeriodSummaries.Count - 1);

    var response = snapshot.PeriodSummaries
        .Select((period, index) => new BudgetPeriodResponse
        {
            Index = index,
            Start = period.PeriodStart,
            Label = index == currentIndex
                ? $"{period.PeriodStart:MM/dd/yyyy} (Current)"
                : period.PeriodStart.ToString("MM/dd/yyyy"),
            IsCurrent = index == currentIndex
        })
        .ToList();

    return TypedResults.Ok((IReadOnlyList<BudgetPeriodResponse>)response);
});

app.MapGet("/api/desktop/budget-workspace", Results<Ok<DesktopBudgetWorkspaceResponse>, BadRequest<ApiErrorResponse>> (
    int? periodCount,
    string? databasePath,
    string? settingsPath,
    IOptionsMonitor<BudgetServerOptions> options) =>
{
    var currentOptions = GetEffectiveBudgetServerOptions(options.CurrentValue);
    var resolved = ResolvePaths(currentOptions, databasePath, settingsPath);
    if (string.IsNullOrWhiteSpace(resolved.DatabasePath) || !File.Exists(resolved.DatabasePath))
    {
        return TypedResults.BadRequest(new ApiErrorResponse
        {
            Error = "A valid budget database could not be found."
        });
    }

    var settings = BuildWorkspaceSettings(currentOptions, resolved.SettingsPath, resolved.DatabasePath);
    var normalizedPeriodCount = periodCount.GetValueOrDefault();
    var snapshot = normalizedPeriodCount > 0
        ? BudgetWorkspaceService.BuildSnapshot(resolved.DatabasePath, settings, normalizedPeriodCount)
        : BudgetWorkspaceService.BuildSnapshot(resolved.DatabasePath, settings);

    return TypedResults.Ok(new DesktopBudgetWorkspaceResponse
    {
        DatabasePath = resolved.DatabasePath,
        SettingsPath = resolved.SettingsPath,
        Settings = settings,
        Snapshot = snapshot
    });
});

app.MapGet("/api/debts", Results<Ok<IReadOnlyList<DebtSummaryResponse>>, BadRequest<ApiErrorResponse>> (
    string? databasePath,
    string? settingsPath,
    IOptionsMonitor<BudgetServerOptions> options) =>
{
    var contextResult = TryBuildSnapshotContext(options.CurrentValue, databasePath, settingsPath);
    if (contextResult.Error is not null)
    {
        return TypedResults.BadRequest(contextResult.Error);
    }

    var resolved = contextResult.Paths!;
    var snapshot = contextResult.Snapshot!;
    var currentIndex = snapshot.PeriodSummaries.Count == 0
        ? 0
        : Math.Clamp(snapshot.CurrentPeriodIndex, 0, snapshot.PeriodSummaries.Count - 1);

    var debts = DebtRepository.LoadDebts(resolved.DatabasePath);
    var response = debts.Select(debt =>
    {
        var key = string.IsNullOrWhiteSpace(debt.Description) ? $"Debt {debt.Id}" : debt.Description.Trim();
        snapshot.DebtRunningBalances.TryGetValue(key, out var runningBalances);
        snapshot.DebtDisplayBalances.TryGetValue(key, out var displayBalances);
        runningBalances ??= Array.Empty<decimal>();
        displayBalances ??= runningBalances;
        var currentBalance = currentIndex < displayBalances.Length
            ? displayBalances[currentIndex]
            : displayBalances.Length > 0 ? displayBalances[^1] : 0m;
        var currentBalanceOverrideAmount = GetCurrentDebtBalanceOverrideAmount(snapshot, resolved.DatabasePath, debt);
        if (currentBalanceOverrideAmount.HasValue)
        {
            currentBalance = currentBalanceOverrideAmount.Value;
        }
        var payoffIndex = FindProjectedPayoffIndex(runningBalances, currentIndex);
        var deleteBlockingReason = BuildDebtDeleteBlockingReason(resolved.DatabasePath, debt.Description);

        return CreateDebtSummaryResponse(debt, currentBalance, payoffIndex, snapshot, deleteBlockingReason);
    })
    .OrderBy(item => item.Description, StringComparer.CurrentCultureIgnoreCase)
    .ToList();

    return TypedResults.Ok((IReadOnlyList<DebtSummaryResponse>)response);
});

app.MapPost("/api/debts/save", async Task<Results<Ok<DebtSummaryResponse>, BadRequest<ApiErrorResponse>>> (
    DebtSaveRequest request,
    IOptionsMonitor<BudgetServerOptions> options) =>
{
    var contextResult = TryBuildSnapshotContext(options.CurrentValue, null, null);
    if (contextResult.Error is not null)
    {
        return TypedResults.BadRequest(contextResult.Error);
    }

    try
    {
        var resolved = contextResult.Paths!;
        var debtRecord = new DebtRecord
        {
            Id = request.Id,
            Description = request.Description?.Trim() ?? string.Empty,
            Category = request.Category?.Trim() ?? string.Empty,
            DebtType = request.DebtType?.Trim() ?? string.Empty,
            Lender = request.Lender?.Trim() ?? string.Empty,
            Apr = request.Apr,
            StartingBalance = request.StartingBalance,
            OriginalPrincipal = request.OriginalPrincipal,
            MinPayment = request.MinPayment,
            DayDue = request.DayDue,
            FromAccount = request.FromAccount?.Trim() ?? string.Empty,
            Hidden = request.IsHidden,
            Active = request.IsActive,
            LoginLink = request.LoginLink?.Trim() ?? string.Empty,
            Notes = request.Notes ?? string.Empty,
            Cadence = request.Cadence?.Trim() ?? string.Empty,
            SameAs = request.SameAs?.Trim() ?? string.Empty,
            StartDate = request.StartDate?.Trim() ?? string.Empty,
            LastPaymentDate = request.LastPaymentDate?.Trim() ?? string.Empty,
            TermMonths = request.TermMonths,
            MaturityDate = request.MaturityDate?.Trim() ?? string.Empty,
            PromoApr = request.PromoApr,
            PromoStartDate = request.PromoStartDate?.Trim() ?? string.Empty,
            PromoAprEndDate = request.PromoAprEndDate?.Trim() ?? string.Empty,
            CreditLimit = request.CreditLimit,
            EscrowIncluded = request.EscrowIncluded,
            EscrowMonthly = request.EscrowMonthly,
            PmiMonthly = request.PmiMonthly,
            DeferredUntil = request.DeferredUntil?.Trim() ?? string.Empty,
            DeferredStatus = request.DeferredStatus,
            Subsidized = request.Subsidized,
            BalloonAmount = request.BalloonAmount,
            BalloonDueDate = request.BalloonDueDate?.Trim() ?? string.Empty,
            InterestOnlyStartDate = request.InterestOnlyStartDate?.Trim() ?? string.Empty,
            InterestOnlyEndDate = request.InterestOnlyEndDate?.Trim() ?? string.Empty,
            ForgivenessDate = request.ForgivenessDate?.Trim() ?? string.Empty,
            StudentRepaymentPlan = request.StudentRepaymentPlan?.Trim() ?? string.Empty,
            RateChangeSchedule = request.RateChangeSchedule?.Trim() ?? string.Empty,
            CustomInterestRule = request.CustomInterestRule?.Trim() ?? string.Empty,
            CustomFeeRule = request.CustomFeeRule?.Trim() ?? string.Empty,
            DayCountBasis = request.DayCountBasis,
            PaymentsPerYear = request.PaymentsPerYear
        };

        if (string.IsNullOrWhiteSpace(debtRecord.Description))
        {
            return TypedResults.BadRequest(new ApiErrorResponse { Error = "Debt description is required." });
        }

        if (string.IsNullOrWhiteSpace(debtRecord.Category))
        {
            return TypedResults.BadRequest(new ApiErrorResponse { Error = "Debt category is required." });
        }

        if (debtRecord.StartingBalance < 0m)
        {
            return TypedResults.BadRequest(new ApiErrorResponse { Error = "Starting balance cannot be negative." });
        }

        var existingDebt = debtRecord.Id > 0
            ? DebtRepository.LoadDebts(resolved.DatabasePath).FirstOrDefault(item => item.Id == debtRecord.Id)
            : null;

        if (debtRecord.Id <= 0 || existingDebt is null)
        {
            debtRecord.Id = DebtRepository.CreateDebt(resolved.DatabasePath, debtRecord);
        }
        else
        {
            DebtRepository.SaveDebt(resolved.DatabasePath, debtRecord);
        }

        var previousDescription = request.PreviousDescription?.Trim() ?? string.Empty;
        var updatedDescription = debtRecord.Description.Trim();
        if (!string.IsNullOrWhiteSpace(previousDescription) &&
            !string.Equals(previousDescription, updatedDescription, StringComparison.CurrentCulture))
        {
            RenameDebtFundingReferences(resolved.DatabasePath, previousDescription, updatedDescription);
        }

        PreservePastDebtValues(resolved.DatabasePath, contextResult.Snapshot!, existingDebt ?? debtRecord, debtRecord);
        ApplyCurrentDebtBalanceOverride(resolved.DatabasePath, contextResult.Snapshot!, debtRecord, request.CurrentPeriodBalanceOverrideAmount);

        var freshContext = TryBuildSnapshotContext(options.CurrentValue, null, null);
        if (freshContext.Error is not null)
        {
            return TypedResults.BadRequest(freshContext.Error);
        }

        var freshPaths = freshContext.Paths!;
        var freshSnapshot = freshContext.Snapshot!;
        var currentIndex = freshSnapshot.PeriodSummaries.Count == 0
            ? 0
            : Math.Clamp(freshSnapshot.CurrentPeriodIndex, 0, freshSnapshot.PeriodSummaries.Count - 1);
        var key = string.IsNullOrWhiteSpace(debtRecord.Description) ? $"Debt {debtRecord.Id}" : debtRecord.Description.Trim();
        freshSnapshot.DebtRunningBalances.TryGetValue(key, out var runningBalances);
        freshSnapshot.DebtDisplayBalances.TryGetValue(key, out var displayBalances);
        runningBalances ??= Array.Empty<decimal>();
        displayBalances ??= runningBalances;
        var currentBalance = currentIndex < displayBalances.Length
            ? displayBalances[currentIndex]
            : displayBalances.Length > 0 ? displayBalances[^1] : 0m;
        var currentBalanceOverrideAmount = GetCurrentDebtBalanceOverrideAmount(freshSnapshot, freshPaths.DatabasePath, debtRecord);
        if (currentBalanceOverrideAmount.HasValue)
        {
            currentBalance = currentBalanceOverrideAmount.Value;
        }

        var payoffIndex = FindProjectedPayoffIndex(runningBalances, currentIndex);
        var deleteBlockingReason = BuildDebtDeleteBlockingReason(freshPaths.DatabasePath, debtRecord.Description);
        return TypedResults.Ok(CreateDebtSummaryResponse(debtRecord, currentBalance, payoffIndex, freshSnapshot, deleteBlockingReason));
    }
    catch (Exception ex)
    {
        ServerAppLogger.LogError("Failed to save debt.", ex);
        return TypedResults.BadRequest(new ApiErrorResponse { Error = ex.Message });
    }
});

app.MapDelete("/api/debts/{debtId:int}", Results<Ok<DebtDeleteResponse>, BadRequest<ApiErrorResponse>> (
    int debtId,
    IOptionsMonitor<BudgetServerOptions> options) =>
{
    var contextResult = TryBuildSnapshotContext(options.CurrentValue, null, null);
    if (contextResult.Error is not null)
    {
        return TypedResults.BadRequest(contextResult.Error);
    }

    try
    {
        var resolved = contextResult.Paths!;
        var debt = DebtRepository.LoadDebts(resolved.DatabasePath).FirstOrDefault(item => item.Id == debtId);
        if (debt is null)
        {
            return TypedResults.BadRequest(new ApiErrorResponse { Error = "The selected debt could not be found on the server." });
        }

        var deleteBlockingReason = BuildDebtDeleteBlockingReason(resolved.DatabasePath, debt.Description);
        if (!string.IsNullOrWhiteSpace(deleteBlockingReason))
        {
            return TypedResults.BadRequest(new ApiErrorResponse { Error = deleteBlockingReason });
        }

        DebtRepository.DeleteDebt(resolved.DatabasePath, debtId);
        BudgetOverrideHistoryRepository.DeleteLiveOverrides(resolved.DatabasePath, 1, BuildDebtOverrideKey(debt), debt.Description);
        BudgetOverrideHistoryRepository.DeleteLiveOverrides(resolved.DatabasePath, -101, BuildDebtOverrideKey(debt), debt.Description);

        return TypedResults.Ok(new DebtDeleteResponse
        {
            Id = debtId,
            Description = debt.Description
        });
    }
    catch (Exception ex)
    {
        ServerAppLogger.LogError("Failed to delete debt.", ex);
        return TypedResults.BadRequest(new ApiErrorResponse { Error = ex.Message });
    }
});

app.MapGet("/api/savings", Results<Ok<IReadOnlyList<SavingsSummaryResponse>>, BadRequest<ApiErrorResponse>> (
    string? databasePath,
    string? settingsPath,
    IOptionsMonitor<BudgetServerOptions> options) =>
{
    var contextResult = TryBuildSnapshotContext(options.CurrentValue, databasePath, settingsPath);
    if (contextResult.Error is not null)
    {
        return TypedResults.BadRequest(contextResult.Error);
    }

    var resolved = contextResult.Paths!;
    var snapshot = contextResult.Snapshot!;
    var currentIndex = snapshot.PeriodSummaries.Count == 0
        ? 0
        : Math.Clamp(snapshot.CurrentPeriodIndex, 0, snapshot.PeriodSummaries.Count - 1);

    var savings = SavingsRepository.LoadSavings(resolved.DatabasePath);
    var response = savings.Select(item =>
    {
        var key = string.IsNullOrWhiteSpace(item.Description) ? $"Savings {item.Id}" : item.Description.Trim();
        snapshot.SavingsRunningBalances.TryGetValue(key, out var balances);
        balances ??= Array.Empty<decimal>();
        var currentBalance = currentIndex < balances.Length
            ? balances[currentIndex]
            : balances.Length > 0 ? balances[^1] : 0m;
        var deleteBlockingReason = BuildSavingsDeleteBlockingReason(resolved.DatabasePath, item.Description);

        return new SavingsSummaryResponse
        {
            Id = item.Id,
            Description = item.Description,
            GoalAmount = item.GoalAmount,
            DepositAmount = item.DepositAmount,
            GoalDate = item.GoalDate,
            HasGoal = item.HasGoal,
            Frequency = item.Frequency,
            OnDay = item.OnDay,
            OnDate = item.OnDate,
            StartDate = item.StartDate,
            EndDate = item.EndDate,
            FromAccount = item.FromAccount,
            SameAs = item.SameAs,
            Category = item.Category,
            CurrentBalance = currentBalance,
            IsHidden = item.Hidden,
            IsActive = item.Active,
            LoginLink = item.LoginLink,
            Notes = item.Notes,
            CanDelete = string.IsNullOrWhiteSpace(deleteBlockingReason),
            DeleteBlockedReason = deleteBlockingReason
        };
    })
    .OrderBy(item => item.Description, StringComparer.CurrentCultureIgnoreCase)
    .ToList();

    return TypedResults.Ok((IReadOnlyList<SavingsSummaryResponse>)response);
});

app.MapPost("/api/savings/save", async Task<Results<Ok<SavingsSummaryResponse>, BadRequest<ApiErrorResponse>>> (
    SavingsSaveRequest request,
    IOptionsMonitor<BudgetServerOptions> options) =>
{
    var contextResult = TryBuildSnapshotContext(options.CurrentValue, null, null);
    if (contextResult.Error is not null)
    {
        return TypedResults.BadRequest(contextResult.Error);
    }

    try
    {
        var resolved = contextResult.Paths!;
        var savingsRecord = new SavingsRecord
        {
            Id = request.Id,
            Description = request.Description?.Trim() ?? string.Empty,
            GoalAmount = request.GoalAmount,
            DepositAmount = request.DepositAmount,
            GoalDate = request.GoalDate?.Trim() ?? string.Empty,
            HasGoal = request.HasGoal,
            Frequency = request.Frequency?.Trim() ?? string.Empty,
            OnDay = request.OnDay,
            OnDate = request.OnDate?.Trim() ?? string.Empty,
            StartDate = request.StartDate?.Trim() ?? string.Empty,
            EndDate = request.EndDate?.Trim() ?? string.Empty,
            FromAccount = request.FromAccount?.Trim() ?? string.Empty,
            SameAs = request.SameAs?.Trim() ?? string.Empty,
            Category = request.Category?.Trim() ?? string.Empty,
            Hidden = request.IsHidden,
            Active = request.IsActive,
            LoginLink = request.LoginLink?.Trim() ?? string.Empty,
            Notes = request.Notes ?? string.Empty
        };

        if (string.IsNullOrWhiteSpace(savingsRecord.Description))
        {
            return TypedResults.BadRequest(new ApiErrorResponse { Error = "Savings description is required." });
        }

        if (string.IsNullOrWhiteSpace(savingsRecord.Category))
        {
            return TypedResults.BadRequest(new ApiErrorResponse { Error = "Savings category is required." });
        }

        if (string.IsNullOrWhiteSpace(savingsRecord.Frequency))
        {
            return TypedResults.BadRequest(new ApiErrorResponse { Error = "Savings frequency is required." });
        }

        if (string.IsNullOrWhiteSpace(savingsRecord.FromAccount))
        {
            return TypedResults.BadRequest(new ApiErrorResponse { Error = "From account is required." });
        }

        if (savingsRecord.GoalAmount < 0m)
        {
            return TypedResults.BadRequest(new ApiErrorResponse { Error = "Goal amount cannot be negative." });
        }

        if (savingsRecord.Id <= 0)
        {
            savingsRecord.Id = SavingsRepository.CreateSavings(resolved.DatabasePath, savingsRecord);
        }
        else
        {
            SavingsRepository.SaveSavings(resolved.DatabasePath, savingsRecord);
        }

        var previousDescription = request.PreviousDescription?.Trim() ?? string.Empty;
        var updatedDescription = savingsRecord.Description.Trim();
        if (!string.IsNullOrWhiteSpace(previousDescription) &&
            !string.Equals(previousDescription, updatedDescription, StringComparison.CurrentCulture))
        {
            RenameSavingsFundingReferences(resolved.DatabasePath, previousDescription, updatedDescription);
        }

        var freshContext = TryBuildSnapshotContext(options.CurrentValue, null, null);
        if (freshContext.Error is not null)
        {
            return TypedResults.BadRequest(freshContext.Error);
        }

        var freshPaths = freshContext.Paths!;
        var freshSnapshot = freshContext.Snapshot!;
        var currentIndex = freshSnapshot.PeriodSummaries.Count == 0
            ? 0
            : Math.Clamp(freshSnapshot.CurrentPeriodIndex, 0, freshSnapshot.PeriodSummaries.Count - 1);
        var key = string.IsNullOrWhiteSpace(savingsRecord.Description) ? $"Savings {savingsRecord.Id}" : savingsRecord.Description.Trim();
        freshSnapshot.SavingsRunningBalances.TryGetValue(key, out var balances);
        balances ??= Array.Empty<decimal>();
        var currentBalance = currentIndex < balances.Length
            ? balances[currentIndex]
            : balances.Length > 0 ? balances[^1] : 0m;
        var deleteBlockingReason = BuildSavingsDeleteBlockingReason(freshPaths.DatabasePath, savingsRecord.Description);

        return TypedResults.Ok(new SavingsSummaryResponse
        {
            Id = savingsRecord.Id,
            Description = savingsRecord.Description,
            GoalAmount = savingsRecord.GoalAmount,
            DepositAmount = savingsRecord.DepositAmount,
            GoalDate = savingsRecord.GoalDate,
            HasGoal = savingsRecord.HasGoal,
            Frequency = savingsRecord.Frequency,
            OnDay = savingsRecord.OnDay,
            OnDate = savingsRecord.OnDate,
            StartDate = savingsRecord.StartDate,
            EndDate = savingsRecord.EndDate,
            FromAccount = savingsRecord.FromAccount,
            SameAs = savingsRecord.SameAs,
            Category = savingsRecord.Category,
            CurrentBalance = currentBalance,
            IsHidden = savingsRecord.Hidden,
            IsActive = savingsRecord.Active,
            LoginLink = savingsRecord.LoginLink,
            Notes = savingsRecord.Notes,
            CanDelete = string.IsNullOrWhiteSpace(deleteBlockingReason),
            DeleteBlockedReason = deleteBlockingReason
        });
    }
    catch (Exception ex)
    {
        ServerAppLogger.LogError("Failed to save savings.", ex);
        return TypedResults.BadRequest(new ApiErrorResponse { Error = ex.Message });
    }
});

app.MapDelete("/api/savings/{savingsId:int}", Results<Ok<SavingsDeleteResponse>, BadRequest<ApiErrorResponse>> (
    int savingsId,
    IOptionsMonitor<BudgetServerOptions> options) =>
{
    var contextResult = TryBuildSnapshotContext(options.CurrentValue, null, null);
    if (contextResult.Error is not null)
    {
        return TypedResults.BadRequest(contextResult.Error);
    }

    try
    {
        var resolved = contextResult.Paths!;
        var savings = SavingsRepository.LoadSavings(resolved.DatabasePath).FirstOrDefault(item => item.Id == savingsId);
        if (savings is null)
        {
            return TypedResults.BadRequest(new ApiErrorResponse { Error = "The selected savings item could not be found on the server." });
        }

        var deleteBlockingReason = BuildSavingsDeleteBlockingReason(resolved.DatabasePath, savings.Description);
        if (!string.IsNullOrWhiteSpace(deleteBlockingReason))
        {
            return TypedResults.BadRequest(new ApiErrorResponse { Error = deleteBlockingReason });
        }

        SavingsRepository.DeleteSavings(resolved.DatabasePath, savingsId);
        return TypedResults.Ok(new SavingsDeleteResponse
        {
            Id = savingsId,
            Description = savings.Description
        });
    }
    catch (Exception ex)
    {
        ServerAppLogger.LogError("Failed to delete savings.", ex);
        return TypedResults.BadRequest(new ApiErrorResponse { Error = ex.Message });
    }
});

app.MapGet("/api/accounts", Results<Ok<IReadOnlyList<AccountSummaryResponse>>, BadRequest<ApiErrorResponse>> (
    string? databasePath,
    string? settingsPath,
    IOptionsMonitor<BudgetServerOptions> options) =>
{
    var contextResult = TryBuildSnapshotContext(options.CurrentValue, databasePath, settingsPath);
    if (contextResult.Error is not null)
    {
        return TypedResults.BadRequest(contextResult.Error);
    }

    var resolved = contextResult.Paths!;
    var snapshot = contextResult.Snapshot!;
    var currentIndex = snapshot.PeriodSummaries.Count == 0
        ? 0
        : Math.Clamp(snapshot.CurrentPeriodIndex, 0, snapshot.PeriodSummaries.Count - 1);

    var accounts = AccountRepository.LoadAccounts(resolved.DatabasePath);
    var response = accounts.Select(account =>
    {
        var key = string.IsNullOrWhiteSpace(account.Description) ? $"Account {account.Id}" : account.Description.Trim();
        snapshot.AccountRunningBalances.TryGetValue(key, out var balances);
        balances ??= Array.Empty<decimal>();
        var currentBalance = currentIndex < balances.Length
            ? balances[currentIndex]
            : balances.Length > 0 ? balances[^1] : 0m;
        var deleteBlockingReason = BuildAccountDeleteBlockingReason(resolved.DatabasePath, account.Description);

        return new AccountSummaryResponse
        {
            Id = account.Id,
            Description = account.Description,
            AccountType = account.AccountType,
            LoginLink = account.LoginLink,
            Notes = account.Notes,
            CurrentBalance = currentBalance,
            SafetyNet = account.SafetyNet,
            IsHidden = account.Hidden,
            IsActive = account.Active,
            CanDelete = string.IsNullOrWhiteSpace(deleteBlockingReason),
            DeleteBlockedReason = deleteBlockingReason
        };
    })
    .OrderBy(item => item.Description, StringComparer.CurrentCultureIgnoreCase)
    .ToList();

    return TypedResults.Ok((IReadOnlyList<AccountSummaryResponse>)response);
});

app.MapPost("/api/accounts/save", async Task<Results<Ok<AccountSummaryResponse>, BadRequest<ApiErrorResponse>>> (
    AccountSaveRequest request,
    IOptionsMonitor<BudgetServerOptions> options) =>
{
    var contextResult = TryBuildSnapshotContext(options.CurrentValue, null, null);
    if (contextResult.Error is not null)
    {
        return TypedResults.BadRequest(contextResult.Error);
    }

    try
    {
        var resolved = contextResult.Paths!;
        var accountRecord = new AccountRecord
        {
            Id = request.Id,
            Description = request.Description?.Trim() ?? string.Empty,
            AccountType = request.AccountType?.Trim() ?? string.Empty,
            LoginLink = request.LoginLink?.Trim() ?? string.Empty,
            Notes = request.Notes ?? string.Empty,
            SafetyNet = request.SafetyNet,
            Hidden = request.IsHidden,
            Active = request.IsActive
        };

        if (string.IsNullOrWhiteSpace(accountRecord.Description))
        {
            return TypedResults.BadRequest(new ApiErrorResponse
            {
                Error = "Account description is required."
            });
        }

        if (string.IsNullOrWhiteSpace(accountRecord.AccountType))
        {
            return TypedResults.BadRequest(new ApiErrorResponse
            {
                Error = "Account type is required."
            });
        }

        if (accountRecord.SafetyNet < 0m)
        {
            return TypedResults.BadRequest(new ApiErrorResponse
            {
                Error = "Safety net cannot be negative."
            });
        }

        if (accountRecord.Id <= 0)
        {
            accountRecord.Id = AccountRepository.CreateAccount(resolved.DatabasePath, accountRecord);
        }
        else
        {
            AccountRepository.SaveAccount(resolved.DatabasePath, accountRecord);
        }

        var previousDescription = request.PreviousDescription?.Trim() ?? string.Empty;
        var updatedDescription = accountRecord.Description.Trim();
        if (!string.IsNullOrWhiteSpace(previousDescription) &&
            !string.Equals(previousDescription, updatedDescription, StringComparison.CurrentCulture))
        {
            AccountRepository.RenameAccountReferences(resolved.DatabasePath, previousDescription, updatedDescription);
        }

        var freshContext = TryBuildSnapshotContext(options.CurrentValue, null, null);
        if (freshContext.Error is not null)
        {
            return TypedResults.BadRequest(freshContext.Error);
        }

        var freshPaths = freshContext.Paths!;
        var freshSnapshot = freshContext.Snapshot!;
        var currentIndex = freshSnapshot.PeriodSummaries.Count == 0
            ? 0
            : Math.Clamp(freshSnapshot.CurrentPeriodIndex, 0, freshSnapshot.PeriodSummaries.Count - 1);
        var accountKey = string.IsNullOrWhiteSpace(accountRecord.Description) ? $"Account {accountRecord.Id}" : accountRecord.Description.Trim();
        freshSnapshot.AccountRunningBalances.TryGetValue(accountKey, out var balances);
        balances ??= Array.Empty<decimal>();
        var currentBalance = currentIndex < balances.Length
            ? balances[currentIndex]
            : balances.Length > 0 ? balances[^1] : 0m;
        var deleteBlockingReason = BuildAccountDeleteBlockingReason(freshPaths.DatabasePath, accountRecord.Description);

        return TypedResults.Ok(new AccountSummaryResponse
        {
            Id = accountRecord.Id,
            Description = accountRecord.Description,
            AccountType = accountRecord.AccountType,
            LoginLink = accountRecord.LoginLink,
            Notes = accountRecord.Notes,
            CurrentBalance = currentBalance,
            SafetyNet = accountRecord.SafetyNet,
            IsHidden = accountRecord.Hidden,
            IsActive = accountRecord.Active,
            CanDelete = string.IsNullOrWhiteSpace(deleteBlockingReason),
            DeleteBlockedReason = deleteBlockingReason
        });
    }
    catch (Exception ex)
    {
        ServerAppLogger.LogError("Failed to save account.", ex);
        return TypedResults.BadRequest(new ApiErrorResponse
        {
            Error = ex.Message
        });
    }
});

app.MapDelete("/api/accounts/{accountId:int}", Results<Ok<AccountDeleteResponse>, BadRequest<ApiErrorResponse>> (
    int accountId,
    IOptionsMonitor<BudgetServerOptions> options) =>
{
    var contextResult = TryBuildSnapshotContext(options.CurrentValue, null, null);
    if (contextResult.Error is not null)
    {
        return TypedResults.BadRequest(contextResult.Error);
    }

    try
    {
        var resolved = contextResult.Paths!;
        var account = AccountRepository.LoadAccounts(resolved.DatabasePath).FirstOrDefault(item => item.Id == accountId);
        if (account is null)
        {
            return TypedResults.BadRequest(new ApiErrorResponse
            {
                Error = "The selected account could not be found on the server."
            });
        }

        var deleteBlockingReason = BuildAccountDeleteBlockingReason(resolved.DatabasePath, account.Description);
        if (!string.IsNullOrWhiteSpace(deleteBlockingReason))
        {
            return TypedResults.BadRequest(new ApiErrorResponse
            {
                Error = deleteBlockingReason
            });
        }

        AccountRepository.DeleteAccount(resolved.DatabasePath, accountId);
        return TypedResults.Ok(new AccountDeleteResponse
        {
            Id = accountId,
            Description = account.Description
        });
    }
    catch (Exception ex)
    {
        ServerAppLogger.LogError("Failed to delete account.", ex);
        return TypedResults.BadRequest(new ApiErrorResponse
        {
            Error = ex.Message
        });
    }
});

app.MapGet("/api/income", Results<Ok<IReadOnlyList<IncomeSummaryResponse>>, BadRequest<ApiErrorResponse>> (
    string? databasePath,
    string? settingsPath,
    IOptionsMonitor<BudgetServerOptions> options) =>
{
    var contextResult = TryBuildSnapshotContext(options.CurrentValue, databasePath, settingsPath);
    if (contextResult.Error is not null)
    {
        return TypedResults.BadRequest(contextResult.Error);
    }

    var resolved = contextResult.Paths!;
    var response = IncomeRepository.LoadIncome(resolved.DatabasePath)
        .Select(item =>
        {
            var deleteBlockingReason = BuildIncomeDeleteBlockingReason(resolved.DatabasePath, item.Id, item.Description);
            return new IncomeSummaryResponse
            {
                Id = item.Id,
                Description = item.Description,
                Amount = item.Amount,
                Cadence = item.Cadence,
                OnDay = item.OnDay,
                OnDate = item.OnDate,
                AutoIncrease = item.AutoIncrease,
                AutoIncreaseOnDate = item.AutoIncreaseOnDate,
                StartDate = item.StartDate,
                EndDate = item.EndDate,
                ToAccount = item.ToAccount,
                SameAs = item.SameAs,
                IsHidden = item.Hidden,
                IsActive = item.Active,
                LoginLink = item.LoginLink,
                Notes = item.Notes,
                CanDelete = string.IsNullOrWhiteSpace(deleteBlockingReason),
                DeleteBlockedReason = deleteBlockingReason
            };
        })
        .OrderBy(item => item.Description, StringComparer.CurrentCultureIgnoreCase)
        .ToList();

    return TypedResults.Ok((IReadOnlyList<IncomeSummaryResponse>)response);
});

app.MapPost("/api/income/save", async Task<Results<Ok<IncomeSummaryResponse>, BadRequest<ApiErrorResponse>>> (
    IncomeSaveRequest request,
    IOptionsMonitor<BudgetServerOptions> options) =>
{
    var contextResult = TryBuildSnapshotContext(options.CurrentValue, null, null);
    if (contextResult.Error is not null)
    {
        return TypedResults.BadRequest(contextResult.Error);
    }

    try
    {
        var resolved = contextResult.Paths!;
        var incomeRecord = new IncomeRecord
        {
            Id = request.Id,
            Description = request.Description?.Trim() ?? string.Empty,
            Amount = request.Amount,
            Cadence = request.Cadence?.Trim() ?? string.Empty,
            OnDay = request.OnDay,
            OnDate = request.OnDate?.Trim() ?? string.Empty,
            AutoIncrease = request.AutoIncrease,
            AutoIncreaseOnDate = request.AutoIncreaseOnDate?.Trim() ?? string.Empty,
            StartDate = request.StartDate?.Trim() ?? string.Empty,
            EndDate = request.EndDate?.Trim() ?? string.Empty,
            ToAccount = request.ToAccount?.Trim() ?? string.Empty,
            SameAs = request.SameAs?.Trim() ?? string.Empty,
            Hidden = request.IsHidden,
            Active = request.IsActive,
            LoginLink = request.LoginLink?.Trim() ?? string.Empty,
            Notes = request.Notes ?? string.Empty
        };

        if (string.IsNullOrWhiteSpace(incomeRecord.Description))
        {
            return TypedResults.BadRequest(new ApiErrorResponse
            {
                Error = "Income description is required."
            });
        }

        if (string.IsNullOrWhiteSpace(incomeRecord.Cadence))
        {
            return TypedResults.BadRequest(new ApiErrorResponse
            {
                Error = "Income cadence is required."
            });
        }

        if (incomeRecord.Amount < 0m)
        {
            return TypedResults.BadRequest(new ApiErrorResponse
            {
                Error = "Income amount cannot be negative."
            });
        }

        if (incomeRecord.AutoIncrease < 0m)
        {
            return TypedResults.BadRequest(new ApiErrorResponse
            {
                Error = "Auto increase cannot be negative."
            });
        }

        if (string.Equals(incomeRecord.Cadence, "Same As", StringComparison.OrdinalIgnoreCase) &&
            !string.IsNullOrWhiteSpace(incomeRecord.Description) &&
            string.Equals(incomeRecord.SameAs, incomeRecord.Description, StringComparison.CurrentCultureIgnoreCase))
        {
            return TypedResults.BadRequest(new ApiErrorResponse
            {
                Error = "Income cannot use itself for Same As."
            });
        }

        var existingIncome = incomeRecord.Id > 0
            ? IncomeRepository.LoadIncome(resolved.DatabasePath).FirstOrDefault(item => item.Id == incomeRecord.Id)
            : null;

        if (incomeRecord.Id <= 0 || existingIncome is null)
        {
            incomeRecord.Id = IncomeRepository.CreateIncome(resolved.DatabasePath, incomeRecord);
        }
        else
        {
            IncomeRepository.SaveIncome(resolved.DatabasePath, incomeRecord);
        }

        var previousDescription = request.PreviousDescription?.Trim() ?? string.Empty;
        var updatedDescription = incomeRecord.Description.Trim();
        if (!string.IsNullOrWhiteSpace(previousDescription) &&
            !string.Equals(previousDescription, updatedDescription, StringComparison.CurrentCulture))
        {
            RenameIncomeSameAsReferences(resolved.DatabasePath, previousDescription, updatedDescription);
        }

        var deleteBlockingReason = BuildIncomeDeleteBlockingReason(resolved.DatabasePath, incomeRecord.Id, incomeRecord.Description);
        return TypedResults.Ok(new IncomeSummaryResponse
        {
            Id = incomeRecord.Id,
            Description = incomeRecord.Description,
            Amount = incomeRecord.Amount,
            Cadence = incomeRecord.Cadence,
            OnDay = incomeRecord.OnDay,
            OnDate = incomeRecord.OnDate,
            AutoIncrease = incomeRecord.AutoIncrease,
            AutoIncreaseOnDate = incomeRecord.AutoIncreaseOnDate,
            StartDate = incomeRecord.StartDate,
            EndDate = incomeRecord.EndDate,
            ToAccount = incomeRecord.ToAccount,
            SameAs = incomeRecord.SameAs,
            IsHidden = incomeRecord.Hidden,
            IsActive = incomeRecord.Active,
            LoginLink = incomeRecord.LoginLink,
            Notes = incomeRecord.Notes,
            CanDelete = string.IsNullOrWhiteSpace(deleteBlockingReason),
            DeleteBlockedReason = deleteBlockingReason
        });
    }
    catch (Exception ex)
    {
        ServerAppLogger.LogError("Failed to save income.", ex);
        return TypedResults.BadRequest(new ApiErrorResponse
        {
            Error = ex.Message
        });
    }
});

app.MapDelete("/api/income/{incomeId:int}", Results<Ok<IncomeDeleteResponse>, BadRequest<ApiErrorResponse>> (
    int incomeId,
    IOptionsMonitor<BudgetServerOptions> options) =>
{
    var contextResult = TryBuildSnapshotContext(options.CurrentValue, null, null);
    if (contextResult.Error is not null)
    {
        return TypedResults.BadRequest(contextResult.Error);
    }

    try
    {
        var resolved = contextResult.Paths!;
        var income = IncomeRepository.LoadIncome(resolved.DatabasePath).FirstOrDefault(item => item.Id == incomeId);
        if (income is null)
        {
            return TypedResults.BadRequest(new ApiErrorResponse
            {
                Error = "The selected income item could not be found on the server."
            });
        }

        var deleteBlockingReason = BuildIncomeDeleteBlockingReason(resolved.DatabasePath, income.Id, income.Description);
        if (!string.IsNullOrWhiteSpace(deleteBlockingReason))
        {
            return TypedResults.BadRequest(new ApiErrorResponse
            {
                Error = deleteBlockingReason
            });
        }

        IncomeRepository.DeleteIncome(resolved.DatabasePath, incomeId);
        return TypedResults.Ok(new IncomeDeleteResponse
        {
            Id = incomeId,
            Description = income.Description
        });
    }
    catch (Exception ex)
    {
        ServerAppLogger.LogError("Failed to delete income.", ex);
        return TypedResults.BadRequest(new ApiErrorResponse
        {
            Error = ex.Message
        });
    }
});

app.MapGet("/api/expenses", Results<Ok<IReadOnlyList<ExpenseSummaryResponse>>, BadRequest<ApiErrorResponse>> (
    string? databasePath,
    string? settingsPath,
    IOptionsMonitor<BudgetServerOptions> options) =>
{
    var contextResult = TryBuildSnapshotContext(options.CurrentValue, databasePath, settingsPath);
    if (contextResult.Error is not null)
    {
        return TypedResults.BadRequest(contextResult.Error);
    }

    var resolved = contextResult.Paths!;
    var response = ExpenseRepository.LoadExpenses(resolved.DatabasePath)
        .Select(item => new ExpenseSummaryResponse
        {
            Id = item.Id,
            Description = item.Description,
            AmountDue = item.AmountDue,
            Cadence = item.Cadence,
            DueDay = item.DueDay,
            DueDate = item.DueDate,
            FromAccount = item.FromAccount,
            SameAs = item.SameAs,
            Category = item.Category,
            IsHidden = item.Hidden,
            IsActive = item.Active,
            LoginLink = item.LoginLink,
            Notes = item.Notes
        })
        .OrderBy(item => item.Category, StringComparer.CurrentCultureIgnoreCase)
        .ThenBy(item => item.Description, StringComparer.CurrentCultureIgnoreCase)
        .ToList();

    return TypedResults.Ok((IReadOnlyList<ExpenseSummaryResponse>)response);
});

app.MapPost("/api/expenses/save", async Task<Results<Ok<ExpenseSummaryResponse>, BadRequest<ApiErrorResponse>>> (
    ExpenseSaveRequest request,
    IOptionsMonitor<BudgetServerOptions> options) =>
{
    var contextResult = TryBuildSnapshotContext(options.CurrentValue, null, null);
    if (contextResult.Error is not null)
    {
        return TypedResults.BadRequest(contextResult.Error);
    }

    try
    {
        var resolved = contextResult.Paths!;
        var expenseRecord = new ExpenseRecord
        {
            Id = request.Id,
            Description = request.Description?.Trim() ?? string.Empty,
            AmountDue = request.AmountDue,
            Cadence = request.Cadence?.Trim() ?? string.Empty,
            DueDay = request.DueDay,
            DueDate = request.DueDate?.Trim() ?? string.Empty,
            FromAccount = request.FromAccount?.Trim() ?? string.Empty,
            SameAs = request.SameAs?.Trim() ?? string.Empty,
            Category = request.Category?.Trim() ?? string.Empty,
            Hidden = request.IsHidden,
            Active = request.IsActive,
            LoginLink = request.LoginLink?.Trim() ?? string.Empty,
            Notes = request.Notes ?? string.Empty
        };

        if (string.IsNullOrWhiteSpace(expenseRecord.Description))
        {
            return TypedResults.BadRequest(new ApiErrorResponse { Error = "Expense description is required." });
        }

        if (string.IsNullOrWhiteSpace(expenseRecord.Cadence))
        {
            return TypedResults.BadRequest(new ApiErrorResponse { Error = "Expense frequency is required." });
        }

        if (string.IsNullOrWhiteSpace(expenseRecord.Category))
        {
            return TypedResults.BadRequest(new ApiErrorResponse { Error = "Expense category is required." });
        }

        if (expenseRecord.AmountDue < 0m)
        {
            return TypedResults.BadRequest(new ApiErrorResponse { Error = "Expense amount cannot be negative." });
        }

        if (string.IsNullOrWhiteSpace(expenseRecord.FromAccount))
        {
            return TypedResults.BadRequest(new ApiErrorResponse { Error = "From account is required." });
        }

        if (string.Equals(expenseRecord.Cadence, "Same As", StringComparison.OrdinalIgnoreCase) &&
            string.IsNullOrWhiteSpace(expenseRecord.SameAs))
        {
            return TypedResults.BadRequest(new ApiErrorResponse { Error = "Select an income source for Same As expenses." });
        }

        if (expenseRecord.Id <= 0)
        {
            expenseRecord.Id = ExpenseRepository.CreateExpense(resolved.DatabasePath, expenseRecord);
        }
        else
        {
            ExpenseRepository.SaveExpense(resolved.DatabasePath, expenseRecord);
        }

        return TypedResults.Ok(new ExpenseSummaryResponse
        {
            Id = expenseRecord.Id,
            Description = expenseRecord.Description,
            AmountDue = expenseRecord.AmountDue,
            Cadence = expenseRecord.Cadence,
            DueDay = expenseRecord.DueDay,
            DueDate = expenseRecord.DueDate,
            FromAccount = expenseRecord.FromAccount,
            SameAs = expenseRecord.SameAs,
            Category = expenseRecord.Category,
            IsHidden = expenseRecord.Hidden,
            IsActive = expenseRecord.Active,
            LoginLink = expenseRecord.LoginLink,
            Notes = expenseRecord.Notes
        });
    }
    catch (Exception ex)
    {
        ServerAppLogger.LogError("Failed to save expense.", ex);
        return TypedResults.BadRequest(new ApiErrorResponse { Error = ex.Message });
    }
});

app.MapDelete("/api/expenses/{expenseId:int}", Results<Ok<ExpenseDeleteResponse>, BadRequest<ApiErrorResponse>> (
    int expenseId,
    IOptionsMonitor<BudgetServerOptions> options) =>
{
    var contextResult = TryBuildSnapshotContext(options.CurrentValue, null, null);
    if (contextResult.Error is not null)
    {
        return TypedResults.BadRequest(contextResult.Error);
    }

    try
    {
        var resolved = contextResult.Paths!;
        var expense = ExpenseRepository.LoadExpenses(resolved.DatabasePath).FirstOrDefault(item => item.Id == expenseId);
        if (expense is null)
        {
            return TypedResults.BadRequest(new ApiErrorResponse
            {
                Error = "The selected expense could not be found on the server."
            });
        }

        ExpenseRepository.DeleteExpense(resolved.DatabasePath, expenseId);
        return TypedResults.Ok(new ExpenseDeleteResponse
        {
            Id = expenseId,
            Description = expense.Description
        });
    }
    catch (Exception ex)
    {
        ServerAppLogger.LogError("Failed to delete expense.", ex);
        return TypedResults.BadRequest(new ApiErrorResponse
        {
            Error = ex.Message
        });
    }
});

app.MapGet("/api/transactions", Results<Ok<IReadOnlyList<TransactionSummaryResponse>>, BadRequest<ApiErrorResponse>> (
    int? limit,
    string? databasePath,
    string? settingsPath,
    IOptions<BudgetServerOptions> options) =>
{
    var contextResult = TryBuildSnapshotContext(options.Value, databasePath, settingsPath);
    if (contextResult.Error is not null)
    {
        return TypedResults.BadRequest(contextResult.Error);
    }

    var resolved = contextResult.Paths!;
    var transactions = TransactionRepository.LoadTransactions(resolved.DatabasePath);
    var assignments = TransactionRepository.LoadTransactionAssignments(resolved.DatabasePath);
    var incomeLookup = IncomeRepository.LoadIncome(resolved.DatabasePath).ToDictionary(item => item.Id, item => item.Description ?? string.Empty);
    var debtLookup = DebtRepository.LoadDebts(resolved.DatabasePath).ToDictionary(item => item.Id, item => item.Description ?? string.Empty);
    var expenseLookup = ExpenseRepository.LoadExpenses(resolved.DatabasePath).ToDictionary(item => item.Id, item => item.Description ?? string.Empty);
    var savingsLookup = SavingsRepository.LoadSavings(resolved.DatabasePath).ToDictionary(item => item.Id, item => item.Description ?? string.Empty);
    var assignmentLookup = assignments
        .GroupBy(item => item.TransactionId)
        .ToDictionary(
            group => group.Key,
            group => (IReadOnlyList<TransactionAssignmentSummaryResponse>)group.Select(item => new TransactionAssignmentSummaryResponse
            {
                Id = item.Id,
                CategoryIndex = item.CatIdx,
                CategoryLabel = GetTransactionAssignmentCategoryLabel(item.CatIdx),
                ItemId = item.ItemId,
                ItemLabel = ResolveTransactionAssignmentItemLabel(item.CatIdx, item.ItemId, incomeLookup, debtLookup, expenseLookup, savingsLookup),
                Amount = item.Amount,
                Notes = item.Notes,
                NeedsReview = item.NeedsReview
            }).ToList());

    var safeLimit = Math.Clamp(limit ?? 50, 1, 250);
    var response = transactions
        .Take(safeLimit)
        .Select(transaction => new TransactionSummaryResponse
        {
            Id = transaction.Id,
            SourceName = transaction.SourceName,
            TransactionDate = transaction.TransactionDate,
            Description = transaction.Description,
            Amount = transaction.Amount,
            Notes = transaction.Notes,
            Assignments = assignmentLookup.TryGetValue(transaction.Id, out var transactionAssignments)
                ? transactionAssignments
                : Array.Empty<TransactionAssignmentSummaryResponse>()
        })
        .ToList();

    return TypedResults.Ok((IReadOnlyList<TransactionSummaryResponse>)response);
});

app.MapPost("/api/transactions/save", Results<Ok<TransactionSummaryResponse>, BadRequest<ApiErrorResponse>> (
    TransactionSaveRequest request,
    string? databasePath,
    string? settingsPath,
    IOptions<BudgetServerOptions> options) =>
{
    var contextResult = TryBuildSnapshotContext(options.Value, databasePath, settingsPath);
    if (contextResult.Error is not null)
    {
        return TypedResults.BadRequest(contextResult.Error);
    }

    var resolved = contextResult.Paths!;
    TransactionRepository.EnsureTransactionsSchema(resolved.DatabasePath);

    var record = new TransactionRecord
    {
        Id = request.Id,
        SourceName = request.SourceName ?? string.Empty,
        TransactionDate = request.TransactionDate ?? string.Empty,
        Description = request.Description ?? string.Empty,
        Amount = request.Amount,
        Notes = request.Notes ?? string.Empty
    };

    var assignments = (request.Assignments ?? Array.Empty<TransactionAssignmentSaveRequest>())
        .Where(item => item is not null && item.ItemId > 0 && item.CategoryIndex >= 0 && Math.Abs(item.Amount) > 0.009m)
        .Select(item => new TransactionAssignmentRecord
        {
            Id = item.Id,
            TransactionId = request.Id,
            CatIdx = item.CategoryIndex,
            ItemId = item.ItemId,
            Amount = Math.Abs(item.Amount),
            Notes = item.Notes ?? string.Empty,
            NeedsReview = item.NeedsReview
        })
        .ToList();

    var transactionId = request.Id;
    var existingTransactionIds = TransactionRepository.LoadTransactions(resolved.DatabasePath)
        .Select(item => item.Id)
        .ToHashSet();
    if (transactionId > 0 && existingTransactionIds.Contains(transactionId))
    {
        TransactionRepository.SaveTransaction(resolved.DatabasePath, record, assignments);
    }
    else
    {
        transactionId = TransactionRepository.CreateTransaction(resolved.DatabasePath, record, assignments);
    }

    BudgetWorkspaceService.SyncTransactionSourceOverrides(resolved.DatabasePath, BuildWorkspaceSettings(options.Value, resolved.SettingsPath, resolved.DatabasePath));

    var transactions = TransactionRepository.LoadTransactions(resolved.DatabasePath);
    var saved = transactions.FirstOrDefault(item => item.Id == transactionId);
    if (saved is null)
    {
        return TypedResults.BadRequest(new ApiErrorResponse { Error = "The saved transaction could not be reloaded from the server." });
    }

    var savedAssignments = TransactionRepository.LoadTransactionAssignments(resolved.DatabasePath)
        .Where(item => item.TransactionId == transactionId)
        .ToList();
    var incomeLookup = IncomeRepository.LoadIncome(resolved.DatabasePath).ToDictionary(item => item.Id, item => item.Description ?? string.Empty);
    var debtLookup = DebtRepository.LoadDebts(resolved.DatabasePath).ToDictionary(item => item.Id, item => item.Description ?? string.Empty);
    var expenseLookup = ExpenseRepository.LoadExpenses(resolved.DatabasePath).ToDictionary(item => item.Id, item => item.Description ?? string.Empty);
    var savingsLookup = SavingsRepository.LoadSavings(resolved.DatabasePath).ToDictionary(item => item.Id, item => item.Description ?? string.Empty);

    return TypedResults.Ok(new TransactionSummaryResponse
    {
        Id = saved.Id,
        SourceName = saved.SourceName,
        TransactionDate = saved.TransactionDate,
        Description = saved.Description,
        Amount = saved.Amount,
        Notes = saved.Notes,
        Assignments = savedAssignments.Select(item => new TransactionAssignmentSummaryResponse
        {
            Id = item.Id,
            CategoryIndex = item.CatIdx,
            CategoryLabel = GetTransactionAssignmentCategoryLabel(item.CatIdx),
            ItemId = item.ItemId,
            ItemLabel = ResolveTransactionAssignmentItemLabel(item.CatIdx, item.ItemId, incomeLookup, debtLookup, expenseLookup, savingsLookup),
            Amount = item.Amount,
            Notes = item.Notes,
            NeedsReview = item.NeedsReview
        }).ToList()
    });
});

app.MapDelete("/api/transactions/{transactionId:int}", Results<Ok<ApiMessageResponse>, BadRequest<ApiErrorResponse>> (
    int transactionId,
    string? databasePath,
    string? settingsPath,
    IOptions<BudgetServerOptions> options) =>
{
    var contextResult = TryBuildSnapshotContext(options.Value, databasePath, settingsPath);
    if (contextResult.Error is not null)
    {
        return TypedResults.BadRequest(contextResult.Error);
    }

    var resolved = contextResult.Paths!;
    if (transactionId <= 0)
    {
        return TypedResults.BadRequest(new ApiErrorResponse { Error = "A valid transaction id is required." });
    }

    TransactionRepository.DeleteTransaction(resolved.DatabasePath, transactionId);
    BudgetWorkspaceService.SyncTransactionSourceOverrides(resolved.DatabasePath, BuildWorkspaceSettings(options.Value, resolved.SettingsPath, resolved.DatabasePath));
    return TypedResults.Ok(new ApiMessageResponse { Message = "Deleted transaction." });
});

app.MapGet("/api/budget/periods/{periodIndex:int}/items", Results<Ok<BudgetPeriodItemsResponse>, BadRequest<ApiErrorResponse>> (
    int periodIndex,
    string? databasePath,
    string? settingsPath,
    IOptions<BudgetServerOptions> options) =>
{
    var contextResult = TryBuildSnapshotContext(options.Value, databasePath, settingsPath);
    if (contextResult.Error is not null)
    {
        return TypedResults.BadRequest(contextResult.Error);
    }

    var resolved = contextResult.Paths!;
    var snapshot = contextResult.Snapshot!;
    if (snapshot.PeriodSummaries.Count == 0)
    {
        return TypedResults.Ok(new BudgetPeriodItemsResponse
        {
            PeriodIndex = 0,
            TotalPeriodCount = 0,
            Items = Array.Empty<BudgetPeriodItemResponse>()
        });
    }

    var safePeriodIndex = Math.Clamp(periodIndex, 0, snapshot.PeriodSummaries.Count - 1);
    var periodStart = snapshot.PeriodSummaries[safePeriodIndex].PeriodStart;
    var periodName = snapshot.BudgetPeriod;

    var incomeLookup = IncomeRepository.LoadIncome(resolved.DatabasePath).ToDictionary(item => item.Id);
    var expenseLookup = ExpenseRepository.LoadExpenses(resolved.DatabasePath).ToDictionary(item => item.Id);
    var savingsLookup = SavingsRepository.LoadSavings(resolved.DatabasePath).ToDictionary(item => item.Id);
    var debtLookup = DebtRepository.LoadDebts(resolved.DatabasePath).ToDictionary(item => item.Id);
    var assignmentTotals = BuildPeriodAssignmentTotals(resolved.DatabasePath, snapshot, safePeriodIndex, periodName);

    var items = snapshot.ItemizedBudgetRows
        .Select(row => BuildBudgetPeriodItemResponse(row, safePeriodIndex, snapshot, incomeLookup, expenseLookup, savingsLookup, debtLookup, assignmentTotals))
        .Where(item => item is not null)
        .Select(item => item!)
        .ToList();

    var existingKeys = new HashSet<string>(items.Select(item => item.SourceKey), StringComparer.OrdinalIgnoreCase);

    foreach (var supplemental in BuildSupplementalBudgetItems(snapshot, safePeriodIndex, incomeLookup, expenseLookup, savingsLookup, debtLookup, existingKeys))
    {
        items.Add(supplemental);
    }

    items = items
        .OrderBy(item => GetBudgetSectionSortOrder(item.SectionName))
        .ThenBy(item => item.GroupName, StringComparer.CurrentCultureIgnoreCase)
        .ThenBy(item => item.Label, StringComparer.CurrentCultureIgnoreCase)
        .ToList();

    return TypedResults.Ok(new BudgetPeriodItemsResponse
    {
        PeriodIndex = safePeriodIndex,
        PeriodStart = periodStart,
        TotalPeriodCount = snapshot.PeriodSummaries.Count,
        CurrentPeriodIndex = snapshot.CurrentPeriodIndex,
        Items = items
    });
});

app.MapPost("/api/budget/additional", Results<Ok<BudgetItemAdditionalUpdateResponse>, BadRequest<ApiErrorResponse>> (
    BudgetItemAdditionalUpdateRequest request,
    string? databasePath,
    string? settingsPath,
    IOptions<BudgetServerOptions> options) =>
{
    var contextResult = TryBuildSnapshotContext(options.Value, databasePath, settingsPath);
    if (contextResult.Error is not null)
    {
        return TypedResults.BadRequest(contextResult.Error);
    }

    var resolved = contextResult.Paths!;
    var snapshot = contextResult.Snapshot!;

    if (snapshot.PeriodSummaries.Count == 0)
    {
        return TypedResults.BadRequest(new ApiErrorResponse { Error = "No budget periods are available." });
    }

    var safePeriodIndex = Math.Clamp(request.PeriodIndex, 0, snapshot.PeriodSummaries.Count - 1);
    var targetRow = snapshot.ItemizedBudgetRows.FirstOrDefault(row =>
        string.Equals(row.SourceKey, request.SourceKey, StringComparison.OrdinalIgnoreCase) &&
        string.Equals(row.Label, request.ItemLabel, StringComparison.OrdinalIgnoreCase));

    var sourceKey = targetRow?.SourceKey ?? request.SourceKey;
    var itemLabel = targetRow?.Label ?? request.ItemLabel;
    var categoryIndex = targetRow is not null
        ? ParseCategoryIndexFromSourceKey(targetRow.SourceKey, targetRow.SectionName)
        : ParseCategoryIndexFromSourceKey(request.SourceKey, string.Empty);
    if (categoryIndex < 0)
    {
        return TypedResults.BadRequest(new ApiErrorResponse { Error = "The selected budget item does not support additional adjustments." });
    }

    var dueDate = snapshot.PeriodSummaries[safePeriodIndex].PeriodStart;
    var sourceAmount = targetRow is not null && safePeriodIndex < targetRow.SourceBaseValues.Length
        ? targetRow.SourceBaseValues[safePeriodIndex]
        : 0m;
    var scheduledAmount = targetRow is not null && safePeriodIndex < targetRow.ScheduledValues.Length
        ? targetRow.ScheduledValues[safePeriodIndex]
        : 0m;
    var currentPaid = targetRow?.PaidIndexes.Contains(safePeriodIndex) ?? false;
    var currentNote = BudgetWorkspaceService.GetSourceOverrideNote(resolved.DatabasePath, dueDate, categoryIndex, sourceKey, itemLabel);
    var existingSelectionMode = BudgetWorkspaceService.GetSourceOverrideSelectionMode(resolved.DatabasePath, dueDate, categoryIndex, sourceKey, itemLabel);
    var existingManualAmount = BudgetWorkspaceService.GetSourceOverrideManualAmount(resolved.DatabasePath, dueDate, categoryIndex, sourceKey, itemLabel);
    var calculatedAmount = sourceAmount + request.Additional;
    var resolvedSelectionMode = existingSelectionMode == 3
        ? 3
        : calculatedAmount > scheduledAmount ? 2 : 1;

    BudgetWorkspaceService.SaveSourceOverride(
        resolved.DatabasePath,
        dueDate,
        categoryIndex,
        sourceKey,
        itemLabel,
        sourceAmount,
        request.Additional,
        currentNote,
        currentPaid,
        resolvedSelectionMode,
        existingManualAmount);

    return TypedResults.Ok(new BudgetItemAdditionalUpdateResponse
    {
        PeriodIndex = safePeriodIndex,
        PeriodStart = dueDate,
        SourceKey = sourceKey,
        ItemLabel = itemLabel,
        Additional = request.Additional,
        Message = $"Saved additional budget for {itemLabel} on {dueDate:MM/dd/yyyy}."
    });
});

app.MapGet("/api/budget/editor-state", Results<Ok<BudgetCellEditorStateResponse>, BadRequest<ApiErrorResponse>> (
    int periodIndex,
    string sourceKey,
    string itemLabel,
    string? databasePath,
    string? settingsPath,
    IOptions<BudgetServerOptions> options) =>
{
    var contextResult = TryBuildSnapshotContext(options.Value, databasePath, settingsPath);
    if (contextResult.Error is not null)
    {
        return TypedResults.BadRequest(contextResult.Error);
    }

    var resolved = contextResult.Paths!;
    var snapshot = contextResult.Snapshot!;
    if (snapshot.PeriodSummaries.Count == 0)
    {
        return TypedResults.BadRequest(new ApiErrorResponse { Error = "No budget periods are available." });
    }

    var safePeriodIndex = Math.Clamp(periodIndex, 0, snapshot.PeriodSummaries.Count - 1);
    var targetRow = snapshot.ItemizedBudgetRows.FirstOrDefault(row =>
        string.Equals(row.SourceKey, sourceKey, StringComparison.OrdinalIgnoreCase) &&
        string.Equals(row.Label, itemLabel, StringComparison.OrdinalIgnoreCase));
    if (targetRow is null)
    {
        return TypedResults.BadRequest(new ApiErrorResponse { Error = "Could not find the selected budget item on the server." });
    }

    var resolvedSourceKey = string.IsNullOrWhiteSpace(targetRow.SourceKey) ? sourceKey : targetRow.SourceKey;
    var resolvedItemLabel = string.IsNullOrWhiteSpace(targetRow.Label) ? itemLabel : targetRow.Label;
    var categoryIndex = ParseCategoryIndexFromSourceKey(targetRow.SourceKey, targetRow.SectionName);
    var dueDate = snapshot.PeriodSummaries[safePeriodIndex].PeriodStart;
    var overrideState = BudgetWorkspaceService.GetBudgetCellEditorOverrideState(
        resolved.DatabasePath,
        dueDate,
        categoryIndex,
        resolvedSourceKey,
        resolvedItemLabel);

    decimal? currentDebtBalanceOverrideAmount = null;
    if (categoryIndex == 1 && string.Equals(targetRow.SectionName, "Debts", StringComparison.OrdinalIgnoreCase))
    {
        var debt = DebtRepository.LoadDebts(resolved.DatabasePath).FirstOrDefault(item =>
            string.Equals(item.Description, resolvedItemLabel, StringComparison.CurrentCultureIgnoreCase));
        if (debt is not null)
        {
            currentDebtBalanceOverrideAmount = GetCurrentDebtBalanceOverrideAmount(snapshot, resolved.DatabasePath, debt);
        }
    }

    return TypedResults.Ok(new BudgetCellEditorStateResponse
    {
        PeriodIndex = safePeriodIndex,
        PeriodStart = dueDate,
        SourceKey = resolvedSourceKey,
        ItemLabel = resolvedItemLabel,
        CategoryIndex = categoryIndex,
        Amount = overrideState.Amount,
        Additional = overrideState.Additional,
        ManualAmount = overrideState.ManualAmount,
        Paid = overrideState.Paid,
        SelectionMode = overrideState.SelectionMode,
        Notes = overrideState.Notes,
        ManualOverrideAmount = overrideState.ManualOverrideAmount,
        ManualOverrideAdditional = overrideState.ManualOverrideAdditional,
        ManualOverridePaid = overrideState.ManualOverridePaid,
        ManualOverrideSelectionMode = overrideState.ManualOverrideSelectionMode,
        ManualOverrideNote = overrideState.ManualOverrideNote,
        CurrentDebtBalanceOverrideAmount = currentDebtBalanceOverrideAmount,
        LinkedTransactions = GetBudgetCellLinkedTransactionLines(snapshot, resolved.DatabasePath, safePeriodIndex, categoryIndex, resolvedItemLabel, resolvedSourceKey)
    });
});

app.MapPost("/api/budget/save", Results<Ok<BudgetCellSaveResponse>, BadRequest<ApiErrorResponse>> (
    BudgetCellSaveRequest request,
    string? databasePath,
    string? settingsPath,
    IOptions<BudgetServerOptions> options) =>
{
    var contextResult = TryBuildSnapshotContext(options.Value, databasePath, settingsPath);
    if (contextResult.Error is not null)
    {
        return TypedResults.BadRequest(contextResult.Error);
    }

    var resolved = contextResult.Paths!;
    var snapshot = contextResult.Snapshot!;
    if (snapshot.PeriodSummaries.Count == 0)
    {
        return TypedResults.BadRequest(new ApiErrorResponse { Error = "No budget periods are available." });
    }

    var safePeriodIndex = Math.Clamp(request.PeriodIndex, 0, snapshot.PeriodSummaries.Count - 1);
    var targetRow = snapshot.ItemizedBudgetRows.FirstOrDefault(row =>
        string.Equals(row.SourceKey, request.SourceKey, StringComparison.OrdinalIgnoreCase) &&
        string.Equals(row.Label, request.ItemLabel, StringComparison.OrdinalIgnoreCase));
    if (targetRow is null)
    {
        return TypedResults.BadRequest(new ApiErrorResponse { Error = "Could not find the selected budget item on the server." });
    }

    var sourceKey = string.IsNullOrWhiteSpace(targetRow.SourceKey) ? request.SourceKey : targetRow.SourceKey;
    var itemLabel = string.IsNullOrWhiteSpace(targetRow.Label) ? request.ItemLabel : targetRow.Label;
    var categoryIndex = ParseCategoryIndexFromSourceKey(targetRow.SourceKey, targetRow.SectionName);
    var dueDate = snapshot.PeriodSummaries[safePeriodIndex].PeriodStart;
    var sourceAmount = safePeriodIndex < targetRow.SourceBaseValues.Length ? targetRow.SourceBaseValues[safePeriodIndex] : 0m;
    var usesSelectionModeRow =
        (categoryIndex == 0 && string.Equals(targetRow.SectionName, "Income", StringComparison.OrdinalIgnoreCase)) ||
        (categoryIndex == 1 && string.Equals(targetRow.SectionName, "Debts", StringComparison.OrdinalIgnoreCase)) ||
        (categoryIndex == 2 && string.Equals(targetRow.SectionName, "Expenses", StringComparison.OrdinalIgnoreCase)) ||
        (categoryIndex == 3 && string.Equals(targetRow.SectionName, "Savings", StringComparison.OrdinalIgnoreCase));
    var existingSourceAmount = BudgetWorkspaceService.GetSourceOverrideAmount(resolved.DatabasePath, dueDate, categoryIndex, sourceKey, itemLabel);
    var existingSourceAdditional = BudgetWorkspaceService.GetSourceOverrideAdditional(resolved.DatabasePath, dueDate, categoryIndex, sourceKey, itemLabel);
    var existingSourcePaid = BudgetWorkspaceService.GetSourceOverridePaid(resolved.DatabasePath, dueDate, categoryIndex, sourceKey, itemLabel);
    var hasTransactionSource = sourceAmount != 0m || targetRow.SourceIndexes.Contains(safePeriodIndex);
    var hasSourceBackedValue = hasTransactionSource || existingSourceAmount.HasValue || existingSourceAdditional != 0m || existingSourcePaid || request.Additional != 0m || request.Paid;

    if (usesSelectionModeRow)
    {
        BudgetWorkspaceService.DeleteManualOverride(
            resolved.DatabasePath,
            dueDate,
            categoryIndex,
            sourceKey,
            itemLabel);
        BudgetWorkspaceService.DeleteSourceOverride(
            resolved.DatabasePath,
            dueDate,
            categoryIndex,
            sourceKey,
            itemLabel);
        BudgetWorkspaceService.SaveSourceOverride(
            resolved.DatabasePath,
            dueDate,
            categoryIndex,
            sourceKey,
            itemLabel,
            sourceAmount,
            request.Additional,
            request.Note,
            request.Paid,
            request.SelectionMode,
            request.ManualAmount);

        if (categoryIndex == 1)
        {
            var debt = DebtRepository.LoadDebts(resolved.DatabasePath).FirstOrDefault(item =>
                string.Equals(item.Description, itemLabel, StringComparison.CurrentCultureIgnoreCase));
            if (debt is not null)
            {
                ApplyCurrentDebtBalanceOverride(resolved.DatabasePath, snapshot, debt, request.CurrentPeriodBalanceOverrideAmount);
            }
        }
    }
    else if (hasSourceBackedValue)
    {
        BudgetWorkspaceService.DeleteManualOverride(
            resolved.DatabasePath,
            dueDate,
            categoryIndex,
            sourceKey,
            itemLabel);
        BudgetWorkspaceService.SaveSourceOverride(
            resolved.DatabasePath,
            dueDate,
            categoryIndex,
            sourceKey,
            itemLabel,
            sourceAmount,
            request.Additional,
            request.Note,
            request.Paid);
    }
    else
    {
        BudgetWorkspaceService.SaveManualOverride(
            resolved.DatabasePath,
            dueDate,
            categoryIndex,
            sourceKey,
            itemLabel,
            request.Amount,
            request.Note,
            request.Additional,
            request.Paid);
        BudgetWorkspaceService.UpdateSourceOverrideMetadata(
            resolved.DatabasePath,
            dueDate,
            categoryIndex,
            sourceKey,
            itemLabel,
            request.Additional,
            paid: request.Paid);
    }

    return TypedResults.Ok(new BudgetCellSaveResponse
    {
        PeriodIndex = safePeriodIndex,
        PeriodStart = dueDate,
        SourceKey = sourceKey,
        ItemLabel = itemLabel,
        Message = $"Saved budget cell for {itemLabel} on {dueDate:MM/dd/yyyy}."
    });
});

app.MapPost("/api/budget/clear", Results<Ok<BudgetCellSaveResponse>, BadRequest<ApiErrorResponse>> (
    BudgetCellClearRequest request,
    string? databasePath,
    string? settingsPath,
    IOptions<BudgetServerOptions> options) =>
{
    var contextResult = TryBuildSnapshotContext(options.Value, databasePath, settingsPath);
    if (contextResult.Error is not null)
    {
        return TypedResults.BadRequest(contextResult.Error);
    }

    var resolved = contextResult.Paths!;
    var snapshot = contextResult.Snapshot!;
    if (snapshot.PeriodSummaries.Count == 0)
    {
        return TypedResults.BadRequest(new ApiErrorResponse { Error = "No budget periods are available." });
    }

    var safePeriodIndex = Math.Clamp(request.PeriodIndex, 0, snapshot.PeriodSummaries.Count - 1);
    var targetRow = snapshot.ItemizedBudgetRows.FirstOrDefault(row =>
        string.Equals(row.SourceKey, request.SourceKey, StringComparison.OrdinalIgnoreCase) &&
        string.Equals(row.Label, request.ItemLabel, StringComparison.OrdinalIgnoreCase));
    if (targetRow is null)
    {
        return TypedResults.BadRequest(new ApiErrorResponse { Error = "Could not find the selected budget item on the server." });
    }

    var sourceKey = string.IsNullOrWhiteSpace(targetRow.SourceKey) ? request.SourceKey : targetRow.SourceKey;
    var itemLabel = string.IsNullOrWhiteSpace(targetRow.Label) ? request.ItemLabel : targetRow.Label;
    var categoryIndex = ParseCategoryIndexFromSourceKey(targetRow.SourceKey, targetRow.SectionName);
    var dueDate = snapshot.PeriodSummaries[safePeriodIndex].PeriodStart;
    var usesSelectionModeRow =
        (categoryIndex == 0 && string.Equals(targetRow.SectionName, "Income", StringComparison.OrdinalIgnoreCase)) ||
        (categoryIndex == 1 && string.Equals(targetRow.SectionName, "Debts", StringComparison.OrdinalIgnoreCase)) ||
        (categoryIndex == 2 && string.Equals(targetRow.SectionName, "Expenses", StringComparison.OrdinalIgnoreCase)) ||
        (categoryIndex == 3 && string.Equals(targetRow.SectionName, "Savings", StringComparison.OrdinalIgnoreCase));

    BudgetWorkspaceService.DeleteManualOverride(
        resolved.DatabasePath,
        dueDate,
        categoryIndex,
        sourceKey,
        itemLabel);
    if (usesSelectionModeRow)
    {
        BudgetWorkspaceService.DeleteSourceOverride(
            resolved.DatabasePath,
            dueDate,
            categoryIndex,
            sourceKey,
            itemLabel);
    }

    BudgetWorkspaceService.SyncTransactionSourceOverrides(resolved.DatabasePath, BuildWorkspaceSettings(options.Value, resolved.SettingsPath, resolved.DatabasePath));

    return TypedResults.Ok(new BudgetCellSaveResponse
    {
        PeriodIndex = safePeriodIndex,
        PeriodStart = dueDate,
        SourceKey = sourceKey,
        ItemLabel = itemLabel,
        Message = usesSelectionModeRow
            ? $"Reset budget source selection for {itemLabel} on {dueDate:MM/dd/yyyy}."
            : $"Cleared manual override for {itemLabel} on {dueDate:MM/dd/yyyy}."
    });
});

try
{
    app.Run();
    ServerAppLogger.LogInfo("Server shut down normally.");
}
catch (IOException ex) when (IsAddressAlreadyInUse(ex))
{
    ServerAppLogger.LogError("Server could not start because the configured port is already in use.", ex);
    Console.Error.WriteLine("JCBudgeting.Server could not start because the configured port is already in use.");
    Console.Error.WriteLine("Stop the other server on that port, or change the port in Server Setup and try again.");
    Environment.ExitCode = 1;
}
catch (Exception ex)
{
    ServerAppLogger.LogError("Fatal server host exception.", ex);
    throw;
}

static void WriteServerStartupBanner(string urls, string databasePath)
{
    ServerConsoleStyle.WriteColoredLine("Starting JCBudgeting Server...", ConsoleColor.Yellow);
    Console.WriteLine();
    Console.WriteLine("Copyright (c) 2026 Cliff Flanders");
    Console.WriteLine("All rights reserved.");
    Console.WriteLine("This software is proprietary and confidential.");
    Console.WriteLine("You may not modify or reverse engineer this software without explicit permission from the author.");
    Console.WriteLine("This software is provided \"as is\", without warranty of any kind.");
    Console.WriteLine("This software is for personal use only and must not be used for commercial purposes.");
    Console.WriteLine();
    Console.WriteLine("For more information, visit: https://github.com/cbrzilla/JC_Budgeting_App");
    Console.WriteLine();

    var accessUrls = GetSuggestedAccessUrls(urls);
    if (accessUrls.Count > 0)
    {
        Console.WriteLine("Sharing on:");
        foreach (var accessUrl in accessUrls)
        {
            if (Uri.TryCreate(accessUrl, UriKind.Absolute, out var uri))
            {
                Console.WriteLine($"  {uri.Host}:{uri.Port}");
            }
        }
    }
    else
    {
        Console.WriteLine("Sharing on: unavailable");
    }

    var setupUrl = ServerSetupUrlHelper.GetSetupLaunchUrl(urls);
    if (!string.IsNullOrWhiteSpace(setupUrl))
    {
        Console.WriteLine($"Setup page: {setupUrl}");
    }

    var databaseName = string.IsNullOrWhiteSpace(databasePath)
        ? "(no database configured)"
        : Path.GetFileName(databasePath);
    Console.WriteLine($"Database: {databaseName}");
    Console.WriteLine();
    ServerAppLogger.LogInfo(
        $"Sharing URLs: {(accessUrls.Count > 0 ? string.Join(", ", accessUrls) : "unavailable")}. Database: {databaseName}.");
}

static ResolvedPaths ResolvePaths(BudgetServerOptions options, string? requestDatabasePath, string? requestSettingsPath)
{
    options = GetEffectiveBudgetServerOptions(options);
    var databasePath = FirstNonBlank(requestDatabasePath, options.DatabasePath);
    var settingsPath = FirstNonBlank(requestSettingsPath, options.SettingsPath);

    databasePath = ResolveAgainstBaseDirectory(databasePath);
    settingsPath = ResolveAgainstBaseDirectory(settingsPath);

    return new ResolvedPaths(databasePath, settingsPath);
}

static string FirstNonBlank(params string?[] values) =>
    values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value))?.Trim() ?? string.Empty;

static string NormalizeIpAddress(IPAddress? address)
{
    if (address is null)
    {
        return string.Empty;
    }

    if (address.IsIPv4MappedToIPv6)
    {
        address = address.MapToIPv4();
    }

    return address.ToString();
}

static bool IsAddressAlreadyInUse(Exception ex)
{
    for (var current = ex; current is not null; current = current.InnerException!)
    {
        if (current is System.Net.Sockets.SocketException socketEx &&
            socketEx.SocketErrorCode == System.Net.Sockets.SocketError.AddressAlreadyInUse)
        {
            return true;
        }

        if (current.Message.Contains("address already in use", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }
    }

    return false;
}

static IReadOnlyList<string> GetSuggestedAccessUrls(string urls)
{
    var raw = FirstNonBlank(urls, "http://0.0.0.0:5099");
    if (!Uri.TryCreate(raw, UriKind.Absolute, out var uri))
    {
        return Array.Empty<string>();
    }

    var port = uri.Port > 0 ? uri.Port : 5099;
    var scheme = string.IsNullOrWhiteSpace(uri.Scheme) ? "http" : uri.Scheme;
    var host = uri.Host ?? string.Empty;
    var results = new List<string>();

    void Add(string candidate)
    {
        if (string.IsNullOrWhiteSpace(candidate))
        {
            return;
        }

        if (!results.Contains(candidate, StringComparer.OrdinalIgnoreCase))
        {
            results.Add(candidate);
        }
    }

    Add($"{scheme}://localhost:{port}");

    if (string.Equals(host, "127.0.0.1", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(host, "localhost", StringComparison.OrdinalIgnoreCase))
    {
        Add($"{scheme}://127.0.0.1:{port}");
        return results;
    }

    foreach (var networkInterface in NetworkInterface.GetAllNetworkInterfaces())
    {
        if (networkInterface.OperationalStatus != OperationalStatus.Up)
        {
            continue;
        }

        if (networkInterface.NetworkInterfaceType == NetworkInterfaceType.Loopback ||
            networkInterface.NetworkInterfaceType == NetworkInterfaceType.Tunnel)
        {
            continue;
        }

        IPInterfaceProperties? properties;
        try
        {
            properties = networkInterface.GetIPProperties();
        }
        catch
        {
            continue;
        }

        foreach (var unicastAddress in properties.UnicastAddresses)
        {
            var address = unicastAddress.Address;
            if (address.AddressFamily != AddressFamily.InterNetwork)
            {
                continue;
            }

            if (IPAddress.IsLoopback(address))
            {
                continue;
            }

            Add($"{scheme}://{address}:{port}");
        }
    }

    return results;
}

static string ResolveAgainstBaseDirectory(string path)
{
    if (string.IsNullOrWhiteSpace(path))
    {
        return string.Empty;
    }

    return Path.IsPathRooted(path)
        ? Path.GetFullPath(path)
        : Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, path));
}

static string GetStandaloneServerConfigPath() =>
    Path.Combine(GetStandaloneServerDataDirectory(), "budgetserver.local.json");

static string GetStandaloneServerDataDirectory()
{
    var configuredRoot = FirstNonBlank(
        Environment.GetEnvironmentVariable("JCBUDGETING_SERVER_DATA_DIR"),
        Environment.GetEnvironmentVariable("APP_DATA_DIR"));
    var path = string.IsNullOrWhiteSpace(configuredRoot)
        ? AppContext.BaseDirectory
        : ResolveAgainstBaseDirectory(configuredRoot);
    Directory.CreateDirectory(path);
    return path;
}

static string GetStandaloneDatabasesDirectory()
{
    var path = Path.Combine(GetStandaloneServerDataDirectory(), "Databases");
    Directory.CreateDirectory(path);
    return path;
}

static BudgetServerOptions GetEffectiveBudgetServerOptions(BudgetServerOptions options)
{
    try
    {
        if (HasExplicitRuntimeBudgetServerOverrides())
        {
            return options;
        }

        var configPath = GetStandaloneServerConfigPath();
        if (!File.Exists(configPath))
        {
            return options;
        }

        var json = File.ReadAllText(configPath);
        if (string.IsNullOrWhiteSpace(json))
        {
            return options;
        }

        var config = JsonSerializer.Deserialize<StandaloneServerConfigFile>(json);
        if (config?.BudgetServer is null)
        {
            return options;
        }

        return new BudgetServerOptions
        {
            DatabasePath = string.IsNullOrWhiteSpace(config.BudgetServer.DatabasePath) ? options.DatabasePath : config.BudgetServer.DatabasePath,
            SettingsPath = string.IsNullOrWhiteSpace(config.BudgetServer.SettingsPath) ? options.SettingsPath : config.BudgetServer.SettingsPath,
            BudgetPeriod = string.IsNullOrWhiteSpace(config.BudgetServer.BudgetPeriod) ? options.BudgetPeriod : config.BudgetServer.BudgetPeriod,
            BudgetStartDate = string.IsNullOrWhiteSpace(config.BudgetServer.BudgetStartDate) ? options.BudgetStartDate : config.BudgetServer.BudgetStartDate,
            BudgetYears = config.BudgetServer.BudgetYears > 0 ? config.BudgetServer.BudgetYears : options.BudgetYears
        };
    }
    catch
    {
        return options;
    }
}

static bool HasExplicitRuntimeBudgetServerOverrides() =>
    !string.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable("BudgetServer__DatabasePath")) ||
    !string.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable("BudgetServer__SettingsPath")) ||
    !string.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable("BudgetServer__BudgetPeriod")) ||
    !string.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable("BudgetServer__BudgetStartDate")) ||
    !string.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable("BudgetServer__BudgetYears"));

static (string ConfigPath, string DatabasePath, string SettingsPath, string Urls) SaveStandaloneServerConfig(ServerConfigSaveRequest request)
{
    var databasePath = ResolveAgainstBaseDirectory(request.DatabasePath);
    var settingsPath = string.Empty;

    if (string.IsNullOrWhiteSpace(databasePath))
    {
        throw new InvalidOperationException("Database path is required.");
    }

    if (Directory.Exists(databasePath))
    {
        throw new InvalidOperationException("Database path must include a .jcbdb file name, not just a folder.");
    }

    var normalizedUrls = string.IsNullOrWhiteSpace(request.Urls)
        ? "http://0.0.0.0:5099"
        : request.Urls.Trim();

    var normalizedBudgetPeriod = string.IsNullOrWhiteSpace(request.BudgetPeriod) ? "Monthly" : request.BudgetPeriod.Trim();
    var normalizedBudgetStartDate = request.BudgetStartDate?.Trim() ?? string.Empty;
    var normalizedBudgetYears = request.BudgetYears <= 0 ? 20 : request.BudgetYears;

    var config = new StandaloneServerConfigFile
    {
        Urls = normalizedUrls,
        BudgetServer = new StandaloneBudgetServerSection
        {
            DatabasePath = databasePath,
            SettingsPath = string.Empty,
            BudgetPeriod = normalizedBudgetPeriod,
            BudgetStartDate = normalizedBudgetStartDate,
            BudgetYears = normalizedBudgetYears
        }
    };

    var configPath = GetStandaloneServerConfigPath();
    var configDirectory = Path.GetDirectoryName(configPath);
    if (!string.IsNullOrWhiteSpace(configDirectory))
    {
        Directory.CreateDirectory(configDirectory);
    }

    var json = JsonSerializer.Serialize(config, new JsonSerializerOptions
    {
        WriteIndented = true
    });
    File.WriteAllText(configPath, json);

    BudgetWorkspaceService.SaveBudgetTimelineSettings(databasePath, new BudgetWorkspaceSettings
    {
        BudgetPeriod = normalizedBudgetPeriod,
        BudgetStartDate = normalizedBudgetStartDate,
        BudgetYears = normalizedBudgetYears
    });

    return (configPath, databasePath, settingsPath, normalizedUrls);
}

static void CreateDatabaseFromPackagedTemplate(string destinationPath)
{
    var templatePath = FindPackagedBlankTemplatePath();
    if (string.IsNullOrWhiteSpace(templatePath))
    {
        throw new InvalidOperationException("The packaged blank database template was not found.");
    }

    var directory = Path.GetDirectoryName(destinationPath);
    if (!string.IsNullOrWhiteSpace(directory))
    {
        Directory.CreateDirectory(directory);
    }

    File.Copy(templatePath, destinationPath, overwrite: false);
}

static string ResolveCreateDatabaseTargetPath(string rawPath)
{
    var resolvedPath = ResolveAgainstBaseDirectory(rawPath);
    if (string.IsNullOrWhiteSpace(resolvedPath))
    {
        return Path.Combine(GetStandaloneDatabasesDirectory(), "new_budget_database.jcbdb");
    }

    if (Directory.Exists(resolvedPath))
    {
        return Path.Combine(resolvedPath, "new_budget_database.jcbdb");
    }

    var hasExtension = !string.IsNullOrWhiteSpace(Path.GetExtension(resolvedPath));
    if (!hasExtension)
    {
        return resolvedPath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar) + ".jcbdb";
    }

    var extension = Path.GetExtension(resolvedPath);
    if (!string.Equals(extension, ".jcbdb", StringComparison.OrdinalIgnoreCase))
    {
        throw new InvalidOperationException("Database file must use the .jcbdb extension.");
    }

    return resolvedPath;
}

static string ResolveUploadDatabaseTargetPath(string sourceFileName)
{
    var safeFileName = Path.GetFileName(sourceFileName ?? string.Empty);
    if (string.IsNullOrWhiteSpace(safeFileName))
    {
        safeFileName = "uploaded_budget_database.jcbdb";
    }

    if (!string.Equals(Path.GetExtension(safeFileName), ".jcbdb", StringComparison.OrdinalIgnoreCase))
    {
        safeFileName = Path.GetFileNameWithoutExtension(safeFileName) + ".jcbdb";
    }

    var targetDirectory = GetStandaloneDatabasesDirectory();
    var targetPath = Path.Combine(targetDirectory, safeFileName);
    return EnsureUniqueDatabasePath(targetPath);
}

static string EnsureUniqueDatabasePath(string path)
{
    if (!File.Exists(path))
    {
        return path;
    }

    var directory = Path.GetDirectoryName(path) ?? GetStandaloneDatabasesDirectory();
    var baseName = Path.GetFileNameWithoutExtension(path);
    var extension = Path.GetExtension(path);
    return Path.Combine(directory, $"{baseName}_{DateTime.Now:yyyyMMdd_HHmmss}{extension}");
}

static IReadOnlyList<ServerDatabaseOption> GetAvailableDatabaseFiles(string selectedDatabasePath)
{
    var selectedFullPath = ResolveAgainstBaseDirectory(selectedDatabasePath);
    var results = new List<ServerDatabaseOption>();

    try
    {
        var databasesDirectory = GetStandaloneDatabasesDirectory();
        foreach (var path in Directory.GetFiles(databasesDirectory, "*.jcbdb", SearchOption.TopDirectoryOnly)
                     .OrderBy(Path.GetFileName, StringComparer.OrdinalIgnoreCase))
        {
            results.Add(new ServerDatabaseOption
            {
                Name = Path.GetFileName(path),
                Path = path
            });
        }
    }
    catch
    {
        // Ignore directory scan failures.
    }

    if (!string.IsNullOrWhiteSpace(selectedFullPath) &&
        !results.Any(item => string.Equals(item.Path, selectedFullPath, StringComparison.OrdinalIgnoreCase)))
    {
        results.Insert(0, new ServerDatabaseOption
        {
            Name = Path.GetFileName(selectedFullPath),
            Path = selectedFullPath
        });
    }

    return results;
}

static string FindPackagedBlankTemplatePath()
{
    var candidatePaths = new[]
    {
        Path.Combine(AppContext.BaseDirectory, "blank-template.jcbdb"),
        Path.Combine(AppContext.BaseDirectory, "Assets", "blank-template.jcbdb")
    };

    return candidatePaths.FirstOrDefault(File.Exists) ?? string.Empty;
}

static SnapshotContext TryBuildSnapshotContext(BudgetServerOptions options, string? requestDatabasePath, string? requestSettingsPath)
{
    options = GetEffectiveBudgetServerOptions(options);
    var resolved = ResolvePaths(options, requestDatabasePath, requestSettingsPath);
    if (string.IsNullOrWhiteSpace(resolved.DatabasePath))
    {
        return new SnapshotContext(null, null, new ApiErrorResponse
        {
            Error = "DatabasePath is required. Configure BudgetServer:DatabasePath or pass ?databasePath=..."
        });
    }

    if (!File.Exists(resolved.DatabasePath))
    {
        return new SnapshotContext(null, null, new ApiErrorResponse
        {
            Error = $"Database was not found: {resolved.DatabasePath}"
        });
    }

    var settings = BuildWorkspaceSettings(options, resolved.SettingsPath, resolved.DatabasePath);
    var snapshot = BudgetWorkspaceService.BuildSnapshot(resolved.DatabasePath, settings);
    return new SnapshotContext(resolved, snapshot, null);
}

static BudgetWorkspaceSettings BuildWorkspaceSettings(BudgetServerOptions options, string resolvedSettingsPath, string resolvedDatabasePath)
{
    options = GetEffectiveBudgetServerOptions(options);
    var settings = !string.IsNullOrWhiteSpace(resolvedSettingsPath) && File.Exists(resolvedSettingsPath)
        ? BudgetWorkspaceService.LoadSettings(resolvedSettingsPath, resolvedDatabasePath)
        : new BudgetWorkspaceSettings();

    if (!string.IsNullOrWhiteSpace(options.BudgetPeriod))
    {
        settings.BudgetPeriod = options.BudgetPeriod.Trim();
    }

    if (!string.IsNullOrWhiteSpace(options.BudgetStartDate))
    {
        settings.BudgetStartDate = options.BudgetStartDate.Trim();
    }

    if (options.BudgetYears > 0)
    {
        settings.BudgetYears = options.BudgetYears;
    }

    BudgetWorkspaceService.TryApplyBudgetTimelineSettingsFromDatabase(resolvedDatabasePath, settings);

    return settings;
}

static int? FindProjectedPayoffIndex(decimal[] runningBalances, int startIndex)
{
    if (runningBalances.Length == 0)
    {
        return null;
    }

    for (var i = Math.Max(0, startIndex); i < runningBalances.Length; i++)
    {
        if (runningBalances[i] <= 0.009m)
        {
            return i;
        }
    }

    return null;
}

static string GetTransactionAssignmentCategoryLabel(int categoryIndex) =>
    categoryIndex switch
    {
        0 => "Income",
        1 => "Debt",
        2 => "Expense",
        3 => "Savings",
        _ => "Other"
    };

static string ResolveTransactionAssignmentItemLabel(
    int categoryIndex,
    int itemId,
    IReadOnlyDictionary<int, string> incomeLookup,
    IReadOnlyDictionary<int, string> debtLookup,
    IReadOnlyDictionary<int, string> expenseLookup,
    IReadOnlyDictionary<int, string> savingsLookup)
{
    return categoryIndex switch
    {
        0 when incomeLookup.TryGetValue(itemId, out var incomeLabel) => string.IsNullOrWhiteSpace(incomeLabel) ? $"Income {itemId}" : incomeLabel,
        1 when debtLookup.TryGetValue(itemId, out var debtLabel) => string.IsNullOrWhiteSpace(debtLabel) ? $"Debt {itemId}" : debtLabel,
        2 when expenseLookup.TryGetValue(itemId, out var expenseLabel) => string.IsNullOrWhiteSpace(expenseLabel) ? $"Expense {itemId}" : expenseLabel,
        3 when savingsLookup.TryGetValue(itemId, out var savingsLabel) => string.IsNullOrWhiteSpace(savingsLabel) ? $"Savings {itemId}" : savingsLabel,
        _ => $"Item {itemId}"
    };
}

static string BuildAccountDeleteBlockingReason(string databasePath, string? accountName)
{
    var normalizedName = accountName?.Trim() ?? string.Empty;
    if (string.IsNullOrWhiteSpace(normalizedName) || string.IsNullOrWhiteSpace(databasePath) || !File.Exists(databasePath))
    {
        return string.Empty;
    }

    var blockingReferences = new List<string>();

    blockingReferences.AddRange(
        SavingsRepository.LoadSavings(databasePath)
            .Where(item => string.Equals(item.FromAccount?.Trim(), normalizedName, StringComparison.CurrentCulture))
            .Select(item => FormatBlockingReference("Savings", item.Category, item.Description, "From Account")));

    blockingReferences.AddRange(
        IncomeRepository.LoadIncome(databasePath)
            .Where(item => string.Equals(item.ToAccount?.Trim(), normalizedName, StringComparison.CurrentCulture))
            .Select(item => FormatBlockingReference("Income", null, item.Description, "To Account")));

    blockingReferences.AddRange(
        ExpenseRepository.LoadExpenses(databasePath)
            .Where(item => string.Equals(item.FromAccount?.Trim(), normalizedName, StringComparison.CurrentCulture))
            .Select(item => FormatBlockingReference("Expenses", item.Category, item.Description, "From Account")));

    blockingReferences.AddRange(
        DebtRepository.LoadDebts(databasePath)
            .Where(item => string.Equals(item.FromAccount?.Trim(), normalizedName, StringComparison.CurrentCulture))
            .Select(item => FormatBlockingReference("Debts", null, item.Description, "From Account")));

    if (blockingReferences.Count == 0)
    {
        return string.Empty;
    }

    var details = blockingReferences
        .OrderBy(item => item, StringComparer.CurrentCultureIgnoreCase)
        .ToList();

    return details.Count > 6
        ? $"Cannot delete this item since it is linked to other items: {string.Join(", ", details)}"
        : $"Cannot delete this item since it is linked to other items:{Environment.NewLine}{Environment.NewLine}{string.Join(Environment.NewLine + Environment.NewLine, details)}";
}

static string BuildIncomeDeleteBlockingReason(string databasePath, int incomeId, string? incomeName)
{
    var normalizedName = incomeName?.Trim() ?? string.Empty;
    if (string.IsNullOrWhiteSpace(normalizedName) || string.IsNullOrWhiteSpace(databasePath) || !File.Exists(databasePath))
    {
        return string.Empty;
    }

    var blockingReferences = new List<string>();

    blockingReferences.AddRange(
        IncomeRepository.LoadIncome(databasePath)
            .Where(item => item.Id != incomeId && string.Equals(item.SameAs?.Trim(), normalizedName, StringComparison.CurrentCulture))
            .Select(item => FormatBlockingReference("Income", null, item.Description, "Same As")));

    blockingReferences.AddRange(
        SavingsRepository.LoadSavings(databasePath)
            .Where(item => string.Equals(item.SameAs?.Trim(), normalizedName, StringComparison.CurrentCulture))
            .Select(item => FormatBlockingReference("Savings", item.Category, item.Description, "Same As")));

    blockingReferences.AddRange(
        ExpenseRepository.LoadExpenses(databasePath)
            .Where(item => string.Equals(item.SameAs?.Trim(), normalizedName, StringComparison.CurrentCulture))
            .Select(item => FormatBlockingReference("Expenses", item.Category, item.Description, "Same As")));

    blockingReferences.AddRange(
        DebtRepository.LoadDebts(databasePath)
            .Where(item => string.Equals(item.SameAs?.Trim(), normalizedName, StringComparison.CurrentCulture))
            .Select(item => FormatBlockingReference("Debts", null, item.Description, "Same As")));

    if (blockingReferences.Count == 0)
    {
        return string.Empty;
    }

    var details = blockingReferences
        .OrderBy(item => item, StringComparer.CurrentCultureIgnoreCase)
        .ToList();

    return details.Count > 6
        ? $"Cannot delete this item since it is linked to other items: {string.Join(", ", details)}"
        : $"Cannot delete this item since it is linked to other items:{Environment.NewLine}{Environment.NewLine}{string.Join(Environment.NewLine + Environment.NewLine, details)}";
}

static string BuildSavingsDeleteBlockingReason(string databasePath, string? savingsName)
{
    var normalizedName = savingsName?.Trim() ?? string.Empty;
    if (string.IsNullOrWhiteSpace(normalizedName) || string.IsNullOrWhiteSpace(databasePath) || !File.Exists(databasePath))
    {
        return string.Empty;
    }

    var savingsFundingName = $"(Savings) {normalizedName}";
    var blockingReferences = new List<string>();

    blockingReferences.AddRange(
        ExpenseRepository.LoadExpenses(databasePath)
            .Where(item => string.Equals(item.FromAccount?.Trim(), savingsFundingName, StringComparison.CurrentCulture))
            .Select(item => FormatBlockingReference("Expenses", item.Category, item.Description, "From Account")));

    blockingReferences.AddRange(
        DebtRepository.LoadDebts(databasePath)
            .Where(item => string.Equals(item.FromAccount?.Trim(), savingsFundingName, StringComparison.CurrentCulture))
            .Select(item => FormatBlockingReference("Debts", null, item.Description, "From Account")));

    if (blockingReferences.Count == 0)
    {
        return string.Empty;
    }

    var details = blockingReferences
        .OrderBy(item => item, StringComparer.CurrentCultureIgnoreCase)
        .ToList();

    return details.Count > 6
        ? $"Cannot delete this item since it is linked to other items: {string.Join(", ", details)}"
        : $"Cannot delete this item since it is linked to other items:{Environment.NewLine}{Environment.NewLine}{string.Join(Environment.NewLine + Environment.NewLine, details)}";
}

static string BuildDebtDeleteBlockingReason(string databasePath, string? debtName)
{
    var normalizedName = debtName?.Trim() ?? string.Empty;
    if (string.IsNullOrWhiteSpace(normalizedName) || string.IsNullOrWhiteSpace(databasePath) || !File.Exists(databasePath))
    {
        return string.Empty;
    }

    var debtFundingName = $"(Debt) {normalizedName}";
    var blockingReferences = ExpenseRepository.LoadExpenses(databasePath)
        .Where(item => string.Equals(item.FromAccount?.Trim(), debtFundingName, StringComparison.CurrentCulture))
        .Select(item => FormatBlockingReference("Expenses", item.Category, item.Description, "From Account"))
        .OrderBy(item => item, StringComparer.CurrentCultureIgnoreCase)
        .ToList();

    if (blockingReferences.Count == 0)
    {
        return string.Empty;
    }

    return blockingReferences.Count > 6
        ? $"Cannot delete this item since it is linked to other items: {string.Join(", ", blockingReferences)}"
        : $"Cannot delete this item since it is linked to other items:{Environment.NewLine}{Environment.NewLine}{string.Join(Environment.NewLine + Environment.NewLine, blockingReferences)}";
}

static void RenameIncomeSameAsReferences(string databasePath, string previousDescription, string newDescription)
{
    if (string.IsNullOrWhiteSpace(databasePath) || !File.Exists(databasePath))
    {
        return;
    }

    var previousValue = previousDescription.Trim();
    var newValue = newDescription.Trim();
    if (string.IsNullOrWhiteSpace(previousValue) ||
        string.IsNullOrWhiteSpace(newValue) ||
        string.Equals(previousValue, newValue, StringComparison.CurrentCulture))
    {
        return;
    }

    using var conn = new SqliteConnection($"Data Source={databasePath}");
    conn.Open();
    using var transaction = conn.BeginTransaction();

    RenameReferenceColumn(conn, transaction, "savings", "SameAs", previousValue, newValue);
    RenameReferenceColumn(conn, transaction, "expenses", "SameAs", previousValue, newValue);
    RenameReferenceColumn(conn, transaction, "debts", "SameAs", previousValue, newValue);
    RenameReferenceColumn(conn, transaction, "income", "SameAs", previousValue, newValue, "AND COALESCE(Cadence,'') = 'Same As'");

    transaction.Commit();
}

static void RenameSavingsFundingReferences(string databasePath, string previousDescription, string newDescription)
{
    if (string.IsNullOrWhiteSpace(databasePath) || !File.Exists(databasePath))
    {
        return;
    }

    var previousValue = previousDescription.Trim();
    var newValue = newDescription.Trim();
    if (string.IsNullOrWhiteSpace(previousValue) ||
        string.IsNullOrWhiteSpace(newValue) ||
        string.Equals(previousValue, newValue, StringComparison.CurrentCulture))
    {
        return;
    }

    var previousFundingSource = $"(Savings) {previousValue}";
    var newFundingSource = $"(Savings) {newValue}";

    using var conn = new SqliteConnection($"Data Source={databasePath}");
    conn.Open();
    using var transaction = conn.BeginTransaction();

    RenameReferenceColumn(conn, transaction, "expenses", "FromAccount", previousFundingSource, newFundingSource);
    RenameReferenceColumn(conn, transaction, "debts", "FromAccount", previousFundingSource, newFundingSource);

    transaction.Commit();
}

static void RenameDebtFundingReferences(string databasePath, string previousDescription, string newDescription)
{
    if (string.IsNullOrWhiteSpace(databasePath) || !File.Exists(databasePath))
    {
        return;
    }

    var previousValue = previousDescription.Trim();
    var newValue = newDescription.Trim();
    if (string.IsNullOrWhiteSpace(previousValue) ||
        string.IsNullOrWhiteSpace(newValue) ||
        string.Equals(previousValue, newValue, StringComparison.CurrentCulture))
    {
        return;
    }

    var previousFundingSource = $"(Debt) {previousValue}";
    var newFundingSource = $"(Debt) {newValue}";

    using var conn = new SqliteConnection($"Data Source={databasePath}");
    conn.Open();
    using var transaction = conn.BeginTransaction();

    RenameReferenceColumn(conn, transaction, "expenses", "FromAccount", previousFundingSource, newFundingSource);

    transaction.Commit();
}

static void RenameReferenceColumn(SqliteConnection conn, SqliteTransaction transaction, string tableName, string columnName, string previousValue, string newValue, string additionalFilter = "")
{
    using var cmd = conn.CreateCommand();
    cmd.Transaction = transaction;
    cmd.CommandText = $"UPDATE {tableName} SET {columnName} = @newValue WHERE {columnName} = @previousValue {additionalFilter}";
    cmd.Parameters.AddWithValue("@newValue", newValue);
    cmd.Parameters.AddWithValue("@previousValue", previousValue);
    cmd.ExecuteNonQuery();
}

static string FormatBlockingReference(string itemType, string? category, string? itemName, string fieldName)
{
    var normalizedItemName = string.IsNullOrWhiteSpace(itemName) ? $"Unnamed {itemType.TrimEnd('s')}" : itemName.Trim();
    var categoryPart = string.IsNullOrWhiteSpace(category) ? string.Empty : $" | {category.Trim()}";
    return $"{itemType}{categoryPart} | {normalizedItemName} | {fieldName}";
}

const int DebtBalanceHistoryCategoryIndex = -101;

static string BuildDebtOverrideKey(DebtRecord debt) =>
    debt.Id > 0 ? $"ID:1:{debt.Id}" : (debt.Description?.Trim() ?? string.Empty);

static decimal? GetCurrentDebtBalanceOverrideAmount(BudgetWorkspaceSnapshot snapshot, string databasePath, DebtRecord debt)
{
    if (snapshot.PeriodSummaries.Count == 0)
    {
        return null;
    }

    var currentIndex = Math.Clamp(snapshot.CurrentPeriodIndex, 0, snapshot.PeriodSummaries.Count - 1);
    var dueDate = snapshot.PeriodSummaries[currentIndex].PeriodStart.Date;
    return BudgetWorkspaceService.GetManualOverrideAmount(
        databasePath,
        dueDate,
        DebtBalanceHistoryCategoryIndex,
        BuildDebtOverrideKey(debt),
        debt.Description);
}

static void ApplyCurrentDebtBalanceOverride(string databasePath, BudgetWorkspaceSnapshot snapshot, DebtRecord debt, decimal? amount)
{
    if (snapshot.PeriodSummaries.Count == 0)
    {
        return;
    }

    var currentIndex = Math.Clamp(snapshot.CurrentPeriodIndex, 0, snapshot.PeriodSummaries.Count - 1);
    var dueDate = snapshot.PeriodSummaries[currentIndex].PeriodStart.Date;
    if (amount.HasValue)
    {
        BudgetWorkspaceService.SaveManualOverride(
            databasePath,
            dueDate,
            DebtBalanceHistoryCategoryIndex,
            BuildDebtOverrideKey(debt),
            debt.Description,
            amount.Value);
        return;
    }

    BudgetWorkspaceService.DeleteManualOverride(
        databasePath,
        dueDate,
        DebtBalanceHistoryCategoryIndex,
        BuildDebtOverrideKey(debt),
        debt.Description);
}

static void PreservePastDebtValues(string databasePath, BudgetWorkspaceSnapshot snapshot, DebtRecord previousDebt, DebtRecord savedDebt)
{
    if (snapshot.PeriodSummaries.Count == 0 || snapshot.CurrentPeriodIndex <= 0)
    {
        return;
    }

    var preservedPeriods = Math.Clamp(snapshot.CurrentPeriodIndex, 0, snapshot.PeriodSummaries.Count);
    var previousDescription = previousDebt.Description?.Trim() ?? string.Empty;
    var previousKey = BuildDebtOverrideKey(previousDebt);
    var savedKey = BuildDebtOverrideKey(savedDebt);
    var savedDescription = savedDebt.Description?.Trim() ?? string.Empty;

    var itemizedRow = snapshot.ItemizedBudgetRows
        .Where(row =>
            string.Equals(row.SectionName, "Debts", StringComparison.OrdinalIgnoreCase) &&
            (string.Equals(row.SourceKey, previousKey, StringComparison.OrdinalIgnoreCase) ||
             string.Equals(string.IsNullOrWhiteSpace(row.SourceLabel) ? row.Label : row.SourceLabel, previousDescription, StringComparison.OrdinalIgnoreCase)))
        .OrderByDescending(row => string.Equals(row.SourceKey, previousKey, StringComparison.OrdinalIgnoreCase))
        .FirstOrDefault();

    if (itemizedRow is not null)
    {
        for (var i = 0; i < Math.Min(preservedPeriods, itemizedRow.Values.Length); i++)
        {
            BudgetWorkspaceService.SaveManualOverride(
                databasePath,
                snapshot.PeriodSummaries[i].PeriodStart.Date,
                1,
                savedKey,
                savedDescription,
                itemizedRow.Values[i]);
        }
    }

    if (!string.IsNullOrWhiteSpace(previousDescription) &&
        snapshot.DebtRunningBalances.TryGetValue(previousDescription, out var runningBalances))
    {
        for (var i = 0; i < Math.Min(preservedPeriods, runningBalances.Length); i++)
        {
            BudgetWorkspaceService.SaveManualOverride(
                databasePath,
                snapshot.PeriodSummaries[i].PeriodStart.Date,
                DebtBalanceHistoryCategoryIndex,
                savedKey,
                savedDescription,
                runningBalances[i]);
        }
    }
}

static DebtSummaryResponse CreateDebtSummaryResponse(
    DebtRecord debt,
    decimal currentBalance,
    int? payoffIndex,
    BudgetWorkspaceSnapshot snapshot,
    string deleteBlockingReason)
{
    return new DebtSummaryResponse
    {
        Id = debt.Id,
        Description = debt.Description,
        Category = debt.Category,
        DebtType = debt.DebtType,
        Lender = debt.Lender,
        Apr = debt.Apr,
        StartingBalance = debt.StartingBalance,
        OriginalPrincipal = debt.OriginalPrincipal,
        MinPayment = debt.MinPayment,
        DayDue = debt.DayDue,
        FromAccount = debt.FromAccount,
        LoginLink = debt.LoginLink,
        Notes = debt.Notes,
        Cadence = debt.Cadence,
        SameAs = debt.SameAs,
        StartDate = debt.StartDate,
        LastPaymentDate = debt.LastPaymentDate,
        TermMonths = debt.TermMonths,
        MaturityDate = debt.MaturityDate,
        PromoApr = debt.PromoApr,
        PromoStartDate = debt.PromoStartDate,
        PromoAprEndDate = debt.PromoAprEndDate,
        CreditLimit = debt.CreditLimit,
        EscrowIncluded = debt.EscrowIncluded,
        EscrowMonthly = debt.EscrowMonthly,
        PmiMonthly = debt.PmiMonthly,
        DeferredUntil = debt.DeferredUntil,
        DeferredStatus = debt.DeferredStatus,
        Subsidized = debt.Subsidized,
        BalloonAmount = debt.BalloonAmount,
        BalloonDueDate = debt.BalloonDueDate,
        InterestOnlyStartDate = debt.InterestOnlyStartDate,
        InterestOnlyEndDate = debt.InterestOnlyEndDate,
        ForgivenessDate = debt.ForgivenessDate,
        StudentRepaymentPlan = debt.StudentRepaymentPlan,
        RateChangeSchedule = debt.RateChangeSchedule,
        CustomInterestRule = debt.CustomInterestRule,
        CustomFeeRule = debt.CustomFeeRule,
        DayCountBasis = debt.DayCountBasis,
        PaymentsPerYear = debt.PaymentsPerYear,
        CurrentBalance = currentBalance,
        ExpectedPayoffDate = payoffIndex.HasValue && payoffIndex.Value < snapshot.PeriodSummaries.Count
            ? snapshot.PeriodSummaries[payoffIndex.Value].PeriodStart
            : null,
        IsHidden = debt.Hidden,
        IsActive = debt.Active,
        CanDelete = string.IsNullOrWhiteSpace(deleteBlockingReason),
        DeleteBlockedReason = deleteBlockingReason
    };
}

static Dictionary<string, decimal> BuildPeriodAssignmentTotals(string databasePath, BudgetWorkspaceSnapshot snapshot, int targetPeriodIndex, string periodName)
{
    var transactions = TransactionRepository.LoadTransactions(databasePath).ToDictionary(item => item.Id);
    var assignments = TransactionRepository.LoadTransactionAssignments(databasePath);
    var periods = snapshot.PeriodSummaries.Select(item => item.PeriodStart).ToList();
    var totals = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);

    foreach (var assignment in assignments)
    {
        if (!transactions.TryGetValue(assignment.TransactionId, out var transaction) || transaction is null)
        {
            continue;
        }

        if (!DateTime.TryParse(transaction.TransactionDate, CultureInfo.CurrentCulture, DateTimeStyles.None, out var transactionDate) &&
            !DateTime.TryParse(transaction.TransactionDate, CultureInfo.InvariantCulture, DateTimeStyles.None, out transactionDate))
        {
            continue;
        }

        var periodIndex = FindPeriodIndexForDate(periods, transactionDate.Date, periodName);
        if (periodIndex != targetPeriodIndex)
        {
            continue;
        }

        var key = $"ID:{assignment.CatIdx}:{assignment.ItemId}";
        if (!totals.ContainsKey(key))
        {
            totals[key] = 0m;
        }

        totals[key] += Math.Abs(assignment.Amount);
    }

    return totals;
}

static IReadOnlyList<string> GetBudgetCellLinkedTransactionLines(
    BudgetWorkspaceSnapshot snapshot,
    string databasePath,
    int targetPeriodIndex,
    int categoryIndex,
    string itemLabel,
    string sourceKey)
{
    if (string.IsNullOrWhiteSpace(databasePath) || !File.Exists(databasePath) || targetPeriodIndex < 0)
    {
        return Array.Empty<string>();
    }

    var periods = snapshot.PeriodSummaries.Select(item => item.PeriodStart).ToList();
    var periodName = snapshot.BudgetPeriod;
    var itemId = ParseItemIdFromSourceKey(sourceKey);
    if (!itemId.HasValue || itemId.Value <= 0)
    {
        return Array.Empty<string>();
    }

    var incomeLookup = IncomeRepository.LoadIncome(databasePath).ToDictionary(item => item.Id, item => item.Description ?? string.Empty);
    var debtLookup = DebtRepository.LoadDebts(databasePath).ToDictionary(item => item.Id, item => item.Description ?? string.Empty);
    var expenseLookup = ExpenseRepository.LoadExpenses(databasePath).ToDictionary(item => item.Id, item => item.Description ?? string.Empty);
    var savingsLookup = SavingsRepository.LoadSavings(databasePath).ToDictionary(item => item.Id, item => item.Description ?? string.Empty);
    var transactions = TransactionRepository.LoadTransactions(databasePath).ToDictionary(item => item.Id);
    var assignments = TransactionRepository.LoadTransactionAssignments(databasePath);
    var lines = new List<(DateTime sortDate, string line)>();

    foreach (var assignment in assignments.Where(item => item.CatIdx == categoryIndex && item.ItemId == itemId.Value))
    {
        if (!transactions.TryGetValue(assignment.TransactionId, out var transaction) || transaction is null)
        {
            continue;
        }

        if (!DateTime.TryParse(transaction.TransactionDate, CultureInfo.CurrentCulture, DateTimeStyles.None, out var transactionDate) &&
            !DateTime.TryParse(transaction.TransactionDate, CultureInfo.InvariantCulture, DateTimeStyles.None, out transactionDate))
        {
            continue;
        }

        var periodIndex = FindPeriodIndexForDate(periods, transactionDate.Date, periodName);
        if (periodIndex != targetPeriodIndex)
        {
            continue;
        }

        var resolvedLabel = ResolveTransactionAssignmentItemLabel(categoryIndex, itemId.Value, incomeLookup, debtLookup, expenseLookup, savingsLookup);
        if (!string.Equals(resolvedLabel, itemLabel, StringComparison.CurrentCultureIgnoreCase) &&
            !string.Equals(sourceKey, $"ID:{categoryIndex}:{itemId.Value}", StringComparison.OrdinalIgnoreCase))
        {
            continue;
        }

        var sourceName = string.IsNullOrWhiteSpace(transaction.SourceName) ? "Unknown Source" : transaction.SourceName.Trim();
        lines.Add((transactionDate, $"{sourceName} - {transactionDate:MM/dd/yyyy} - {transaction.Description} - {transaction.Amount:C2}"));
    }

    return lines
        .OrderBy(item => item.sortDate)
        .ThenBy(item => item.line, StringComparer.CurrentCultureIgnoreCase)
        .Select(item => item.line)
        .ToList();
}

static BudgetPeriodItemResponse? BuildBudgetPeriodItemResponse(
    BudgetWorkspaceItemSeries row,
    int periodIndex,
    BudgetWorkspaceSnapshot snapshot,
    IReadOnlyDictionary<int, IncomeRecord> incomeLookup,
    IReadOnlyDictionary<int, ExpenseRecord> expenseLookup,
    IReadOnlyDictionary<int, SavingsRecord> savingsLookup,
    IReadOnlyDictionary<int, DebtRecord> debtLookup,
    IReadOnlyDictionary<string, decimal> assignmentTotals)
{
    if (periodIndex < 0)
    {
        return null;
    }

    var categoryIndex = ParseCategoryIndexFromSourceKey(row.SourceKey, row.SectionName);
    if (categoryIndex < 0)
    {
        return null;
    }

    var itemId = ParseItemIdFromSourceKey(row.SourceKey);
    var scheduled = periodIndex < row.ScheduledValues.Length ? row.ScheduledValues[periodIndex] : 0m;
    var effective = periodIndex < row.Values.Length ? row.Values[periodIndex] : scheduled;
    var sourceAdditional = periodIndex < row.SourceAdditionalValues.Length ? row.SourceAdditionalValues[periodIndex] : 0m;
    var manualAdditional = periodIndex < row.ManualAdditionalValues.Length ? row.ManualAdditionalValues[periodIndex] : 0m;
    var additional = manualAdditional != 0m ? manualAdditional : sourceAdditional;
    var actual = assignmentTotals.TryGetValue(row.SourceKey, out var total) ? total : 0m;
    var difference = effective - actual;
    var isSavingsLinked = false;
    decimal? linkedSavingsAvailable = null;
    string linkedFundingSource = string.Empty;

    if (categoryIndex == 2 && itemId.HasValue && expenseLookup.TryGetValue(itemId.Value, out var expense))
    {
        linkedFundingSource = expense.FromAccount ?? string.Empty;
        var savingsName = TryGetSavingsSourceName(expense.FromAccount);
        if (!string.IsNullOrWhiteSpace(savingsName) && snapshot.SavingsRunningBalances.TryGetValue(savingsName, out var balances))
        {
            isSavingsLinked = true;
            linkedSavingsAvailable = periodIndex < balances.Length
                ? balances[periodIndex]
                : balances.Length > 0 ? balances[^1] : 0m;
        }
    }
    else if (categoryIndex == 3 && itemId.HasValue && savingsLookup.TryGetValue(itemId.Value, out var savings))
    {
        linkedFundingSource = savings.FromAccount ?? string.Empty;
        var savingsKey = string.IsNullOrWhiteSpace(savings.Description) ? $"Savings {savings.Id}" : savings.Description.Trim();
        if (snapshot.SavingsRunningBalances.TryGetValue(savingsKey, out var balances))
        {
            linkedSavingsAvailable = periodIndex < balances.Length
                ? balances[periodIndex]
                : balances.Length > 0 ? balances[^1] : 0m;
        }
    }
    else if (categoryIndex == 1 && itemId.HasValue && debtLookup.TryGetValue(itemId.Value, out var debt))
    {
        linkedFundingSource = debt.FromAccount ?? string.Empty;
    }

    var status = BuildBudgetItemStatusText(isSavingsLinked, linkedSavingsAvailable, effective, actual, difference);
    var hasRecognizedBudgetItem = categoryIndex is 0 or 1 or 2 or 3;
    var isVisible = hasRecognizedBudgetItem || scheduled != 0m || effective != 0m || actual != 0m || additional != 0m || isSavingsLinked;
    if (!isVisible)
    {
        return null;
    }

    return new BudgetPeriodItemResponse
    {
        PeriodIndex = periodIndex,
        SectionName = row.SectionName,
        GroupName = row.GroupName,
        Label = row.Label,
        SourceKey = row.SourceKey,
        CategoryIndex = categoryIndex,
        ItemId = itemId,
        ScheduledAmount = scheduled,
        EffectiveAmount = effective,
        Additional = additional,
        ActualAmount = actual,
        DifferenceAmount = difference,
        IsPaid = row.PaidIndexes.Contains(periodIndex),
        IsSavingsLinked = isSavingsLinked,
        LinkedSavingsAvailable = linkedSavingsAvailable,
        FundingSource = linkedFundingSource,
        StatusText = status
    };
}

static IEnumerable<BudgetPeriodItemResponse> BuildSupplementalBudgetItems(
    BudgetWorkspaceSnapshot snapshot,
    int periodIndex,
    IReadOnlyDictionary<int, IncomeRecord> incomeLookup,
    IReadOnlyDictionary<int, ExpenseRecord> expenseLookup,
    IReadOnlyDictionary<int, SavingsRecord> savingsLookup,
    IReadOnlyDictionary<int, DebtRecord> debtLookup,
    ISet<string> existingKeys)
{
    foreach (var income in incomeLookup.Values.Where(item => item.Active && !item.Hidden))
    {
        var sourceKey = $"ID:0:{income.Id}";
        if (existingKeys.Contains(sourceKey))
        {
            continue;
        }

        yield return new BudgetPeriodItemResponse
        {
            PeriodIndex = periodIndex,
            SectionName = "Income",
            GroupName = "Income",
            Label = string.IsNullOrWhiteSpace(income.Description) ? $"Income {income.Id}" : income.Description.Trim(),
            SourceKey = sourceKey,
            CategoryIndex = 0,
            ItemId = income.Id,
            ScheduledAmount = 0m,
            EffectiveAmount = 0m,
            Additional = 0m,
            ActualAmount = 0m,
            DifferenceAmount = 0m,
            IsPaid = false,
            IsSavingsLinked = false,
            LinkedSavingsAvailable = null,
            FundingSource = income.ToAccount ?? string.Empty,
            StatusText = "Budgeted: $0.00"
        };
    }

    foreach (var expense in expenseLookup.Values.Where(item => item.Active && !item.Hidden))
    {
        var sourceKey = $"ID:2:{expense.Id}";
        if (existingKeys.Contains(sourceKey))
        {
            continue;
        }

        var savingsName = TryGetSavingsSourceName(expense.FromAccount);
        decimal[]? savingsBalances = null;
        var isSavingsLinked = !string.IsNullOrWhiteSpace(savingsName) &&
                              snapshot.SavingsRunningBalances.TryGetValue(savingsName, out savingsBalances);
        decimal? linkedSavingsAvailable = null;
        if (isSavingsLinked && savingsBalances is not null)
        {
            linkedSavingsAvailable = periodIndex < savingsBalances.Length
                ? savingsBalances[periodIndex]
                : savingsBalances.Length > 0 ? savingsBalances[^1] : 0m;
        }

        yield return new BudgetPeriodItemResponse
        {
            PeriodIndex = periodIndex,
            SectionName = "Expenses",
            GroupName = string.IsNullOrWhiteSpace(expense.Category) ? "Expenses" : expense.Category.Trim(),
            Label = string.IsNullOrWhiteSpace(expense.Description) ? $"Expense {expense.Id}" : expense.Description.Trim(),
            SourceKey = sourceKey,
            CategoryIndex = 2,
            ItemId = expense.Id,
            ScheduledAmount = 0m,
            EffectiveAmount = 0m,
            Additional = 0m,
            ActualAmount = 0m,
            DifferenceAmount = 0m,
            IsPaid = false,
            IsSavingsLinked = isSavingsLinked,
            LinkedSavingsAvailable = linkedSavingsAvailable,
            FundingSource = expense.FromAccount ?? string.Empty,
            StatusText = isSavingsLinked && linkedSavingsAvailable.HasValue
                ? $"Available in savings: {linkedSavingsAvailable.Value:C}"
                : "Budgeted: $0.00"
        };
    }

    foreach (var debt in debtLookup.Values.Where(item => item.Active && !item.Hidden))
    {
        var sourceKey = $"ID:1:{debt.Id}";
        if (existingKeys.Contains(sourceKey))
        {
            continue;
        }

        yield return new BudgetPeriodItemResponse
        {
            PeriodIndex = periodIndex,
            SectionName = "Debts",
            GroupName = string.IsNullOrWhiteSpace(debt.Category) ? "Debts" : debt.Category.Trim(),
            Label = string.IsNullOrWhiteSpace(debt.Description) ? $"Debt {debt.Id}" : debt.Description.Trim(),
            SourceKey = sourceKey,
            CategoryIndex = 1,
            ItemId = debt.Id,
            ScheduledAmount = 0m,
            EffectiveAmount = 0m,
            Additional = 0m,
            ActualAmount = 0m,
            DifferenceAmount = 0m,
            IsPaid = false,
            IsSavingsLinked = false,
            LinkedSavingsAvailable = null,
            FundingSource = debt.FromAccount ?? string.Empty,
            StatusText = "Budgeted: $0.00"
        };
    }

    foreach (var savings in savingsLookup.Values.Where(item => item.Active && !item.Hidden))
    {
        var sourceKey = $"ID:3:{savings.Id}";
        if (existingKeys.Contains(sourceKey))
        {
            continue;
        }

        decimal? currentSavingsBalance = null;
        var savingsKey = string.IsNullOrWhiteSpace(savings.Description) ? $"Savings {savings.Id}" : savings.Description.Trim();
        if (snapshot.SavingsRunningBalances.TryGetValue(savingsKey, out var balances))
        {
            currentSavingsBalance = periodIndex < balances.Length
                ? balances[periodIndex]
                : balances.Length > 0 ? balances[^1] : 0m;
        }

        yield return new BudgetPeriodItemResponse
        {
            PeriodIndex = periodIndex,
            SectionName = "Savings",
            GroupName = string.IsNullOrWhiteSpace(savings.Category) ? "Savings" : savings.Category.Trim(),
            Label = string.IsNullOrWhiteSpace(savings.Description) ? $"Savings {savings.Id}" : savings.Description.Trim(),
            SourceKey = sourceKey,
            CategoryIndex = 3,
            ItemId = savings.Id,
            ScheduledAmount = 0m,
            EffectiveAmount = 0m,
            Additional = 0m,
            ActualAmount = 0m,
            DifferenceAmount = 0m,
            IsPaid = false,
            IsSavingsLinked = false,
            LinkedSavingsAvailable = currentSavingsBalance,
            FundingSource = savings.FromAccount ?? string.Empty,
            StatusText = "Budgeted: $0.00"
        };
    }
}

static string BuildBudgetItemStatusText(bool isSavingsLinked, decimal? linkedSavingsAvailable, decimal effective, decimal actual, decimal difference)
{
    if (isSavingsLinked && linkedSavingsAvailable.HasValue)
    {
        return $"Available in savings: {linkedSavingsAvailable.Value:C}";
    }

    if (actual == 0m)
    {
        return $"Budgeted: {effective:C}";
    }

    if (difference < 0m)
    {
        return $"Over by {Math.Abs(difference):C}";
    }

    if (difference > 0m)
    {
        return $"Remaining: {difference:C}";
    }

    return "On budget";
}

static int ParseCategoryIndexFromSourceKey(string sourceKey, string sectionName)
{
    if (!string.IsNullOrWhiteSpace(sourceKey) && sourceKey.StartsWith("ID:", StringComparison.OrdinalIgnoreCase))
    {
        var parts = sourceKey.Split(':', StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length >= 3 && int.TryParse(parts[1], out var parsedCategory))
        {
            return parsedCategory;
        }
    }

    return sectionName switch
    {
        "Income" => 0,
        "Debts" => 1,
        "Expenses" => 2,
        "Savings" => 3,
        _ => -1
    };
}

static int? ParseItemIdFromSourceKey(string sourceKey)
{
    if (string.IsNullOrWhiteSpace(sourceKey) || !sourceKey.StartsWith("ID:", StringComparison.OrdinalIgnoreCase))
    {
        return null;
    }

    var parts = sourceKey.Split(':', StringSplitOptions.RemoveEmptyEntries);
    if (parts.Length >= 3 && int.TryParse(parts[2], out var parsedId))
    {
        return parsedId;
    }

    return null;
}

static string TryGetSavingsSourceName(string? fundingSource)
{
    if (string.IsNullOrWhiteSpace(fundingSource))
    {
        return string.Empty;
    }

    const string prefix = "(Savings) ";
    return fundingSource.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)
        ? fundingSource[prefix.Length..].Trim()
        : string.Empty;
}

static int GetBudgetSectionSortOrder(string sectionName) =>
    sectionName switch
    {
        "Expenses" => 0,
        "Debts" => 1,
        "Savings" => 2,
        "Income" => 3,
        _ => 9
    };

static int FindPeriodIndexForDate(IReadOnlyList<DateTime> periods, DateTime targetDate, string periodName)
{
    if (periods.Count == 0)
    {
        return -1;
    }

    for (var index = 0; index < periods.Count; index++)
    {
        var start = periods[index].Date;
        var endExclusive = index + 1 < periods.Count
            ? periods[index + 1].Date
            : periodName switch
            {
                "Weekly" => start.AddDays(7),
                "Bi-Weekly" => start.AddDays(14),
                _ => start.AddMonths(1)
            };

        if (targetDate >= start && targetDate < endExclusive)
        {
            return index;
        }
    }

    return -1;
}

internal sealed record BudgetServerOptions
{
    public string DatabasePath { get; init; } = string.Empty;
    public string SettingsPath { get; init; } = string.Empty;
    public string BudgetPeriod { get; init; } = "Monthly";
    public string BudgetStartDate { get; init; } = string.Empty;
    public int BudgetYears { get; init; } = 20;
}

internal sealed record ResolvedPaths(string DatabasePath, string SettingsPath);

internal sealed record SnapshotContext(ResolvedPaths? Paths, BudgetWorkspaceSnapshot? Snapshot, ApiErrorResponse? Error);

public sealed class OverviewSummaryResponse
{
    public string DatabasePath { get; init; } = string.Empty;
    public string SettingsPath { get; init; } = string.Empty;
    public DateTimeOffset GeneratedAt { get; init; }
    public string BudgetPeriod { get; init; } = string.Empty;
    public int BudgetYears { get; init; }
    public int CurrentPeriodIndex { get; init; }
    public DateTime? CurrentPeriodStart { get; init; }
    public int TotalPeriodCount { get; init; }
    public decimal IncomeTotal { get; init; }
    public decimal PlannedOutflow { get; init; }
    public decimal ExpenseTotal { get; init; }
    public decimal DebtPaymentTotal { get; init; }
    public decimal SavingsContributionTotal { get; init; }
    public decimal NetFlow { get; init; }
    public decimal CurrentDebtBalance { get; init; }
    public int AccountCount { get; init; }
    public int SavingsCount { get; init; }
    public int DebtCount { get; init; }
    public OverviewSummaryAccountLow? LowestAccount { get; init; }
}

public sealed class OverviewSummaryAccountLow
{
    public string Label { get; init; } = string.Empty;
    public decimal Value { get; init; }
    public DateTime PeriodStart { get; init; }
}

public sealed class ApiErrorResponse
{
    public string Error { get; init; } = string.Empty;
}

public sealed class ServerConfigResponse
{
    public string ConfigPath { get; init; } = string.Empty;
    public string DatabasePath { get; init; } = string.Empty;
    public string SettingsPath { get; init; } = string.Empty;
    public string BudgetPeriod { get; init; } = "Monthly";
    public string BudgetStartDate { get; init; } = string.Empty;
    public int BudgetYears { get; init; }
    public string Urls { get; init; } = string.Empty;
    public IReadOnlyList<string> AccessUrls { get; init; } = Array.Empty<string>();
    public bool ConfigFileExists { get; init; }
    public bool DatabaseExists { get; init; }
    public bool SettingsExists { get; init; }
    public IReadOnlyList<ServerDatabaseOption> Databases { get; init; } = Array.Empty<ServerDatabaseOption>();
}

public sealed class ServerConfigSaveRequest
{
    public string DatabasePath { get; init; } = string.Empty;
    public string SettingsPath { get; init; } = string.Empty;
    public string BudgetPeriod { get; init; } = "Monthly";
    public string BudgetStartDate { get; init; } = string.Empty;
    public int BudgetYears { get; init; } = 20;
    public string Urls { get; init; } = "http://0.0.0.0:5099";
}

public sealed class ServerConfigSaveResponse
{
    public string Message { get; init; } = string.Empty;
    public string ConfigPath { get; init; } = string.Empty;
    public string DatabasePath { get; init; } = string.Empty;
    public string SettingsPath { get; init; } = string.Empty;
    public string Urls { get; init; } = string.Empty;
    public IReadOnlyList<string> AccessUrls { get; init; } = Array.Empty<string>();
}

public sealed class ServerConfigCreateDatabaseResponse
{
    public string Message { get; init; } = string.Empty;
    public string ConfigPath { get; init; } = string.Empty;
    public string DatabasePath { get; init; } = string.Empty;
    public string SettingsPath { get; init; } = string.Empty;
    public string Urls { get; init; } = string.Empty;
    public IReadOnlyList<string> AccessUrls { get; init; } = Array.Empty<string>();
}

public sealed class ServerConfigDatabaseTransferResponse
{
    public string Message { get; init; } = string.Empty;
    public string DatabasePath { get; init; } = string.Empty;
}

public sealed class ServerDatabaseOption
{
    public string Name { get; init; } = string.Empty;
    public string Path { get; init; } = string.Empty;
}

public sealed class ChangeTokenResponse
{
    public string DatabasePath { get; init; } = string.Empty;
    public DateTime LastWriteUtc { get; init; }
    public long Length { get; init; }
    public string Token { get; init; } = string.Empty;
}

internal sealed class ServerRuntimeConsole : IDisposable
{
    private readonly object _watcherGate = new();
    private readonly ConcurrentDictionary<string, DateTime> _recentClientLogTimesUtc = new(StringComparer.OrdinalIgnoreCase);
    private FileSystemWatcher? _databaseWatcher;
    private string _watchedDatabasePath = string.Empty;
    private DateTime _lastDatabaseEventLoggedUtc = DateTime.MinValue;

    public void LogClientConnection(HttpContext context)
    {
        var path = context.Request.Path;
        if (!path.StartsWithSegments("/api", StringComparison.OrdinalIgnoreCase) &&
            !path.Equals("/", StringComparison.OrdinalIgnoreCase) &&
            !path.Equals("/setup", StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        var remoteIp = NormalizeIpAddress(context.Connection.RemoteIpAddress);
        if (string.IsNullOrWhiteSpace(remoteIp))
        {
            return;
        }

        var nowUtc = DateTime.UtcNow;
        if (_recentClientLogTimesUtc.TryGetValue(remoteIp, out var previousUtc) &&
            nowUtc - previousUtc < TimeSpan.FromMinutes(10))
        {
            return;
        }

        _recentClientLogTimesUtc[remoteIp] = nowUtc;
        ServerConsoleStyle.WriteColoredLine($"Client connected: {remoteIp}", ConsoleColor.Green);
        ServerAppLogger.LogInfo($"Client connected: {remoteIp}");
    }

    public void WatchDatabase(string databasePath)
    {
        var resolvedPath = string.IsNullOrWhiteSpace(databasePath)
            ? string.Empty
            : Path.GetFullPath(databasePath);

        lock (_watcherGate)
        {
            if (string.Equals(_watchedDatabasePath, resolvedPath, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            DisposeWatcher();
            _watchedDatabasePath = resolvedPath;

            if (string.IsNullOrWhiteSpace(resolvedPath))
            {
                ServerConsoleStyle.WriteColoredLine("Database selected: (none configured)", ConsoleColor.Blue);
                ServerAppLogger.LogInfo("Database selected: (none configured)");
                return;
            }

            ServerConsoleStyle.WriteColoredLine($"Database selected: {Path.GetFileName(resolvedPath)}", ConsoleColor.Blue);
            ServerAppLogger.LogInfo($"Database selected: {Path.GetFileName(resolvedPath)}");

            var directory = Path.GetDirectoryName(resolvedPath);
            var fileName = Path.GetFileName(resolvedPath);
            if (string.IsNullOrWhiteSpace(directory) || string.IsNullOrWhiteSpace(fileName) || !Directory.Exists(directory))
            {
                return;
            }

            _databaseWatcher = new FileSystemWatcher(directory, fileName)
            {
                NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite | NotifyFilters.Size | NotifyFilters.CreationTime,
                IncludeSubdirectories = false,
                EnableRaisingEvents = true
            };
            _databaseWatcher.Changed += (_, _) => LogDatabaseEvent("changed");
            _databaseWatcher.Created += (_, _) => LogDatabaseEvent("created");
            _databaseWatcher.Deleted += (_, _) => LogDatabaseEvent("deleted");
            _databaseWatcher.Renamed += (_, _) => LogDatabaseEvent("renamed");
        }
    }

    public void Dispose()
    {
        lock (_watcherGate)
        {
            DisposeWatcher();
        }
    }

    private void LogDatabaseEvent(string eventName)
    {
        lock (_watcherGate)
        {
            var nowUtc = DateTime.UtcNow;
            if (nowUtc - _lastDatabaseEventLoggedUtc < TimeSpan.FromSeconds(1))
            {
                return;
            }

            _lastDatabaseEventLoggedUtc = nowUtc;
            var databaseName = string.IsNullOrWhiteSpace(_watchedDatabasePath)
                ? "(unknown)"
                : Path.GetFileName(_watchedDatabasePath);
            Console.WriteLine($"Source database {eventName}: {databaseName}");
            ServerAppLogger.LogInfo($"Source database {eventName}: {databaseName}");
        }
    }

    private void DisposeWatcher()
    {
        if (_databaseWatcher is null)
        {
            return;
        }

        _databaseWatcher.EnableRaisingEvents = false;
        _databaseWatcher.Dispose();
        _databaseWatcher = null;
        _lastDatabaseEventLoggedUtc = DateTime.MinValue;
    }

    private static string NormalizeIpAddress(IPAddress? address)
    {
        if (address is null)
        {
            return string.Empty;
        }

        if (address.IsIPv4MappedToIPv6)
        {
            address = address.MapToIPv4();
        }

        return address.ToString();
    }
}

internal static class ServerConsoleStyle
{
    public static void WriteColoredLine(string message, ConsoleColor color)
    {
        var originalColor = Console.ForegroundColor;
        try
        {
            Console.ForegroundColor = color;
            Console.WriteLine(message);
        }
        finally
        {
            Console.ForegroundColor = originalColor;
        }
    }
}

internal static class ServerSetupUrlHelper
{
    public static string GetSetupLaunchUrl(string urls)
    {
        var raw = string.IsNullOrWhiteSpace(urls)
            ? "http://localhost:5099"
            : urls.Trim();

        if (!Uri.TryCreate(raw, UriKind.Absolute, out var uri))
        {
            return string.Empty;
        }

        var port = uri.Port > 0 ? uri.Port : 5099;
        var scheme = string.IsNullOrWhiteSpace(uri.Scheme) ? "http" : uri.Scheme;
        return $"{scheme}://localhost:{port}/setup";
    }
}

public sealed class BudgetPeriodResponse
{
    public int Index { get; init; }
    public DateTime Start { get; init; }
    public string Label { get; init; } = string.Empty;
    public bool IsCurrent { get; init; }
}

public sealed class DesktopBudgetWorkspaceResponse
{
    public string DatabasePath { get; init; } = string.Empty;
    public string SettingsPath { get; init; } = string.Empty;
    public BudgetWorkspaceSettings Settings { get; init; } = new();
    public BudgetWorkspaceSnapshot Snapshot { get; init; } = new();
}

public sealed class DebtSummaryResponse
{
    public int Id { get; init; }
    public string Description { get; init; } = string.Empty;
    public string Category { get; init; } = string.Empty;
    public string DebtType { get; init; } = string.Empty;
    public string Lender { get; init; } = string.Empty;
    public decimal Apr { get; init; }
    public decimal StartingBalance { get; init; }
    public decimal OriginalPrincipal { get; init; }
    public decimal MinPayment { get; init; }
    public int? DayDue { get; init; }
    public string FromAccount { get; init; } = string.Empty;
    public string LoginLink { get; init; } = string.Empty;
    public string Notes { get; init; } = string.Empty;
    public string Cadence { get; init; } = string.Empty;
    public string SameAs { get; init; } = string.Empty;
    public string StartDate { get; init; } = string.Empty;
    public string LastPaymentDate { get; init; } = string.Empty;
    public int? TermMonths { get; init; }
    public string MaturityDate { get; init; } = string.Empty;
    public decimal PromoApr { get; init; }
    public string PromoStartDate { get; init; } = string.Empty;
    public string PromoAprEndDate { get; init; } = string.Empty;
    public decimal CreditLimit { get; init; }
    public bool EscrowIncluded { get; init; }
    public decimal EscrowMonthly { get; init; }
    public decimal PmiMonthly { get; init; }
    public string DeferredUntil { get; init; } = string.Empty;
    public bool DeferredStatus { get; init; }
    public bool Subsidized { get; init; }
    public decimal BalloonAmount { get; init; }
    public string BalloonDueDate { get; init; } = string.Empty;
    public string InterestOnlyStartDate { get; init; } = string.Empty;
    public string InterestOnlyEndDate { get; init; } = string.Empty;
    public string ForgivenessDate { get; init; } = string.Empty;
    public string StudentRepaymentPlan { get; init; } = string.Empty;
    public string RateChangeSchedule { get; init; } = string.Empty;
    public string CustomInterestRule { get; init; } = string.Empty;
    public string CustomFeeRule { get; init; } = string.Empty;
    public int? DayCountBasis { get; init; }
    public int? PaymentsPerYear { get; init; }
    public decimal CurrentBalance { get; init; }
    public DateTime? ExpectedPayoffDate { get; init; }
    public bool IsHidden { get; init; }
    public bool IsActive { get; init; }
    public bool CanDelete { get; init; }
    public string DeleteBlockedReason { get; init; } = string.Empty;
}

public sealed class DebtSaveRequest
{
    public int Id { get; init; }
    public string Description { get; init; } = string.Empty;
    public string Category { get; init; } = string.Empty;
    public string DebtType { get; init; } = string.Empty;
    public string Lender { get; init; } = string.Empty;
    public decimal Apr { get; init; }
    public decimal StartingBalance { get; init; }
    public decimal OriginalPrincipal { get; init; }
    public decimal MinPayment { get; init; }
    public int? DayDue { get; init; }
    public string FromAccount { get; init; } = string.Empty;
    public bool IsHidden { get; init; }
    public bool IsActive { get; init; } = true;
    public string LoginLink { get; init; } = string.Empty;
    public string Notes { get; init; } = string.Empty;
    public string Cadence { get; init; } = string.Empty;
    public string SameAs { get; init; } = string.Empty;
    public string StartDate { get; init; } = string.Empty;
    public string LastPaymentDate { get; init; } = string.Empty;
    public int? TermMonths { get; init; }
    public string MaturityDate { get; init; } = string.Empty;
    public decimal PromoApr { get; init; }
    public string PromoStartDate { get; init; } = string.Empty;
    public string PromoAprEndDate { get; init; } = string.Empty;
    public decimal CreditLimit { get; init; }
    public bool EscrowIncluded { get; init; }
    public decimal EscrowMonthly { get; init; }
    public decimal PmiMonthly { get; init; }
    public string DeferredUntil { get; init; } = string.Empty;
    public bool DeferredStatus { get; init; }
    public bool Subsidized { get; init; }
    public decimal BalloonAmount { get; init; }
    public string BalloonDueDate { get; init; } = string.Empty;
    public string InterestOnlyStartDate { get; init; } = string.Empty;
    public string InterestOnlyEndDate { get; init; } = string.Empty;
    public string ForgivenessDate { get; init; } = string.Empty;
    public string StudentRepaymentPlan { get; init; } = string.Empty;
    public string RateChangeSchedule { get; init; } = string.Empty;
    public string CustomInterestRule { get; init; } = string.Empty;
    public string CustomFeeRule { get; init; } = string.Empty;
    public int? DayCountBasis { get; init; }
    public int? PaymentsPerYear { get; init; }
    public string PreviousDescription { get; init; } = string.Empty;
    public decimal? CurrentPeriodBalanceOverrideAmount { get; init; }
}

public sealed class DebtDeleteResponse
{
    public int Id { get; init; }
    public string Description { get; init; } = string.Empty;
}

public sealed class SavingsSummaryResponse
{
    public int Id { get; init; }
    public string Description { get; init; } = string.Empty;
    public decimal GoalAmount { get; init; }
    public decimal DepositAmount { get; init; }
    public string GoalDate { get; init; } = string.Empty;
    public bool HasGoal { get; init; }
    public string Frequency { get; init; } = string.Empty;
    public int? OnDay { get; init; }
    public string OnDate { get; init; } = string.Empty;
    public string StartDate { get; init; } = string.Empty;
    public string EndDate { get; init; } = string.Empty;
    public string FromAccount { get; init; } = string.Empty;
    public string SameAs { get; init; } = string.Empty;
    public string Category { get; init; } = string.Empty;
    public decimal CurrentBalance { get; init; }
    public bool IsHidden { get; init; }
    public bool IsActive { get; init; }
    public string LoginLink { get; init; } = string.Empty;
    public string Notes { get; init; } = string.Empty;
    public bool CanDelete { get; init; }
    public string DeleteBlockedReason { get; init; } = string.Empty;
}

public sealed class SavingsSaveRequest
{
    public int Id { get; init; }
    public string Description { get; init; } = string.Empty;
    public decimal GoalAmount { get; init; }
    public decimal DepositAmount { get; init; }
    public string GoalDate { get; init; } = string.Empty;
    public bool HasGoal { get; init; }
    public string Frequency { get; init; } = string.Empty;
    public int? OnDay { get; init; }
    public string OnDate { get; init; } = string.Empty;
    public string StartDate { get; init; } = string.Empty;
    public string EndDate { get; init; } = string.Empty;
    public string FromAccount { get; init; } = string.Empty;
    public string SameAs { get; init; } = string.Empty;
    public string Category { get; init; } = string.Empty;
    public bool IsHidden { get; init; }
    public bool IsActive { get; init; } = true;
    public string LoginLink { get; init; } = string.Empty;
    public string Notes { get; init; } = string.Empty;
    public string PreviousDescription { get; init; } = string.Empty;
}

public sealed class SavingsDeleteResponse
{
    public int Id { get; init; }
    public string Description { get; init; } = string.Empty;
}

public sealed class AccountSummaryResponse
{
    public int Id { get; init; }
    public string Description { get; init; } = string.Empty;
    public string AccountType { get; init; } = string.Empty;
    public string LoginLink { get; init; } = string.Empty;
    public string Notes { get; init; } = string.Empty;
    public decimal CurrentBalance { get; init; }
    public decimal SafetyNet { get; init; }
    public bool IsHidden { get; init; }
    public bool IsActive { get; init; }
    public bool CanDelete { get; init; }
    public string DeleteBlockedReason { get; init; } = string.Empty;
}

public sealed class AccountSaveRequest
{
    public int Id { get; init; }
    public string Description { get; init; } = string.Empty;
    public string AccountType { get; init; } = string.Empty;
    public string LoginLink { get; init; } = string.Empty;
    public string Notes { get; init; } = string.Empty;
    public decimal SafetyNet { get; init; }
    public bool IsHidden { get; init; }
    public bool IsActive { get; init; } = true;
    public string PreviousDescription { get; init; } = string.Empty;
}

public sealed class AccountDeleteResponse
{
    public int Id { get; init; }
    public string Description { get; init; } = string.Empty;
}

public sealed class IncomeSummaryResponse
{
    public int Id { get; init; }
    public string Description { get; init; } = string.Empty;
    public decimal Amount { get; init; }
    public string Cadence { get; init; } = string.Empty;
    public int? OnDay { get; init; }
    public string OnDate { get; init; } = string.Empty;
    public decimal AutoIncrease { get; init; }
    public string AutoIncreaseOnDate { get; init; } = string.Empty;
    public string StartDate { get; init; } = string.Empty;
    public string EndDate { get; init; } = string.Empty;
    public string ToAccount { get; init; } = string.Empty;
    public string SameAs { get; init; } = string.Empty;
    public bool IsHidden { get; init; }
    public bool IsActive { get; init; }
    public string LoginLink { get; init; } = string.Empty;
    public string Notes { get; init; } = string.Empty;
    public bool CanDelete { get; init; }
    public string DeleteBlockedReason { get; init; } = string.Empty;
}

public sealed class IncomeSaveRequest
{
    public int Id { get; init; }
    public string Description { get; init; } = string.Empty;
    public decimal Amount { get; init; }
    public string Cadence { get; init; } = string.Empty;
    public int? OnDay { get; init; }
    public string OnDate { get; init; } = string.Empty;
    public decimal AutoIncrease { get; init; }
    public string AutoIncreaseOnDate { get; init; } = string.Empty;
    public string StartDate { get; init; } = string.Empty;
    public string EndDate { get; init; } = string.Empty;
    public string ToAccount { get; init; } = string.Empty;
    public string SameAs { get; init; } = string.Empty;
    public bool IsHidden { get; init; }
    public bool IsActive { get; init; } = true;
    public string LoginLink { get; init; } = string.Empty;
    public string Notes { get; init; } = string.Empty;
    public string PreviousDescription { get; init; } = string.Empty;
}

public sealed class IncomeDeleteResponse
{
    public int Id { get; init; }
    public string Description { get; init; } = string.Empty;
}

public sealed class ExpenseSummaryResponse
{
    public int Id { get; init; }
    public string Description { get; init; } = string.Empty;
    public decimal AmountDue { get; init; }
    public string Cadence { get; init; } = string.Empty;
    public int? DueDay { get; init; }
    public string DueDate { get; init; } = string.Empty;
    public string FromAccount { get; init; } = string.Empty;
    public string SameAs { get; init; } = string.Empty;
    public string Category { get; init; } = string.Empty;
    public bool IsHidden { get; init; }
    public bool IsActive { get; init; }
    public string LoginLink { get; init; } = string.Empty;
    public string Notes { get; init; } = string.Empty;
}

public sealed class ExpenseSaveRequest
{
    public int Id { get; init; }
    public string Description { get; init; } = string.Empty;
    public decimal AmountDue { get; init; }
    public string Cadence { get; init; } = string.Empty;
    public int? DueDay { get; init; }
    public string DueDate { get; init; } = string.Empty;
    public string FromAccount { get; init; } = string.Empty;
    public string SameAs { get; init; } = string.Empty;
    public string Category { get; init; } = string.Empty;
    public bool IsHidden { get; init; }
    public bool IsActive { get; init; } = true;
    public string LoginLink { get; init; } = string.Empty;
    public string Notes { get; init; } = string.Empty;
}

public sealed class ExpenseDeleteResponse
{
    public int Id { get; init; }
    public string Description { get; init; } = string.Empty;
}

public sealed class TransactionSummaryResponse
{
    public int Id { get; init; }
    public string SourceName { get; init; } = string.Empty;
    public string TransactionDate { get; init; } = string.Empty;
    public string Description { get; init; } = string.Empty;
    public decimal Amount { get; init; }
    public string Notes { get; init; } = string.Empty;
    public IReadOnlyList<TransactionAssignmentSummaryResponse> Assignments { get; init; } = Array.Empty<TransactionAssignmentSummaryResponse>();
}

public sealed class TransactionAssignmentSummaryResponse
{
    public int Id { get; init; }
    public int CategoryIndex { get; init; }
    public string CategoryLabel { get; init; } = string.Empty;
    public int ItemId { get; init; }
    public string ItemLabel { get; init; } = string.Empty;
    public decimal Amount { get; init; }
    public string Notes { get; init; } = string.Empty;
    public bool NeedsReview { get; init; }
}

public sealed class TransactionSaveRequest
{
    public int Id { get; init; }
    public string SourceName { get; init; } = string.Empty;
    public string TransactionDate { get; init; } = string.Empty;
    public string Description { get; init; } = string.Empty;
    public decimal Amount { get; init; }
    public string Notes { get; init; } = string.Empty;
    public IReadOnlyList<TransactionAssignmentSaveRequest> Assignments { get; init; } = Array.Empty<TransactionAssignmentSaveRequest>();
}

public sealed class TransactionAssignmentSaveRequest
{
    public int Id { get; init; }
    public int CategoryIndex { get; init; }
    public int ItemId { get; init; }
    public decimal Amount { get; init; }
    public string Notes { get; init; } = string.Empty;
    public bool NeedsReview { get; init; }
}

public sealed class ApiMessageResponse
{
    public string Message { get; init; } = string.Empty;
}

public sealed class BudgetPeriodItemsResponse
{
    public int PeriodIndex { get; init; }
    public DateTime? PeriodStart { get; init; }
    public int CurrentPeriodIndex { get; init; }
    public int TotalPeriodCount { get; init; }
    public IReadOnlyList<BudgetPeriodItemResponse> Items { get; init; } = Array.Empty<BudgetPeriodItemResponse>();
}

public sealed class BudgetPeriodItemResponse
{
    public int PeriodIndex { get; init; }
    public string SectionName { get; init; } = string.Empty;
    public string GroupName { get; init; } = string.Empty;
    public string Label { get; init; } = string.Empty;
    public string SourceKey { get; init; } = string.Empty;
    public int CategoryIndex { get; init; }
    public int? ItemId { get; init; }
    public decimal ScheduledAmount { get; init; }
    public decimal EffectiveAmount { get; init; }
    public decimal Additional { get; init; }
    public decimal ActualAmount { get; init; }
    public decimal DifferenceAmount { get; init; }
    public bool IsPaid { get; init; }
    public bool IsSavingsLinked { get; init; }
    public decimal? LinkedSavingsAvailable { get; init; }
    public string FundingSource { get; init; } = string.Empty;
    public string StatusText { get; init; } = string.Empty;
}

public sealed class BudgetItemAdditionalUpdateRequest
{
    public int PeriodIndex { get; init; }
    public string SourceKey { get; init; } = string.Empty;
    public string ItemLabel { get; init; } = string.Empty;
    public decimal Additional { get; init; }
}

public sealed class BudgetItemAdditionalUpdateResponse
{
    public int PeriodIndex { get; init; }
    public DateTime PeriodStart { get; init; }
    public string SourceKey { get; init; } = string.Empty;
    public string ItemLabel { get; init; } = string.Empty;
    public decimal Additional { get; init; }
    public string Message { get; init; } = string.Empty;
}

public sealed class BudgetCellEditorStateResponse
{
    public int PeriodIndex { get; init; }
    public DateTime PeriodStart { get; init; }
    public string SourceKey { get; init; } = string.Empty;
    public string ItemLabel { get; init; } = string.Empty;
    public int CategoryIndex { get; init; }
    public decimal? Amount { get; init; }
    public decimal Additional { get; init; }
    public decimal ManualAmount { get; init; }
    public bool Paid { get; init; }
    public int SelectionMode { get; init; }
    public string Notes { get; init; } = string.Empty;
    public decimal? ManualOverrideAmount { get; init; }
    public decimal ManualOverrideAdditional { get; init; }
    public bool ManualOverridePaid { get; init; }
    public int ManualOverrideSelectionMode { get; init; }
    public string ManualOverrideNote { get; init; } = string.Empty;
    public decimal? CurrentDebtBalanceOverrideAmount { get; init; }
    public IReadOnlyList<string> LinkedTransactions { get; init; } = Array.Empty<string>();
}

public sealed class BudgetCellSaveRequest
{
    public int PeriodIndex { get; init; }
    public string SourceKey { get; init; } = string.Empty;
    public string ItemLabel { get; init; } = string.Empty;
    public decimal Amount { get; init; }
    public decimal Additional { get; init; }
    public string Note { get; init; } = string.Empty;
    public bool Paid { get; init; }
    public int SelectionMode { get; init; }
    public decimal ManualAmount { get; init; }
    public decimal? CurrentPeriodBalanceOverrideAmount { get; init; }
}

public sealed class BudgetCellClearRequest
{
    public int PeriodIndex { get; init; }
    public string SourceKey { get; init; } = string.Empty;
    public string ItemLabel { get; init; } = string.Empty;
}

public sealed class BudgetCellSaveResponse
{
    public int PeriodIndex { get; init; }
    public DateTime PeriodStart { get; init; }
    public string SourceKey { get; init; } = string.Empty;
    public string ItemLabel { get; init; } = string.Empty;
    public string Message { get; init; } = string.Empty;
}

internal sealed class StandaloneServerConfigFile
{
    public string Urls { get; init; } = "http://0.0.0.0:5099";
    public StandaloneBudgetServerSection BudgetServer { get; init; } = new();
}

internal sealed class StandaloneBudgetServerSection
{
    public string DatabasePath { get; init; } = string.Empty;
    public string SettingsPath { get; init; } = string.Empty;
    public string BudgetPeriod { get; init; } = "Monthly";
    public string BudgetStartDate { get; init; } = string.Empty;
    public int BudgetYears { get; init; } = 20;
}
