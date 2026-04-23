Imports Microsoft.Data.Sqlite

Namespace Global.JCBudgeting.Core

    Public Class DebtRecord
        Public Property Id As Integer
        Public Property Description As String = String.Empty
        Public Property Category As String = String.Empty
        Public Property DebtType As String = String.Empty
        Public Property Lender As String = String.Empty
        Public Property Apr As Decimal
        Public Property StartingBalance As Decimal
        Public Property OriginalPrincipal As Decimal
        Public Property MinPayment As Decimal
        Public Property DayDue As Integer?
        Public Property FromAccount As String = String.Empty
        Public Property Hidden As Boolean
        Public Property Active As Boolean = True
        Public Property LoginLink As String = String.Empty
        Public Property Notes As String = String.Empty
        Public Property Cadence As String = String.Empty
        Public Property SameAs As String = String.Empty
        Public Property StartDate As String = String.Empty
        Public Property LastPaymentDate As String = String.Empty
        Public Property TermMonths As Integer?
        Public Property MaturityDate As String = String.Empty
        Public Property PromoApr As Decimal
        Public Property PromoStartDate As String = String.Empty
        Public Property PromoAprEndDate As String = String.Empty
        Public Property CreditLimit As Decimal
        Public Property EscrowIncluded As Boolean
        Public Property EscrowMonthly As Decimal
        Public Property PmiMonthly As Decimal
        Public Property DeferredUntil As String = String.Empty
        Public Property DeferredStatus As Boolean
        Public Property Subsidized As Boolean
        Public Property BalloonAmount As Decimal
        Public Property BalloonDueDate As String = String.Empty
        Public Property InterestOnlyStartDate As String = String.Empty
        Public Property InterestOnlyEndDate As String = String.Empty
        Public Property ForgivenessDate As String = String.Empty
        Public Property StudentRepaymentPlan As String = String.Empty
        Public Property RateChangeSchedule As String = String.Empty
        Public Property CustomInterestRule As String = String.Empty
        Public Property CustomFeeRule As String = String.Empty
        Public Property DayCountBasis As Integer?
        Public Property PaymentsPerYear As Integer?
    End Class

    Public Module DebtRepository

        Public Function LoadDebts(databasePath As String) As List(Of DebtRecord)
            Dim results As New List(Of DebtRecord)()
            If String.IsNullOrWhiteSpace(databasePath) OrElse Not IO.File.Exists(databasePath) Then
                Return results
            End If

            EnsureDebtSchema(databasePath)

            Using conn As New SqliteConnection("Data Source=" & databasePath.Trim())
                conn.Open()

                Using cmd = conn.CreateCommand()
                    cmd.CommandText =
                        "SELECT Id, Description, Category, DebtType, Lender, APR, StartingBalance, OriginalPrincipal, PaymentAmount, DayDue, FromAccount, Hidden, COALESCE(Active,1), LoginLink, Notes, Frequency, SameAs, StartDate, LastPaymentDate, TermMonths, MaturityDate, PromoApr, PromoStartDate, PromoAprEndDate, CreditLimit, EscrowIncluded, EscrowAmount, PmiMonthly, DeferredUntil, DeferredStatus, Subsidized, BalloonAmount, BalloonDueDate, InterestOnlyStartDate, InterestOnlyEndDate, ForgivenessDate, StudentRepaymentPlan, RateChangeSchedule, CustomInterestRule, CustomFeeRule, DayCountBasis, PaymentsPerYear " &
                        "FROM debts ORDER BY COALESCE(Hidden,0), COALESCE(Active,1) DESC, Category COLLATE NOCASE, Description COLLATE NOCASE"

                    Using reader = cmd.ExecuteReader()
                        While reader.Read()
                            results.Add(New DebtRecord With {
                                .Id = If(reader.IsDBNull(0), 0, reader.GetInt32(0)),
                                .Description = If(reader.IsDBNull(1), String.Empty, Convert.ToString(reader.GetValue(1))),
                                .Category = If(reader.IsDBNull(2), String.Empty, Convert.ToString(reader.GetValue(2))),
                                .DebtType = If(reader.IsDBNull(3), String.Empty, Convert.ToString(reader.GetValue(3))),
                                .Lender = If(reader.IsDBNull(4), String.Empty, Convert.ToString(reader.GetValue(4))),
                                .Apr = If(reader.IsDBNull(5), 0D, Convert.ToDecimal(reader.GetValue(5))),
                                .StartingBalance = If(reader.IsDBNull(6), 0D, Convert.ToDecimal(reader.GetValue(6))),
                                .OriginalPrincipal = If(reader.IsDBNull(7), 0D, Convert.ToDecimal(reader.GetValue(7))),
                                .MinPayment = If(reader.IsDBNull(8), 0D, Convert.ToDecimal(reader.GetValue(8))),
                                .DayDue = If(reader.IsDBNull(9), CType(Nothing, Integer?), Convert.ToInt32(reader.GetValue(9))),
                                .FromAccount = If(reader.IsDBNull(10), String.Empty, Convert.ToString(reader.GetValue(10))),
                                .Hidden = Not reader.IsDBNull(11) AndAlso Convert.ToInt32(reader.GetValue(11)) <> 0,
                                .Active = reader.IsDBNull(12) OrElse Convert.ToInt32(reader.GetValue(12)) <> 0,
                                .LoginLink = If(reader.IsDBNull(13), String.Empty, Convert.ToString(reader.GetValue(13))),
                                .Notes = If(reader.IsDBNull(14), String.Empty, Convert.ToString(reader.GetValue(14))),
                                .Cadence = If(reader.IsDBNull(15), String.Empty, Convert.ToString(reader.GetValue(15))),
                                .SameAs = If(reader.IsDBNull(16), String.Empty, Convert.ToString(reader.GetValue(16))),
                                .StartDate = If(reader.IsDBNull(17), String.Empty, Convert.ToString(reader.GetValue(17))),
                                .LastPaymentDate = If(reader.IsDBNull(18), String.Empty, Convert.ToString(reader.GetValue(18))),
                                .TermMonths = If(reader.IsDBNull(19), CType(Nothing, Integer?), Convert.ToInt32(reader.GetValue(19))),
                                .MaturityDate = If(reader.IsDBNull(20), String.Empty, Convert.ToString(reader.GetValue(20))),
                                .PromoApr = If(reader.IsDBNull(21), 0D, Convert.ToDecimal(reader.GetValue(21))),
                                .PromoStartDate = If(reader.IsDBNull(22), String.Empty, Convert.ToString(reader.GetValue(22))),
                                .PromoAprEndDate = If(reader.IsDBNull(23), String.Empty, Convert.ToString(reader.GetValue(23))),
                                .CreditLimit = If(reader.IsDBNull(24), 0D, Convert.ToDecimal(reader.GetValue(24))),
                                .EscrowIncluded = Not reader.IsDBNull(25) AndAlso Convert.ToInt32(reader.GetValue(25)) <> 0,
                                .EscrowMonthly = If(reader.IsDBNull(26), 0D, Convert.ToDecimal(reader.GetValue(26))),
                                .PmiMonthly = If(reader.IsDBNull(27), 0D, Convert.ToDecimal(reader.GetValue(27))),
                                .DeferredUntil = If(reader.IsDBNull(28), String.Empty, Convert.ToString(reader.GetValue(28))),
                                .DeferredStatus = Not reader.IsDBNull(29) AndAlso Convert.ToInt32(reader.GetValue(29)) <> 0,
                                .Subsidized = Not reader.IsDBNull(30) AndAlso Convert.ToInt32(reader.GetValue(30)) <> 0,
                                .BalloonAmount = If(reader.IsDBNull(31), 0D, Convert.ToDecimal(reader.GetValue(31))),
                                .BalloonDueDate = If(reader.IsDBNull(32), String.Empty, Convert.ToString(reader.GetValue(32))),
                                .InterestOnlyStartDate = If(reader.IsDBNull(33), String.Empty, Convert.ToString(reader.GetValue(33))),
                                .InterestOnlyEndDate = If(reader.IsDBNull(34), String.Empty, Convert.ToString(reader.GetValue(34))),
                                .ForgivenessDate = If(reader.IsDBNull(35), String.Empty, Convert.ToString(reader.GetValue(35))),
                                .StudentRepaymentPlan = If(reader.IsDBNull(36), String.Empty, Convert.ToString(reader.GetValue(36))),
                                .RateChangeSchedule = If(reader.IsDBNull(37), String.Empty, Convert.ToString(reader.GetValue(37))),
                                .CustomInterestRule = If(reader.IsDBNull(38), String.Empty, Convert.ToString(reader.GetValue(38))),
                                .CustomFeeRule = If(reader.IsDBNull(39), String.Empty, Convert.ToString(reader.GetValue(39))),
                                .DayCountBasis = If(reader.IsDBNull(40), CType(Nothing, Integer?), Convert.ToInt32(reader.GetValue(40))),
                                .PaymentsPerYear = If(reader.IsDBNull(41), CType(Nothing, Integer?), Convert.ToInt32(reader.GetValue(41)))
                            })
                        End While
                    End Using
                End Using
            End Using

            Return results
        End Function

        Public Sub SaveDebt(databasePath As String, debt As DebtRecord)
            If String.IsNullOrWhiteSpace(databasePath) OrElse Not IO.File.Exists(databasePath) Then
                Throw New IO.FileNotFoundException("Budget database was not found.", databasePath)
            End If

            If debt Is Nothing Then
                Throw New ArgumentNullException(NameOf(debt))
            End If

            If debt.Id <= 0 Then
                Throw New InvalidOperationException("Only existing debts can be saved in this migration pass.")
            End If

            EnsureDebtSchema(databasePath)

            Using conn As New SqliteConnection("Data Source=" & databasePath.Trim())
                conn.Open()

                Using cmd = conn.CreateCommand()
                    cmd.CommandText =
                        "UPDATE debts SET " &
                        "Description = @Description, Category = @Category, DebtType = @DebtType, Lender = @Lender, APR = @APR, StartingBalance = @StartingBalance, OriginalPrincipal = @OriginalPrincipal, PaymentAmount = @PaymentAmount, DayDue = @DayDue, FromAccount = @FromAccount, Hidden = @Hidden, Active = @Active, LoginLink = @LoginLink, Notes = @Notes, Frequency = @Frequency, SameAs = @SameAs, StartDate = @StartDate, LastPaymentDate = @LastPaymentDate, TermMonths = @TermMonths, MaturityDate = @MaturityDate, PromoApr = @PromoApr, PromoStartDate = @PromoStartDate, PromoAprEndDate = @PromoAprEndDate, CreditLimit = @CreditLimit, EscrowIncluded = @EscrowIncluded, EscrowAmount = @EscrowAmount, PmiMonthly = @PmiMonthly, DeferredUntil = @DeferredUntil, DeferredStatus = @DeferredStatus, Subsidized = @Subsidized, BalloonAmount = @BalloonAmount, BalloonDueDate = @BalloonDueDate, InterestOnlyStartDate = @InterestOnlyStartDate, InterestOnlyEndDate = @InterestOnlyEndDate, ForgivenessDate = @ForgivenessDate, StudentRepaymentPlan = @StudentRepaymentPlan, RateChangeSchedule = @RateChangeSchedule, CustomInterestRule = @CustomInterestRule, CustomFeeRule = @CustomFeeRule, DayCountBasis = @DayCountBasis, PaymentsPerYear = @PaymentsPerYear " &
                        "WHERE Id = @Id"

                    AddDebtParameters(cmd, debt)
                    cmd.Parameters.AddWithValue("@Id", debt.Id)

                    Dim rowsChanged = cmd.ExecuteNonQuery()
                    If rowsChanged <= 0 Then
                        Throw New InvalidOperationException("The selected debt could not be saved.")
                    End If
                End Using
            End Using
        End Sub

        Public Function CreateDebt(databasePath As String, debt As DebtRecord) As Integer
            If String.IsNullOrWhiteSpace(databasePath) OrElse Not IO.File.Exists(databasePath) Then
                Throw New IO.FileNotFoundException("Budget database was not found.", databasePath)
            End If

            If debt Is Nothing Then
                Throw New ArgumentNullException(NameOf(debt))
            End If

            EnsureDebtSchema(databasePath)

            Using conn As New SqliteConnection("Data Source=" & databasePath.Trim())
                conn.Open()

                Using cmd = conn.CreateCommand()
                    cmd.CommandText =
                        "INSERT INTO debts (" &
                        "Description, Category, DebtType, Lender, APR, StartingBalance, OriginalPrincipal, PaymentAmount, DayDue, FromAccount, Hidden, Active, LoginLink, Notes, Frequency, SameAs, StartDate, LastPaymentDate, TermMonths, MaturityDate, PromoApr, PromoStartDate, PromoAprEndDate, CreditLimit, EscrowIncluded, EscrowAmount, PmiMonthly, DeferredUntil, DeferredStatus, Subsidized, BalloonAmount, BalloonDueDate, InterestOnlyStartDate, InterestOnlyEndDate, ForgivenessDate, StudentRepaymentPlan, RateChangeSchedule, CustomInterestRule, CustomFeeRule, DayCountBasis, PaymentsPerYear" &
                        ") VALUES (" &
                        "@Description, @Category, @DebtType, @Lender, @APR, @StartingBalance, @OriginalPrincipal, @PaymentAmount, @DayDue, @FromAccount, @Hidden, @Active, @LoginLink, @Notes, @Frequency, @SameAs, @StartDate, @LastPaymentDate, @TermMonths, @MaturityDate, @PromoApr, @PromoStartDate, @PromoAprEndDate, @CreditLimit, @EscrowIncluded, @EscrowAmount, @PmiMonthly, @DeferredUntil, @DeferredStatus, @Subsidized, @BalloonAmount, @BalloonDueDate, @InterestOnlyStartDate, @InterestOnlyEndDate, @ForgivenessDate, @StudentRepaymentPlan, @RateChangeSchedule, @CustomInterestRule, @CustomFeeRule, @DayCountBasis, @PaymentsPerYear" &
                        ");" &
                        "SELECT last_insert_rowid();"

                    AddDebtParameters(cmd, debt)

                    Dim rawId = cmd.ExecuteScalar()
                    Return Convert.ToInt32(rawId)
                End Using
            End Using
        End Function

        Public Sub DeleteDebt(databasePath As String, debtId As Integer)
            If String.IsNullOrWhiteSpace(databasePath) OrElse Not IO.File.Exists(databasePath) Then
                Throw New IO.FileNotFoundException("Budget database was not found.", databasePath)
            End If

            If debtId <= 0 Then
                Throw New ArgumentOutOfRangeException(NameOf(debtId))
            End If

            EnsureDebtSchema(databasePath)

            Using conn As New SqliteConnection("Data Source=" & databasePath.Trim())
                conn.Open()

                Using cmd = conn.CreateCommand()
                    cmd.CommandText = "DELETE FROM debts WHERE Id = @Id"
                    cmd.Parameters.AddWithValue("@Id", debtId)

                    Dim rowsChanged = cmd.ExecuteNonQuery()
                    If rowsChanged <= 0 Then
                        Throw New InvalidOperationException("The selected debt could not be deleted.")
                    End If
                End Using
            End Using
        End Sub

        Public Sub RestoreDebt(databasePath As String, debt As DebtRecord)
            If String.IsNullOrWhiteSpace(databasePath) OrElse Not IO.File.Exists(databasePath) Then
                Throw New IO.FileNotFoundException("Budget database was not found.", databasePath)
            End If

            If debt Is Nothing Then
                Throw New ArgumentNullException(NameOf(debt))
            End If

            If debt.Id <= 0 Then
                Throw New InvalidOperationException("A valid debt id is required for restore.")
            End If

            EnsureDebtSchema(databasePath)

            Using conn As New SqliteConnection("Data Source=" & databasePath.Trim())
                conn.Open()

                Using cmd = conn.CreateCommand()
                    cmd.CommandText =
                        "INSERT OR REPLACE INTO debts (" &
                        "Id, Description, Category, DebtType, Lender, APR, StartingBalance, OriginalPrincipal, PaymentAmount, DayDue, FromAccount, Hidden, Active, LoginLink, Notes, Frequency, SameAs, StartDate, LastPaymentDate, TermMonths, MaturityDate, PromoApr, PromoStartDate, PromoAprEndDate, CreditLimit, EscrowIncluded, EscrowAmount, PmiMonthly, DeferredUntil, DeferredStatus, Subsidized, BalloonAmount, BalloonDueDate, InterestOnlyStartDate, InterestOnlyEndDate, ForgivenessDate, StudentRepaymentPlan, RateChangeSchedule, CustomInterestRule, CustomFeeRule, DayCountBasis, PaymentsPerYear" &
                        ") VALUES (" &
                        "@Id, @Description, @Category, @DebtType, @Lender, @APR, @StartingBalance, @OriginalPrincipal, @PaymentAmount, @DayDue, @FromAccount, @Hidden, @Active, @LoginLink, @Notes, @Frequency, @SameAs, @StartDate, @LastPaymentDate, @TermMonths, @MaturityDate, @PromoApr, @PromoStartDate, @PromoAprEndDate, @CreditLimit, @EscrowIncluded, @EscrowAmount, @PmiMonthly, @DeferredUntil, @DeferredStatus, @Subsidized, @BalloonAmount, @BalloonDueDate, @InterestOnlyStartDate, @InterestOnlyEndDate, @ForgivenessDate, @StudentRepaymentPlan, @RateChangeSchedule, @CustomInterestRule, @CustomFeeRule, @DayCountBasis, @PaymentsPerYear" &
                        ")"

                    AddDebtParameters(cmd, debt)
                    cmd.Parameters.AddWithValue("@Id", debt.Id)
                    cmd.ExecuteNonQuery()
                End Using
            End Using
        End Sub

        Private Sub AddDebtParameters(cmd As SqliteCommand, debt As DebtRecord)
            cmd.Parameters.AddWithValue("@Description", If(debt.Description, String.Empty))
            cmd.Parameters.AddWithValue("@Category", If(debt.Category, String.Empty))
            cmd.Parameters.AddWithValue("@DebtType", If(debt.DebtType, String.Empty))
            cmd.Parameters.AddWithValue("@Lender", If(debt.Lender, String.Empty))
            cmd.Parameters.AddWithValue("@APR", debt.Apr)
            cmd.Parameters.AddWithValue("@StartingBalance", debt.StartingBalance)
            cmd.Parameters.AddWithValue("@OriginalPrincipal", debt.OriginalPrincipal)
            cmd.Parameters.AddWithValue("@PaymentAmount", debt.MinPayment)

            Dim dayDueParam = cmd.CreateParameter()
            dayDueParam.ParameterName = "@DayDue"
            dayDueParam.Value = If(debt.DayDue.HasValue, debt.DayDue.Value, DBNull.Value)
            cmd.Parameters.Add(dayDueParam)

            cmd.Parameters.AddWithValue("@FromAccount", If(debt.FromAccount, String.Empty))
            cmd.Parameters.AddWithValue("@Hidden", If(debt.Hidden, 1, 0))
            cmd.Parameters.AddWithValue("@Active", If(debt.Active, 1, 0))
            cmd.Parameters.AddWithValue("@LoginLink", If(debt.LoginLink, String.Empty))
            cmd.Parameters.AddWithValue("@Notes", If(debt.Notes, String.Empty))
            cmd.Parameters.AddWithValue("@Frequency", If(debt.Cadence, String.Empty))
            cmd.Parameters.AddWithValue("@SameAs", If(debt.SameAs, String.Empty))
            cmd.Parameters.AddWithValue("@StartDate", If(debt.StartDate, String.Empty))
            cmd.Parameters.AddWithValue("@LastPaymentDate", If(debt.LastPaymentDate, String.Empty))

            Dim termMonthsParam = cmd.CreateParameter()
            termMonthsParam.ParameterName = "@TermMonths"
            termMonthsParam.Value = If(debt.TermMonths.HasValue, debt.TermMonths.Value, DBNull.Value)
            cmd.Parameters.Add(termMonthsParam)

            cmd.Parameters.AddWithValue("@MaturityDate", If(debt.MaturityDate, String.Empty))
            cmd.Parameters.AddWithValue("@PromoApr", debt.PromoApr)
            cmd.Parameters.AddWithValue("@PromoStartDate", If(debt.PromoStartDate, String.Empty))
            cmd.Parameters.AddWithValue("@PromoAprEndDate", If(debt.PromoAprEndDate, String.Empty))
            cmd.Parameters.AddWithValue("@CreditLimit", debt.CreditLimit)
            cmd.Parameters.AddWithValue("@EscrowIncluded", If(debt.EscrowIncluded, 1, 0))
            cmd.Parameters.AddWithValue("@EscrowAmount", debt.EscrowMonthly)
            cmd.Parameters.AddWithValue("@PmiMonthly", debt.PmiMonthly)
            cmd.Parameters.AddWithValue("@DeferredUntil", If(debt.DeferredUntil, String.Empty))
            cmd.Parameters.AddWithValue("@DeferredStatus", If(debt.DeferredStatus, 1, 0))
            cmd.Parameters.AddWithValue("@Subsidized", If(debt.Subsidized, 1, 0))
            cmd.Parameters.AddWithValue("@BalloonAmount", debt.BalloonAmount)
            cmd.Parameters.AddWithValue("@BalloonDueDate", If(debt.BalloonDueDate, String.Empty))
            cmd.Parameters.AddWithValue("@InterestOnlyStartDate", If(debt.InterestOnlyStartDate, String.Empty))
            cmd.Parameters.AddWithValue("@InterestOnlyEndDate", If(debt.InterestOnlyEndDate, String.Empty))
            cmd.Parameters.AddWithValue("@ForgivenessDate", If(debt.ForgivenessDate, String.Empty))
            cmd.Parameters.AddWithValue("@StudentRepaymentPlan", If(debt.StudentRepaymentPlan, String.Empty))
            cmd.Parameters.AddWithValue("@RateChangeSchedule", If(debt.RateChangeSchedule, String.Empty))
            cmd.Parameters.AddWithValue("@CustomInterestRule", If(debt.CustomInterestRule, String.Empty))
            cmd.Parameters.AddWithValue("@CustomFeeRule", If(debt.CustomFeeRule, String.Empty))

            Dim dayCountBasisParam = cmd.CreateParameter()
            dayCountBasisParam.ParameterName = "@DayCountBasis"
            dayCountBasisParam.Value = If(debt.DayCountBasis.HasValue, debt.DayCountBasis.Value, DBNull.Value)
            cmd.Parameters.Add(dayCountBasisParam)

            Dim paymentsPerYearParam = cmd.CreateParameter()
            paymentsPerYearParam.ParameterName = "@PaymentsPerYear"
            paymentsPerYearParam.Value = If(debt.PaymentsPerYear.HasValue, debt.PaymentsPerYear.Value, DBNull.Value)
            cmd.Parameters.Add(paymentsPerYearParam)
        End Sub

        Private Sub EnsureDebtSchema(databasePath As String)
            Using conn As New SqliteConnection("Data Source=" & databasePath.Trim())
                conn.Open()

                If HasTable(conn, "debts") AndAlso HasTable(conn, "debts_new") Then
                    Using cleanupCmd = conn.CreateCommand()
                        cleanupCmd.CommandText = "DROP TABLE IF EXISTS debts_new"
                        cleanupCmd.ExecuteNonQuery()
                    End Using
                End If

                Dim existingColumns As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                Using existsCmd = conn.CreateCommand()
                    existsCmd.CommandText = "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='debts'"
                    Dim tableExists = Convert.ToInt32(existsCmd.ExecuteScalar()) > 0
                    If tableExists Then
                        Using pragmaCmd = conn.CreateCommand()
                            pragmaCmd.CommandText = "PRAGMA table_info(debts)"
                            Using reader = pragmaCmd.ExecuteReader()
                                While reader.Read()
                                    existingColumns.Add(Convert.ToString(reader.GetValue(1)))
                                End While
                            End Using
                        End Using
                    End If
                End Using

                Dim requiredColumns = {
                    "Id", "Description", "Category", "DebtType", "Lender", "APR", "StartingBalance", "OriginalPrincipal",
                    "PaymentAmount", "DayDue", "FromAccount", "Hidden", "Active", "LoginLink", "Notes", "Frequency",
                    "SameAs", "StartDate", "LastPaymentDate", "TermMonths", "MaturityDate", "PromoApr", "PromoStartDate",
                    "PromoAprEndDate", "CreditLimit", "EscrowIncluded", "EscrowAmount", "PmiMonthly",
                    "DeferredUntil", "DeferredStatus", "Subsidized",
                    "BalloonAmount", "BalloonDueDate", "InterestOnlyStartDate", "InterestOnlyEndDate", "ForgivenessDate",
                    "StudentRepaymentPlan", "RateChangeSchedule", "CustomInterestRule", "CustomFeeRule", "DayCountBasis",
                    "PaymentsPerYear"
                }

                Dim needsReset = existingColumns.Count = 0
                If Not needsReset Then
                    For Each required In requiredColumns
                        If Not existingColumns.Contains(required) Then
                            needsReset = True
                            Exit For
                        End If
                    Next
                End If

                If Not needsReset Then
                    Return
                End If

                Using createCmd = conn.CreateCommand()
                    createCmd.CommandText =
                        "CREATE TABLE debts_new (" &
                        "Id INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT, " &
                        "Description TEXT NOT NULL DEFAULT '', " &
                        "Category TEXT NULL, " &
                        "DebtType TEXT NULL, " &
                        "Lender TEXT NULL, " &
                        "APR REAL NOT NULL DEFAULT 0, " &
                        "StartingBalance REAL NOT NULL DEFAULT 0, " &
                        "OriginalPrincipal REAL NOT NULL DEFAULT 0, " &
                        "PaymentAmount REAL NOT NULL DEFAULT 0, " &
                        "DayDue INTEGER NULL, " &
                        "FromAccount TEXT NULL, " &
                        "Hidden INTEGER NOT NULL DEFAULT 0, " &
                        "Active INTEGER NOT NULL DEFAULT 1, " &
                        "LoginLink TEXT NULL, " &
                        "Notes TEXT NULL, " &
                        "Frequency TEXT NULL, " &
                        "SameAs TEXT NULL, " &
                        "StartDate TEXT NULL, " &
                        "LastPaymentDate TEXT NULL, " &
                        "TermMonths INTEGER NULL, " &
                        "MaturityDate TEXT NULL, " &
                        "PromoApr REAL NOT NULL DEFAULT 0, " &
                        "PromoStartDate TEXT NULL, " &
                        "PromoAprEndDate TEXT NULL, " &
                        "CreditLimit REAL NOT NULL DEFAULT 0, " &
                        "EscrowIncluded INTEGER NOT NULL DEFAULT 0, " &
                        "EscrowAmount REAL NOT NULL DEFAULT 0, " &
                        "PmiMonthly REAL NOT NULL DEFAULT 0, " &
                        "DeferredUntil TEXT NULL, " &
                        "DeferredStatus INTEGER NOT NULL DEFAULT 0, " &
                        "Subsidized INTEGER NOT NULL DEFAULT 0, " &
                        "BalloonAmount REAL NOT NULL DEFAULT 0, " &
                        "BalloonDueDate TEXT NULL, " &
                        "InterestOnlyStartDate TEXT NULL, " &
                        "InterestOnlyEndDate TEXT NULL, " &
                        "ForgivenessDate TEXT NULL, " &
                        "StudentRepaymentPlan TEXT NULL, " &
                        "RateChangeSchedule TEXT NULL, " &
                        "CustomInterestRule TEXT NULL, " &
                        "CustomFeeRule TEXT NULL, " &
                        "DayCountBasis INTEGER NULL, " &
                        "PaymentsPerYear INTEGER NULL" &
                        ")"
                    createCmd.ExecuteNonQuery()
                End Using

                If existingColumns.Count > 0 Then
                    Using copyCmd = conn.CreateCommand()
                        copyCmd.CommandText =
                            "INSERT INTO debts_new (" &
                            "Id, Description, Category, DebtType, Lender, APR, StartingBalance, OriginalPrincipal, PaymentAmount, DayDue, FromAccount, Hidden, Active, LoginLink, Notes, Frequency, SameAs, StartDate, LastPaymentDate, TermMonths, MaturityDate, PromoApr, PromoStartDate, PromoAprEndDate, CreditLimit, EscrowIncluded, EscrowAmount, PmiMonthly, DeferredUntil, DeferredStatus, Subsidized, BalloonAmount, BalloonDueDate, InterestOnlyStartDate, InterestOnlyEndDate, ForgivenessDate, StudentRepaymentPlan, RateChangeSchedule, CustomInterestRule, CustomFeeRule, DayCountBasis, PaymentsPerYear" &
                            ") SELECT " &
                            "Id, Description, Category, DebtType, Lender, APR, StartingBalance, OriginalPrincipal, PaymentAmount, DayDue, FromAccount, Hidden, COALESCE(Active,1), LoginLink, Notes, Frequency, SameAs, StartDate, LastPaymentDate, TermMonths, MaturityDate, PromoApr, PromoStartDate, PromoAprEndDate, CreditLimit, EscrowIncluded, EscrowAmount, PmiMonthly, DeferredUntil, DeferredStatus, Subsidized, BalloonAmount, BalloonDueDate, InterestOnlyStartDate, InterestOnlyEndDate, ForgivenessDate, StudentRepaymentPlan, RateChangeSchedule, CustomInterestRule, CustomFeeRule, DayCountBasis, PaymentsPerYear " &
                            "FROM debts"
                        copyCmd.ExecuteNonQuery()
                    End Using
                End If

                Using dropCmd = conn.CreateCommand()
                    dropCmd.CommandText = "DROP TABLE IF EXISTS debts"
                    dropCmd.ExecuteNonQuery()
                End Using

                Using renameCmd = conn.CreateCommand()
                    renameCmd.CommandText = "ALTER TABLE debts_new RENAME TO debts"
                    renameCmd.ExecuteNonQuery()
                End Using
            End Using
        End Sub

        Private Function HasTable(conn As SqliteConnection, tableName As String) As Boolean
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name=@name"
                cmd.Parameters.AddWithValue("@name", tableName)
                Return Convert.ToInt32(cmd.ExecuteScalar()) > 0
            End Using
        End Function

    End Module

End Namespace
