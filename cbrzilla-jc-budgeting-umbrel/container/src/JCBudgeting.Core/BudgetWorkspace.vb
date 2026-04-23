Imports System.Globalization
Imports Microsoft.Data.Sqlite
Imports System.Linq

Namespace Global.JCBudgeting.Core

    Public Class BudgetWorkspaceSettings
        Public Property BudgetPeriod As String = "Monthly"
        Public Property BudgetStartDate As String = String.Empty
        Public Property BudgetYears As Integer = 20
        Public Property AppTheme As String = "Light"
        Public Property EnableAnimations As Boolean = True
        Public Property ThemeColor As String = String.Empty
        Public Property BudgetZoomPercent As Integer = 100
        Public Property TransactionsZoomPercent As Integer = 100
        Public Property BudgetDistributionZoomPercent As Integer = 100
        Public Property SavingsDistributionZoomPercent As Integer = 100
        Public Property ServerMode As String = "Off"
        Public Property ServerPort As Integer = 5099
        Public Property ExternalServerHost As String = String.Empty
        Public Property ExternalServerPort As Integer = 5099
        Public Property ExternalHostOfflineCachePath As String = String.Empty
        Public Property ExternalHostOfflineBaselineChangeToken As String = String.Empty
        Public Property PendingExternalHostOfflineSync As Boolean = False
        Public Property GuidedTourCompleted As Boolean = False
        Public Property AccountsDetailTourCompleted As Boolean = False
        Public Property SavingsDetailTourCompleted As Boolean = False
        Public Property DebtsDetailTourCompleted As Boolean = False
        Public Property ExpensesDetailTourCompleted As Boolean = False
        Public Property TransactionsDetailTourCompleted As Boolean = False
    End Class

    Public Class BudgetWorkspacePeriodSummary
        Public Property PeriodStart As DateTime
        Public Property IncomeTotal As Decimal
        Public Property DebtTotal As Decimal
        Public Property ExpenseTotal As Decimal
        Public Property SavingsTotal As Decimal
        Public ReadOnly Property NetCashFlow As Decimal
            Get
                Return IncomeTotal - DebtTotal - ExpenseTotal - SavingsTotal
            End Get
        End Property
    End Class

    Public Class BudgetWorkspaceSnapshot
        Public Property BudgetPeriod As String = "Monthly"
        Public Property BudgetStart As DateTime
        Public Property BudgetEndExclusive As DateTime
        Public Property BudgetYears As Integer
        Public Property TotalPeriodCount As Integer
        Public Property CurrentPeriodIndex As Integer
        Public Property PeriodSummaries As New List(Of BudgetWorkspacePeriodSummary)()
        Public Property AccountRunningBalances As New Dictionary(Of String, Decimal())(StringComparer.OrdinalIgnoreCase)
        Public Property SavingsRunningBalances As New Dictionary(Of String, Decimal())(StringComparer.OrdinalIgnoreCase)
        Public Property DebtRunningBalances As New Dictionary(Of String, Decimal())(StringComparer.OrdinalIgnoreCase)
        Public Property DebtDisplayBalances As New Dictionary(Of String, Decimal())(StringComparer.OrdinalIgnoreCase)
        Public Property ItemizedBudgetRows As New List(Of BudgetWorkspaceItemSeries)()
    End Class

    Public Class BudgetWorkspaceItemSeries
        Public Property SectionName As String = String.Empty
        Public Property GroupName As String = String.Empty
        Public Property Label As String = String.Empty
        Public Property SourceLabel As String = String.Empty
        Public Property SourceKey As String = String.Empty
        Public Property Hidden As Boolean
        Public Property StatusText As String = String.Empty
        Public Property ScheduledValues As Decimal() = Array.Empty(Of Decimal)()
        Public Property SourceBaseValues As Decimal() = Array.Empty(Of Decimal)()
        Public Property SourceAdditionalValues As Decimal() = Array.Empty(Of Decimal)()
        Public Property ManualAdditionalValues As Decimal() = Array.Empty(Of Decimal)()
        Public Property PaidIndexes As Integer() = Array.Empty(Of Integer)()
        Public Property Values As Decimal() = Array.Empty(Of Decimal)()
        Public Property ManualIndexes As New List(Of Integer)()
        Public Property SourceIndexes As New List(Of Integer)()
    End Class

    Friend Class DebtProjectionResult
        Public Property PaymentValues As Decimal() = Array.Empty(Of Decimal)()
        Public Property BalanceValues As Decimal() = Array.Empty(Of Decimal)()
        Public Property DisplayBalanceValues As Decimal() = Array.Empty(Of Decimal)()
    End Class

    Friend Class BudgetRoutingSnapshot
        Public Property FromAccount As String = String.Empty
        Public Property ToAccount As String = String.Empty
        Public Property SameAs As String = String.Empty
        Public Property FromAccountId As Integer?
        Public Property ToAccountId As Integer?
        Public Property SameAsId As Integer?
        Public Property FromSavingsId As Integer?
        Public Property FromDebtId As Integer?
    End Class

    Public Module BudgetWorkspaceService
        Private Const DebtBalanceHistoryCategory As Integer = -101
        Private ReadOnly BudgetOverrideCleanupSignatures As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        Private ReadOnly BudgetOverrideItemLabelColumnPresence As New Dictionary(Of String, Boolean)(StringComparer.OrdinalIgnoreCase)
        Private Function ClampZoomPercent(value As Integer) As Integer
            Return Math.Max(75, Math.Min(200, value))
        End Function

        Public Sub SaveManualOverride(databasePath As String, dueDate As DateTime, categoryIndex As Integer, itemKey As String, itemDescription As String, amount As Decimal, Optional note As String = Nothing, Optional additional As Decimal = 0D, Optional paid As Boolean = False, Optional selectionMode As Integer = 3, Optional fromAccountSnapshot As String = Nothing, Optional toAccountSnapshot As String = Nothing, Optional sameAsSnapshot As String = Nothing, Optional fromAccountSnapshotId As Integer? = Nothing, Optional toAccountSnapshotId As Integer? = Nothing, Optional sameAsSnapshotId As Integer? = Nothing, Optional fromSavingsSnapshotId As Integer? = Nothing, Optional fromDebtSnapshotId As Integer? = Nothing)
            If String.IsNullOrWhiteSpace(databasePath) Then
                Throw New ArgumentException("A database path is required.", NameOf(databasePath))
            End If

            If String.IsNullOrWhiteSpace(itemKey) OrElse String.IsNullOrWhiteSpace(itemDescription) Then
                Throw New ArgumentException("A valid item key and item description are required.")
            End If

            Using conn As New SqliteConnection("Data Source=" & databasePath.Trim())
                conn.Open()
                EnsureBudgetDateOverridesTable(conn)
                Using cmd = conn.CreateCommand()
                    cmd.CommandText = "INSERT OR REPLACE INTO budgetDateOverrides (DueDate, CatIdx, ItemDesc, ItemLabel, Amount, Additional, ManualAmount, Paid, Notes, SelectionMode, FromAccountSnapshot, ToAccountSnapshot, SameAsSnapshot, FromAccountSnapshotId, ToAccountSnapshotId, SameAsSnapshotId, FromSavingsSnapshotId, FromDebtSnapshotId) VALUES (@d, @c, @n, @l, @a, @additional, @manualAmount, @paid, @notes, @selectionMode, @fromAccountSnapshot, @toAccountSnapshot, @sameAsSnapshot, @fromAccountSnapshotId, @toAccountSnapshotId, @sameAsSnapshotId, @fromSavingsSnapshotId, @fromDebtSnapshotId)"
                    cmd.Parameters.AddWithValue("@d", dueDate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture))
                    cmd.Parameters.AddWithValue("@c", categoryIndex)
                    cmd.Parameters.AddWithValue("@n", itemKey.Trim())
                    cmd.Parameters.AddWithValue("@l", itemDescription.Trim())
                    cmd.Parameters.AddWithValue("@a", amount)
                    cmd.Parameters.AddWithValue("@additional", additional)
                    cmd.Parameters.AddWithValue("@manualAmount", amount)
                    cmd.Parameters.AddWithValue("@paid", If(paid, 1, 0))
                    cmd.Parameters.AddWithValue("@notes", If(note, String.Empty).Trim())
                    cmd.Parameters.AddWithValue("@selectionMode", If(selectionMode = 0, 3, selectionMode))
                    cmd.Parameters.AddWithValue("@fromAccountSnapshot", If(fromAccountSnapshot, String.Empty).Trim())
                    cmd.Parameters.AddWithValue("@toAccountSnapshot", If(toAccountSnapshot, String.Empty).Trim())
                    cmd.Parameters.AddWithValue("@sameAsSnapshot", If(sameAsSnapshot, String.Empty).Trim())
                    cmd.Parameters.AddWithValue("@fromAccountSnapshotId", If(fromAccountSnapshotId.HasValue, CType(fromAccountSnapshotId.Value, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@toAccountSnapshotId", If(toAccountSnapshotId.HasValue, CType(toAccountSnapshotId.Value, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@sameAsSnapshotId", If(sameAsSnapshotId.HasValue, CType(sameAsSnapshotId.Value, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@fromSavingsSnapshotId", If(fromSavingsSnapshotId.HasValue, CType(fromSavingsSnapshotId.Value, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@fromDebtSnapshotId", If(fromDebtSnapshotId.HasValue, CType(fromDebtSnapshotId.Value, Object), DBNull.Value))
                    cmd.ExecuteNonQuery()
                End Using
            End Using
        End Sub

        Public Sub DeleteManualOverride(databasePath As String, dueDate As DateTime, categoryIndex As Integer, itemKey As String, itemDescription As String)
            If String.IsNullOrWhiteSpace(databasePath) Then
                Throw New ArgumentException("A database path is required.", NameOf(databasePath))
            End If

            If String.IsNullOrWhiteSpace(itemKey) OrElse String.IsNullOrWhiteSpace(itemDescription) Then
                Throw New ArgumentException("A valid item key and item description are required.")
            End If

            Using conn As New SqliteConnection("Data Source=" & databasePath.Trim())
                conn.Open()
                EnsureBudgetDateOverridesTable(conn)
                Using cmd = conn.CreateCommand()
                    cmd.CommandText = "DELETE FROM budgetDateOverrides WHERE DueDate=@d AND CatIdx=@c AND (ItemDesc=@n OR ItemDesc=@legacy)"
                    cmd.Parameters.AddWithValue("@d", dueDate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture))
                    cmd.Parameters.AddWithValue("@c", categoryIndex)
                    cmd.Parameters.AddWithValue("@n", itemKey.Trim())
                    cmd.Parameters.AddWithValue("@legacy", itemDescription.Trim())
                    cmd.ExecuteNonQuery()
                End Using
            End Using
        End Sub

        Public Sub DeleteSourceOverride(databasePath As String, dueDate As DateTime, categoryIndex As Integer, itemKey As String, itemDescription As String)
            If String.IsNullOrWhiteSpace(databasePath) Then
                Throw New ArgumentException("A database path is required.", NameOf(databasePath))
            End If

            If String.IsNullOrWhiteSpace(itemKey) OrElse String.IsNullOrWhiteSpace(itemDescription) Then
                Throw New ArgumentException("A valid item key and item description are required.")
            End If

            Using conn As New SqliteConnection("Data Source=" & databasePath.Trim())
                conn.Open()
                EnsureBudgetDateOverridesTable(conn)
                Using cmd = conn.CreateCommand()
                    cmd.CommandText = "DELETE FROM budgetDateOverrides WHERE DueDate=@d AND CatIdx=@c AND (ItemDesc=@n OR ItemDesc=@legacy)"
                    cmd.Parameters.AddWithValue("@d", dueDate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture))
                    cmd.Parameters.AddWithValue("@c", categoryIndex)
                    cmd.Parameters.AddWithValue("@n", itemKey.Trim())
                    cmd.Parameters.AddWithValue("@legacy", itemDescription.Trim())
                    cmd.ExecuteNonQuery()
                End Using
            End Using
        End Sub

        Public Structure BudgetCellEditorOverrideState
            Public Property HasRow As Boolean
            Public Property Amount As Decimal?
            Public Property Additional As Decimal
            Public Property ManualAmount As Decimal
            Public Property Paid As Boolean
            Public Property SelectionMode As Integer
            Public Property Notes As String
            Public Property ManualOverrideAmount As Decimal?
            Public Property ManualOverrideAdditional As Decimal
            Public Property ManualOverridePaid As Boolean
            Public Property ManualOverrideSelectionMode As Integer
            Public Property ManualOverrideNote As String
        End Structure

        Public Function GetBudgetCellEditorOverrideState(databasePath As String, dueDate As DateTime, categoryIndex As Integer, itemKey As String, itemDescription As String) As BudgetCellEditorOverrideState
            If String.IsNullOrWhiteSpace(databasePath) Then
                Throw New ArgumentException("A database path is required.", NameOf(databasePath))
            End If

            If String.IsNullOrWhiteSpace(itemKey) OrElse String.IsNullOrWhiteSpace(itemDescription) Then
                Throw New ArgumentException("A valid item key and item description are required.")
            End If

            Dim result As New BudgetCellEditorOverrideState With {
                .HasRow = False,
                .Amount = Nothing,
                .Additional = 0D,
                .ManualAmount = 0D,
                .Paid = False,
                .SelectionMode = 0,
                .Notes = String.Empty,
                .ManualOverrideAmount = Nothing,
                .ManualOverrideAdditional = 0D,
                .ManualOverridePaid = False,
                .ManualOverrideSelectionMode = 0,
                .ManualOverrideNote = String.Empty
            }

            Using conn As New SqliteConnection("Data Source=" & databasePath.Trim())
                conn.Open()
                EnsureBudgetDateOverridesTable(conn)
                Using cmd = conn.CreateCommand()
                    cmd.CommandText = "SELECT Amount, COALESCE(Additional,0), COALESCE(ManualAmount,0), COALESCE(Paid,0), COALESCE(SelectionMode,0), COALESCE(Notes,'') FROM budgetDateOverrides WHERE DueDate=@d AND CatIdx=@c AND (ItemDesc=@n OR ItemDesc=@legacy) ORDER BY CASE WHEN ItemDesc=@n THEN 0 ELSE 1 END LIMIT 1"
                    cmd.Parameters.AddWithValue("@d", dueDate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture))
                    cmd.Parameters.AddWithValue("@c", categoryIndex)
                    cmd.Parameters.AddWithValue("@n", itemKey.Trim())
                    cmd.Parameters.AddWithValue("@legacy", itemDescription.Trim())
                    Using reader = cmd.ExecuteReader()
                        If Not reader.Read() Then
                            Return result
                        End If

                        result.HasRow = True
                        If Not reader.IsDBNull(0) Then
                            result.Amount = Convert.ToDecimal(reader.GetValue(0), CultureInfo.InvariantCulture)
                        End If

                        result.Additional = If(reader.IsDBNull(1), 0D, Convert.ToDecimal(reader.GetValue(1), CultureInfo.InvariantCulture))
                        result.ManualAmount = If(reader.IsDBNull(2), 0D, Convert.ToDecimal(reader.GetValue(2), CultureInfo.InvariantCulture))
                        result.Paid = Not reader.IsDBNull(3) AndAlso Convert.ToInt32(reader.GetValue(3), CultureInfo.InvariantCulture) <> 0
                        result.SelectionMode = If(reader.IsDBNull(4), 0, Convert.ToInt32(reader.GetValue(4), CultureInfo.InvariantCulture))
                        result.Notes = If(reader.IsDBNull(5), String.Empty, Convert.ToString(reader.GetValue(5))?.Trim())
                    End Using
                End Using
            End Using

            If result.SelectionMode = 3 Then
                result.ManualOverrideAmount = If(result.Amount.HasValue, result.ManualAmount, CType(Nothing, Decimal?))
                result.ManualOverrideAdditional = result.Additional
                result.ManualOverridePaid = result.Paid
                result.ManualOverrideSelectionMode = result.SelectionMode
                result.ManualOverrideNote = result.Notes
            End If

            Return result
        End Function

        Public Function GetManualOverrideAmount(databasePath As String, dueDate As DateTime, categoryIndex As Integer, itemKey As String, itemDescription As String) As Decimal?
            If String.IsNullOrWhiteSpace(databasePath) Then
                Throw New ArgumentException("A database path is required.", NameOf(databasePath))
            End If

            If String.IsNullOrWhiteSpace(itemKey) OrElse String.IsNullOrWhiteSpace(itemDescription) Then
                Throw New ArgumentException("A valid item key and item description are required.")
            End If

            Using conn As New SqliteConnection("Data Source=" & databasePath.Trim())
                conn.Open()
                EnsureBudgetDateOverridesTable(conn)
                Using cmd = conn.CreateCommand()
                    cmd.CommandText = "SELECT CASE WHEN COALESCE(SelectionMode,0)=3 THEN COALESCE(ManualAmount, Amount) ELSE NULL END FROM budgetDateOverrides WHERE DueDate=@d AND CatIdx=@c AND (ItemDesc=@n OR ItemDesc=@legacy) ORDER BY CASE WHEN ItemDesc=@n THEN 0 ELSE 1 END LIMIT 1"
                    cmd.Parameters.AddWithValue("@d", dueDate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture))
                    cmd.Parameters.AddWithValue("@c", categoryIndex)
                    cmd.Parameters.AddWithValue("@n", itemKey.Trim())
                    cmd.Parameters.AddWithValue("@legacy", itemDescription.Trim())
                    Dim result = cmd.ExecuteScalar()
                    If result Is Nothing OrElse result Is DBNull.Value Then
                        Return Nothing
                    End If

                    Return Convert.ToDecimal(result, CultureInfo.InvariantCulture)
                End Using
            End Using
        End Function

        Public Function GetManualOverrideNote(databasePath As String, dueDate As DateTime, categoryIndex As Integer, itemKey As String, itemDescription As String) As String
            If String.IsNullOrWhiteSpace(databasePath) Then
                Throw New ArgumentException("A database path is required.", NameOf(databasePath))
            End If

            If String.IsNullOrWhiteSpace(itemKey) OrElse String.IsNullOrWhiteSpace(itemDescription) Then
                Throw New ArgumentException("A valid item key and item description are required.")
            End If

            Using conn As New SqliteConnection("Data Source=" & databasePath.Trim())
                conn.Open()
                EnsureBudgetDateOverridesTable(conn)
                Using cmd = conn.CreateCommand()
                    cmd.CommandText = "SELECT CASE WHEN COALESCE(SelectionMode,0)=3 THEN COALESCE(Notes, '') ELSE '' END FROM budgetDateOverrides WHERE DueDate=@d AND CatIdx=@c AND (ItemDesc=@n OR ItemDesc=@legacy) ORDER BY CASE WHEN ItemDesc=@n THEN 0 ELSE 1 END LIMIT 1"
                    cmd.Parameters.AddWithValue("@d", dueDate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture))
                    cmd.Parameters.AddWithValue("@c", categoryIndex)
                    cmd.Parameters.AddWithValue("@n", itemKey.Trim())
                    cmd.Parameters.AddWithValue("@legacy", itemDescription.Trim())
                    Dim result = cmd.ExecuteScalar()
                    If result Is Nothing OrElse result Is DBNull.Value Then
                        Return String.Empty
                    End If

                    Return Convert.ToString(result)?.Trim()
                End Using
            End Using
        End Function

        Public Function GetManualOverrideAdditional(databasePath As String, dueDate As DateTime, categoryIndex As Integer, itemKey As String, itemDescription As String) As Decimal
            If String.IsNullOrWhiteSpace(databasePath) Then
                Throw New ArgumentException("A database path is required.", NameOf(databasePath))
            End If

            If String.IsNullOrWhiteSpace(itemKey) OrElse String.IsNullOrWhiteSpace(itemDescription) Then
                Throw New ArgumentException("A valid item key and item description are required.")
            End If

            Using conn As New SqliteConnection("Data Source=" & databasePath.Trim())
                conn.Open()
                EnsureBudgetDateOverridesTable(conn)
                Using cmd = conn.CreateCommand()
                    cmd.CommandText = "SELECT CASE WHEN COALESCE(SelectionMode,0)=3 THEN COALESCE(Additional,0) ELSE 0 END FROM budgetDateOverrides WHERE DueDate=@d AND CatIdx=@c AND (ItemDesc=@n OR ItemDesc=@legacy) ORDER BY CASE WHEN ItemDesc=@n THEN 0 ELSE 1 END LIMIT 1"
                    cmd.Parameters.AddWithValue("@d", dueDate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture))
                    cmd.Parameters.AddWithValue("@c", categoryIndex)
                    cmd.Parameters.AddWithValue("@n", itemKey.Trim())
                    cmd.Parameters.AddWithValue("@legacy", itemDescription.Trim())
                    Dim result = cmd.ExecuteScalar()
                    If result Is Nothing OrElse result Is DBNull.Value Then
                        Return 0D
                    End If

                    Return Convert.ToDecimal(result, CultureInfo.InvariantCulture)
                End Using
            End Using
        End Function

        Public Function GetManualOverridePaid(databasePath As String, dueDate As DateTime, categoryIndex As Integer, itemKey As String, itemDescription As String) As Boolean
            If String.IsNullOrWhiteSpace(databasePath) Then
                Throw New ArgumentException("A database path is required.", NameOf(databasePath))
            End If

            If String.IsNullOrWhiteSpace(itemKey) OrElse String.IsNullOrWhiteSpace(itemDescription) Then
                Throw New ArgumentException("A valid item key and item description are required.")
            End If

            Using conn As New SqliteConnection("Data Source=" & databasePath.Trim())
                conn.Open()
                EnsureBudgetDateOverridesTable(conn)
                Using cmd = conn.CreateCommand()
                    cmd.CommandText = "SELECT CASE WHEN COALESCE(SelectionMode,0)=3 THEN COALESCE(Paid,0) ELSE 0 END FROM budgetDateOverrides WHERE DueDate=@d AND CatIdx=@c AND (ItemDesc=@n OR ItemDesc=@legacy) ORDER BY CASE WHEN ItemDesc=@n THEN 0 ELSE 1 END LIMIT 1"
                    cmd.Parameters.AddWithValue("@d", dueDate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture))
                    cmd.Parameters.AddWithValue("@c", categoryIndex)
                    cmd.Parameters.AddWithValue("@n", itemKey.Trim())
                    cmd.Parameters.AddWithValue("@legacy", itemDescription.Trim())
                    Dim result = cmd.ExecuteScalar()
                    Return result IsNot Nothing AndAlso result IsNot DBNull.Value AndAlso Convert.ToInt32(result, CultureInfo.InvariantCulture) <> 0
                End Using
            End Using
        End Function

        Public Function GetManualOverrideSelectionMode(databasePath As String, dueDate As DateTime, categoryIndex As Integer, itemKey As String, itemDescription As String) As Integer
            If String.IsNullOrWhiteSpace(databasePath) Then
                Throw New ArgumentException("A database path is required.", NameOf(databasePath))
            End If

            If String.IsNullOrWhiteSpace(itemKey) OrElse String.IsNullOrWhiteSpace(itemDescription) Then
                Throw New ArgumentException("A valid item key and item description are required.")
            End If

            Using conn As New SqliteConnection("Data Source=" & databasePath.Trim())
                conn.Open()
                EnsureBudgetDateOverridesTable(conn)
                Using cmd = conn.CreateCommand()
                    cmd.CommandText = "SELECT CASE WHEN COALESCE(SelectionMode,0)=3 THEN COALESCE(SelectionMode,0) ELSE 0 END FROM budgetDateOverrides WHERE DueDate=@d AND CatIdx=@c AND (ItemDesc=@n OR ItemDesc=@legacy) ORDER BY CASE WHEN ItemDesc=@n THEN 0 ELSE 1 END LIMIT 1"
                    cmd.Parameters.AddWithValue("@d", dueDate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture))
                    cmd.Parameters.AddWithValue("@c", categoryIndex)
                    cmd.Parameters.AddWithValue("@n", itemKey.Trim())
                    cmd.Parameters.AddWithValue("@legacy", itemDescription.Trim())
                    Dim result = cmd.ExecuteScalar()
                    If result Is Nothing OrElse result Is DBNull.Value Then
                        Return 0
                    End If

                    Return Convert.ToInt32(result, CultureInfo.InvariantCulture)
                End Using
            End Using
        End Function

        Public Function GetSourceOverrideAmount(databasePath As String, dueDate As DateTime, categoryIndex As Integer, itemKey As String, itemDescription As String) As Decimal?
            If String.IsNullOrWhiteSpace(databasePath) Then
                Throw New ArgumentException("A database path is required.", NameOf(databasePath))
            End If

            If String.IsNullOrWhiteSpace(itemKey) OrElse String.IsNullOrWhiteSpace(itemDescription) Then
                Throw New ArgumentException("A valid item key and item description are required.")
            End If

            Using conn As New SqliteConnection("Data Source=" & databasePath.Trim())
                conn.Open()
                EnsureBudgetDateOverridesTable(conn)
                Using cmd = conn.CreateCommand()
                    cmd.CommandText = "SELECT Amount FROM budgetDateOverrides WHERE DueDate=@d AND CatIdx=@c AND (ItemDesc=@n OR ItemDesc=@legacy) ORDER BY CASE WHEN ItemDesc=@n THEN 0 ELSE 1 END LIMIT 1"
                    cmd.Parameters.AddWithValue("@d", dueDate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture))
                    cmd.Parameters.AddWithValue("@c", categoryIndex)
                    cmd.Parameters.AddWithValue("@n", itemKey.Trim())
                    cmd.Parameters.AddWithValue("@legacy", itemDescription.Trim())
                    Dim result = cmd.ExecuteScalar()
                    If result Is Nothing OrElse result Is DBNull.Value Then
                        Return Nothing
                    End If

                    Return Convert.ToDecimal(result, CultureInfo.InvariantCulture)
                End Using
            End Using
        End Function

        Public Function GetSourceOverrideAdditional(databasePath As String, dueDate As DateTime, categoryIndex As Integer, itemKey As String, itemDescription As String) As Decimal
            If String.IsNullOrWhiteSpace(databasePath) Then
                Throw New ArgumentException("A database path is required.", NameOf(databasePath))
            End If

            If String.IsNullOrWhiteSpace(itemKey) OrElse String.IsNullOrWhiteSpace(itemDescription) Then
                Throw New ArgumentException("A valid item key and item description are required.")
            End If

            Using conn As New SqliteConnection("Data Source=" & databasePath.Trim())
                conn.Open()
                EnsureBudgetDateOverridesTable(conn)
                Using cmd = conn.CreateCommand()
                    cmd.CommandText = "SELECT COALESCE(Additional,0) FROM budgetDateOverrides WHERE DueDate=@d AND CatIdx=@c AND (ItemDesc=@n OR ItemDesc=@legacy) ORDER BY CASE WHEN ItemDesc=@n THEN 0 ELSE 1 END LIMIT 1"
                    cmd.Parameters.AddWithValue("@d", dueDate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture))
                    cmd.Parameters.AddWithValue("@c", categoryIndex)
                    cmd.Parameters.AddWithValue("@n", itemKey.Trim())
                    cmd.Parameters.AddWithValue("@legacy", itemDescription.Trim())
                    Dim result = cmd.ExecuteScalar()
                    If result Is Nothing OrElse result Is DBNull.Value Then
                        Return 0D
                    End If

                    Return Convert.ToDecimal(result, CultureInfo.InvariantCulture)
                End Using
            End Using
        End Function

        Public Function GetSourceOverrideManualAmount(databasePath As String, dueDate As DateTime, categoryIndex As Integer, itemKey As String, itemDescription As String) As Decimal
            If String.IsNullOrWhiteSpace(databasePath) Then
                Throw New ArgumentException("A database path is required.", NameOf(databasePath))
            End If

            If String.IsNullOrWhiteSpace(itemKey) OrElse String.IsNullOrWhiteSpace(itemDescription) Then
                Throw New ArgumentException("A valid item key and item description are required.")
            End If

            Using conn As New SqliteConnection("Data Source=" & databasePath.Trim())
                conn.Open()
                EnsureBudgetDateOverridesTable(conn)
                Using cmd = conn.CreateCommand()
                    cmd.CommandText = "SELECT COALESCE(ManualAmount,0) FROM budgetDateOverrides WHERE DueDate=@d AND CatIdx=@c AND (ItemDesc=@n OR ItemDesc=@legacy) ORDER BY CASE WHEN ItemDesc=@n THEN 0 ELSE 1 END LIMIT 1"
                    cmd.Parameters.AddWithValue("@d", dueDate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture))
                    cmd.Parameters.AddWithValue("@c", categoryIndex)
                    cmd.Parameters.AddWithValue("@n", itemKey.Trim())
                    cmd.Parameters.AddWithValue("@legacy", itemDescription.Trim())
                    Dim result = cmd.ExecuteScalar()
                    If result Is Nothing OrElse result Is DBNull.Value Then
                        Return 0D
                    End If

                    Return Convert.ToDecimal(result, CultureInfo.InvariantCulture)
                End Using
            End Using
        End Function

        Public Function GetSourceOverridePaid(databasePath As String, dueDate As DateTime, categoryIndex As Integer, itemKey As String, itemDescription As String) As Boolean
            If String.IsNullOrWhiteSpace(databasePath) Then
                Throw New ArgumentException("A database path is required.", NameOf(databasePath))
            End If

            If String.IsNullOrWhiteSpace(itemKey) OrElse String.IsNullOrWhiteSpace(itemDescription) Then
                Throw New ArgumentException("A valid item key and item description are required.")
            End If

            Using conn As New SqliteConnection("Data Source=" & databasePath.Trim())
                conn.Open()
                EnsureBudgetDateOverridesTable(conn)
                Using cmd = conn.CreateCommand()
                    cmd.CommandText = "SELECT COALESCE(Paid,0) FROM budgetDateOverrides WHERE DueDate=@d AND CatIdx=@c AND (ItemDesc=@n OR ItemDesc=@legacy) ORDER BY CASE WHEN ItemDesc=@n THEN 0 ELSE 1 END LIMIT 1"
                    cmd.Parameters.AddWithValue("@d", dueDate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture))
                    cmd.Parameters.AddWithValue("@c", categoryIndex)
                    cmd.Parameters.AddWithValue("@n", itemKey.Trim())
                    cmd.Parameters.AddWithValue("@legacy", itemDescription.Trim())
                    Dim result = cmd.ExecuteScalar()
                    Return result IsNot Nothing AndAlso result IsNot DBNull.Value AndAlso Convert.ToInt32(result, CultureInfo.InvariantCulture) <> 0
                End Using
            End Using
        End Function

        Public Function GetSourceOverrideSelectionMode(databasePath As String, dueDate As DateTime, categoryIndex As Integer, itemKey As String, itemDescription As String) As Integer
            If String.IsNullOrWhiteSpace(databasePath) Then
                Throw New ArgumentException("A database path is required.", NameOf(databasePath))
            End If

            If String.IsNullOrWhiteSpace(itemKey) OrElse String.IsNullOrWhiteSpace(itemDescription) Then
                Throw New ArgumentException("A valid item key and item description are required.")
            End If

            Using conn As New SqliteConnection("Data Source=" & databasePath.Trim())
                conn.Open()
                EnsureBudgetDateOverridesTable(conn)
                Using cmd = conn.CreateCommand()
                    cmd.CommandText = "SELECT COALESCE(SelectionMode,0) FROM budgetDateOverrides WHERE DueDate=@d AND CatIdx=@c AND (ItemDesc=@n OR ItemDesc=@legacy) ORDER BY CASE WHEN ItemDesc=@n THEN 0 ELSE 1 END LIMIT 1"
                    cmd.Parameters.AddWithValue("@d", dueDate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture))
                    cmd.Parameters.AddWithValue("@c", categoryIndex)
                    cmd.Parameters.AddWithValue("@n", itemKey.Trim())
                    cmd.Parameters.AddWithValue("@legacy", itemDescription.Trim())
                    Dim result = cmd.ExecuteScalar()
                    If result Is Nothing OrElse result Is DBNull.Value Then
                        Return 0
                    End If

                    Return Convert.ToInt32(result, CultureInfo.InvariantCulture)
                End Using
            End Using
        End Function

        Public Function GetSourceOverrideNote(databasePath As String, dueDate As DateTime, categoryIndex As Integer, itemKey As String, itemDescription As String) As String
            If String.IsNullOrWhiteSpace(databasePath) Then
                Throw New ArgumentException("A database path is required.", NameOf(databasePath))
            End If

            If String.IsNullOrWhiteSpace(itemKey) OrElse String.IsNullOrWhiteSpace(itemDescription) Then
                Throw New ArgumentException("A valid item key and item description are required.")
            End If

            Using conn As New SqliteConnection("Data Source=" & databasePath.Trim())
                conn.Open()
                EnsureBudgetDateOverridesTable(conn)
                Using cmd = conn.CreateCommand()
                    cmd.CommandText = "SELECT COALESCE(Notes,'') FROM budgetDateOverrides WHERE DueDate=@d AND CatIdx=@c AND (ItemDesc=@n OR ItemDesc=@legacy) ORDER BY CASE WHEN ItemDesc=@n THEN 0 ELSE 1 END LIMIT 1"
                    cmd.Parameters.AddWithValue("@d", dueDate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture))
                    cmd.Parameters.AddWithValue("@c", categoryIndex)
                    cmd.Parameters.AddWithValue("@n", itemKey.Trim())
                    cmd.Parameters.AddWithValue("@legacy", itemDescription.Trim())
                    Dim result = cmd.ExecuteScalar()
                    If result Is Nothing OrElse result Is DBNull.Value Then
                        Return String.Empty
                    End If

                    Return Convert.ToString(result)?.Trim()
                End Using
            End Using
        End Function

        Public Sub SaveSourceOverride(databasePath As String, dueDate As DateTime, categoryIndex As Integer, itemKey As String, itemDescription As String, amount As Decimal, additional As Decimal, Optional note As String = Nothing, Optional paid As Boolean = False, Optional selectionMode As Integer = 0, Optional manualAmount As Decimal = 0D, Optional fromAccountSnapshot As String = Nothing, Optional toAccountSnapshot As String = Nothing, Optional sameAsSnapshot As String = Nothing, Optional fromAccountSnapshotId As Integer? = Nothing, Optional toAccountSnapshotId As Integer? = Nothing, Optional sameAsSnapshotId As Integer? = Nothing)
            If String.IsNullOrWhiteSpace(databasePath) Then
                Throw New ArgumentException("A database path is required.", NameOf(databasePath))
            End If

            If String.IsNullOrWhiteSpace(itemKey) OrElse String.IsNullOrWhiteSpace(itemDescription) Then
                Throw New ArgumentException("A valid item key and item description are required.")
            End If

            Using conn As New SqliteConnection("Data Source=" & databasePath.Trim())
                conn.Open()
                EnsureBudgetDateOverridesTable(conn)
                Using cmd = conn.CreateCommand()
                    cmd.CommandText = "INSERT OR REPLACE INTO budgetDateOverrides (DueDate, CatIdx, ItemDesc, ItemLabel, Amount, Additional, ManualAmount, Paid, Notes, SelectionMode, FromAccountSnapshot, ToAccountSnapshot, SameAsSnapshot, FromAccountSnapshotId, ToAccountSnapshotId, SameAsSnapshotId) VALUES (@d, @c, @n, @l, @a, @additional, @manualAmount, @paid, @notes, @selectionMode, @fromAccountSnapshot, @toAccountSnapshot, @sameAsSnapshot, @fromAccountSnapshotId, @toAccountSnapshotId, @sameAsSnapshotId)"
                    cmd.Parameters.AddWithValue("@d", dueDate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture))
                    cmd.Parameters.AddWithValue("@c", categoryIndex)
                    cmd.Parameters.AddWithValue("@n", itemKey.Trim())
                    cmd.Parameters.AddWithValue("@l", itemDescription.Trim())
                    cmd.Parameters.AddWithValue("@a", amount)
                    cmd.Parameters.AddWithValue("@additional", additional)
                    cmd.Parameters.AddWithValue("@manualAmount", manualAmount)
                    cmd.Parameters.AddWithValue("@paid", If(paid, 1, 0))
                    cmd.Parameters.AddWithValue("@notes", If(note, String.Empty).Trim())
                    cmd.Parameters.AddWithValue("@selectionMode", selectionMode)
                    cmd.Parameters.AddWithValue("@fromAccountSnapshot", If(fromAccountSnapshot, String.Empty).Trim())
                    cmd.Parameters.AddWithValue("@toAccountSnapshot", If(toAccountSnapshot, String.Empty).Trim())
                    cmd.Parameters.AddWithValue("@sameAsSnapshot", If(sameAsSnapshot, String.Empty).Trim())
                    cmd.Parameters.AddWithValue("@fromAccountSnapshotId", If(fromAccountSnapshotId.HasValue, CType(fromAccountSnapshotId.Value, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@toAccountSnapshotId", If(toAccountSnapshotId.HasValue, CType(toAccountSnapshotId.Value, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@sameAsSnapshotId", If(sameAsSnapshotId.HasValue, CType(sameAsSnapshotId.Value, Object), DBNull.Value))
                    cmd.ExecuteNonQuery()
                End Using
            End Using
        End Sub

        Public Sub UpdateSourceOverrideMetadata(databasePath As String, dueDate As DateTime, categoryIndex As Integer, itemKey As String, itemDescription As String, additional As Decimal, Optional note As String = Nothing, Optional paid As Boolean = False)
            If String.IsNullOrWhiteSpace(databasePath) Then
                Throw New ArgumentException("A database path is required.", NameOf(databasePath))
            End If

            If String.IsNullOrWhiteSpace(itemKey) OrElse String.IsNullOrWhiteSpace(itemDescription) Then
                Throw New ArgumentException("A valid item key and item description are required.")
            End If

            Using conn As New SqliteConnection("Data Source=" & databasePath.Trim())
                conn.Open()
                EnsureBudgetDateOverridesTable(conn)
                Using cmd = conn.CreateCommand()
                    cmd.CommandText = "UPDATE budgetDateOverrides SET Additional=@additional, Paid=@paid, Notes=@notes WHERE DueDate=@d AND CatIdx=@c AND (ItemDesc=@n OR ItemDesc=@legacy)"
                    cmd.Parameters.AddWithValue("@additional", additional)
                    cmd.Parameters.AddWithValue("@paid", If(paid, 1, 0))
                    cmd.Parameters.AddWithValue("@notes", If(note, String.Empty).Trim())
                    cmd.Parameters.AddWithValue("@d", dueDate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture))
                    cmd.Parameters.AddWithValue("@c", categoryIndex)
                    cmd.Parameters.AddWithValue("@n", itemKey.Trim())
                    cmd.Parameters.AddWithValue("@legacy", itemDescription.Trim())
                    cmd.ExecuteNonQuery()
                End Using
            End Using
        End Sub

        Public Sub SyncTransactionSourceOverrides(databasePath As String, settings As BudgetWorkspaceSettings)
            If String.IsNullOrWhiteSpace(databasePath) OrElse Not IO.File.Exists(databasePath) Then
                Return
            End If

            If settings Is Nothing Then
                settings = New BudgetWorkspaceSettings()
            End If

            Dim startDate As DateTime
            If String.IsNullOrWhiteSpace(settings.BudgetStartDate) OrElse
               Not DateTime.TryParse(settings.BudgetStartDate, CultureInfo.CurrentCulture, DateTimeStyles.None, startDate) Then
                startDate = DateTime.Today
            End If

            Dim budgetYears = Math.Max(1, settings.BudgetYears)
            Dim periodName = If(String.IsNullOrWhiteSpace(settings.BudgetPeriod), "Monthly", settings.BudgetPeriod)
            Dim endExclusive = startDate.AddYears(budgetYears)
            Dim periods = GeneratePeriods(startDate, endExclusive, periodName)

            Dim incomes = IncomeRepository.LoadIncome(databasePath)
            Dim expenses = ExpenseRepository.LoadExpenses(databasePath)
            Dim savings = SavingsRepository.LoadSavings(databasePath)
            Dim debts = DebtRepository.LoadDebts(databasePath)
            TransactionRepository.CleanupOrphanedAssignments(databasePath, incomes, debts, expenses, savings)

            Dim itemLookup As New Dictionary(Of String, (Key As String, Label As String))(StringComparer.OrdinalIgnoreCase)
            Dim preservedSourceMetadata As New Dictionary(Of String, (Additional As Decimal, ManualAmount As Decimal, Paid As Boolean, Note As String, SelectionMode As Integer, FromAccountSnapshot As String, ToAccountSnapshot As String, SameAsSnapshot As String))(StringComparer.OrdinalIgnoreCase)
            For Each item In incomes
                Dim label = If(String.IsNullOrWhiteSpace(item.Description), $"Income {item.Id}", item.Description.Trim())
                itemLookup($"0|{item.Id}") = (BuildOverrideStorageKey(0, item.Id, label), label)
            Next

            For Each item In debts
                Dim label = If(String.IsNullOrWhiteSpace(item.Description), $"Debt {item.Id}", item.Description.Trim())
                itemLookup($"1|{item.Id}") = (BuildOverrideStorageKey(1, item.Id, label), label)
            Next

            For Each item In expenses
                Dim label = If(String.IsNullOrWhiteSpace(item.Description), $"Expense {item.Id}", item.Description.Trim())
                itemLookup($"2|{item.Id}") = (BuildOverrideStorageKey(2, item.Id, label), label)
            Next

            For Each item In savings
                Dim label = If(String.IsNullOrWhiteSpace(item.Description), $"Savings {item.Id}", item.Description.Trim())
                itemLookup($"3|{item.Id}") = (BuildOverrideStorageKey(3, item.Id, label), label)
            Next

            Dim transactions = TransactionRepository.LoadTransactions(databasePath).ToDictionary(Function(x) x.Id)
            Dim assignments = TransactionRepository.LoadTransactionAssignments(databasePath)
            Dim aggregates As New Dictionary(Of String, Decimal)(StringComparer.OrdinalIgnoreCase)
            Dim assignedSourceKeys As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            For Each assignment In assignments
                If assignment Is Nothing OrElse assignment.TransactionId <= 0 Then
                    Continue For
                End If

                Dim transaction As TransactionRecord = Nothing
                If Not transactions.TryGetValue(assignment.TransactionId, transaction) OrElse transaction Is Nothing Then
                    Continue For
                End If

                Dim txDate = ParseOptionalDate(transaction.TransactionDate)
                If Not txDate.HasValue Then
                    Continue For
                End If

                Dim periodIndex = FindPeriodIndexForDate(periods, txDate.Value, periodName)
                If periodIndex < 0 OrElse periodIndex >= periods.Count Then
                    Continue For
                End If

                Dim lookupKey = $"{assignment.CatIdx}|{assignment.ItemId}"
                If Not itemLookup.ContainsKey(lookupKey) Then
                    Continue For
                End If

                Dim target = itemLookup(lookupKey)
                assignedSourceKeys.Add($"{assignment.CatIdx}|{target.Key}")
                Dim dueText = periods(periodIndex).ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
                Dim aggregateKey = $"{dueText}|{assignment.CatIdx}|{target.Key}|{target.Label}"
                If Not aggregates.ContainsKey(aggregateKey) Then
                    aggregates(aggregateKey) = 0D
                End If

                Dim signedAssignmentAmount = Math.Abs(assignment.Amount)
                If transaction.Amount < 0D Then
                    signedAssignmentAmount = -signedAssignmentAmount
                End If
                If assignment.CatIdx <> 0 Then
                    signedAssignmentAmount = -signedAssignmentAmount
                End If

                aggregates(aggregateKey) += signedAssignmentAmount
            Next

            Using conn As New SqliteConnection("Data Source=" & databasePath.Trim())
                conn.Open()
                EnsureBudgetDateOverridesTable(conn)
                Using selectCmd = conn.CreateCommand()
                    selectCmd.CommandText = "SELECT DueDate, CatIdx, ItemDesc, COALESCE(Additional,0), COALESCE(ManualAmount,0), COALESCE(Paid,0), COALESCE(Notes,''), COALESCE(SelectionMode,0), COALESCE(FromAccountSnapshot,''), COALESCE(ToAccountSnapshot,''), COALESCE(SameAsSnapshot,'') FROM budgetDateOverrides WHERE CatIdx>=0"
                    Using rdr = selectCmd.ExecuteReader()
                        While rdr.Read()
                            Dim dueText = If(rdr.IsDBNull(0), String.Empty, Convert.ToString(rdr.GetValue(0)))
                            Dim catIdx = If(rdr.IsDBNull(1), -1, Convert.ToInt32(rdr.GetValue(1), CultureInfo.InvariantCulture))
                            Dim itemDesc = If(rdr.IsDBNull(2), String.Empty, Convert.ToString(rdr.GetValue(2))).Trim()
                            If catIdx < 0 OrElse String.IsNullOrWhiteSpace(dueText) OrElse String.IsNullOrWhiteSpace(itemDesc) Then
                                Continue While
                            End If

                            Dim metadataKey = dueText & "|" & catIdx.ToString(CultureInfo.InvariantCulture) & "|" & itemDesc
                            preservedSourceMetadata(metadataKey) = (
                                If(rdr.IsDBNull(3), 0D, Convert.ToDecimal(rdr.GetValue(3), CultureInfo.InvariantCulture)),
                                If(rdr.IsDBNull(4), 0D, Convert.ToDecimal(rdr.GetValue(4), CultureInfo.InvariantCulture)),
                                Not rdr.IsDBNull(5) AndAlso Convert.ToInt32(rdr.GetValue(5), CultureInfo.InvariantCulture) <> 0,
                                If(rdr.IsDBNull(6), String.Empty, Convert.ToString(rdr.GetValue(6))).Trim(),
                                If(rdr.IsDBNull(7), 0, Convert.ToInt32(rdr.GetValue(7), CultureInfo.InvariantCulture)),
                                If(rdr.IsDBNull(8), String.Empty, Convert.ToString(rdr.GetValue(8))).Trim(),
                                If(rdr.IsDBNull(9), String.Empty, Convert.ToString(rdr.GetValue(9))).Trim(),
                                If(rdr.IsDBNull(10), String.Empty, Convert.ToString(rdr.GetValue(10))).Trim())
                        End While
                    End Using
                End Using

                Using deleteCmd = conn.CreateCommand()
                    deleteCmd.CommandText = "DELETE FROM budgetDateOverrides WHERE CatIdx>=0"
                    deleteCmd.ExecuteNonQuery()
                End Using

                Dim insertedKeys As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

                For Each entry In aggregates
                    Dim parts = entry.Key.Split("|"c)
                    If parts.Length < 4 Then
                        Continue For
                    End If

                    Using insertCmd = conn.CreateCommand()
                        Dim metadataKey = parts(0) & "|" & parts(1) & "|" & parts(2)
                        Dim metadata As (Additional As Decimal, ManualAmount As Decimal, Paid As Boolean, Note As String, SelectionMode As Integer, FromAccountSnapshot As String, ToAccountSnapshot As String, SameAsSnapshot As String) =
                            If(preservedSourceMetadata.ContainsKey(metadataKey),
                               preservedSourceMetadata(metadataKey),
                               (Additional:=0D, ManualAmount:=0D, Paid:=False, Note:=String.Empty, SelectionMode:=0, FromAccountSnapshot:=String.Empty, ToAccountSnapshot:=String.Empty, SameAsSnapshot:=String.Empty))
                        insertCmd.CommandText =
"INSERT OR REPLACE INTO budgetDateOverrides (DueDate, CatIdx, ItemDesc, ItemLabel, Amount, Additional, ManualAmount, Paid, Notes, SelectionMode, FromAccountSnapshot, ToAccountSnapshot, SameAsSnapshot) " &
"VALUES (@d, @c, @n, @l, @a, @additional, @manualAmount, @paid, @notes, @selectionMode, @fromAccountSnapshot, @toAccountSnapshot, @sameAsSnapshot)"
                        insertCmd.Parameters.AddWithValue("@d", parts(0))
                        insertCmd.Parameters.AddWithValue("@c", Convert.ToInt32(parts(1), CultureInfo.InvariantCulture))
                        insertCmd.Parameters.AddWithValue("@n", parts(2))
                        insertCmd.Parameters.AddWithValue("@l", parts(3))
                        insertCmd.Parameters.AddWithValue("@a", entry.Value)
                        insertCmd.Parameters.AddWithValue("@additional", metadata.Additional)
                        insertCmd.Parameters.AddWithValue("@manualAmount", metadata.ManualAmount)
                        insertCmd.Parameters.AddWithValue("@paid", If(metadata.Paid, 1, 0))
                        insertCmd.Parameters.AddWithValue("@notes", metadata.Note)
                        insertCmd.Parameters.AddWithValue("@selectionMode", metadata.SelectionMode)
                        insertCmd.Parameters.AddWithValue("@fromAccountSnapshot", metadata.FromAccountSnapshot)
                        insertCmd.Parameters.AddWithValue("@toAccountSnapshot", metadata.ToAccountSnapshot)
                        insertCmd.Parameters.AddWithValue("@sameAsSnapshot", metadata.SameAsSnapshot)
                        insertCmd.ExecuteNonQuery()
                    End Using

                    insertedKeys.Add(parts(0) & "|" & parts(1) & "|" & parts(2))
                Next

                For Each metadataEntry In preservedSourceMetadata
                    Dim metadata = metadataEntry.Value
                    If metadata.Additional = 0D AndAlso metadata.ManualAmount = 0D AndAlso Not metadata.Paid AndAlso String.IsNullOrWhiteSpace(metadata.Note) AndAlso metadata.SelectionMode = 0 AndAlso String.IsNullOrWhiteSpace(metadata.FromAccountSnapshot) AndAlso String.IsNullOrWhiteSpace(metadata.ToAccountSnapshot) AndAlso String.IsNullOrWhiteSpace(metadata.SameAsSnapshot) Then
                        Continue For
                    End If

                    If insertedKeys.Contains(metadataEntry.Key) Then
                        Continue For
                    End If

                    Dim parts = metadataEntry.Key.Split("|"c)
                    If parts.Length < 3 Then
                        Continue For
                    End If

                    Dim dueText = parts(0)
                    Dim catIdxText = parts(1)
                    Dim itemDesc = parts(2)
                    Dim itemLabel = itemDesc
                    Dim lookupKey As String = Nothing

                    For Each candidate In itemLookup
                        If String.Equals(candidate.Value.Key, itemDesc, StringComparison.OrdinalIgnoreCase) Then
                            lookupKey = candidate.Key
                            itemLabel = candidate.Value.Label
                            Exit For
                        End If
                    Next

                    Using insertCmd = conn.CreateCommand()
                        insertCmd.CommandText =
"INSERT OR REPLACE INTO budgetDateOverrides (DueDate, CatIdx, ItemDesc, ItemLabel, Amount, Additional, ManualAmount, Paid, Notes, SelectionMode, FromAccountSnapshot, ToAccountSnapshot, SameAsSnapshot) " &
"VALUES (@d, @c, @n, @l, 0, @additional, @manualAmount, @paid, @notes, @selectionMode, @fromAccountSnapshot, @toAccountSnapshot, @sameAsSnapshot)"
                        insertCmd.Parameters.AddWithValue("@d", dueText)
                        insertCmd.Parameters.AddWithValue("@c", Convert.ToInt32(catIdxText, CultureInfo.InvariantCulture))
                        insertCmd.Parameters.AddWithValue("@n", itemDesc)
                        insertCmd.Parameters.AddWithValue("@l", itemLabel)
                        insertCmd.Parameters.AddWithValue("@additional", metadata.Additional)
                        insertCmd.Parameters.AddWithValue("@manualAmount", metadata.ManualAmount)
                        insertCmd.Parameters.AddWithValue("@paid", If(metadata.Paid, 1, 0))
                        insertCmd.Parameters.AddWithValue("@notes", metadata.Note)
                        insertCmd.Parameters.AddWithValue("@selectionMode", metadata.SelectionMode)
                        insertCmd.Parameters.AddWithValue("@fromAccountSnapshot", metadata.FromAccountSnapshot)
                        insertCmd.Parameters.AddWithValue("@toAccountSnapshot", metadata.ToAccountSnapshot)
                        insertCmd.Parameters.AddWithValue("@sameAsSnapshot", metadata.SameAsSnapshot)
                        insertCmd.ExecuteNonQuery()
                    End Using
                Next
            End Using
        End Sub

        Public Function BackfillOverrideRoutingSnapshots(databasePath As String) As Integer
            If String.IsNullOrWhiteSpace(databasePath) OrElse Not IO.File.Exists(databasePath) Then
                Return 0
            End If

            Dim updatedRows = 0
            Dim routingLookup As New Dictionary(Of String, BudgetRoutingSnapshot)(StringComparer.OrdinalIgnoreCase)

            For Each item In IncomeRepository.LoadIncome(databasePath)
                Dim label = If(String.IsNullOrWhiteSpace(item.Description), $"Income {item.Id}", item.Description.Trim())
                Dim snapshot As New BudgetRoutingSnapshot With {
                    .ToAccount = If(item.ToAccount, String.Empty).Trim(),
                    .ToAccountId = ResolveAccountIdByDescription(databasePath, item.ToAccount)
                }
                routingLookup($"0|{BuildOverrideStorageKey(0, item.Id, label)}") = snapshot
                If Not String.IsNullOrWhiteSpace(label) Then
                    routingLookup($"0|{label}") = snapshot
                End If
            Next

            For Each item In DebtRepository.LoadDebts(databasePath)
                Dim label = If(String.IsNullOrWhiteSpace(item.Description), $"Debt {item.Id}", item.Description.Trim())
                Dim snapshot As New BudgetRoutingSnapshot With {
                    .FromAccount = If(item.FromAccount, String.Empty).Trim(),
                    .SameAs = If(item.SameAs, String.Empty).Trim(),
                    .FromAccountId = ResolveAccountIdByDescription(databasePath, item.FromAccount),
                    .FromSavingsId = ResolveSavingsIdByFundingSource(databasePath, item.FromAccount),
                    .FromDebtId = ResolveDebtIdByFundingSource(databasePath, item.FromAccount),
                    .SameAsId = ResolveIncomeIdByDescription(databasePath, item.SameAs)
                }
                routingLookup($"1|{BuildOverrideStorageKey(1, item.Id, label)}") = snapshot
                If Not String.IsNullOrWhiteSpace(label) Then
                    routingLookup($"1|{label}") = snapshot
                End If
            Next

            For Each item In ExpenseRepository.LoadExpenses(databasePath)
                Dim label = If(String.IsNullOrWhiteSpace(item.Description), $"Expense {item.Id}", item.Description.Trim())
                Dim snapshot As New BudgetRoutingSnapshot With {
                    .FromAccount = If(item.FromAccount, String.Empty).Trim(),
                    .SameAs = If(item.SameAs, String.Empty).Trim(),
                    .FromAccountId = ResolveAccountIdByDescription(databasePath, item.FromAccount),
                    .FromSavingsId = ResolveSavingsIdByFundingSource(databasePath, item.FromAccount),
                    .FromDebtId = ResolveDebtIdByFundingSource(databasePath, item.FromAccount),
                    .SameAsId = ResolveIncomeIdByDescription(databasePath, item.SameAs)
                }
                routingLookup($"2|{BuildOverrideStorageKey(2, item.Id, label)}") = snapshot
                If Not String.IsNullOrWhiteSpace(label) Then
                    routingLookup($"2|{label}") = snapshot
                End If
            Next

            For Each item In SavingsRepository.LoadSavings(databasePath)
                Dim label = If(String.IsNullOrWhiteSpace(item.Description), $"Savings {item.Id}", item.Description.Trim())
                Dim snapshot As New BudgetRoutingSnapshot With {
                    .FromAccount = If(item.FromAccount, String.Empty).Trim(),
                    .SameAs = If(item.SameAs, String.Empty).Trim(),
                    .FromAccountId = ResolveAccountIdByDescription(databasePath, item.FromAccount),
                    .FromSavingsId = ResolveSavingsIdByFundingSource(databasePath, item.FromAccount),
                    .FromDebtId = ResolveDebtIdByFundingSource(databasePath, item.FromAccount),
                    .SameAsId = ResolveIncomeIdByDescription(databasePath, item.SameAs)
                }
                routingLookup($"3|{BuildOverrideStorageKey(3, item.Id, label)}") = snapshot
                If Not String.IsNullOrWhiteSpace(label) Then
                    routingLookup($"3|{label}") = snapshot
                End If
            Next

            Using conn As New SqliteConnection("Data Source=" & databasePath.Trim())
                conn.Open()
                EnsureBudgetDateOverridesTable(conn)

                Dim rowsToUpdate As New List(Of (DueDate As String, CatIdx As Integer, ItemDesc As String, FromAccount As String, ToAccount As String, SameAs As String, FromAccountId As Integer?, ToAccountId As Integer?, SameAsId As Integer?, FromSavingsId As Integer?, FromDebtId As Integer?))()

                Using selectCmd = conn.CreateCommand()
                    selectCmd.CommandText = "SELECT DueDate, CatIdx, ItemDesc, COALESCE(ItemLabel, ItemDesc), COALESCE(FromAccountSnapshot,''), COALESCE(ToAccountSnapshot,''), COALESCE(SameAsSnapshot,''), FromAccountSnapshotId, ToAccountSnapshotId, SameAsSnapshotId, FromSavingsSnapshotId, FromDebtSnapshotId FROM budgetDateOverrides WHERE CatIdx BETWEEN 0 AND 3"
                    Using reader = selectCmd.ExecuteReader()
                        While reader.Read()
                            Dim dueDate = If(reader.IsDBNull(0), String.Empty, Convert.ToString(reader.GetValue(0), CultureInfo.InvariantCulture))
                            Dim catIdx = If(reader.IsDBNull(1), -1, Convert.ToInt32(reader.GetValue(1), CultureInfo.InvariantCulture))
                            Dim itemDesc = If(reader.IsDBNull(2), String.Empty, Convert.ToString(reader.GetValue(2), CultureInfo.InvariantCulture)).Trim()
                            Dim itemLabel = If(reader.IsDBNull(3), itemDesc, Convert.ToString(reader.GetValue(3), CultureInfo.InvariantCulture)).Trim()
                            Dim fromAccountSnapshot = If(reader.IsDBNull(4), String.Empty, Convert.ToString(reader.GetValue(4), CultureInfo.InvariantCulture)).Trim()
                            Dim toAccountSnapshot = If(reader.IsDBNull(5), String.Empty, Convert.ToString(reader.GetValue(5), CultureInfo.InvariantCulture)).Trim()
                            Dim sameAsSnapshot = If(reader.IsDBNull(6), String.Empty, Convert.ToString(reader.GetValue(6), CultureInfo.InvariantCulture)).Trim()
                            Dim fromAccountSnapshotId = If(reader.IsDBNull(7), CType(Nothing, Integer?), Convert.ToInt32(reader.GetValue(7), CultureInfo.InvariantCulture))
                            Dim toAccountSnapshotId = If(reader.IsDBNull(8), CType(Nothing, Integer?), Convert.ToInt32(reader.GetValue(8), CultureInfo.InvariantCulture))
                            Dim sameAsSnapshotId = If(reader.IsDBNull(9), CType(Nothing, Integer?), Convert.ToInt32(reader.GetValue(9), CultureInfo.InvariantCulture))
                            Dim fromSavingsSnapshotId = If(reader.IsDBNull(10), CType(Nothing, Integer?), Convert.ToInt32(reader.GetValue(10), CultureInfo.InvariantCulture))
                            Dim fromDebtSnapshotId = If(reader.IsDBNull(11), CType(Nothing, Integer?), Convert.ToInt32(reader.GetValue(11), CultureInfo.InvariantCulture))

                            If String.IsNullOrWhiteSpace(itemDesc) Then
                                Continue While
                            End If

                            Dim routing As BudgetRoutingSnapshot = Nothing
                            If Not routingLookup.TryGetValue($"{catIdx}|{itemDesc}", routing) AndAlso
                               Not routingLookup.TryGetValue($"{catIdx}|{itemLabel}", routing) Then
                                Continue While
                            End If

                            Dim newFrom = If(String.IsNullOrWhiteSpace(fromAccountSnapshot), routing.FromAccount, fromAccountSnapshot)
                            Dim newTo = If(String.IsNullOrWhiteSpace(toAccountSnapshot), routing.ToAccount, toAccountSnapshot)
                            Dim newSameAs = If(String.IsNullOrWhiteSpace(sameAsSnapshot), routing.SameAs, sameAsSnapshot)
                            Dim newFromId = If(fromAccountSnapshotId.HasValue, fromAccountSnapshotId, routing.FromAccountId)
                            Dim newToId = If(toAccountSnapshotId.HasValue, toAccountSnapshotId, routing.ToAccountId)
                            Dim newSameAsId = If(sameAsSnapshotId.HasValue, sameAsSnapshotId, routing.SameAsId)
                            Dim newFromSavingsId = If(fromSavingsSnapshotId.HasValue, fromSavingsSnapshotId, routing.FromSavingsId)
                            Dim newFromDebtId = If(fromDebtSnapshotId.HasValue, fromDebtSnapshotId, routing.FromDebtId)

                            If String.Equals(newFrom, fromAccountSnapshot, StringComparison.Ordinal) AndAlso
                               String.Equals(newTo, toAccountSnapshot, StringComparison.Ordinal) AndAlso
                               String.Equals(newSameAs, sameAsSnapshot, StringComparison.Ordinal) AndAlso
                               Nullable.Equals(newFromId, fromAccountSnapshotId) AndAlso
                               Nullable.Equals(newToId, toAccountSnapshotId) AndAlso
                               Nullable.Equals(newSameAsId, sameAsSnapshotId) AndAlso
                               Nullable.Equals(newFromSavingsId, fromSavingsSnapshotId) AndAlso
                               Nullable.Equals(newFromDebtId, fromDebtSnapshotId) Then
                                Continue While
                            End If

                            rowsToUpdate.Add((dueDate, catIdx, itemDesc, newFrom, newTo, newSameAs, newFromId, newToId, newSameAsId, newFromSavingsId, newFromDebtId))
                        End While
                    End Using
                End Using

                For Each row In rowsToUpdate
                    Using updateCmd = conn.CreateCommand()
                        updateCmd.CommandText = "UPDATE budgetDateOverrides SET FromAccountSnapshot=@fromAccountSnapshot, ToAccountSnapshot=@toAccountSnapshot, SameAsSnapshot=@sameAsSnapshot, FromAccountSnapshotId=@fromAccountSnapshotId, ToAccountSnapshotId=@toAccountSnapshotId, SameAsSnapshotId=@sameAsSnapshotId, FromSavingsSnapshotId=@fromSavingsSnapshotId, FromDebtSnapshotId=@fromDebtSnapshotId WHERE DueDate=@d AND CatIdx=@c AND ItemDesc=@n"
                        updateCmd.Parameters.AddWithValue("@fromAccountSnapshot", row.FromAccount)
                        updateCmd.Parameters.AddWithValue("@toAccountSnapshot", row.ToAccount)
                        updateCmd.Parameters.AddWithValue("@sameAsSnapshot", row.SameAs)
                        updateCmd.Parameters.AddWithValue("@fromAccountSnapshotId", If(row.FromAccountId.HasValue, CType(row.FromAccountId.Value, Object), DBNull.Value))
                        updateCmd.Parameters.AddWithValue("@toAccountSnapshotId", If(row.ToAccountId.HasValue, CType(row.ToAccountId.Value, Object), DBNull.Value))
                        updateCmd.Parameters.AddWithValue("@sameAsSnapshotId", If(row.SameAsId.HasValue, CType(row.SameAsId.Value, Object), DBNull.Value))
                        updateCmd.Parameters.AddWithValue("@fromSavingsSnapshotId", If(row.FromSavingsId.HasValue, CType(row.FromSavingsId.Value, Object), DBNull.Value))
                        updateCmd.Parameters.AddWithValue("@fromDebtSnapshotId", If(row.FromDebtId.HasValue, CType(row.FromDebtId.Value, Object), DBNull.Value))
                        updateCmd.Parameters.AddWithValue("@d", row.DueDate)
                        updateCmd.Parameters.AddWithValue("@c", row.CatIdx)
                        updateCmd.Parameters.AddWithValue("@n", row.ItemDesc)
                        updatedRows += updateCmd.ExecuteNonQuery()
                    End Using
                Next
            End Using

            Return updatedRows
        End Function

        Private Function ResolveAccountIdByDescription(databasePath As String, accountDescription As String) As Integer?
            Dim normalized = If(accountDescription, String.Empty).Trim()
            If normalized = String.Empty Then
                Return Nothing
            End If

            For Each account In AccountRepository.LoadAccounts(databasePath)
                If String.Equals(If(account.Description, String.Empty).Trim(), normalized, StringComparison.OrdinalIgnoreCase) Then
                    Return account.Id
                End If
            Next

            Return Nothing
        End Function

        Private Function ResolveIncomeIdByDescription(databasePath As String, incomeDescription As String) As Integer?
            Dim normalized = If(incomeDescription, String.Empty).Trim()
            If normalized = String.Empty Then
                Return Nothing
            End If

            For Each income In IncomeRepository.LoadIncome(databasePath)
                If String.Equals(If(income.Description, String.Empty).Trim(), normalized, StringComparison.OrdinalIgnoreCase) Then
                    Return income.Id
                End If
            Next

            Return Nothing
        End Function

        Private Function ResolveSavingsIdByFundingSource(databasePath As String, fundingSource As String) As Integer?
            Dim savingsName = TryGetSavingsSourceName(fundingSource)
            If String.IsNullOrWhiteSpace(savingsName) Then
                Return Nothing
            End If

            For Each savings In SavingsRepository.LoadSavings(databasePath)
                If String.Equals(If(savings.Description, String.Empty).Trim(), savingsName, StringComparison.OrdinalIgnoreCase) Then
                    Return savings.Id
                End If
            Next

            Return Nothing
        End Function

        Private Function ResolveDebtIdByFundingSource(databasePath As String, fundingSource As String) As Integer?
            Dim debtName = TryGetDebtSourceName(fundingSource)
            If String.IsNullOrWhiteSpace(debtName) Then
                Return Nothing
            End If

            For Each debt In DebtRepository.LoadDebts(databasePath)
                If String.Equals(If(debt.Description, String.Empty).Trim(), debtName, StringComparison.OrdinalIgnoreCase) Then
                    Return debt.Id
                End If
            Next

            Return Nothing
        End Function

        Public Function LoadSettings(settingsPath As String, Optional databasePath As String = Nothing) As BudgetWorkspaceSettings
            Dim settings As New BudgetWorkspaceSettings()
            If Not String.IsNullOrWhiteSpace(settingsPath) AndAlso IO.File.Exists(settingsPath) Then
                For Each line In IO.File.ReadAllLines(settingsPath)
                    Dim parts = line.Split(New Char() {"="c}, 2)
                    If parts.Length <> 2 Then Continue For
                    Dim key = parts(0).Trim()
                    Dim value = parts(1).Trim()
                    Select Case key
                        Case "BudgetPeriod"
                            If Not String.IsNullOrWhiteSpace(value) Then settings.BudgetPeriod = value
                        Case "BudgetStartDate"
                            settings.BudgetStartDate = value
                        Case "BudgetYears"
                            Dim yrs As Integer
                            If Integer.TryParse(value, yrs) Then settings.BudgetYears = Math.Max(1, yrs)
                        Case "LastTheme"
                            If Not String.IsNullOrWhiteSpace(value) Then settings.AppTheme = value
                        Case "EnableAnimations"
                            Dim enabled As Boolean
                            If Boolean.TryParse(value, enabled) Then settings.EnableAnimations = enabled
                        Case "ThemeColor"
                            settings.ThemeColor = value
                        Case "BudgetZoomPercent"
                            Dim budgetZoom As Integer
                            If Integer.TryParse(value, budgetZoom) Then settings.BudgetZoomPercent = ClampZoomPercent(budgetZoom)
                        Case "TransactionsZoomPercent"
                            Dim transactionsZoom As Integer
                            If Integer.TryParse(value, transactionsZoom) Then settings.TransactionsZoomPercent = ClampZoomPercent(transactionsZoom)
                        Case "BudgetDistributionZoomPercent"
                            Dim budgetDistributionZoom As Integer
                            If Integer.TryParse(value, budgetDistributionZoom) Then settings.BudgetDistributionZoomPercent = ClampZoomPercent(budgetDistributionZoom)
                        Case "SavingsDistributionZoomPercent"
                            Dim savingsDistributionZoom As Integer
                            If Integer.TryParse(value, savingsDistributionZoom) Then settings.SavingsDistributionZoomPercent = ClampZoomPercent(savingsDistributionZoom)
                        Case "ServerMode"
                            If Not String.IsNullOrWhiteSpace(value) Then settings.ServerMode = value
                        Case "ServerPort"
                            Dim serverPort As Integer
                            If Integer.TryParse(value, serverPort) Then settings.ServerPort = Math.Max(1, serverPort)
                        Case "ExternalServerHost"
                            settings.ExternalServerHost = value
                        Case "ExternalServerPort"
                            Dim externalServerPort As Integer
                            If Integer.TryParse(value, externalServerPort) Then settings.ExternalServerPort = Math.Max(1, externalServerPort)
                        Case "ExternalHostOfflineCachePath"
                            settings.ExternalHostOfflineCachePath = value
                        Case "ExternalHostOfflineBaselineChangeToken"
                            settings.ExternalHostOfflineBaselineChangeToken = value
                        Case "PendingExternalHostOfflineSync"
                            Dim pendingOfflineSync As Boolean
                            If Boolean.TryParse(value, pendingOfflineSync) Then settings.PendingExternalHostOfflineSync = pendingOfflineSync
                        Case "GuidedTourCompleted"
                            Dim guidedTourCompleted As Boolean
                            If Boolean.TryParse(value, guidedTourCompleted) Then settings.GuidedTourCompleted = guidedTourCompleted
                        Case "AccountsDetailTourCompleted"
                            Dim accountsDetailTourCompleted As Boolean
                            If Boolean.TryParse(value, accountsDetailTourCompleted) Then settings.AccountsDetailTourCompleted = accountsDetailTourCompleted
                        Case "SavingsDetailTourCompleted"
                            Dim savingsDetailTourCompleted As Boolean
                            If Boolean.TryParse(value, savingsDetailTourCompleted) Then settings.SavingsDetailTourCompleted = savingsDetailTourCompleted
                        Case "DebtsDetailTourCompleted"
                            Dim debtsDetailTourCompleted As Boolean
                            If Boolean.TryParse(value, debtsDetailTourCompleted) Then settings.DebtsDetailTourCompleted = debtsDetailTourCompleted
                        Case "ExpensesDetailTourCompleted"
                            Dim expensesDetailTourCompleted As Boolean
                            If Boolean.TryParse(value, expensesDetailTourCompleted) Then settings.ExpensesDetailTourCompleted = expensesDetailTourCompleted
                        Case "TransactionsDetailTourCompleted"
                            Dim transactionsDetailTourCompleted As Boolean
                            If Boolean.TryParse(value, transactionsDetailTourCompleted) Then settings.TransactionsDetailTourCompleted = transactionsDetailTourCompleted
                    End Select
                Next
            End If

            If Not String.Equals(settings.ServerMode, "ExternalHost", StringComparison.OrdinalIgnoreCase) Then
                Dim appliedFromDatabase = False
                Try
                    appliedFromDatabase = TryApplyBudgetTimelineSettingsFromDatabase(databasePath, settings)
                Catch
                    appliedFromDatabase = False
                End Try

                If Not appliedFromDatabase AndAlso
                   Not String.IsNullOrWhiteSpace(databasePath) AndAlso
                   IO.File.Exists(databasePath) Then
                    Try
                        SaveBudgetTimelineSettings(databasePath, settings)
                    Catch
                        ' Ignore timeline migration failures during passive settings load.
                    End Try
                End If
            End If

            Return settings
        End Function

        Public Sub SaveSettings(settingsPath As String, settings As BudgetWorkspaceSettings, Optional databasePath As String = Nothing, Optional saveBudgetTimelineToDatabase As Boolean = True)
            If String.IsNullOrWhiteSpace(settingsPath) Then
                Throw New ArgumentException("A settings path is required.", NameOf(settingsPath))
            End If

            Dim lines As New List(Of String)
            If IO.File.Exists(settingsPath) Then
                lines.AddRange(IO.File.ReadAllLines(settingsPath))
            Else
                Dim dir = IO.Path.GetDirectoryName(settingsPath)
                If Not String.IsNullOrWhiteSpace(dir) AndAlso Not IO.Directory.Exists(dir) Then
                    IO.Directory.CreateDirectory(dir)
                End If
            End If

            RemoveSetting(lines, "BudgetPeriod")
            RemoveSetting(lines, "BudgetStartDate")
            RemoveSetting(lines, "BudgetYears")
            UpsertSetting(lines, "LastTheme", If(String.IsNullOrWhiteSpace(settings.AppTheme), "Light", settings.AppTheme.Trim()))
            UpsertSetting(lines, "EnableAnimations", settings.EnableAnimations.ToString())
            UpsertSetting(lines, "ThemeColor", If(settings.ThemeColor, String.Empty))
            UpsertSetting(lines, "BudgetZoomPercent", ClampZoomPercent(settings.BudgetZoomPercent).ToString(CultureInfo.InvariantCulture))
            UpsertSetting(lines, "TransactionsZoomPercent", ClampZoomPercent(settings.TransactionsZoomPercent).ToString(CultureInfo.InvariantCulture))
            UpsertSetting(lines, "BudgetDistributionZoomPercent", ClampZoomPercent(settings.BudgetDistributionZoomPercent).ToString(CultureInfo.InvariantCulture))
            UpsertSetting(lines, "SavingsDistributionZoomPercent", ClampZoomPercent(settings.SavingsDistributionZoomPercent).ToString(CultureInfo.InvariantCulture))
            UpsertSetting(lines, "ServerMode", If(String.IsNullOrWhiteSpace(settings.ServerMode), "Off", settings.ServerMode.Trim()))
            UpsertSetting(lines, "ServerPort", Math.Max(1, settings.ServerPort).ToString(CultureInfo.InvariantCulture))
            UpsertSetting(lines, "ExternalServerHost", If(settings.ExternalServerHost, String.Empty))
            UpsertSetting(lines, "ExternalServerPort", Math.Max(1, settings.ExternalServerPort).ToString(CultureInfo.InvariantCulture))
            UpsertSetting(lines, "ExternalHostOfflineCachePath", If(settings.ExternalHostOfflineCachePath, String.Empty))
            UpsertSetting(lines, "ExternalHostOfflineBaselineChangeToken", If(settings.ExternalHostOfflineBaselineChangeToken, String.Empty))
            UpsertSetting(lines, "PendingExternalHostOfflineSync", settings.PendingExternalHostOfflineSync.ToString())
            UpsertSetting(lines, "GuidedTourCompleted", settings.GuidedTourCompleted.ToString())
            UpsertSetting(lines, "AccountsDetailTourCompleted", settings.AccountsDetailTourCompleted.ToString())
            UpsertSetting(lines, "SavingsDetailTourCompleted", settings.SavingsDetailTourCompleted.ToString())
            UpsertSetting(lines, "DebtsDetailTourCompleted", settings.DebtsDetailTourCompleted.ToString())
            UpsertSetting(lines, "ExpensesDetailTourCompleted", settings.ExpensesDetailTourCompleted.ToString())
            UpsertSetting(lines, "TransactionsDetailTourCompleted", settings.TransactionsDetailTourCompleted.ToString())

            IO.File.WriteAllLines(settingsPath, lines)

            If saveBudgetTimelineToDatabase AndAlso
               Not String.Equals(settings.ServerMode, "ExternalHost", StringComparison.OrdinalIgnoreCase) Then
                SaveBudgetTimelineSettings(databasePath, settings)
            End If
        End Sub

        Public Sub SaveBudgetTimelineSettings(databasePath As String, settings As BudgetWorkspaceSettings)
            If String.IsNullOrWhiteSpace(databasePath) OrElse Not IO.File.Exists(databasePath) Then
                Return
            End If

            Using conn As New SqliteConnection("Data Source=" & databasePath.Trim())
                conn.Open()
                EnsureWorkspaceSettingsTable(conn)
                UpsertWorkspaceSetting(conn, "BudgetPeriod", If(String.IsNullOrWhiteSpace(settings.BudgetPeriod), "Monthly", settings.BudgetPeriod.Trim()))
                UpsertWorkspaceSetting(conn, "BudgetStartDate", settings.BudgetStartDate)
                UpsertWorkspaceSetting(conn, "BudgetYears", Math.Max(1, settings.BudgetYears).ToString(CultureInfo.InvariantCulture))
            End Using
        End Sub

        Public Function BuildSnapshot(databasePath As String, settings As BudgetWorkspaceSettings, Optional maxPeriods As Integer = Integer.MaxValue) As BudgetWorkspaceSnapshot
            Dim snapshot As New BudgetWorkspaceSnapshot()
            Dim startDate As DateTime
            If String.IsNullOrWhiteSpace(settings.BudgetStartDate) OrElse
               Not DateTime.TryParse(settings.BudgetStartDate, CultureInfo.CurrentCulture, DateTimeStyles.None, startDate) Then
                startDate = DateTime.Today
            End If

            Dim budgetYears = Math.Max(1, settings.BudgetYears)
            Dim periodName = If(String.IsNullOrWhiteSpace(settings.BudgetPeriod), "Monthly", settings.BudgetPeriod)
            Dim endExclusive = startDate.AddYears(budgetYears)
            Dim allPeriods = GeneratePeriods(startDate, endExclusive, periodName)
            Dim currentPeriodIndex = FindCurrentPeriodIndex(allPeriods, periodName)
            Dim periodsToLoad = Math.Min(allPeriods.Count, Math.Max(1, Math.Max(maxPeriods, currentPeriodIndex + 1)))
            Dim periods = allPeriods.Take(periodsToLoad).ToList()
            Dim debtBalanceHistoryOverrides = LoadHistoricalBudgetOverrides(databasePath, periods, periodName, DebtBalanceHistoryCategory)

            snapshot.BudgetPeriod = periodName
            snapshot.BudgetStart = startDate
            snapshot.BudgetEndExclusive = endExclusive
            snapshot.BudgetYears = budgetYears
            snapshot.TotalPeriodCount = allPeriods.Count
            snapshot.CurrentPeriodIndex = currentPeriodIndex
            snapshot.PeriodSummaries = periods.Select(Function(p) New BudgetWorkspacePeriodSummary With {.PeriodStart = p}).ToList()

            If String.IsNullOrWhiteSpace(databasePath) OrElse Not IO.File.Exists(databasePath) OrElse snapshot.PeriodSummaries.Count = 0 Then
                Return snapshot
            End If

            Dim incomes = IncomeRepository.LoadIncome(databasePath)
            Dim expenses = ExpenseRepository.LoadExpenses(databasePath)
            Dim savings = SavingsRepository.LoadSavings(databasePath)
            Dim debts = DebtRepository.LoadDebts(databasePath)
            Dim accounts = AccountRepository.LoadAccounts(databasePath)
            Dim sourceOverrides = LoadSourceBudgetOverrides(databasePath, periods, periodName)
            Dim manualOverrides = LoadManualBudgetOverrides(databasePath, periods, periodName)
            Dim handledManualKeys As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            Dim knownGroupLookup As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            Dim legacyOverrideOwners = BuildLegacyOverrideOwners(incomes, expenses, savings, debts)
            Dim incomeTotals(periods.Count - 1) As Decimal
            Dim debtTotals(periods.Count - 1) As Decimal
            Dim expenseTotals(periods.Count - 1) As Decimal
            Dim savingsTotals(periods.Count - 1) As Decimal
            Dim savingsWithdrawalDeltas As New Dictionary(Of String, Decimal())(StringComparer.OrdinalIgnoreCase)
            Dim debtChargeDeltas As New Dictionary(Of String, Decimal())(StringComparer.OrdinalIgnoreCase)

            For Each account In accounts
                Dim key = NormalizeAccountKey(account.Description)
                If Not snapshot.AccountRunningBalances.ContainsKey(key) Then
                    snapshot.AccountRunningBalances(key) = New Decimal(periods.Count - 1) {}
                End If
            Next
            If Not snapshot.AccountRunningBalances.ContainsKey(String.Empty) Then
                snapshot.AccountRunningBalances(String.Empty) = New Decimal(periods.Count - 1) {}
            End If

            For Each item In savings
                Dim key = If(item.Description, String.Empty).Trim()
                If key <> String.Empty AndAlso Not snapshot.SavingsRunningBalances.ContainsKey(key) Then
                    snapshot.SavingsRunningBalances(key) = New Decimal(periods.Count - 1) {}
                End If
                If key <> String.Empty AndAlso Not savingsWithdrawalDeltas.ContainsKey(key) Then
                    savingsWithdrawalDeltas(key) = New Decimal(periods.Count - 1) {}
                End If
            Next

            For Each item In debts
                If Not IsDebtFundingSource(item) Then Continue For
                Dim key = If(item.Description, String.Empty).Trim()
                If key <> String.Empty AndAlso Not debtChargeDeltas.ContainsKey(key) Then
                    debtChargeDeltas(key) = New Decimal(periods.Count - 1) {}
                End If
            Next

            Dim incomeByName As New Dictionary(Of String, IncomeRecord)(StringComparer.OrdinalIgnoreCase)
            Dim incomeNameById As New Dictionary(Of Integer, String)()
            For Each income In incomes
                Dim key = If(income.Description, String.Empty).Trim()
                If key = String.Empty Then Continue For
                If Not incomeByName.ContainsKey(key) Then
                    incomeByName(key) = income
                End If
                If income.Id > 0 AndAlso Not incomeNameById.ContainsKey(income.Id) Then
                    incomeNameById(income.Id) = key
                End If
            Next

            Dim accountNameById As New Dictionary(Of Integer, String)()
            For Each account In accounts
                Dim key = If(account.Description, String.Empty).Trim()
                If account.Id > 0 AndAlso key <> String.Empty AndAlso Not accountNameById.ContainsKey(account.Id) Then
                    accountNameById(account.Id) = key
                End If
            Next

            Dim savingsNameById As New Dictionary(Of Integer, String)()
            For Each savingsItem In savings
                Dim key = If(savingsItem.Description, String.Empty).Trim()
                If savingsItem.Id > 0 AndAlso key <> String.Empty AndAlso Not savingsNameById.ContainsKey(savingsItem.Id) Then
                    savingsNameById(savingsItem.Id) = key
                End If
            Next

            Dim debtNameById As New Dictionary(Of Integer, String)()
            For Each debtItem In debts
                Dim key = If(debtItem.Description, String.Empty).Trim()
                If debtItem.Id > 0 AndAlso key <> String.Empty AndAlso Not debtNameById.ContainsKey(debtItem.Id) Then
                    debtNameById(debtItem.Id) = key
                End If
            Next

            For Each item In incomes
                Dim incomeLabel = If(String.IsNullOrWhiteSpace(item.Description), $"Income {item.Id}", item.Description.Trim())
                Dim incomeSourceKey = BuildOverrideStorageKey(0, item.Id, incomeLabel)
                Dim incomeEndText = item.EndDate
                If Not item.Active Then
                    incomeEndText = snapshot.PeriodSummaries(Math.Max(0, snapshot.CurrentPeriodIndex)).PeriodStart.AddDays(-1).ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
                End If

                Dim currentIncomeBaselineDate = periods(Math.Min(Math.Max(snapshot.CurrentPeriodIndex, 0), periods.Count - 1))
                Dim incomeValues = BuildIncomeScheduledAmountSeries(periods, periodName, item, incomeEndText, currentIncomeBaselineDate, incomeByName)
                Dim incomeScheduledValues = CType(incomeValues.Clone(), Decimal())
                Dim incomeSourceBaseValues(periods.Count - 1) As Decimal
                Dim incomeSourceAdditionalValues(periods.Count - 1) As Decimal
                Dim incomeManualAdditionalValues(periods.Count - 1) As Decimal
                Dim incomeManualIndexes As New List(Of Integer)()
                Dim incomeSourceIndexes As New List(Of Integer)()
                Dim incomePaidIndexes As New List(Of Integer)()
                ApplySourceOverrides(0, incomeLabel, incomeSourceKey, incomeValues, sourceOverrides, legacyOverrideOwners, incomeSourceIndexes, snapshot.CurrentPeriodIndex, incomeSourceBaseValues, incomeSourceAdditionalValues, paidIndexes:=incomePaidIndexes)
                ApplyManualOverrides(0, incomeLabel, incomeSourceKey, incomeValues, manualOverrides, handledManualKeys, incomeManualIndexes, legacyOverrideOwners, incomeSourceIndexes, incomeManualAdditionalValues, incomePaidIndexes)
                ApplySeriesToAccountDeltaByPeriod(
                    snapshot.AccountRunningBalances,
                    incomeValues,
                    True,
                    Function(periodIndex)
                        Return ResolveHistoricalRoutingSnapshot(manualOverrides, sourceOverrides, 0, incomeLabel, incomeSourceKey, periodIndex, String.Empty, item.ToAccount, String.Empty, accountNameById, incomeNameById, savingsNameById, debtNameById).ToAccount
                    End Function)
                snapshot.ItemizedBudgetRows.Add(New BudgetWorkspaceItemSeries With {
                    .SectionName = "Income",
                    .GroupName = "Income",
                    .Label = incomeLabel,
                    .SourceLabel = incomeLabel,
                    .SourceKey = incomeSourceKey,
                    .Hidden = item.Hidden,
                    .StatusText = BuildIncomeStatusText(item),
                    .ScheduledValues = incomeScheduledValues,
                    .SourceBaseValues = incomeSourceBaseValues,
                    .SourceAdditionalValues = incomeSourceAdditionalValues,
                    .ManualAdditionalValues = incomeManualAdditionalValues,
                    .PaidIndexes = incomePaidIndexes.ToArray(),
                    .Values = incomeValues,
                    .ManualIndexes = incomeManualIndexes,
                    .SourceIndexes = incomeSourceIndexes
                })
                knownGroupLookup("0|" & incomeLabel) = "Income"
                AddSeriesToTotals(incomeTotals, incomeValues)
            Next

            For Each item In expenses
                Dim cadence = item.Cadence
                Dim dueDay = item.DueDay
                Dim dueDate = item.DueDate
                Dim startText As String = String.Empty
                Dim endText As String = String.Empty

                If (String.Equals(cadence, "Due Monthly", StringComparison.OrdinalIgnoreCase) OrElse
                    String.Equals(cadence, "Monthly on Day", StringComparison.OrdinalIgnoreCase)) AndAlso
                    ParseOptionalDate(dueDate).HasValue Then
                    startText = dueDate
                End If

                If Not item.Active Then
                    endText = snapshot.PeriodSummaries(Math.Max(0, snapshot.CurrentPeriodIndex)).PeriodStart.AddDays(-1).ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
                End If

                If String.Equals(cadence, "Same As", StringComparison.OrdinalIgnoreCase) AndAlso Not String.IsNullOrWhiteSpace(item.SameAs) AndAlso incomeByName.ContainsKey(item.SameAs) Then
                    Dim income = incomeByName(item.SameAs)
                    cadence = income.Cadence
                    dueDay = income.OnDay
                    dueDate = income.OnDate
                    startText = ResolveScheduleStartText(startText, income.StartDate)
                    endText = ResolveScheduleEndText(endText, income.EndDate)
                    If Not income.Active Then
                        Dim inactiveCutoff = snapshot.PeriodSummaries(Math.Max(0, snapshot.CurrentPeriodIndex)).PeriodStart.AddDays(-1).ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
                        endText = ResolveScheduleEndText(endText, inactiveCutoff)
                    End If
                End If

                Dim expenseLabel = If(String.IsNullOrWhiteSpace(item.Description), $"Expense {item.Id}", item.Description.Trim())
                Dim expenseSourceKey = BuildOverrideStorageKey(2, item.Id, expenseLabel)
                Dim expenseValues = BuildScheduledAmountSeries(periods, periodName, item.AmountDue, cadence, dueDay, dueDate, startText, endText)
                Dim expenseScheduledValues = CType(expenseValues.Clone(), Decimal())
                Dim expenseSourceBaseValues(periods.Count - 1) As Decimal
                Dim expenseSourceAdditionalValues(periods.Count - 1) As Decimal
                Dim expenseManualAdditionalValues(periods.Count - 1) As Decimal
                Dim expenseManualIndexes As New List(Of Integer)()
                Dim expenseSourceIndexes As New List(Of Integer)()
                Dim expensePaidIndexes As New List(Of Integer)()
                ApplySourceOverrides(2, expenseLabel, expenseSourceKey, expenseValues, sourceOverrides, legacyOverrideOwners, expenseSourceIndexes, snapshot.CurrentPeriodIndex, expenseSourceBaseValues, expenseSourceAdditionalValues, paidIndexes:=expensePaidIndexes)
                ApplyManualOverrides(2, expenseLabel, expenseSourceKey, expenseValues, manualOverrides, handledManualKeys, expenseManualIndexes, legacyOverrideOwners, expenseSourceIndexes, expenseManualAdditionalValues, expensePaidIndexes)
                ApplySeriesToAccountDeltaByPeriod(
                    snapshot.AccountRunningBalances,
                    expenseValues,
                    False,
                    Function(periodIndex)
                        Return ResolveHistoricalRoutingSnapshot(manualOverrides, sourceOverrides, 2, expenseLabel, expenseSourceKey, periodIndex, item.FromAccount, String.Empty, item.SameAs, accountNameById, incomeNameById, savingsNameById, debtNameById).FromAccount
                    End Function)
                ApplySeriesToSavingsWithdrawalByPeriod(
                    savingsWithdrawalDeltas,
                    expenseValues,
                    Function(periodIndex)
                        Return ResolveHistoricalRoutingSnapshot(manualOverrides, sourceOverrides, 2, expenseLabel, expenseSourceKey, periodIndex, item.FromAccount, String.Empty, item.SameAs, accountNameById, incomeNameById, savingsNameById, debtNameById).FromAccount
                    End Function)
                ApplySeriesToDebtChargeByPeriod(
                    debtChargeDeltas,
                    expenseValues,
                    Function(periodIndex)
                        Return ResolveHistoricalRoutingSnapshot(manualOverrides, sourceOverrides, 2, expenseLabel, expenseSourceKey, periodIndex, item.FromAccount, String.Empty, item.SameAs, accountNameById, incomeNameById, savingsNameById, debtNameById).FromAccount
                    End Function)
                snapshot.ItemizedBudgetRows.Add(New BudgetWorkspaceItemSeries With {
                    .SectionName = "Expenses",
                    .GroupName = NormalizeGroupName(item.Category, "Uncategorized"),
                    .Label = expenseLabel,
                    .SourceLabel = expenseLabel,
                    .SourceKey = expenseSourceKey,
                    .Hidden = item.Hidden,
                    .StatusText = BuildExpenseStatusText(item),
                    .ScheduledValues = expenseScheduledValues,
                    .SourceBaseValues = expenseSourceBaseValues,
                    .SourceAdditionalValues = expenseSourceAdditionalValues,
                    .ManualAdditionalValues = expenseManualAdditionalValues,
                    .PaidIndexes = expensePaidIndexes.ToArray(),
                    .Values = expenseValues,
                    .ManualIndexes = expenseManualIndexes,
                    .SourceIndexes = expenseSourceIndexes
                })
                knownGroupLookup("2|" & expenseLabel) = NormalizeGroupName(item.Category, "Uncategorized")
                AddSeriesToTotals(expenseTotals, expenseValues)
            Next

            For Each item In savings
                Dim cadence = item.Frequency
                Dim onDay = item.OnDay
                Dim onDate = item.OnDate
                Dim startText = item.StartDate
                Dim endText = item.EndDate

                If String.Equals(cadence, "Same As", StringComparison.OrdinalIgnoreCase) AndAlso Not String.IsNullOrWhiteSpace(item.SameAs) AndAlso incomeByName.ContainsKey(item.SameAs) Then
                    Dim income = incomeByName(item.SameAs)
                    cadence = income.Cadence
                    onDay = income.OnDay
                    onDate = income.OnDate
                    startText = ResolveScheduleStartText(startText, income.StartDate)
                    endText = ResolveScheduleEndText(endText, income.EndDate)
                    If Not income.Active Then
                        Dim inactiveCutoff = snapshot.PeriodSummaries(Math.Max(0, snapshot.CurrentPeriodIndex)).PeriodStart.AddDays(-1).ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
                        endText = ResolveScheduleEndText(endText, inactiveCutoff)
                    End If
                End If

                Dim savingsLabel = If(String.IsNullOrWhiteSpace(item.Description), $"Savings {item.Id}", item.Description.Trim())
                Dim savingsSourceKey = BuildOverrideStorageKey(3, item.Id, savingsLabel)
                Dim savingsValues = BuildScheduledAmountSeries(periods, periodName, item.DepositAmount, cadence, onDay, onDate, startText, endText)
                Dim savingsScheduledValues = CType(savingsValues.Clone(), Decimal())
                Dim savingsSourceBaseValues(periods.Count - 1) As Decimal
                Dim savingsSourceAdditionalValues(periods.Count - 1) As Decimal
                Dim savingsManualAdditionalValues(periods.Count - 1) As Decimal
                Dim savingsManualIndexes As New List(Of Integer)()
                Dim savingsSourceIndexes As New List(Of Integer)()
                Dim savingsPaidIndexes As New List(Of Integer)()
                ApplySourceOverrides(3, savingsLabel, savingsSourceKey, savingsValues, sourceOverrides, legacyOverrideOwners, savingsSourceIndexes, snapshot.CurrentPeriodIndex, savingsSourceBaseValues, savingsSourceAdditionalValues, paidIndexes:=savingsPaidIndexes)
                ApplyManualOverrides(3, savingsLabel, savingsSourceKey, savingsValues, manualOverrides, handledManualKeys, savingsManualIndexes, legacyOverrideOwners, savingsSourceIndexes, savingsManualAdditionalValues, savingsPaidIndexes)
                ApplySeriesToAccountDeltaByPeriod(
                    snapshot.AccountRunningBalances,
                    savingsValues,
                    False,
                    Function(periodIndex)
                        Return ResolveHistoricalRoutingSnapshot(manualOverrides, sourceOverrides, 3, savingsLabel, savingsSourceKey, periodIndex, item.FromAccount, String.Empty, item.SameAs, accountNameById, incomeNameById, savingsNameById, debtNameById).FromAccount
                    End Function)
                ApplySeriesToSavingsDelta(snapshot.SavingsRunningBalances, item.Description, savingsValues, True)
                ApplySeriesToSavingsWithdrawalByPeriod(
                    savingsWithdrawalDeltas,
                    savingsValues,
                    Function(periodIndex)
                        Return ResolveHistoricalRoutingSnapshot(manualOverrides, sourceOverrides, 3, savingsLabel, savingsSourceKey, periodIndex, item.FromAccount, String.Empty, item.SameAs, accountNameById, incomeNameById, savingsNameById, debtNameById).FromAccount
                    End Function)
                snapshot.ItemizedBudgetRows.Add(New BudgetWorkspaceItemSeries With {
                    .SectionName = "Savings",
                    .GroupName = NormalizeGroupName(item.Category, "Savings"),
                    .Label = savingsLabel,
                    .SourceLabel = savingsLabel,
                    .SourceKey = savingsSourceKey,
                    .Hidden = item.Hidden,
                    .StatusText = BuildSavingsStatusText(item),
                    .ScheduledValues = savingsScheduledValues,
                    .SourceBaseValues = savingsSourceBaseValues,
                    .SourceAdditionalValues = savingsSourceAdditionalValues,
                    .ManualAdditionalValues = savingsManualAdditionalValues,
                    .PaidIndexes = savingsPaidIndexes.ToArray(),
                    .Values = savingsValues,
                    .ManualIndexes = savingsManualIndexes,
                    .SourceIndexes = savingsSourceIndexes
                })
                knownGroupLookup("3|" & savingsLabel) = NormalizeGroupName(item.Category, "Savings")
                AddSeriesToTotals(savingsTotals, savingsValues)
            Next

            For Each item In debts
                Dim cadence = item.Cadence
                Dim dueDay = item.DayDue
                Dim dueDate As String = String.Empty
                Dim startText = item.StartDate
                Dim endText As String = String.Empty

                If String.Equals(cadence, "Yearly on Date", StringComparison.OrdinalIgnoreCase) OrElse
                   String.Equals(cadence, "Due Yearly", StringComparison.OrdinalIgnoreCase) Then
                    dueDate = item.StartDate
                    startText = String.Empty
                End If

                If String.Equals(cadence, "Same As", StringComparison.OrdinalIgnoreCase) AndAlso Not String.IsNullOrWhiteSpace(item.SameAs) AndAlso incomeByName.ContainsKey(item.SameAs) Then
                    Dim income = incomeByName(item.SameAs)
                    cadence = income.Cadence
                    dueDay = income.OnDay
                    dueDate = income.OnDate
                    startText = ResolveScheduleStartText(item.StartDate, income.StartDate)
                    endText = ResolveScheduleEndText(endText, income.EndDate)
                    If Not income.Active Then
                        Dim inactiveCutoff = snapshot.PeriodSummaries(Math.Max(0, snapshot.CurrentPeriodIndex)).PeriodStart.AddDays(-1).ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
                        endText = ResolveScheduleEndText(endText, inactiveCutoff)
                    End If
                End If

                Dim debtLabel = If(String.IsNullOrWhiteSpace(item.Description), $"Debt {item.Id}", item.Description.Trim())
                Dim debtSourceKey = BuildOverrideStorageKey(1, item.Id, debtLabel)
                Dim debtValues = BuildScheduledAmountSeries(periods, periodName, item.MinPayment, cadence, dueDay, dueDate, startText, endText)
                Dim debtScheduledValues = CType(debtValues.Clone(), Decimal())
                Dim debtSourceBaseValues(periods.Count - 1) As Decimal
                Dim debtSourceAdditionalValues(periods.Count - 1) As Decimal
                Dim debtManualAdditionalValues(periods.Count - 1) As Decimal
                Dim debtManualIndexes As New List(Of Integer)()
                Dim debtSourceIndexes As New List(Of Integer)()
                Dim debtPaidIndexes As New List(Of Integer)()
                ApplySourceOverrides(1, debtLabel, debtSourceKey, debtValues, sourceOverrides, legacyOverrideOwners, debtSourceIndexes, snapshot.CurrentPeriodIndex, debtSourceBaseValues, debtSourceAdditionalValues, paidIndexes:=debtPaidIndexes)
                ApplyManualOverrides(1, debtLabel, debtSourceKey, debtValues, manualOverrides, handledManualKeys, debtManualIndexes, legacyOverrideOwners, debtSourceIndexes, debtManualAdditionalValues, debtPaidIndexes)
                Dim debtChargeValues = If(debtChargeDeltas.ContainsKey(debtLabel), debtChargeDeltas(debtLabel), New Decimal(periods.Count - 1) {})
                Dim debtProjection = BuildDebtProjection(periods, periodName, item, cadence, dueDay, dueDate, startText, endText, debtValues, debtChargeValues, debtManualIndexes, debtSourceIndexes)
                debtValues = debtProjection.PaymentValues
                Dim debtResetStartingBalance = Math.Max(0D, item.StartingBalance)
                Dim debtResetFromPeriod = Math.Min(snapshot.CurrentPeriodIndex + 1, Math.Max(0, periods.Count - 1))
                If snapshot.CurrentPeriodIndex >= 0 AndAlso snapshot.CurrentPeriodIndex < debtProjection.BalanceValues.Length Then
                    debtResetStartingBalance = Math.Max(0D, debtProjection.BalanceValues(snapshot.CurrentPeriodIndex))

                    Dim latestHistoricalPeriodIndex = -1
                    Dim latestHistoricalBalance = FindLatestHistoricalOverrideValueAtOrBefore(
                        DebtBalanceHistoryCategory,
                        debtLabel,
                        debtSourceKey,
                        snapshot.CurrentPeriodIndex,
                        debtBalanceHistoryOverrides,
                        latestHistoricalPeriodIndex)

                    If latestHistoricalBalance.HasValue AndAlso latestHistoricalPeriodIndex >= 0 Then
                        debtResetStartingBalance = Math.Max(0D, latestHistoricalBalance.Value)
                        debtResetFromPeriod = latestHistoricalPeriodIndex
                    End If
                End If
                If debtResetFromPeriod < periods.Count Then
                    ResetDebtProjectionFromPeriod(periods, periodName, item, cadence, dueDay, dueDate, startText, endText, debtProjection, debtChargeValues, debtManualIndexes, debtSourceIndexes, debtResetFromPeriod, debtResetStartingBalance)
                End If
                ApplyHistoricalOverrides(DebtBalanceHistoryCategory, debtLabel, debtSourceKey, debtProjection.BalanceValues, debtBalanceHistoryOverrides, Nothing)
                ApplyHistoricalOverrides(DebtBalanceHistoryCategory, debtLabel, debtSourceKey, debtProjection.DisplayBalanceValues, debtBalanceHistoryOverrides, Nothing)
                ApplyManualOverrides(1, debtLabel, debtSourceKey, debtProjection.PaymentValues, manualOverrides, handledManualKeys, debtManualIndexes, legacyOverrideOwners, debtSourceIndexes, debtManualAdditionalValues, debtPaidIndexes)
                ApplySeriesToAccountDeltaByPeriod(
                    snapshot.AccountRunningBalances,
                    debtValues,
                    False,
                    Function(periodIndex)
                        Return ResolveHistoricalRoutingSnapshot(manualOverrides, sourceOverrides, 1, debtLabel, debtSourceKey, periodIndex, item.FromAccount, String.Empty, item.SameAs, accountNameById, incomeNameById, savingsNameById, debtNameById).FromAccount
                    End Function)
                ApplySeriesToSavingsWithdrawalByPeriod(
                    savingsWithdrawalDeltas,
                    debtValues,
                    Function(periodIndex)
                        Return ResolveHistoricalRoutingSnapshot(manualOverrides, sourceOverrides, 1, debtLabel, debtSourceKey, periodIndex, item.FromAccount, String.Empty, item.SameAs, accountNameById, incomeNameById, savingsNameById, debtNameById).FromAccount
                    End Function)
                snapshot.DebtRunningBalances(If(String.IsNullOrWhiteSpace(item.Description), $"Debt {item.Id}", item.Description.Trim())) =
                    debtProjection.BalanceValues
                snapshot.DebtDisplayBalances(If(String.IsNullOrWhiteSpace(item.Description), $"Debt {item.Id}", item.Description.Trim())) =
                    debtProjection.DisplayBalanceValues
                snapshot.ItemizedBudgetRows.Add(New BudgetWorkspaceItemSeries With {
                    .SectionName = "Debts",
                    .GroupName = NormalizeGroupName(item.Category, "Other"),
                    .Label = debtLabel,
                    .SourceLabel = debtLabel,
                    .SourceKey = debtSourceKey,
                    .Hidden = item.Hidden,
                    .StatusText = BuildDebtStatusText(item),
                    .ScheduledValues = debtScheduledValues,
                    .SourceBaseValues = debtSourceBaseValues,
                    .SourceAdditionalValues = debtSourceAdditionalValues,
                    .ManualAdditionalValues = debtManualAdditionalValues,
                    .PaidIndexes = debtPaidIndexes.ToArray(),
                    .Values = debtValues,
                    .ManualIndexes = debtManualIndexes,
                    .SourceIndexes = debtSourceIndexes
                })
                knownGroupLookup("1|" & debtLabel) = NormalizeGroupName(item.Category, "Other")
                AddSeriesToTotals(debtTotals, debtValues)
            Next

            AddManualOnlyRows(snapshot, manualOverrides, handledManualKeys, knownGroupLookup, incomeTotals, debtTotals, expenseTotals, savingsTotals)
            ApplySavingsWithdrawals(snapshot.SavingsRunningBalances, savingsWithdrawalDeltas)

            For i = 0 To snapshot.PeriodSummaries.Count - 1
                snapshot.PeriodSummaries(i).IncomeTotal = incomeTotals(i)
                snapshot.PeriodSummaries(i).DebtTotal = debtTotals(i)
                snapshot.PeriodSummaries(i).ExpenseTotal = expenseTotals(i)
                snapshot.PeriodSummaries(i).SavingsTotal = savingsTotals(i)
            Next

            ConvertDeltasToRunningBalances(snapshot.AccountRunningBalances)
            ConvertDeltasToRunningBalances(snapshot.SavingsRunningBalances)

            Return snapshot
        End Function

        Public Sub PurgeOrphanedBudgetOverrides(databasePath As String)
            ' Intentionally disabled.
            '
            ' Override cleanup used to physically delete live manual override rows based on
            ' key matching heuristics. That behavior is too risky to run automatically or
            ' on demand until it is rebuilt as an explicit preview/repair workflow.
            Return
        End Sub

        Private Sub ApplySeriesToAccountDelta(
            accountBalances As Dictionary(Of String, Decimal()),
            accountName As String,
            values As Decimal(),
            isPositive As Boolean)

            If accountBalances Is Nothing OrElse values Is Nothing Then Return
            If TryGetSavingsSourceName(accountName) <> String.Empty Then Return
            If TryGetDebtSourceName(accountName) <> String.Empty Then Return
            Dim key = NormalizeAccountKey(accountName)
            If Not accountBalances.ContainsKey(key) Then
                accountBalances(key) = New Decimal(values.Length - 1) {}
            End If

            Dim target = accountBalances(key)
            For i = 0 To Math.Min(target.Length, values.Length) - 1
                If isPositive Then
                    target(i) += values(i)
                Else
                    target(i) -= values(i)
                End If
            Next
        End Sub

        Private Sub ApplySeriesToAccountDeltaByPeriod(
            accountBalances As Dictionary(Of String, Decimal()),
            values As Decimal(),
            isPositive As Boolean,
            accountNameResolver As Func(Of Integer, String))

            If accountBalances Is Nothing OrElse values Is Nothing OrElse accountNameResolver Is Nothing Then Return

            For i = 0 To values.Length - 1
                Dim accountName = accountNameResolver(i)
                If TryGetSavingsSourceName(accountName) <> String.Empty Then Continue For
                If TryGetDebtSourceName(accountName) <> String.Empty Then Continue For

                Dim key = NormalizeAccountKey(accountName)
                If Not accountBalances.ContainsKey(key) Then
                    accountBalances(key) = New Decimal(values.Length - 1) {}
                End If

                Dim target = accountBalances(key)
                If isPositive Then
                    target(i) += values(i)
                Else
                    target(i) -= values(i)
                End If
            Next
        End Sub

        Private Sub ApplySeriesToSavingsDelta(
            savingsBalances As Dictionary(Of String, Decimal()),
            savingsName As String,
            values As Decimal(),
            isPositive As Boolean)

            Dim key = If(savingsName, String.Empty).Trim()
            If savingsBalances Is Nothing OrElse values Is Nothing OrElse key = String.Empty Then Return
            If Not savingsBalances.ContainsKey(key) Then
                savingsBalances(key) = New Decimal(values.Length - 1) {}
            End If

            Dim target = savingsBalances(key)
            For i = 0 To Math.Min(target.Length, values.Length) - 1
                If isPositive Then
                    target(i) += values(i)
                Else
                    target(i) -= values(i)
                End If
            Next
        End Sub

        Private Sub ApplySeriesToSavingsWithdrawal(
            savingsWithdrawals As Dictionary(Of String, Decimal()),
            accountName As String,
            values As Decimal())

            If savingsWithdrawals Is Nothing OrElse values Is Nothing Then Return
            Dim savingsName = TryGetSavingsSourceName(accountName)
            If savingsName = String.Empty Then Return
            If Not savingsWithdrawals.ContainsKey(savingsName) Then
                savingsWithdrawals(savingsName) = New Decimal(values.Length - 1) {}
            End If

            Dim target = savingsWithdrawals(savingsName)
            For i = 0 To Math.Min(target.Length, values.Length) - 1
                target(i) += values(i)
            Next
        End Sub

        Private Sub ApplySeriesToSavingsWithdrawalByPeriod(
            savingsWithdrawals As Dictionary(Of String, Decimal()),
            values As Decimal(),
            accountNameResolver As Func(Of Integer, String))

            If savingsWithdrawals Is Nothing OrElse values Is Nothing OrElse accountNameResolver Is Nothing Then Return

            For i = 0 To values.Length - 1
                Dim savingsName = TryGetSavingsSourceName(accountNameResolver(i))
                If savingsName = String.Empty Then Continue For
                If Not savingsWithdrawals.ContainsKey(savingsName) Then
                    savingsWithdrawals(savingsName) = New Decimal(values.Length - 1) {}
                End If

                Dim target = savingsWithdrawals(savingsName)
                target(i) += values(i)
            Next
        End Sub

        Private Sub ApplySeriesToDebtCharge(
            debtCharges As Dictionary(Of String, Decimal()),
            fundingSource As String,
            values As Decimal())

            If debtCharges Is Nothing OrElse values Is Nothing Then Return
            Dim debtName = TryGetDebtSourceName(fundingSource)
            If debtName = String.Empty OrElse Not debtCharges.ContainsKey(debtName) Then Return

            Dim target = debtCharges(debtName)
            For i = 0 To Math.Min(target.Length, values.Length) - 1
                target(i) += values(i)
            Next
        End Sub

        Private Sub ApplySeriesToDebtChargeByPeriod(
            debtCharges As Dictionary(Of String, Decimal()),
            values As Decimal(),
            fundingSourceResolver As Func(Of Integer, String))

            If debtCharges Is Nothing OrElse values Is Nothing OrElse fundingSourceResolver Is Nothing Then Return

            For i = 0 To values.Length - 1
                Dim debtName = TryGetDebtSourceName(fundingSourceResolver(i))
                If debtName = String.Empty OrElse Not debtCharges.ContainsKey(debtName) Then Continue For

                Dim target = debtCharges(debtName)
                target(i) += values(i)
            Next
        End Sub

        Private Sub ApplySavingsWithdrawals(
            savingsBalances As Dictionary(Of String, Decimal()),
            savingsWithdrawals As Dictionary(Of String, Decimal()))

            If savingsBalances Is Nothing OrElse savingsWithdrawals Is Nothing Then Return
            For Each key In savingsWithdrawals.Keys
                If Not savingsBalances.ContainsKey(key) Then Continue For
                Dim target = savingsBalances(key)
                Dim withdrawals = savingsWithdrawals(key)
                For i = 0 To Math.Min(target.Length, withdrawals.Length) - 1
                    target(i) -= withdrawals(i)
                Next
            Next
        End Sub

        Private Sub AddScheduledAccountDelta(
            accountBalances As Dictionary(Of String, Decimal()),
            periods As List(Of DateTime),
            budgetPeriod As String,
            amount As Decimal,
            cadence As String,
            dayNumber As Integer?,
            dateText As String,
            startText As String,
            endText As String,
            accountName As String,
            isPositive As Boolean)

            If accountBalances Is Nothing Then Return
            Dim key = NormalizeAccountKey(accountName)
            If Not accountBalances.ContainsKey(key) Then
                accountBalances(key) = New Decimal(periods.Count - 1) {}
            End If

            AddScheduledAmountToArray(accountBalances(key), periods, budgetPeriod, amount, cadence, dayNumber, dateText, startText, endText, isPositive)
        End Sub

        Private Sub AddScheduledSavingsDelta(
            savingsBalances As Dictionary(Of String, Decimal()),
            periods As List(Of DateTime),
            budgetPeriod As String,
            amount As Decimal,
            cadence As String,
            dayNumber As Integer?,
            dateText As String,
            startText As String,
            endText As String,
            savingsName As String)

            Dim key = If(savingsName, String.Empty).Trim()
            If key = String.Empty Then Return
            If Not savingsBalances.ContainsKey(key) Then
                savingsBalances(key) = New Decimal(periods.Count - 1) {}
            End If

            AddScheduledAmountToArray(savingsBalances(key), periods, budgetPeriod, amount, cadence, dayNumber, dateText, startText, endText, True)
        End Sub

        Private Sub AddScheduledAmountToArray(
            target As Decimal(),
            periods As List(Of DateTime),
            budgetPeriod As String,
            amount As Decimal,
            cadence As String,
            dayNumber As Integer?,
            dateText As String,
            startText As String,
            endText As String,
            isPositive As Boolean)

            If target Is Nothing OrElse periods Is Nothing OrElse periods.Count = 0 Then Return
            AddScheduledAmount(
                periods.Select(Function(p) New BudgetWorkspacePeriodSummary With {.PeriodStart = p}).ToList(),
                periods,
                budgetPeriod,
                amount,
                cadence,
                dayNumber,
                dateText,
                startText,
                endText,
                Sub(summary, scheduledAmount)
                    Dim idx = FindPeriodIndexForDate(periods, summary.PeriodStart, budgetPeriod)
                    If idx >= 0 AndAlso idx < target.Length Then
                        If isPositive Then
                            target(idx) += scheduledAmount
                        Else
                            target(idx) -= scheduledAmount
                        End If
                    End If
                End Sub)
        End Sub

        Private Function BuildScheduledAmountSeries(
            periods As List(Of DateTime),
            budgetPeriod As String,
            amount As Decimal,
            cadence As String,
            dayNumber As Integer?,
            dateText As String,
            startText As String,
            endText As String) As Decimal()

            Dim values(periods.Count - 1) As Decimal
            AddScheduledAmountToArray(values, periods, budgetPeriod, amount, cadence, dayNumber, dateText, startText, endText, True)
            Return values
        End Function

        Private Function BuildIncomeScheduledAmountSeries(
            periods As List(Of DateTime),
            budgetPeriod As String,
            item As IncomeRecord,
            endText As String,
            currentBaselineDate As DateTime,
            Optional incomeByName As Dictionary(Of String, IncomeRecord) = Nothing) As Decimal()

            Dim values(periods.Count - 1) As Decimal
            If item Is Nothing OrElse periods.Count = 0 Then
                Return values
            End If

            Dim effectiveCadence = If(item.Cadence, String.Empty).Trim()
            Dim effectiveOnDay = item.OnDay
            Dim effectiveOnDate = item.OnDate
            Dim effectiveStartText = item.StartDate
            Dim effectiveEndText = endText

            If String.Equals(effectiveCadence, "Same As", StringComparison.OrdinalIgnoreCase) AndAlso
               Not String.IsNullOrWhiteSpace(item.SameAs) AndAlso
               incomeByName IsNot Nothing AndAlso
               incomeByName.ContainsKey(item.SameAs) Then

                Dim linkedIncome = incomeByName(item.SameAs)
                effectiveCadence = If(linkedIncome.Cadence, String.Empty).Trim()
                effectiveOnDay = linkedIncome.OnDay
                effectiveOnDate = linkedIncome.OnDate
                effectiveStartText = ResolveScheduleStartText(item.StartDate, linkedIncome.StartDate)
                effectiveEndText = ResolveScheduleEndText(endText, linkedIncome.EndDate)

                If Not linkedIncome.Active Then
                    Return values
                End If
            End If

            If String.IsNullOrWhiteSpace(effectiveCadence) OrElse
               String.Equals(effectiveCadence, "Manually Entered", StringComparison.OrdinalIgnoreCase) Then
                Return values
            End If

            Dim itemStart = ParseDateOrDefault(effectiveStartText, periods(0))
            Dim itemEnd = ParseDateOrDefault(effectiveEndText, periods(periods.Count - 1).AddYears(50))

            If String.Equals(effectiveCadence, "Per Month", StringComparison.OrdinalIgnoreCase) OrElse
               String.Equals(effectiveCadence, "Per Year", StringComparison.OrdinalIgnoreCase) Then
                For i = 0 To periods.Count - 1
                    Dim pStart = periods(i)
                    If pStart.Date < itemStart.Date OrElse pStart.Date > itemEnd.Date Then
                        Continue For
                    End If

                    Dim periodEvaluationDate = GetPeriodEndExclusive(periods, i, budgetPeriod).AddDays(-1)
                    Dim adjustedAmount = GetIncomeAmountForForecastDate(item, periodEvaluationDate, currentBaselineDate)
                    values(i) += GetDistributedPerPeriodAmount(adjustedAmount, effectiveCadence, budgetPeriod)
                Next

                Return values
            End If

            For Each dueDate In GenerateDueDates(effectiveCadence, If(effectiveOnDay, 1), effectiveOnDate, itemStart, itemEnd, periods(0), GetPeriodEndExclusive(periods, periods.Count - 1, budgetPeriod))
                Dim idx = FindPeriodIndexForDate(periods, dueDate, budgetPeriod)
                If idx >= 0 AndAlso idx < values.Length Then
                    values(idx) += GetIncomeAmountForForecastDate(item, dueDate, currentBaselineDate)
                End If
            Next

            Return values
        End Function

        Private Sub AddSeriesToTotals(target As Decimal(), values As Decimal())
            If target Is Nothing OrElse values Is Nothing Then Return
            For i = 0 To Math.Min(target.Length, values.Length) - 1
                target(i) += values(i)
            Next
        End Sub

        Private Class ManualOverrideSeries
            Public Property StorageKey As String = String.Empty
            Public Property DisplayLabel As String = String.Empty
            Public Property Values As New Dictionary(Of Integer, Decimal)()
            Public Property Notes As New Dictionary(Of Integer, String)()
            Public Property Additionals As New Dictionary(Of Integer, Decimal)()
            Public Property ManualAmounts As New Dictionary(Of Integer, Decimal)()
            Public Property SelectionModes As New Dictionary(Of Integer, Integer)()
            Public Property PaidIndexes As New HashSet(Of Integer)()
            Public Property FromAccountSnapshots As New Dictionary(Of Integer, String)()
            Public Property ToAccountSnapshots As New Dictionary(Of Integer, String)()
            Public Property SameAsSnapshots As New Dictionary(Of Integer, String)()
            Public Property FromAccountSnapshotIds As New Dictionary(Of Integer, Integer)()
            Public Property ToAccountSnapshotIds As New Dictionary(Of Integer, Integer)()
            Public Property SameAsSnapshotIds As New Dictionary(Of Integer, Integer)()
            Public Property FromSavingsSnapshotIds As New Dictionary(Of Integer, Integer)()
            Public Property FromDebtSnapshotIds As New Dictionary(Of Integer, Integer)()
        End Class

        Private Function LoadManualBudgetOverrides(databasePath As String, periods As List(Of DateTime), budgetPeriod As String) As Dictionary(Of String, ManualOverrideSeries)
            Dim results As New Dictionary(Of String, ManualOverrideSeries)(StringComparer.OrdinalIgnoreCase)
            If String.IsNullOrWhiteSpace(databasePath) OrElse Not IO.File.Exists(databasePath) Then
                Return results
            End If

            Using conn As New SqliteConnection("Data Source=" & databasePath.Trim() & ";Mode=ReadOnly")
                conn.Open()
                Dim hasItemLabelColumn = HasBudgetDateOverridesColumn(conn, "ItemLabel")
                Dim hasNotesColumn = HasBudgetDateOverridesColumn(conn, "Notes")
                Dim hasAdditionalColumn = HasBudgetDateOverridesColumn(conn, "Additional")
                Dim hasPaidColumn = HasBudgetDateOverridesColumn(conn, "Paid")
                Dim hasSelectionModeColumn = HasBudgetDateOverridesColumn(conn, "SelectionMode")
                Dim hasManualAmountColumn = HasBudgetDateOverridesColumn(conn, "ManualAmount")
                Dim hasFromAccountSnapshotColumn = HasBudgetDateOverridesColumn(conn, "FromAccountSnapshot")
                Dim hasToAccountSnapshotColumn = HasBudgetDateOverridesColumn(conn, "ToAccountSnapshot")
                Dim hasSameAsSnapshotColumn = HasBudgetDateOverridesColumn(conn, "SameAsSnapshot")
                Dim hasFromAccountSnapshotIdColumn = HasBudgetDateOverridesColumn(conn, "FromAccountSnapshotId")
                Dim hasToAccountSnapshotIdColumn = HasBudgetDateOverridesColumn(conn, "ToAccountSnapshotId")
                Dim hasSameAsSnapshotIdColumn = HasBudgetDateOverridesColumn(conn, "SameAsSnapshotId")
                Dim hasFromSavingsSnapshotIdColumn = HasBudgetDateOverridesColumn(conn, "FromSavingsSnapshotId")
                Dim hasFromDebtSnapshotIdColumn = HasBudgetDateOverridesColumn(conn, "FromDebtSnapshotId")
                Using cmd = conn.CreateCommand()
                    cmd.CommandText =
                        If(hasItemLabelColumn AndAlso hasNotesColumn AndAlso hasAdditionalColumn AndAlso hasPaidColumn AndAlso hasSelectionModeColumn AndAlso hasManualAmountColumn,
                           "SELECT DueDate, CatIdx, ItemDesc, COALESCE(ItemLabel, ItemDesc), CASE WHEN COALESCE(SelectionMode,0)=3 THEN COALESCE(ManualAmount, Amount) ELSE Amount END, COALESCE(Notes, ''), COALESCE(Additional,0), COALESCE(Paid,0), COALESCE(SelectionMode,0), " &
                           If(hasFromAccountSnapshotColumn, "COALESCE(FromAccountSnapshot,'')", "''") & ", " &
                           If(hasToAccountSnapshotColumn, "COALESCE(ToAccountSnapshot,'')", "''") & ", " &
                           If(hasSameAsSnapshotColumn, "COALESCE(SameAsSnapshot,'')", "''") & ", " &
                           If(hasFromAccountSnapshotIdColumn, "FromAccountSnapshotId", "NULL") & ", " &
                           If(hasToAccountSnapshotIdColumn, "ToAccountSnapshotId", "NULL") & ", " &
                           If(hasSameAsSnapshotIdColumn, "SameAsSnapshotId", "NULL") & ", " &
                           If(hasFromSavingsSnapshotIdColumn, "FromSavingsSnapshotId", "NULL") & ", " &
                           If(hasFromDebtSnapshotIdColumn, "FromDebtSnapshotId", "NULL") &
                           " FROM budgetDateOverrides WHERE COALESCE(SelectionMode,0)=3 AND CatIdx>=0",
                           If(hasItemLabelColumn AndAlso hasNotesColumn AndAlso hasAdditionalColumn AndAlso hasPaidColumn AndAlso hasSelectionModeColumn,
                              "SELECT DueDate, CatIdx, ItemDesc, COALESCE(ItemLabel, ItemDesc), Amount, COALESCE(Notes, ''), COALESCE(Additional,0), COALESCE(Paid,0), COALESCE(SelectionMode,0), " &
                              If(hasFromAccountSnapshotColumn, "COALESCE(FromAccountSnapshot,'')", "''") & ", " &
                              If(hasToAccountSnapshotColumn, "COALESCE(ToAccountSnapshot,'')", "''") & ", " &
                              If(hasSameAsSnapshotColumn, "COALESCE(SameAsSnapshot,'')", "''") & ", " &
                              If(hasFromAccountSnapshotIdColumn, "FromAccountSnapshotId", "NULL") & ", " &
                              If(hasToAccountSnapshotIdColumn, "ToAccountSnapshotId", "NULL") & ", " &
                              If(hasSameAsSnapshotIdColumn, "SameAsSnapshotId", "NULL") & ", " &
                              If(hasFromSavingsSnapshotIdColumn, "FromSavingsSnapshotId", "NULL") & ", " &
                              If(hasFromDebtSnapshotIdColumn, "FromDebtSnapshotId", "NULL") &
                              " FROM budgetDateOverrides WHERE COALESCE(SelectionMode,0)=3 AND CatIdx>=0",
                              If(hasItemLabelColumn AndAlso hasNotesColumn AndAlso hasAdditionalColumn AndAlso hasPaidColumn,
                                 "SELECT DueDate, CatIdx, ItemDesc, COALESCE(ItemLabel, ItemDesc), Amount, COALESCE(Notes, ''), COALESCE(Additional,0), COALESCE(Paid,0), 3, '', '', '', NULL, NULL, NULL, NULL, NULL FROM budgetDateOverrides WHERE CatIdx>=0",
                                 If(hasItemLabelColumn AndAlso hasNotesColumn AndAlso hasAdditionalColumn,
                                    "SELECT DueDate, CatIdx, ItemDesc, COALESCE(ItemLabel, ItemDesc), Amount, COALESCE(Notes, ''), COALESCE(Additional,0), 0, 3, '', '', '', NULL, NULL, NULL, NULL, NULL FROM budgetDateOverrides WHERE CatIdx>=0",
                                    If(hasItemLabelColumn AndAlso hasNotesColumn,
                                       "SELECT DueDate, CatIdx, ItemDesc, COALESCE(ItemLabel, ItemDesc), Amount, COALESCE(Notes, ''), 0, 0, 3, '', '', '', NULL, NULL, NULL, NULL, NULL FROM budgetDateOverrides WHERE CatIdx>=0",
                                       If(hasItemLabelColumn,
                                          "SELECT DueDate, CatIdx, ItemDesc, COALESCE(ItemLabel, ItemDesc), Amount, '', 0, 0, 3, '', '', '', NULL, NULL, NULL, NULL, NULL FROM budgetDateOverrides WHERE CatIdx>=0",
                                          "SELECT DueDate, CatIdx, ItemDesc, ItemDesc, Amount, '', 0, 0, 3, '', '', '', NULL, NULL, NULL, NULL, NULL FROM budgetDateOverrides WHERE CatIdx>=0"))))))
                    Using rdr = cmd.ExecuteReader()
                        While rdr.Read()
                            Dim dueText = If(rdr.IsDBNull(0), String.Empty, Convert.ToString(rdr.GetValue(0)))
                            Dim catIdx = If(rdr.IsDBNull(1), -1, Convert.ToInt32(rdr.GetValue(1)))
                            Dim itemDesc = If(rdr.IsDBNull(2), String.Empty, Convert.ToString(rdr.GetValue(2))).Trim()
                            Dim itemLabel = If(rdr.IsDBNull(3), itemDesc, Convert.ToString(rdr.GetValue(3))).Trim()
                            Dim amount = If(rdr.IsDBNull(4), 0D, Convert.ToDecimal(rdr.GetValue(4)))
                            Dim note = If(rdr.IsDBNull(5), String.Empty, Convert.ToString(rdr.GetValue(5))).Trim()
                            Dim additional = If(rdr.IsDBNull(6), 0D, Convert.ToDecimal(rdr.GetValue(6), CultureInfo.InvariantCulture))
                            Dim paid = Not rdr.IsDBNull(7) AndAlso Convert.ToInt32(rdr.GetValue(7), CultureInfo.InvariantCulture) <> 0
                            Dim selectionMode = If(rdr.IsDBNull(8), 0, Convert.ToInt32(rdr.GetValue(8), CultureInfo.InvariantCulture))
                            If catIdx < 0 OrElse String.IsNullOrWhiteSpace(itemDesc) Then Continue While

                            Dim dueDate = ParseOptionalDate(dueText)
                            If Not dueDate.HasValue Then Continue While

                            Dim idx = FindPeriodIndexForDate(periods, dueDate.Value, budgetPeriod)
                            If idx < 0 Then Continue While

                            Dim key = catIdx.ToString(CultureInfo.InvariantCulture) & "|" & itemDesc
                            If Not results.ContainsKey(key) Then
                                results(key) = New ManualOverrideSeries With {
                                    .StorageKey = itemDesc,
                                    .DisplayLabel = If(String.IsNullOrWhiteSpace(itemLabel), itemDesc, itemLabel)
                                }
                            End If

                            results(key).Values(idx) = amount
                            results(key).Notes(idx) = note
                            results(key).Additionals(idx) = additional
                            results(key).SelectionModes(idx) = selectionMode
                            Dim fromAccountSnapshot = If(rdr.IsDBNull(9), String.Empty, Convert.ToString(rdr.GetValue(9))).Trim()
                            Dim toAccountSnapshot = If(rdr.IsDBNull(10), String.Empty, Convert.ToString(rdr.GetValue(10))).Trim()
                            Dim sameAsSnapshot = If(rdr.IsDBNull(11), String.Empty, Convert.ToString(rdr.GetValue(11))).Trim()
                            If Not String.IsNullOrWhiteSpace(fromAccountSnapshot) Then
                                results(key).FromAccountSnapshots(idx) = fromAccountSnapshot
                            End If
                            If Not String.IsNullOrWhiteSpace(toAccountSnapshot) Then
                                results(key).ToAccountSnapshots(idx) = toAccountSnapshot
                            End If
                            If Not String.IsNullOrWhiteSpace(sameAsSnapshot) Then
                                results(key).SameAsSnapshots(idx) = sameAsSnapshot
                            End If
                            If Not rdr.IsDBNull(12) Then
                                results(key).FromAccountSnapshotIds(idx) = Convert.ToInt32(rdr.GetValue(12), CultureInfo.InvariantCulture)
                            End If
                            If Not rdr.IsDBNull(13) Then
                                results(key).ToAccountSnapshotIds(idx) = Convert.ToInt32(rdr.GetValue(13), CultureInfo.InvariantCulture)
                            End If
                            If Not rdr.IsDBNull(14) Then
                                results(key).SameAsSnapshotIds(idx) = Convert.ToInt32(rdr.GetValue(14), CultureInfo.InvariantCulture)
                            End If
                            If Not rdr.IsDBNull(15) Then
                                results(key).FromSavingsSnapshotIds(idx) = Convert.ToInt32(rdr.GetValue(15), CultureInfo.InvariantCulture)
                            End If
                            If Not rdr.IsDBNull(16) Then
                                results(key).FromDebtSnapshotIds(idx) = Convert.ToInt32(rdr.GetValue(16), CultureInfo.InvariantCulture)
                            End If
                            If paid Then
                                results(key).PaidIndexes.Add(idx)
                            End If
                        End While
                    End Using
                End Using
            End Using

            Return results
        End Function

        Private Function LoadSourceBudgetOverrides(databasePath As String, periods As List(Of DateTime), budgetPeriod As String) As Dictionary(Of String, ManualOverrideSeries)
            Dim results As New Dictionary(Of String, ManualOverrideSeries)(StringComparer.OrdinalIgnoreCase)
            If String.IsNullOrWhiteSpace(databasePath) OrElse Not IO.File.Exists(databasePath) Then
                Return results
            End If

            Using conn As New SqliteConnection("Data Source=" & databasePath.Trim() & ";Mode=ReadOnly")
                conn.Open()
                Dim hasItemLabelColumn = HasBudgetDateOverridesColumn(conn, "ItemLabel")
                Dim hasAdditionalColumn = HasBudgetDateOverridesColumn(conn, "Additional")
                Dim hasPaidColumn = HasBudgetDateOverridesColumn(conn, "Paid")
                Dim hasSelectionModeColumn = HasBudgetDateOverridesColumn(conn, "SelectionMode")
                Dim hasManualAmountColumn = HasBudgetDateOverridesColumn(conn, "ManualAmount")
                Dim hasFromAccountSnapshotColumn = HasBudgetDateOverridesColumn(conn, "FromAccountSnapshot")
                Dim hasToAccountSnapshotColumn = HasBudgetDateOverridesColumn(conn, "ToAccountSnapshot")
                Dim hasSameAsSnapshotColumn = HasBudgetDateOverridesColumn(conn, "SameAsSnapshot")
                Dim hasFromAccountSnapshotIdColumn = HasBudgetDateOverridesColumn(conn, "FromAccountSnapshotId")
                Dim hasToAccountSnapshotIdColumn = HasBudgetDateOverridesColumn(conn, "ToAccountSnapshotId")
                Dim hasSameAsSnapshotIdColumn = HasBudgetDateOverridesColumn(conn, "SameAsSnapshotId")
                Dim hasFromSavingsSnapshotIdColumn = HasBudgetDateOverridesColumn(conn, "FromSavingsSnapshotId")
                Dim hasFromDebtSnapshotIdColumn = HasBudgetDateOverridesColumn(conn, "FromDebtSnapshotId")
                Using cmd = conn.CreateCommand()
                    cmd.CommandText =
                        If(hasItemLabelColumn AndAlso hasAdditionalColumn AndAlso hasPaidColumn AndAlso hasSelectionModeColumn AndAlso hasManualAmountColumn,
                           "SELECT DueDate, CatIdx, ItemDesc, COALESCE(ItemLabel, ItemDesc), Amount, COALESCE(Additional,0), COALESCE(ManualAmount,0), COALESCE(Paid,0), COALESCE(SelectionMode,0), " &
                           If(hasFromAccountSnapshotColumn, "COALESCE(FromAccountSnapshot,'')", "''") & ", " &
                           If(hasToAccountSnapshotColumn, "COALESCE(ToAccountSnapshot,'')", "''") & ", " &
                           If(hasSameAsSnapshotColumn, "COALESCE(SameAsSnapshot,'')", "''") & ", " &
                           If(hasFromAccountSnapshotIdColumn, "FromAccountSnapshotId", "NULL") & ", " &
                           If(hasToAccountSnapshotIdColumn, "ToAccountSnapshotId", "NULL") & ", " &
                           If(hasSameAsSnapshotIdColumn, "SameAsSnapshotId", "NULL") & ", " &
                           If(hasFromSavingsSnapshotIdColumn, "FromSavingsSnapshotId", "NULL") & ", " &
                           If(hasFromDebtSnapshotIdColumn, "FromDebtSnapshotId", "NULL") &
                           " FROM budgetDateOverrides WHERE COALESCE(SelectionMode,0)<>3 AND CatIdx>=0",
                           If(hasItemLabelColumn AndAlso hasAdditionalColumn AndAlso hasPaidColumn AndAlso hasSelectionModeColumn,
                              "SELECT DueDate, CatIdx, ItemDesc, COALESCE(ItemLabel, ItemDesc), Amount, COALESCE(Additional,0), 0, COALESCE(Paid,0), COALESCE(SelectionMode,0), " &
                              If(hasFromAccountSnapshotColumn, "COALESCE(FromAccountSnapshot,'')", "''") & ", " &
                              If(hasToAccountSnapshotColumn, "COALESCE(ToAccountSnapshot,'')", "''") & ", " &
                              If(hasSameAsSnapshotColumn, "COALESCE(SameAsSnapshot,'')", "''") & ", " &
                              If(hasFromAccountSnapshotIdColumn, "FromAccountSnapshotId", "NULL") & ", " &
                              If(hasToAccountSnapshotIdColumn, "ToAccountSnapshotId", "NULL") & ", " &
                              If(hasSameAsSnapshotIdColumn, "SameAsSnapshotId", "NULL") & ", " &
                              If(hasFromSavingsSnapshotIdColumn, "FromSavingsSnapshotId", "NULL") & ", " &
                              If(hasFromDebtSnapshotIdColumn, "FromDebtSnapshotId", "NULL") &
                              " FROM budgetDateOverrides WHERE COALESCE(SelectionMode,0)<>3 AND CatIdx>=0",
                              If(hasItemLabelColumn AndAlso hasAdditionalColumn AndAlso hasPaidColumn,
                                 "SELECT DueDate, CatIdx, ItemDesc, COALESCE(ItemLabel, ItemDesc), Amount, COALESCE(Additional,0), 0, COALESCE(Paid,0), 0, '', '', '', NULL, NULL, NULL, NULL, NULL FROM budgetDateOverrides WHERE CatIdx>=0",
                                 If(hasItemLabelColumn AndAlso hasAdditionalColumn,
                                    "SELECT DueDate, CatIdx, ItemDesc, COALESCE(ItemLabel, ItemDesc), Amount, COALESCE(Additional,0), 0, 0, 0, '', '', '', NULL, NULL, NULL, NULL, NULL FROM budgetDateOverrides WHERE CatIdx>=0",
                                    If(hasItemLabelColumn,
                                       "SELECT DueDate, CatIdx, ItemDesc, COALESCE(ItemLabel, ItemDesc), Amount, 0, 0, 0, 0, '', '', '', NULL, NULL, NULL, NULL, NULL FROM budgetDateOverrides WHERE CatIdx>=0",
                                       "SELECT DueDate, CatIdx, ItemDesc, ItemDesc, Amount, 0, 0, 0, 0, '', '', '', NULL, NULL, NULL, NULL, NULL FROM budgetDateOverrides WHERE CatIdx>=0")))))
                    Using rdr = cmd.ExecuteReader()
                        While rdr.Read()
                            Dim dueText = If(rdr.IsDBNull(0), String.Empty, Convert.ToString(rdr.GetValue(0)))
                            Dim catIdx = If(rdr.IsDBNull(1), -1, Convert.ToInt32(rdr.GetValue(1)))
                            Dim itemDesc = If(rdr.IsDBNull(2), String.Empty, Convert.ToString(rdr.GetValue(2))).Trim()
                            Dim itemLabel = If(rdr.IsDBNull(3), itemDesc, Convert.ToString(rdr.GetValue(3))).Trim()
                            Dim amount = If(rdr.IsDBNull(4), 0D, Convert.ToDecimal(rdr.GetValue(4)))
                            Dim additional = If(rdr.IsDBNull(5), 0D, Convert.ToDecimal(rdr.GetValue(5)))
                            Dim manualAmount = If(rdr.IsDBNull(6), 0D, Convert.ToDecimal(rdr.GetValue(6)))
                            Dim paid = Not rdr.IsDBNull(7) AndAlso Convert.ToInt32(rdr.GetValue(7), CultureInfo.InvariantCulture) <> 0
                            Dim selectionMode = If(rdr.IsDBNull(8), 0, Convert.ToInt32(rdr.GetValue(8), CultureInfo.InvariantCulture))
                            If catIdx < 0 OrElse String.IsNullOrWhiteSpace(itemDesc) Then Continue While

                            Dim dueDate = ParseOptionalDate(dueText)
                            If Not dueDate.HasValue Then Continue While

                            Dim idx = FindPeriodIndexForDate(periods, dueDate.Value, budgetPeriod)
                            If idx < 0 Then Continue While

                            Dim key = catIdx.ToString(CultureInfo.InvariantCulture) & "|" & itemDesc
                            If Not results.ContainsKey(key) Then
                                results(key) = New ManualOverrideSeries With {
                                    .StorageKey = itemDesc,
                                    .DisplayLabel = If(String.IsNullOrWhiteSpace(itemLabel), itemDesc, itemLabel)
                                }
                            End If

                            results(key).Values(idx) = amount
                            results(key).Additionals(idx) = additional
                            results(key).ManualAmounts(idx) = manualAmount
                            results(key).SelectionModes(idx) = selectionMode
                            Dim fromAccountSnapshot = If(rdr.IsDBNull(9), String.Empty, Convert.ToString(rdr.GetValue(9))).Trim()
                            Dim toAccountSnapshot = If(rdr.IsDBNull(10), String.Empty, Convert.ToString(rdr.GetValue(10))).Trim()
                            Dim sameAsSnapshot = If(rdr.IsDBNull(11), String.Empty, Convert.ToString(rdr.GetValue(11))).Trim()
                            If Not String.IsNullOrWhiteSpace(fromAccountSnapshot) Then
                                results(key).FromAccountSnapshots(idx) = fromAccountSnapshot
                            End If
                            If Not String.IsNullOrWhiteSpace(toAccountSnapshot) Then
                                results(key).ToAccountSnapshots(idx) = toAccountSnapshot
                            End If
                            If Not String.IsNullOrWhiteSpace(sameAsSnapshot) Then
                                results(key).SameAsSnapshots(idx) = sameAsSnapshot
                            End If
                            If Not rdr.IsDBNull(12) Then
                                results(key).FromAccountSnapshotIds(idx) = Convert.ToInt32(rdr.GetValue(12), CultureInfo.InvariantCulture)
                            End If
                            If Not rdr.IsDBNull(13) Then
                                results(key).ToAccountSnapshotIds(idx) = Convert.ToInt32(rdr.GetValue(13), CultureInfo.InvariantCulture)
                            End If
                            If Not rdr.IsDBNull(14) Then
                                results(key).SameAsSnapshotIds(idx) = Convert.ToInt32(rdr.GetValue(14), CultureInfo.InvariantCulture)
                            End If
                            If Not rdr.IsDBNull(15) Then
                                results(key).FromSavingsSnapshotIds(idx) = Convert.ToInt32(rdr.GetValue(15), CultureInfo.InvariantCulture)
                            End If
                            If Not rdr.IsDBNull(16) Then
                                results(key).FromDebtSnapshotIds(idx) = Convert.ToInt32(rdr.GetValue(16), CultureInfo.InvariantCulture)
                            End If
                            If paid Then
                                results(key).PaidIndexes.Add(idx)
                            End If
                        End While
                    End Using
                End Using
            End Using

            Return results
        End Function

        Private Function LoadHistoricalBudgetOverrides(databasePath As String, periods As List(Of DateTime), budgetPeriod As String, categoryIndex As Integer) As Dictionary(Of String, ManualOverrideSeries)
            Dim results As New Dictionary(Of String, ManualOverrideSeries)(StringComparer.OrdinalIgnoreCase)
            If String.IsNullOrWhiteSpace(databasePath) OrElse Not IO.File.Exists(databasePath) Then
                Return results
            End If

            Using conn As New SqliteConnection("Data Source=" & databasePath.Trim() & ";Mode=ReadOnly")
                conn.Open()
                Dim hasItemLabelColumn = HasBudgetDateOverridesColumn(conn, "ItemLabel")
                Dim hasNotesColumn = HasBudgetDateOverridesColumn(conn, "Notes")
                Using cmd = conn.CreateCommand()
                    cmd.CommandText =
                        If(hasItemLabelColumn AndAlso hasNotesColumn,
                           "SELECT DueDate, CatIdx, ItemDesc, COALESCE(ItemLabel, ItemDesc), CASE WHEN COALESCE(SelectionMode,0)=3 THEN COALESCE(ManualAmount, Amount) ELSE Amount END, COALESCE(Notes, '') FROM budgetDateOverrides WHERE COALESCE(SelectionMode,0)=3 AND CatIdx=@c",
                           If(hasItemLabelColumn,
                              "SELECT DueDate, CatIdx, ItemDesc, COALESCE(ItemLabel, ItemDesc), CASE WHEN COALESCE(SelectionMode,0)=3 THEN COALESCE(ManualAmount, Amount) ELSE Amount END, '' FROM budgetDateOverrides WHERE COALESCE(SelectionMode,0)=3 AND CatIdx=@c",
                              "SELECT DueDate, CatIdx, ItemDesc, ItemDesc, CASE WHEN COALESCE(SelectionMode,0)=3 THEN COALESCE(ManualAmount, Amount) ELSE Amount END, '' FROM budgetDateOverrides WHERE COALESCE(SelectionMode,0)=3 AND CatIdx=@c"))
                    cmd.Parameters.AddWithValue("@c", categoryIndex)
                    Using rdr = cmd.ExecuteReader()
                        While rdr.Read()
                            Dim dueText = If(rdr.IsDBNull(0), String.Empty, Convert.ToString(rdr.GetValue(0)))
                            Dim catIdx = If(rdr.IsDBNull(1), Int32.MinValue, Convert.ToInt32(rdr.GetValue(1)))
                            Dim itemDesc = If(rdr.IsDBNull(2), String.Empty, Convert.ToString(rdr.GetValue(2))).Trim()
                            Dim itemLabel = If(rdr.IsDBNull(3), itemDesc, Convert.ToString(rdr.GetValue(3))).Trim()
                            Dim amount = If(rdr.IsDBNull(4), 0D, Convert.ToDecimal(rdr.GetValue(4)))
                            Dim note = If(rdr.IsDBNull(5), String.Empty, Convert.ToString(rdr.GetValue(5))).Trim()
                            If catIdx <> categoryIndex OrElse String.IsNullOrWhiteSpace(itemDesc) Then Continue While

                            Dim dueDate = ParseOptionalDate(dueText)
                            If Not dueDate.HasValue Then Continue While

                            Dim idx = FindPeriodIndexForDate(periods, dueDate.Value, budgetPeriod)
                            If idx < 0 Then Continue While

                            Dim key = catIdx.ToString(CultureInfo.InvariantCulture) & "|" & itemDesc
                            If Not results.ContainsKey(key) Then
                                results(key) = New ManualOverrideSeries With {
                                    .StorageKey = itemDesc,
                                    .DisplayLabel = If(String.IsNullOrWhiteSpace(itemLabel), itemDesc, itemLabel)
                                }
                            End If

                            results(key).Values(idx) = amount
                            results(key).Notes(idx) = note
                        End While
                    End Using
                End Using
            End Using

            Return results
        End Function

        Private Function ResolveHistoricalRoutingSnapshot(
            manualOverrides As Dictionary(Of String, ManualOverrideSeries),
            sourceOverrides As Dictionary(Of String, ManualOverrideSeries),
            categoryIndex As Integer,
            label As String,
            sourceKey As String,
            periodIndex As Integer,
            fallbackFrom As String,
            fallbackTo As String,
            fallbackSameAs As String,
            accountNameById As Dictionary(Of Integer, String),
            incomeNameById As Dictionary(Of Integer, String),
            savingsNameById As Dictionary(Of Integer, String),
            debtNameById As Dictionary(Of Integer, String)) As BudgetRoutingSnapshot

            Dim snapshot As New BudgetRoutingSnapshot With {
                .FromAccount = If(fallbackFrom, String.Empty),
                .ToAccount = If(fallbackTo, String.Empty),
                .SameAs = If(fallbackSameAs, String.Empty)
            }

            ApplyRoutingSnapshotFromOverrideSeries(sourceOverrides, categoryIndex, label, sourceKey, periodIndex, snapshot, accountNameById, incomeNameById, savingsNameById, debtNameById)
            ApplyRoutingSnapshotFromOverrideSeries(manualOverrides, categoryIndex, label, sourceKey, periodIndex, snapshot, accountNameById, incomeNameById, savingsNameById, debtNameById)

            Return snapshot
        End Function

        Private Sub ApplyRoutingSnapshotFromOverrideSeries(
            overrideMap As Dictionary(Of String, ManualOverrideSeries),
            categoryIndex As Integer,
            label As String,
            sourceKey As String,
            periodIndex As Integer,
            snapshot As BudgetRoutingSnapshot,
            accountNameById As Dictionary(Of Integer, String),
            incomeNameById As Dictionary(Of Integer, String),
            savingsNameById As Dictionary(Of Integer, String),
            debtNameById As Dictionary(Of Integer, String))

            If overrideMap Is Nothing OrElse snapshot Is Nothing OrElse periodIndex < 0 Then
                Return
            End If

            Dim series = FindOverrideSeries(overrideMap, categoryIndex, label, sourceKey)
            If series Is Nothing Then
                Return
            End If

            If series.FromAccountSnapshots.ContainsKey(periodIndex) Then
                snapshot.FromAccount = series.FromAccountSnapshots(periodIndex)
            End If

            If series.ToAccountSnapshots.ContainsKey(periodIndex) Then
                snapshot.ToAccount = series.ToAccountSnapshots(periodIndex)
            End If

            If series.SameAsSnapshots.ContainsKey(periodIndex) Then
                snapshot.SameAs = series.SameAsSnapshots(periodIndex)
            End If

            If series.FromAccountSnapshotIds.ContainsKey(periodIndex) Then
                snapshot.FromAccountId = series.FromAccountSnapshotIds(periodIndex)
                If accountNameById IsNot Nothing AndAlso accountNameById.ContainsKey(snapshot.FromAccountId.Value) Then
                    snapshot.FromAccount = accountNameById(snapshot.FromAccountId.Value)
                End If
            End If

            If series.FromSavingsSnapshotIds.ContainsKey(periodIndex) Then
                snapshot.FromSavingsId = series.FromSavingsSnapshotIds(periodIndex)
                If savingsNameById IsNot Nothing AndAlso savingsNameById.ContainsKey(snapshot.FromSavingsId.Value) Then
                    snapshot.FromAccount = "(Savings) " & savingsNameById(snapshot.FromSavingsId.Value)
                End If
            End If

            If series.FromDebtSnapshotIds.ContainsKey(periodIndex) Then
                snapshot.FromDebtId = series.FromDebtSnapshotIds(periodIndex)
                If debtNameById IsNot Nothing AndAlso debtNameById.ContainsKey(snapshot.FromDebtId.Value) Then
                    snapshot.FromAccount = "(Debt) " & debtNameById(snapshot.FromDebtId.Value)
                End If
            End If

            If series.ToAccountSnapshotIds.ContainsKey(periodIndex) Then
                snapshot.ToAccountId = series.ToAccountSnapshotIds(periodIndex)
                If accountNameById IsNot Nothing AndAlso accountNameById.ContainsKey(snapshot.ToAccountId.Value) Then
                    snapshot.ToAccount = accountNameById(snapshot.ToAccountId.Value)
                End If
            End If

            If series.SameAsSnapshotIds.ContainsKey(periodIndex) Then
                snapshot.SameAsId = series.SameAsSnapshotIds(periodIndex)
                If incomeNameById IsNot Nothing AndAlso incomeNameById.ContainsKey(snapshot.SameAsId.Value) Then
                    snapshot.SameAs = incomeNameById(snapshot.SameAsId.Value)
                End If
            End If
        End Sub

        Private Function FindOverrideSeries(
            overrideMap As Dictionary(Of String, ManualOverrideSeries),
            categoryIndex As Integer,
            label As String,
            sourceKey As String) As ManualOverrideSeries

            If overrideMap Is Nothing Then
                Return Nothing
            End If

            Dim primaryKey = categoryIndex.ToString(CultureInfo.InvariantCulture) & "|" & sourceKey
            If overrideMap.ContainsKey(primaryKey) Then
                Return overrideMap(primaryKey)
            End If

            Dim legacyKey = categoryIndex.ToString(CultureInfo.InvariantCulture) & "|" & label
            If overrideMap.ContainsKey(legacyKey) Then
                Return overrideMap(legacyKey)
            End If

            Return Nothing
        End Function

        Private Sub EnsureBudgetDateOverridesTable(conn As SqliteConnection)
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "CREATE TABLE IF NOT EXISTS budgetDateOverrides (DueDate TEXT NOT NULL, CatIdx INTEGER NOT NULL, ItemDesc TEXT NOT NULL, ItemLabel TEXT NULL, Amount REAL NOT NULL, Additional REAL NOT NULL DEFAULT 0, ManualAmount REAL NOT NULL DEFAULT 0, Paid INTEGER NOT NULL DEFAULT 0, Notes TEXT NULL, SelectionMode INTEGER NOT NULL DEFAULT 0, FromAccountSnapshot TEXT NULL, ToAccountSnapshot TEXT NULL, SameAsSnapshot TEXT NULL, FromAccountSnapshotId INTEGER NULL, ToAccountSnapshotId INTEGER NULL, SameAsSnapshotId INTEGER NULL, FromSavingsSnapshotId INTEGER NULL, FromDebtSnapshotId INTEGER NULL, PRIMARY KEY(DueDate, CatIdx, ItemDesc))"
                cmd.ExecuteNonQuery()
            End Using
            MigrateBudgetDateOverridesLegacyColumns(conn)
            EnsureBudgetDateOverridesColumn(conn, "ItemLabel", "TEXT NULL")
            EnsureBudgetDateOverridesColumn(conn, "Notes", "TEXT NULL")
            EnsureBudgetDateOverridesColumn(conn, "Additional", "REAL NOT NULL DEFAULT 0")
            EnsureBudgetDateOverridesColumn(conn, "ManualAmount", "REAL NOT NULL DEFAULT 0")
            EnsureBudgetDateOverridesColumn(conn, "Paid", "INTEGER NOT NULL DEFAULT 0")
            EnsureBudgetDateOverridesColumn(conn, "SelectionMode", "INTEGER NOT NULL DEFAULT 0")
            EnsureBudgetDateOverridesColumn(conn, "FromAccountSnapshot", "TEXT NULL")
            EnsureBudgetDateOverridesColumn(conn, "ToAccountSnapshot", "TEXT NULL")
            EnsureBudgetDateOverridesColumn(conn, "SameAsSnapshot", "TEXT NULL")
            EnsureBudgetDateOverridesColumn(conn, "FromAccountSnapshotId", "INTEGER NULL")
            EnsureBudgetDateOverridesColumn(conn, "ToAccountSnapshotId", "INTEGER NULL")
            EnsureBudgetDateOverridesColumn(conn, "SameAsSnapshotId", "INTEGER NULL")
            EnsureBudgetDateOverridesColumn(conn, "FromSavingsSnapshotId", "INTEGER NULL")
            EnsureBudgetDateOverridesColumn(conn, "FromDebtSnapshotId", "INTEGER NULL")
        End Sub

        Private Sub MigrateBudgetDateOverridesLegacyColumns(conn As SqliteConnection)
            If conn Is Nothing Then
                Return
            End If

            Dim hasIgnoreBudgetFloor = HasBudgetDateOverridesColumn(conn, "IgnoreBudgetFloor")
            Dim hasIsManual = HasBudgetDateOverridesColumn(conn, "IsManual")
            If Not hasIgnoreBudgetFloor AndAlso Not hasIsManual Then
                Return
            End If

            Using tx = conn.BeginTransaction()
                Using createCmd = conn.CreateCommand()
                    createCmd.Transaction = tx
                    createCmd.CommandText = "CREATE TABLE budgetDateOverrides_new (DueDate TEXT NOT NULL, CatIdx INTEGER NOT NULL, ItemDesc TEXT NOT NULL, ItemLabel TEXT NULL, Amount REAL NOT NULL, Additional REAL NOT NULL DEFAULT 0, ManualAmount REAL NOT NULL DEFAULT 0, Paid INTEGER NOT NULL DEFAULT 0, Notes TEXT NULL, SelectionMode INTEGER NOT NULL DEFAULT 0, FromAccountSnapshot TEXT NULL, ToAccountSnapshot TEXT NULL, SameAsSnapshot TEXT NULL, FromAccountSnapshotId INTEGER NULL, ToAccountSnapshotId INTEGER NULL, SameAsSnapshotId INTEGER NULL, FromSavingsSnapshotId INTEGER NULL, FromDebtSnapshotId INTEGER NULL, PRIMARY KEY(DueDate, CatIdx, ItemDesc))"
                    createCmd.ExecuteNonQuery()
                End Using

                Dim hasItemLabel = HasBudgetDateOverridesColumn(conn, "ItemLabel")
                Dim hasAdditional = HasBudgetDateOverridesColumn(conn, "Additional")
                Dim hasManualAmount = HasBudgetDateOverridesColumn(conn, "ManualAmount")
                Dim hasPaid = HasBudgetDateOverridesColumn(conn, "Paid")
                Dim hasNotes = HasBudgetDateOverridesColumn(conn, "Notes")
                Dim hasSelectionMode = HasBudgetDateOverridesColumn(conn, "SelectionMode")
                Dim hasFromAccountSnapshot = HasBudgetDateOverridesColumn(conn, "FromAccountSnapshot")
                Dim hasToAccountSnapshot = HasBudgetDateOverridesColumn(conn, "ToAccountSnapshot")
                Dim hasSameAsSnapshot = HasBudgetDateOverridesColumn(conn, "SameAsSnapshot")
                Dim hasFromAccountSnapshotId = HasBudgetDateOverridesColumn(conn, "FromAccountSnapshotId")
                Dim hasToAccountSnapshotId = HasBudgetDateOverridesColumn(conn, "ToAccountSnapshotId")
                Dim hasSameAsSnapshotId = HasBudgetDateOverridesColumn(conn, "SameAsSnapshotId")
                Dim hasFromSavingsSnapshotId = HasBudgetDateOverridesColumn(conn, "FromSavingsSnapshotId")
                Dim hasFromDebtSnapshotId = HasBudgetDateOverridesColumn(conn, "FromDebtSnapshotId")

                Using selectCmd = conn.CreateCommand()
                    selectCmd.Transaction = tx
                    selectCmd.CommandText = "SELECT * FROM budgetDateOverrides"
                    Using reader = selectCmd.ExecuteReader()
                        While reader.Read()
                            Dim dueDate = If(reader("DueDate") Is DBNull.Value, String.Empty, Convert.ToString(reader("DueDate"), CultureInfo.InvariantCulture))
                            Dim catIdx = If(reader("CatIdx") Is DBNull.Value, 0, Convert.ToInt32(reader("CatIdx"), CultureInfo.InvariantCulture))
                            Dim itemDesc = If(reader("ItemDesc") Is DBNull.Value, String.Empty, Convert.ToString(reader("ItemDesc"), CultureInfo.InvariantCulture))
                            Dim amount = If(reader("Amount") Is DBNull.Value, 0D, Convert.ToDecimal(reader("Amount"), CultureInfo.InvariantCulture))
                            Dim itemLabel = If(hasItemLabel AndAlso reader("ItemLabel") IsNot DBNull.Value, Convert.ToString(reader("ItemLabel"), CultureInfo.InvariantCulture), itemDesc)
                            Dim additional = If(hasAdditional AndAlso reader("Additional") IsNot DBNull.Value, Convert.ToDecimal(reader("Additional"), CultureInfo.InvariantCulture), 0D)
                            Dim manualAmount = If(hasManualAmount AndAlso reader("ManualAmount") IsNot DBNull.Value, Convert.ToDecimal(reader("ManualAmount"), CultureInfo.InvariantCulture), 0D)
                            Dim paid = If(hasPaid AndAlso reader("Paid") IsNot DBNull.Value, Convert.ToInt32(reader("Paid"), CultureInfo.InvariantCulture), 0)
                            Dim notes = If(hasNotes AndAlso reader("Notes") IsNot DBNull.Value, Convert.ToString(reader("Notes"), CultureInfo.InvariantCulture), String.Empty)
                            Dim selectionMode = If(hasSelectionMode AndAlso reader("SelectionMode") IsNot DBNull.Value, Convert.ToInt32(reader("SelectionMode"), CultureInfo.InvariantCulture), 0)
                            Dim fromAccountSnapshot = If(hasFromAccountSnapshot AndAlso reader("FromAccountSnapshot") IsNot DBNull.Value, Convert.ToString(reader("FromAccountSnapshot"), CultureInfo.InvariantCulture), String.Empty)
                            Dim toAccountSnapshot = If(hasToAccountSnapshot AndAlso reader("ToAccountSnapshot") IsNot DBNull.Value, Convert.ToString(reader("ToAccountSnapshot"), CultureInfo.InvariantCulture), String.Empty)
                            Dim sameAsSnapshot = If(hasSameAsSnapshot AndAlso reader("SameAsSnapshot") IsNot DBNull.Value, Convert.ToString(reader("SameAsSnapshot"), CultureInfo.InvariantCulture), String.Empty)
                            Dim fromAccountSnapshotId As Object = If(hasFromAccountSnapshotId AndAlso reader("FromAccountSnapshotId") IsNot DBNull.Value, Convert.ToInt32(reader("FromAccountSnapshotId"), CultureInfo.InvariantCulture), DBNull.Value)
                            Dim toAccountSnapshotId As Object = If(hasToAccountSnapshotId AndAlso reader("ToAccountSnapshotId") IsNot DBNull.Value, Convert.ToInt32(reader("ToAccountSnapshotId"), CultureInfo.InvariantCulture), DBNull.Value)
                            Dim sameAsSnapshotId As Object = If(hasSameAsSnapshotId AndAlso reader("SameAsSnapshotId") IsNot DBNull.Value, Convert.ToInt32(reader("SameAsSnapshotId"), CultureInfo.InvariantCulture), DBNull.Value)
                            Dim fromSavingsSnapshotId As Object = If(hasFromSavingsSnapshotId AndAlso reader("FromSavingsSnapshotId") IsNot DBNull.Value, Convert.ToInt32(reader("FromSavingsSnapshotId"), CultureInfo.InvariantCulture), DBNull.Value)
                            Dim fromDebtSnapshotId As Object = If(hasFromDebtSnapshotId AndAlso reader("FromDebtSnapshotId") IsNot DBNull.Value, Convert.ToInt32(reader("FromDebtSnapshotId"), CultureInfo.InvariantCulture), DBNull.Value)
                            Dim legacyIsManual = If(hasIsManual AndAlso reader("IsManual") IsNot DBNull.Value, Convert.ToInt32(reader("IsManual"), CultureInfo.InvariantCulture) <> 0, False)

                            If legacyIsManual Then
                                If manualAmount = 0D Then
                                    manualAmount = amount
                                End If

                                If selectionMode = 0 Then
                                    selectionMode = 3
                                End If
                            End If

                            Using insertCmd = conn.CreateCommand()
                                insertCmd.Transaction = tx
                                insertCmd.CommandText = "INSERT INTO budgetDateOverrides_new (DueDate, CatIdx, ItemDesc, ItemLabel, Amount, Additional, ManualAmount, Paid, Notes, SelectionMode, FromAccountSnapshot, ToAccountSnapshot, SameAsSnapshot, FromAccountSnapshotId, ToAccountSnapshotId, SameAsSnapshotId, FromSavingsSnapshotId, FromDebtSnapshotId) VALUES (@d, @c, @n, @l, @a, @additional, @manualAmount, @paid, @notes, @selectionMode, @fromAccountSnapshot, @toAccountSnapshot, @sameAsSnapshot, @fromAccountSnapshotId, @toAccountSnapshotId, @sameAsSnapshotId, @fromSavingsSnapshotId, @fromDebtSnapshotId)"
                                insertCmd.Parameters.AddWithValue("@d", dueDate)
                                insertCmd.Parameters.AddWithValue("@c", catIdx)
                                insertCmd.Parameters.AddWithValue("@n", itemDesc)
                                insertCmd.Parameters.AddWithValue("@l", If(String.IsNullOrWhiteSpace(itemLabel), itemDesc, itemLabel))
                                insertCmd.Parameters.AddWithValue("@a", amount)
                                insertCmd.Parameters.AddWithValue("@additional", additional)
                                insertCmd.Parameters.AddWithValue("@manualAmount", manualAmount)
                                insertCmd.Parameters.AddWithValue("@paid", paid)
                                insertCmd.Parameters.AddWithValue("@notes", If(notes, String.Empty))
                                insertCmd.Parameters.AddWithValue("@selectionMode", selectionMode)
                                insertCmd.Parameters.AddWithValue("@fromAccountSnapshot", If(fromAccountSnapshot, String.Empty))
                                insertCmd.Parameters.AddWithValue("@toAccountSnapshot", If(toAccountSnapshot, String.Empty))
                                insertCmd.Parameters.AddWithValue("@sameAsSnapshot", If(sameAsSnapshot, String.Empty))
                                insertCmd.Parameters.AddWithValue("@fromAccountSnapshotId", fromAccountSnapshotId)
                                insertCmd.Parameters.AddWithValue("@toAccountSnapshotId", toAccountSnapshotId)
                                insertCmd.Parameters.AddWithValue("@sameAsSnapshotId", sameAsSnapshotId)
                                insertCmd.Parameters.AddWithValue("@fromSavingsSnapshotId", fromSavingsSnapshotId)
                                insertCmd.Parameters.AddWithValue("@fromDebtSnapshotId", fromDebtSnapshotId)
                                insertCmd.ExecuteNonQuery()
                            End Using
                        End While
                    End Using
                End Using

                Using dropCmd = conn.CreateCommand()
                    dropCmd.Transaction = tx
                    dropCmd.CommandText = "DROP TABLE budgetDateOverrides"
                    dropCmd.ExecuteNonQuery()
                End Using

                Using renameCmd = conn.CreateCommand()
                    renameCmd.Transaction = tx
                    renameCmd.CommandText = "ALTER TABLE budgetDateOverrides_new RENAME TO budgetDateOverrides"
                    renameCmd.ExecuteNonQuery()
                End Using

                tx.Commit()
            End Using
        End Sub

        Private Sub ApplyManualOverrides(catIdx As Integer, label As String, sourceKey As String, values As Decimal(), manualOverrides As Dictionary(Of String, ManualOverrideSeries), handledKeys As HashSet(Of String), manualIndexes As List(Of Integer), legacyOverrideOwners As Dictionary(Of String, String), Optional sourceIndexes As IEnumerable(Of Integer) = Nothing, Optional manualAdditionalValues As Decimal() = Nothing, Optional paidIndexes As List(Of Integer) = Nothing)
            Dim exactKey = catIdx.ToString(CultureInfo.InvariantCulture) & "|" & sourceKey.Trim()
            Dim legacyKey = catIdx.ToString(CultureInfo.InvariantCulture) & "|" & label.Trim()
            Dim key As String
            If manualOverrides.ContainsKey(exactKey) Then
                key = exactKey
            ElseIf manualOverrides.ContainsKey(legacyKey) AndAlso ShouldUseLegacyOverride(legacyKey, sourceKey, legacyOverrideOwners) Then
                key = legacyKey
            Else
                Return
            End If

            Dim sourceIndexSet As New HashSet(Of Integer)(If(sourceIndexes, Enumerable.Empty(Of Integer)()))

            For Each entry In manualOverrides(key).Values
                If entry.Key >= 0 AndAlso entry.Key < values.Length Then
                    Dim selectionMode = If(manualOverrides(key).SelectionModes.ContainsKey(entry.Key), manualOverrides(key).SelectionModes(entry.Key), 0)
                    If (catIdx = 0 OrElse catIdx = 1 OrElse catIdx = 2 OrElse catIdx = 3) AndAlso selectionMode = 1 Then
                        If paidIndexes IsNot Nothing AndAlso manualOverrides(key).PaidIndexes.Contains(entry.Key) AndAlso Not paidIndexes.Contains(entry.Key) Then
                            paidIndexes.Add(entry.Key)
                        End If
                        Continue For
                    End If
                    If sourceIndexSet.Contains(entry.Key) Then
                        Continue For
                    End If
                    values(entry.Key) = entry.Value
                    If manualAdditionalValues IsNot Nothing AndAlso entry.Key < manualAdditionalValues.Length Then
                        manualAdditionalValues(entry.Key) = If(manualOverrides(key).Additionals.ContainsKey(entry.Key), manualOverrides(key).Additionals(entry.Key), 0D)
                    End If
                    If paidIndexes IsNot Nothing AndAlso manualOverrides(key).PaidIndexes.Contains(entry.Key) AndAlso Not paidIndexes.Contains(entry.Key) Then
                        paidIndexes.Add(entry.Key)
                    End If
                    If Not manualIndexes.Contains(entry.Key) Then manualIndexes.Add(entry.Key)
                End If
            Next

            handledKeys.Add(key)
        End Sub

        Private Sub ApplySourceOverrides(catIdx As Integer, label As String, sourceKey As String, values As Decimal(), sourceOverrides As Dictionary(Of String, ManualOverrideSeries), legacyOverrideOwners As Dictionary(Of String, String), Optional sourceIndexes As List(Of Integer) = Nothing, Optional currentPeriodIndex As Integer = -1, Optional sourceBaseValues As Decimal() = Nothing, Optional sourceAdditionalValues As Decimal() = Nothing, Optional paidIndexes As List(Of Integer) = Nothing)
            Dim exactKey = catIdx.ToString(CultureInfo.InvariantCulture) & "|" & sourceKey.Trim()
            Dim legacyKey = catIdx.ToString(CultureInfo.InvariantCulture) & "|" & label.Trim()
            Dim key As String
            If sourceOverrides.ContainsKey(exactKey) Then
                key = exactKey
            ElseIf sourceOverrides.ContainsKey(legacyKey) AndAlso ShouldUseLegacyOverride(legacyKey, sourceKey, legacyOverrideOwners) Then
                key = legacyKey
            Else
                Return
            End If

            For Each entry In sourceOverrides(key).Values
                If entry.Key >= 0 AndAlso entry.Key < values.Length Then
                    Dim selectionMode = If(sourceOverrides(key).SelectionModes.ContainsKey(entry.Key), sourceOverrides(key).SelectionModes(entry.Key), 0)
                    Dim additional = If(sourceOverrides(key).Additionals.ContainsKey(entry.Key), sourceOverrides(key).Additionals(entry.Key), 0D)
                    Dim manualAmount = If(sourceOverrides(key).ManualAmounts.ContainsKey(entry.Key), sourceOverrides(key).ManualAmounts(entry.Key), 0D)
                    Dim sourceValue = entry.Value + additional
                    If sourceBaseValues IsNot Nothing AndAlso entry.Key < sourceBaseValues.Length Then
                        sourceBaseValues(entry.Key) = entry.Value
                    End If
                    If sourceAdditionalValues IsNot Nothing AndAlso entry.Key < sourceAdditionalValues.Length Then
                        sourceAdditionalValues(entry.Key) = additional
                    End If
                    If paidIndexes IsNot Nothing AndAlso sourceOverrides(key).PaidIndexes.Contains(entry.Key) AndAlso Not paidIndexes.Contains(entry.Key) Then
                        paidIndexes.Add(entry.Key)
                    End If
                    Dim isPastPeriod = currentPeriodIndex >= 0 AndAlso entry.Key < currentPeriodIndex
                    If catIdx = 0 OrElse catIdx = 1 OrElse catIdx = 2 OrElse catIdx = 3 Then
                        If selectionMode = 1 Then
                            Continue For
                        ElseIf selectionMode = 3 Then
                            values(entry.Key) = manualAmount
                        Else
                            values(entry.Key) = sourceValue
                        End If
                    ElseIf isPastPeriod Then
                        values(entry.Key) = sourceValue
                    Else
                        values(entry.Key) = Math.Max(values(entry.Key), sourceValue)
                    End If
                    If sourceIndexes IsNot Nothing AndAlso Not sourceIndexes.Contains(entry.Key) Then
                        sourceIndexes.Add(entry.Key)
                    End If
                End If
            Next
        End Sub

        Private Sub ApplyHistoricalOverrides(catIdx As Integer, label As String, sourceKey As String, values As Decimal(), historicalOverrides As Dictionary(Of String, ManualOverrideSeries), handledKeys As HashSet(Of String))
            Dim exactKey = catIdx.ToString(CultureInfo.InvariantCulture) & "|" & sourceKey.Trim()
            Dim legacyKey = catIdx.ToString(CultureInfo.InvariantCulture) & "|" & label.Trim()
            Dim key As String = Nothing
            If historicalOverrides.ContainsKey(exactKey) Then
                key = exactKey
            ElseIf historicalOverrides.ContainsKey(legacyKey) Then
                key = legacyKey
            End If

            If String.IsNullOrWhiteSpace(key) Then Return

            For Each entry In historicalOverrides(key).Values
                If entry.Key >= 0 AndAlso entry.Key < values.Length Then
                    values(entry.Key) = entry.Value
                End If
            Next

            If handledKeys IsNot Nothing Then
                handledKeys.Add(key)
            End If
        End Sub

        Private Function FindHistoricalOverrideValue(catIdx As Integer, label As String, sourceKey As String, periodIndex As Integer, historicalOverrides As Dictionary(Of String, ManualOverrideSeries)) As Decimal?
            If historicalOverrides Is Nothing OrElse periodIndex < 0 Then
                Return Nothing
            End If

            Dim exactKey = catIdx.ToString(CultureInfo.InvariantCulture) & "|" & sourceKey.Trim()
            Dim legacyKey = catIdx.ToString(CultureInfo.InvariantCulture) & "|" & label.Trim()
            Dim key As String = Nothing
            If historicalOverrides.ContainsKey(exactKey) Then
                key = exactKey
            ElseIf historicalOverrides.ContainsKey(legacyKey) Then
                key = legacyKey
            End If

            If String.IsNullOrWhiteSpace(key) Then
                Return Nothing
            End If

            Dim amount As Decimal
            If historicalOverrides(key).Values.TryGetValue(periodIndex, amount) Then
                Return amount
            End If

            Return Nothing
        End Function

        Private Function FindLatestHistoricalOverrideValueAtOrBefore(
            catIdx As Integer,
            label As String,
            sourceKey As String,
            maxPeriodIndex As Integer,
            historicalOverrides As Dictionary(Of String, ManualOverrideSeries),
            ByRef resolvedPeriodIndex As Integer) As Decimal?

            resolvedPeriodIndex = -1
            If historicalOverrides Is Nothing OrElse maxPeriodIndex < 0 Then
                Return Nothing
            End If

            Dim exactKey = catIdx.ToString(CultureInfo.InvariantCulture) & "|" & sourceKey.Trim()
            Dim legacyKey = catIdx.ToString(CultureInfo.InvariantCulture) & "|" & label.Trim()
            Dim key As String = Nothing
            If historicalOverrides.ContainsKey(exactKey) Then
                key = exactKey
            ElseIf historicalOverrides.ContainsKey(legacyKey) Then
                key = legacyKey
            End If

            If String.IsNullOrWhiteSpace(key) Then
                Return Nothing
            End If

            Dim latestAmount As Decimal? = Nothing
            For Each entry In historicalOverrides(key).Values
                If entry.Key <= maxPeriodIndex AndAlso entry.Key >= 0 AndAlso entry.Key >= resolvedPeriodIndex Then
                    resolvedPeriodIndex = entry.Key
                    latestAmount = entry.Value
                End If
            Next

            Return latestAmount
        End Function

        Private Sub AddManualOnlyRows(
            snapshot As BudgetWorkspaceSnapshot,
            manualOverrides As Dictionary(Of String, ManualOverrideSeries),
            handledKeys As HashSet(Of String),
            knownGroupLookup As Dictionary(Of String, String),
            incomeTotals As Decimal(),
            debtTotals As Decimal(),
            expenseTotals As Decimal(),
            savingsTotals As Decimal())

            For Each key In manualOverrides.Keys.OrderBy(Function(k) k, StringComparer.OrdinalIgnoreCase)
                If handledKeys.Contains(key) Then Continue For

                Dim pieces = key.Split({"|"c}, 2)
                If pieces.Length <> 2 Then Continue For

                Dim catIdx As Integer
                If Not Integer.TryParse(pieces(0), catIdx) Then Continue For

                Dim label = manualOverrides(key).DisplayLabel
                Dim values(snapshot.PeriodSummaries.Count - 1) As Decimal
                For Each entry In manualOverrides(key).Values
                    If entry.Key >= 0 AndAlso entry.Key < values.Length Then
                        values(entry.Key) = entry.Value
                    End If
                Next
                Dim manualIndexes = manualOverrides(key).Values.Keys.OrderBy(Function(i) i).ToList()
                Dim paidIndexes = manualOverrides(key).PaidIndexes.OrderBy(Function(i) i).ToArray()

                Dim sectionName = GetSectionNameForCategory(catIdx)
                If String.IsNullOrWhiteSpace(sectionName) Then Continue For
                Dim knownGroupKey = catIdx.ToString(CultureInfo.InvariantCulture) & "|" & label
                Dim groupName = If(knownGroupLookup.ContainsKey(knownGroupKey), knownGroupLookup(knownGroupKey), InferManualGroupName(catIdx))
                Dim existingRow = snapshot.ItemizedBudgetRows.Any(Function(row) _
                    String.Equals(row.SectionName, sectionName, StringComparison.OrdinalIgnoreCase) AndAlso
                    (String.Equals(row.SourceKey, manualOverrides(key).StorageKey, StringComparison.OrdinalIgnoreCase) OrElse
                     String.Equals(row.SourceLabel, label, StringComparison.OrdinalIgnoreCase)))
                If existingRow Then Continue For

                snapshot.ItemizedBudgetRows.Add(New BudgetWorkspaceItemSeries With {
                    .SectionName = sectionName,
                    .GroupName = groupName,
                    .Label = label & " (Manual)",
                    .SourceLabel = label,
                    .SourceKey = manualOverrides(key).StorageKey,
                    .Hidden = False,
                    .StatusText = "Manual budget override row",
                    .ScheduledValues = values,
                    .PaidIndexes = paidIndexes,
                    .Values = values,
                    .ManualIndexes = manualIndexes
                })

                Select Case catIdx
                    Case 0
                        AddSeriesToTotals(incomeTotals, values)
                    Case 1
                        AddSeriesToTotals(debtTotals, values)
                    Case 2
                        AddSeriesToTotals(expenseTotals, values)
                    Case 3
                        AddSeriesToTotals(savingsTotals, values)
                End Select
            Next
        End Sub

        Private Function BuildOverrideStorageKey(categoryIndex As Integer, sourceId As Integer, sourceLabel As String) As String
            If sourceId > 0 Then
                Return $"ID:{categoryIndex}:{sourceId}"
            End If

            Return If(sourceLabel, String.Empty).Trim()
        End Function

        Private Function ShouldUseLegacyOverride(legacyKey As String, sourceKey As String, legacyOverrideOwners As Dictionary(Of String, String)) As Boolean
            If Not sourceKey.StartsWith("ID:", StringComparison.OrdinalIgnoreCase) Then
                Return True
            End If

            Return legacyOverrideOwners.ContainsKey(legacyKey) AndAlso
                   String.Equals(legacyOverrideOwners(legacyKey), sourceKey, StringComparison.OrdinalIgnoreCase)
        End Function

        Private Function BuildLegacyOverrideOwners(
            incomes As IEnumerable(Of IncomeRecord),
            expenses As IEnumerable(Of ExpenseRecord),
            savings As IEnumerable(Of SavingsRecord),
            debts As IEnumerable(Of DebtRecord)) As Dictionary(Of String, String)

            Dim owners As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

            For Each item In incomes.OrderBy(Function(x) x.Id)
                AddLegacyOverrideOwner(owners, 0, item.Id, item.Description, $"Income {item.Id}")
            Next

            For Each item In debts.OrderBy(Function(x) x.Id)
                AddLegacyOverrideOwner(owners, 1, item.Id, item.Description, $"Debt {item.Id}")
            Next

            For Each item In expenses.OrderBy(Function(x) x.Id)
                AddLegacyOverrideOwner(owners, 2, item.Id, item.Description, $"Expense {item.Id}")
            Next

            For Each item In savings.OrderBy(Function(x) x.Id)
                AddLegacyOverrideOwner(owners, 3, item.Id, item.Description, $"Savings {item.Id}")
            Next

            Return owners
        End Function

        Private Sub AddLegacyOverrideOwner(owners As Dictionary(Of String, String), categoryIndex As Integer, sourceId As Integer, description As String, fallbackLabel As String)
            Dim label = If(String.IsNullOrWhiteSpace(description), fallbackLabel, description.Trim())
            Dim legacyKey = categoryIndex.ToString(CultureInfo.InvariantCulture) & "|" & label
            If owners.ContainsKey(legacyKey) Then
                Return
            End If

            owners(legacyKey) = BuildOverrideStorageKey(categoryIndex, sourceId, label)
        End Sub

        Private Sub CleanupOrphanedManualOverrides(
            databasePath As String,
            incomes As IEnumerable(Of IncomeRecord),
            expenses As IEnumerable(Of ExpenseRecord),
            savings As IEnumerable(Of SavingsRecord),
            debts As IEnumerable(Of DebtRecord),
            Optional force As Boolean = False)

            If String.IsNullOrWhiteSpace(databasePath) OrElse Not IO.File.Exists(databasePath) Then
                Return
            End If

            Dim normalizedPath = databasePath.Trim()

            Dim incomeList = incomes.OrderBy(Function(x) x.Id).ToList()
            Dim debtList = debts.OrderBy(Function(x) x.Id).ToList()
            Dim expenseList = expenses.OrderBy(Function(x) x.Id).ToList()
            Dim savingsList = savings.OrderBy(Function(x) x.Id).ToList()
            Dim combinedSignature =
                "0:" & BuildOverrideCleanupSignature(incomeList.Select(Function(x) (x.Id, x.Description))) & "|" &
                "1:" & BuildOverrideCleanupSignature(debtList.Select(Function(x) (x.Id, x.Description))) & "|" &
                "2:" & BuildOverrideCleanupSignature(expenseList.Select(Function(x) (x.Id, x.Description))) & "|" &
                "3:" & BuildOverrideCleanupSignature(savingsList.Select(Function(x) (x.Id, x.Description)))

            If Not force Then
                SyncLock BudgetOverrideCleanupSignatures
                    Dim cachedSignature As String = Nothing
                    If BudgetOverrideCleanupSignatures.TryGetValue(normalizedPath, cachedSignature) AndAlso
                       String.Equals(cachedSignature, combinedSignature, StringComparison.Ordinal) Then
                        Return
                    End If
                End SyncLock
            End If

            Dim validKeysByCategory As New Dictionary(Of Integer, HashSet(Of String)) From {
                {0, BuildValidOverrideKeys(0, incomeList.Select(Function(x) (x.Id, x.Description)), "Income")},
                {1, BuildValidOverrideKeys(1, debtList.Select(Function(x) (x.Id, x.Description)), "Debt")},
                {2, BuildValidOverrideKeys(2, expenseList.Select(Function(x) (x.Id, x.Description)), "Expense")},
                {3, BuildValidOverrideKeys(3, savingsList.Select(Function(x) (x.Id, x.Description)), "Savings")}
            }

            Using conn As New SqliteConnection("Data Source=" & normalizedPath)
                conn.Open()
                EnsureBudgetDateOverridesTable(conn)

                For Each categoryIndex In validKeysByCategory.Keys
                    Dim existingKeys As New List(Of String)()
                    Using selectCmd = conn.CreateCommand()
                        selectCmd.CommandText = "SELECT DISTINCT ItemDesc FROM budgetDateOverrides WHERE CatIdx=@c AND COALESCE(SelectionMode,0)=3"
                        selectCmd.Parameters.AddWithValue("@c", categoryIndex)
                        Using reader = selectCmd.ExecuteReader()
                            While reader.Read()
                                Dim key = If(reader.IsDBNull(0), String.Empty, Convert.ToString(reader.GetValue(0))).Trim()
                                If key <> String.Empty Then
                                    existingKeys.Add(key)
                                End If
                            End While
                        End Using
                    End Using

                    For Each existingKey In existingKeys
                        If validKeysByCategory(categoryIndex).Contains(existingKey) Then
                            Continue For
                        End If

                        Using deleteCmd = conn.CreateCommand()
                            deleteCmd.CommandText = "DELETE FROM budgetDateOverrides WHERE CatIdx=@c AND ItemDesc=@n"
                            deleteCmd.Parameters.AddWithValue("@c", categoryIndex)
                            deleteCmd.Parameters.AddWithValue("@n", existingKey)
                            deleteCmd.ExecuteNonQuery()
                        End Using
                    Next
                Next
            End Using

            SyncLock BudgetOverrideCleanupSignatures
                BudgetOverrideCleanupSignatures(normalizedPath) = combinedSignature
            End SyncLock
        End Sub

        Private Function BuildOverrideCleanupSignature(items As IEnumerable(Of (Id As Integer, Description As String))) As String
            Return String.Join("|", items.Select(Function(x) $"{x.Id}:{If(x.Description, String.Empty).Trim()}"))
        End Function

        Private Function BuildValidOverrideKeys(
            categoryIndex As Integer,
            items As IEnumerable(Of (Id As Integer, Description As String)),
            fallbackPrefix As String) As HashSet(Of String)

            Dim results As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            Dim orderedItems = items.OrderBy(Function(x) x.Id).ToList()

            For Each item In orderedItems
                Dim label = If(String.IsNullOrWhiteSpace(item.Description), $"{fallbackPrefix} {item.Id}", item.Description.Trim())
                results.Add(BuildOverrideStorageKey(categoryIndex, item.Id, label))
            Next

            For Each group In orderedItems.
                GroupBy(Function(x) If(String.IsNullOrWhiteSpace(x.Description), $"{fallbackPrefix} {x.Id}", x.Description.Trim()),
                        StringComparer.OrdinalIgnoreCase)
                Dim owner = group.OrderBy(Function(x) x.Id).FirstOrDefault()
                If owner.Id <= 0 Then
                    Continue For
                End If

                Dim label = If(String.IsNullOrWhiteSpace(owner.Description), $"{fallbackPrefix} {owner.Id}", owner.Description.Trim())
                results.Add(label)
            Next

            Return results
        End Function

        Private Sub EnsureBudgetDateOverridesColumn(conn As SqliteConnection, columnName As String, columnDefinition As String)
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "ALTER TABLE budgetDateOverrides ADD COLUMN " & columnName & " " & columnDefinition
                Try
                    cmd.ExecuteNonQuery()
                Catch ex As SqliteException When ex.SqliteErrorCode = 1 AndAlso ex.Message.IndexOf("duplicate column", StringComparison.OrdinalIgnoreCase) >= 0
                    ' Column already exists.
                End Try
            End Using
        End Sub

        Private Sub EnsureWorkspaceSettingsTable(conn As SqliteConnection)
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "CREATE TABLE IF NOT EXISTS workspaceSettings (SettingKey TEXT NOT NULL PRIMARY KEY, SettingValue TEXT NOT NULL DEFAULT '')"
                cmd.ExecuteNonQuery()
            End Using
        End Sub

        Private Sub UpsertWorkspaceSetting(conn As SqliteConnection, key As String, value As String)
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "INSERT INTO workspaceSettings (SettingKey, SettingValue) VALUES (@key, @value) " &
                                  "ON CONFLICT(SettingKey) DO UPDATE SET SettingValue = excluded.SettingValue"
                cmd.Parameters.AddWithValue("@key", If(key, String.Empty).Trim())
                cmd.Parameters.AddWithValue("@value", If(value, String.Empty).Trim())
                cmd.ExecuteNonQuery()
            End Using
        End Sub

        Public Function TryApplyBudgetTimelineSettingsFromDatabase(databasePath As String, settings As BudgetWorkspaceSettings) As Boolean
            If String.IsNullOrWhiteSpace(databasePath) OrElse Not IO.File.Exists(databasePath) Then
                Return False
            End If

            Dim applied As Boolean = False
            Try
                Using conn As New SqliteConnection("Data Source=" & databasePath.Trim() & ";Mode=ReadOnly")
                    conn.Open()

                    Using cmd = conn.CreateCommand()
                        cmd.CommandText = "SELECT name FROM sqlite_master WHERE type='table' AND name='workspaceSettings'"
                        Dim tableExists = cmd.ExecuteScalar() IsNot Nothing
                        If Not tableExists Then
                            Return False
                        End If
                    End Using

                    Using cmd = conn.CreateCommand()
                        cmd.CommandText = "SELECT SettingKey, SettingValue FROM workspaceSettings WHERE SettingKey IN ('BudgetPeriod','BudgetStartDate','BudgetYears')"
                        Using reader = cmd.ExecuteReader()
                            While reader.Read()
                                Dim key = If(reader.IsDBNull(0), String.Empty, Convert.ToString(reader.GetValue(0), CultureInfo.InvariantCulture))
                                Dim value = If(reader.IsDBNull(1), String.Empty, Convert.ToString(reader.GetValue(1), CultureInfo.InvariantCulture))
                                Select Case key
                                    Case "BudgetPeriod"
                                        If Not String.IsNullOrWhiteSpace(value) Then
                                            settings.BudgetPeriod = value.Trim()
                                            applied = True
                                        End If
                                    Case "BudgetStartDate"
                                        settings.BudgetStartDate = value.Trim()
                                        applied = True
                                    Case "BudgetYears"
                                        Dim years As Integer
                                        If Integer.TryParse(value, years) Then
                                            settings.BudgetYears = Math.Max(1, years)
                                            applied = True
                                        End If
                                End Select
                            End While
                        End Using
                    End Using
                End Using
            Catch
                Return False
            End Try
            Return applied
        End Function

        Private Function HasBudgetDateOverridesColumn(conn As SqliteConnection, columnName As String) As Boolean
            Dim cacheKey = conn.DataSource & "|" & columnName
            SyncLock BudgetOverrideItemLabelColumnPresence
                Dim cached As Boolean
                If BudgetOverrideItemLabelColumnPresence.TryGetValue(cacheKey, cached) Then
                    Return cached
                End If
            End SyncLock

            Using cmd = conn.CreateCommand()
                cmd.CommandText = "PRAGMA table_info(budgetDateOverrides)"
                Using reader = cmd.ExecuteReader()
                    While reader.Read()
                        Dim currentName = If(reader.IsDBNull(1), String.Empty, Convert.ToString(reader.GetValue(1)))
                        If String.Equals(currentName, columnName, StringComparison.OrdinalIgnoreCase) Then
                            SyncLock BudgetOverrideItemLabelColumnPresence
                                BudgetOverrideItemLabelColumnPresence(cacheKey) = True
                            End SyncLock
                            Return True
                        End If
                    End While
                End Using
            End Using

            SyncLock BudgetOverrideItemLabelColumnPresence
                BudgetOverrideItemLabelColumnPresence(cacheKey) = False
            End SyncLock
            Return False
        End Function

        Private Function GetSectionNameForCategory(catIdx As Integer) As String
            Select Case catIdx
                Case 0
                    Return "Income"
                Case 1
                    Return "Debts"
                Case 2
                    Return "Expenses"
                Case 3
                    Return "Savings"
                Case Else
                    Return String.Empty
            End Select
        End Function

        Private Function NormalizeGroupName(value As String, fallbackValue As String) As String
            Dim result = If(value, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(result) Then Return fallbackValue
            Return result
        End Function

        Private Function InferManualGroupName(catIdx As Integer) As String
            Select Case catIdx
                Case 0
                    Return "Income"
                Case 1
                    Return "Other"
                Case 2
                    Return "Uncategorized"
                Case 3
                    Return "Savings"
                Case Else
                    Return "Manual"
            End Select
        End Function

        Private Function BuildIncomeStatusText(item As IncomeRecord) As String
            Dim parts As New List(Of String)
            If Not String.IsNullOrWhiteSpace(item.Cadence) Then parts.Add(item.Cadence.Trim())
            If Not String.IsNullOrWhiteSpace(item.ToAccount) Then parts.Add("To: " & item.ToAccount.Trim())
            If Not String.IsNullOrWhiteSpace(item.StartDate) Then parts.Add("Start: " & item.StartDate.Trim())
            If Not String.IsNullOrWhiteSpace(item.EndDate) Then parts.Add("End: " & item.EndDate.Trim())
            Return String.Join(" | ", parts)
        End Function

        Private Function BuildExpenseStatusText(item As ExpenseRecord) As String
            Dim parts As New List(Of String)
            If Not String.IsNullOrWhiteSpace(item.Cadence) Then parts.Add(item.Cadence.Trim())
            If Not String.IsNullOrWhiteSpace(item.FromAccount) Then parts.Add("From: " & item.FromAccount.Trim())
            If Not String.IsNullOrWhiteSpace(item.SameAs) Then parts.Add("Same As: " & item.SameAs.Trim())
            If Not item.Active Then parts.Add("Inactive")
            If Not String.IsNullOrWhiteSpace(item.Notes) Then parts.Add(item.Notes.Trim())
            Return String.Join(" | ", parts)
        End Function

        Private Function BuildSavingsStatusText(item As SavingsRecord) As String
            Dim parts As New List(Of String)
            If Not String.IsNullOrWhiteSpace(item.Frequency) Then parts.Add(item.Frequency.Trim())
            If Not String.IsNullOrWhiteSpace(item.FromAccount) Then parts.Add("From: " & item.FromAccount.Trim())
            If Not item.Active Then parts.Add("Inactive")
            If item.HasGoal AndAlso item.GoalAmount > 0D Then parts.Add("Goal: " & item.GoalAmount.ToString("C2"))
            If item.HasGoal AndAlso Not String.IsNullOrWhiteSpace(item.GoalDate) Then parts.Add("By: " & item.GoalDate.Trim())
            Return String.Join(" | ", parts)
        End Function

        Private Function BuildDebtStatusText(item As DebtRecord) As String
            Dim parts As New List(Of String)
            If Not String.IsNullOrWhiteSpace(item.DebtType) Then parts.Add(item.DebtType.Trim())
            If Not String.IsNullOrWhiteSpace(item.Cadence) Then parts.Add(item.Cadence.Trim())
            If Not String.IsNullOrWhiteSpace(item.Lender) Then parts.Add("Lender: " & item.Lender.Trim())
            If Not String.IsNullOrWhiteSpace(item.FromAccount) Then parts.Add("From: " & item.FromAccount.Trim())
            If item.Apr > 0D Then parts.Add("APR: " & (item.Apr * 100D).ToString("0.##") & "%")
            If item.MinPayment > 0D Then parts.Add("Payment: " & item.MinPayment.ToString("C2"))
            Return String.Join(" | ", parts)
        End Function

        Private Sub UpsertSetting(lines As List(Of String), key As String, value As String)
            Dim prefix = key & "="
            For i = 0 To lines.Count - 1
                If lines(i).StartsWith(prefix, StringComparison.OrdinalIgnoreCase) Then
                    lines(i) = prefix & value
                    Return
                End If
            Next

            lines.Add(prefix & value)
        End Sub

        Private Sub RemoveSetting(lines As List(Of String), key As String)
            Dim prefix = key & "="
            For i = lines.Count - 1 To 0 Step -1
                If lines(i).StartsWith(prefix, StringComparison.OrdinalIgnoreCase) Then
                    lines.RemoveAt(i)
                End If
            Next
        End Sub

        Private Sub ConvertDeltasToRunningBalances(target As Dictionary(Of String, Decimal()))
            If target Is Nothing Then Return
            For Each key In target.Keys.ToList()
                Dim arr = target(key)
                Dim running As Decimal = 0D
                For i = 0 To arr.Length - 1
                    running += arr(i)
                    arr(i) = running
                Next
            Next
        End Sub

        Private Function NormalizeAccountKey(accountName As String) As String
            Dim key = If(accountName, String.Empty).Trim()
            If String.Equals(key, "Unassigned", StringComparison.OrdinalIgnoreCase) Then Return String.Empty
            Return key
        End Function

        Private Function TryGetSavingsSourceName(accountValue As String) As String
            Dim value = If(accountValue, String.Empty).Trim()
            Const Prefix As String = "Savings: "
            Const AlternatePrefix As String = "(Savings) "
            If value.StartsWith(Prefix, StringComparison.OrdinalIgnoreCase) Then
                Return value.Substring(Prefix.Length).Trim()
            End If

            If value.StartsWith(AlternatePrefix, StringComparison.OrdinalIgnoreCase) Then
                Return value.Substring(AlternatePrefix.Length).Trim()
            End If

            Return String.Empty
        End Function

        Private Function TryGetDebtSourceName(accountValue As String) As String
            Dim value = If(accountValue, String.Empty).Trim()
            Const Prefix As String = "(Debt) "
            If value.StartsWith(Prefix, StringComparison.OrdinalIgnoreCase) Then
                Return value.Substring(Prefix.Length).Trim()
            End If

            Return String.Empty
        End Function

        Private Function IsDebtFundingSource(item As DebtRecord) As Boolean
            If item Is Nothing Then Return False

            Dim debtType = If(item.DebtType, String.Empty).Trim()
            Dim debtCategory = If(item.Category, String.Empty).Trim()
            Return String.Equals(debtType, "Revolving Credit", StringComparison.OrdinalIgnoreCase) OrElse
                   String.Equals(debtCategory, "Credit Cards", StringComparison.OrdinalIgnoreCase)
        End Function

        Private Function BuildDebtProjection(
            periods As List(Of DateTime),
            budgetPeriod As String,
            item As DebtRecord,
            cadence As String,
            dueDay As Integer?,
            dueDate As String,
            startText As String,
            endText As String,
            paymentValues As Decimal(),
            chargeValues As Decimal(),
            manualIndexes As IEnumerable(Of Integer),
            sourceIndexes As IEnumerable(Of Integer)) As DebtProjectionResult

            Dim initialPayments = If(paymentValues, Array.Empty(Of Decimal)()).ToArray()
            Dim manualIndexSet As New HashSet(Of Integer)(If(manualIndexes, Enumerable.Empty(Of Integer)()))
            Dim sourceIndexSet As New HashSet(Of Integer)(If(sourceIndexes, Enumerable.Empty(Of Integer)()))
            For i = 0 To initialPayments.Length - 1
                If sourceIndexSet.Contains(i) AndAlso Not manualIndexSet.Contains(i) Then
                    initialPayments(i) = Math.Max(0D, initialPayments(i))
                End If
            Next

            Dim result As New DebtProjectionResult With {
                .PaymentValues = initialPayments,
                .BalanceValues = New Decimal(periods.Count - 1) {},
                .DisplayBalanceValues = New Decimal(periods.Count - 1) {}
            }
            RecalculateDebtProjectionRange(
                periods,
                budgetPeriod,
                item,
                cadence,
                dueDay,
                dueDate,
                startText,
                endText,
                result.PaymentValues,
                If(chargeValues, Array.Empty(Of Decimal)()),
                result.BalanceValues,
                result.DisplayBalanceValues,
                manualIndexes,
                sourceIndexes,
                0,
                Math.Max(0D, item.StartingBalance),
                True)
            Return result
        End Function

        Private Sub ResetDebtProjectionFromPeriod(
            periods As List(Of DateTime),
            budgetPeriod As String,
            item As DebtRecord,
            cadence As String,
            dueDay As Integer?,
            dueDate As String,
            startText As String,
            endText As String,
            projection As DebtProjectionResult,
            chargeValues As Decimal(),
            manualIndexes As IEnumerable(Of Integer),
            sourceIndexes As IEnumerable(Of Integer),
            fromPeriod As Integer,
            startingBalanceOverride As Decimal?)

            If projection Is Nothing OrElse projection.BalanceValues Is Nothing OrElse projection.PaymentValues Is Nothing Then Return
            If periods Is Nothing OrElse periods.Count = 0 OrElse fromPeriod < 0 OrElse fromPeriod >= periods.Count Then Return

            RecalculateDebtProjectionRange(
                periods,
                budgetPeriod,
                item,
                cadence,
                dueDay,
                dueDate,
                startText,
                endText,
                projection.PaymentValues,
                If(chargeValues, Array.Empty(Of Decimal)()),
                projection.BalanceValues,
                projection.DisplayBalanceValues,
                manualIndexes,
                sourceIndexes,
                fromPeriod,
                Math.Max(0D, If(startingBalanceOverride, item.StartingBalance)),
                True)
        End Sub

        Private Sub RecalculateDebtProjectionRange(
            periods As List(Of DateTime),
            budgetPeriod As String,
            item As DebtRecord,
            cadence As String,
            dueDay As Integer?,
            dueDate As String,
            startText As String,
            endText As String,
            payments As Decimal(),
            charges As Decimal(),
            balances As Decimal(),
            displayBalances As Decimal(),
            manualIndexes As IEnumerable(Of Integer),
            sourceIndexes As IEnumerable(Of Integer),
            startPeriod As Integer,
            startingBalance As Decimal,
            overwriteNonManualPayments As Boolean)

            If periods Is Nothing OrElse periods.Count = 0 OrElse payments Is Nothing OrElse balances Is Nothing OrElse displayBalances Is Nothing Then Return

            Dim itemStart = ParseDateOrDefault(startText, periods(Math.Max(0, startPeriod)))
            Dim itemEnd = ParseDateOrDefault(endText, periods(periods.Count - 1).AddYears(50))
            Dim manualIndexSet As New HashSet(Of Integer)(If(manualIndexes, Enumerable.Empty(Of Integer)()))
            Dim sourceIndexSet As New HashSet(Of Integer)(If(sourceIndexes, Enumerable.Empty(Of Integer)()))
            Dim dueDatesByPeriod = BuildDebtDueDatesByPeriod(periods, budgetPeriod, cadence, dueDay, dueDate, itemStart, itemEnd, startPeriod)
            Dim balance = Math.Max(0D, startingBalance)
            Dim graceCharges = 0D
            Dim lastAccrualDate = ResolveDebtInitialAccrualDate(item, itemStart, periods(Math.Max(0, startPeriod)))

            For i = startPeriod To periods.Count - 1
                Dim periodEnd = GetPeriodEndExclusive(periods, i, budgetPeriod).AddDays(-1).Date
                Dim eventDates As List(Of DateTime) = Nothing
                If Not dueDatesByPeriod.TryGetValue(i, eventDates) Then
                    eventDates = New List(Of DateTime)()
                End If

                Dim periodCharges = If(charges IsNot Nothing AndAlso i < charges.Length, Math.Max(0D, charges(i)), 0D)

                Dim isManual = manualIndexSet.Contains(i)
                Dim isSourceBacked = sourceIndexSet.Contains(i)
                Dim isLockedPayment = isManual OrElse isSourceBacked
                Dim lockedTotal = If(isLockedPayment AndAlso i < payments.Length, Math.Max(0D, payments(i)), 0D)
                If eventDates.Count = 0 AndAlso lockedTotal > 0D Then
                    eventDates.Add(periodEnd)
                End If

                If eventDates.Count = 0 Then
                    If graceCharges > 0D Then
                        balance += graceCharges
                        graceCharges = 0D
                    End If
                    If periodCharges > 0D Then
                        graceCharges += periodCharges
                    End If
                    Dim endingBalance = Decimal.Round(Math.Max(0D, balance + graceCharges), 2)
                    balances(i) = endingBalance
                    displayBalances(i) = endingBalance
                    Continue For
                End If

                Dim totalPeriodPayment As Decimal = 0D
                Dim prePaymentBalanceForDisplay As Decimal? = Nothing
                Dim isRevolvingCredit = String.Equals(If(item.DebtType, String.Empty).Trim(), "Revolving Credit", StringComparison.OrdinalIgnoreCase)
                For eventIndex = 0 To eventDates.Count - 1
                    Dim eventDate = eventDates(eventIndex)
                    Dim interest = CalculateDebtInterest(item, balance, eventDate, cadence, lastAccrualDate)
                    If Not prePaymentBalanceForDisplay.HasValue Then
                        Dim displayInterest = If(isRevolvingCredit, 0D, interest)
                        prePaymentBalanceForDisplay = Decimal.Round(Math.Max(0D, balance + displayInterest + graceCharges), 2)
                    End If
                    Dim scheduledPayment = CalculateDebtScheduledPayment(item, balance, interest, eventDate, cadence)
                    Dim actualPayment = If(isLockedPayment,
                        SplitManualDebtPayment(lockedTotal, eventDates.Count, eventIndex),
                        scheduledPayment)
                    Dim outcome = ApplyDebtPayment(item, balance, interest, actualPayment, eventDate, cadence, lastAccrualDate)
                    balance = outcome.NewBalance
                    Dim appliedGracePayment = 0D
                    If graceCharges > 0D Then
                        Dim graceBeforePayment = graceCharges
                        Dim remainingGracePayment = Math.Max(0D, actualPayment - outcome.PaymentOutflow)
                        If remainingGracePayment > 0D Then
                            graceCharges = Math.Max(0D, graceCharges - remainingGracePayment)
                            appliedGracePayment = Math.Min(graceBeforePayment, remainingGracePayment)
                        End If
                    End If
                    totalPeriodPayment += outcome.PaymentOutflow + appliedGracePayment
                    lastAccrualDate = eventDate
                Next

                If isLockedPayment Then
                    payments(i) = Decimal.Round(Math.Max(0D, lockedTotal), 2)
                ElseIf overwriteNonManualPayments AndAlso i < payments.Length Then
                    payments(i) = totalPeriodPayment
                End If

                If graceCharges > 0D Then
                    balance += graceCharges
                    graceCharges = 0D
                End If

                If periodCharges > 0D Then
                    graceCharges += periodCharges
                End If

                Dim periodEndingBalance = Decimal.Round(Math.Max(0D, balance + graceCharges), 2)
                balances(i) = periodEndingBalance
                displayBalances(i) = If(totalPeriodPayment > 0D AndAlso prePaymentBalanceForDisplay.HasValue,
                    prePaymentBalanceForDisplay.Value,
                    periodEndingBalance)
            Next
        End Sub

        Private Function BuildDebtDueDatesByPeriod(
            periods As List(Of DateTime),
            budgetPeriod As String,
            cadence As String,
            dueDay As Integer?,
            dueDate As String,
            itemStart As DateTime,
            itemEnd As DateTime,
            startPeriod As Integer) As Dictionary(Of Integer, List(Of DateTime))

            Dim result As New Dictionary(Of Integer, List(Of DateTime))()
            If periods Is Nothing OrElse periods.Count = 0 Then Return result

            Dim horizonStart = periods(Math.Max(0, startPeriod))
            Dim horizonEndExclusive = GetPeriodEndExclusive(periods, periods.Count - 1, budgetPeriod)

            For Each due In GenerateDueDates(cadence, If(dueDay, 1), dueDate, itemStart, itemEnd, horizonStart, horizonEndExclusive)
                Dim idx = FindPeriodIndexForDate(periods, due, budgetPeriod)
                If idx < startPeriod OrElse idx >= periods.Count Then Continue For
                If Not result.ContainsKey(idx) Then
                    result(idx) = New List(Of DateTime)()
                End If

                result(idx).Add(due.Date)
            Next

            Return result
        End Function

        Private Function ResolveDebtInitialAccrualDate(item As DebtRecord, itemStart As DateTime, horizonStart As DateTime) As DateTime
            Dim lastPayment = ParseOptionalDate(item.LastPaymentDate)
            If lastPayment.HasValue Then
                Return lastPayment.Value.Date
            End If

            If itemStart > horizonStart Then
                Return itemStart.Date
            End If

            Return horizonStart.Date
        End Function

        Private Function CalculateDebtInterest(item As DebtRecord, balance As Decimal, eventDate As DateTime, cadence As String, previousEventDate As DateTime) As Decimal
            If balance <= 0D Then Return 0D

            Dim normalizedType = If(item.DebtType, String.Empty).Trim()
            Dim rateApr = GetEffectiveDebtApr(item, eventDate)
            If rateApr <= 0D Then Return 0D

            If String.Equals(normalizedType, "Deferred-Interest Promotional Financing", StringComparison.OrdinalIgnoreCase) Then
                Dim promoEnd = ParseOptionalDate(item.PromoAprEndDate)
                If promoEnd.HasValue AndAlso eventDate.Date <= promoEnd.Value.Date Then
                    Return 0D
                End If
            End If

            If String.Equals(normalizedType, "Student Loan With Special Repayment Rules", StringComparison.OrdinalIgnoreCase) AndAlso IsStudentDebtDeferred(item, eventDate) Then
                If item.Subsidized Then
                    Return 0D
                End If
            End If

            If String.Equals(normalizedType, "True 0 Percent Installment Loan", StringComparison.OrdinalIgnoreCase) Then
                Return 0D
            End If

            If String.Equals(normalizedType, "Simple Interest Loan", StringComparison.OrdinalIgnoreCase) OrElse
               String.Equals(normalizedType, "Daily Interest Loan", StringComparison.OrdinalIgnoreCase) Then
                Dim dayBasis = If(item.DayCountBasis.HasValue AndAlso item.DayCountBasis.Value > 0, item.DayCountBasis.Value, 365)
                Dim priorDate = If(previousEventDate = Date.MinValue, eventDate.Date, previousEventDate.Date)
                Dim dayCount = Math.Max(0, CInt((eventDate.Date - priorDate).TotalDays))
                Return Decimal.Round(balance * rateApr / dayBasis * dayCount, 10)
            End If

            Dim paymentsPerYear = GetDebtPaymentsPerYear(item, cadence)
            Return Decimal.Round(balance * rateApr / paymentsPerYear, 10)
        End Function

        Private Function GetEffectiveDebtApr(item As DebtRecord, eventDate As DateTime) As Decimal
            Dim normalizedType = If(item.DebtType, String.Empty).Trim()

            If String.Equals(normalizedType, "Deferred-Interest Promotional Financing", StringComparison.OrdinalIgnoreCase) Then
                Dim promoEnd = ParseOptionalDate(item.PromoAprEndDate)
                If promoEnd.HasValue AndAlso eventDate.Date <= promoEnd.Value.Date Then
                    Return 0D
                End If

                If item.PromoApr > 0D Then
                    Return item.PromoApr
                End If
            End If

            If String.Equals(normalizedType, "Variable-Rate Amortizing Loan", StringComparison.OrdinalIgnoreCase) Then
                Dim scheduledRate = TryGetVariableDebtRate(item.RateChangeSchedule, eventDate)
                If scheduledRate.HasValue Then
                    Return scheduledRate.Value
                End If
            End If

            Return Math.Max(0D, item.Apr)
        End Function

        Private Function TryGetVariableDebtRate(scheduleText As String, eventDate As DateTime) As Decimal?
            If String.IsNullOrWhiteSpace(scheduleText) Then Return Nothing

            Dim bestDate As DateTime? = Nothing
            Dim bestRate As Decimal? = Nothing
            Dim parts = scheduleText.Split(New Char() {";"c}, StringSplitOptions.RemoveEmptyEntries)
            For Each rawPart In parts
                Dim pair = rawPart.Split(New Char() {"="c}, 2, StringSplitOptions.None)
                If pair.Length <> 2 Then Continue For

                Dim effectiveDate = ParseOptionalDate(pair(0).Trim())
                Dim rate As Decimal
                If Not effectiveDate.HasValue OrElse
                   Not Decimal.TryParse(pair(1).Trim(), NumberStyles.Number, CultureInfo.InvariantCulture, rate) Then
                    Continue For
                End If

                If effectiveDate.Value.Date <= eventDate.Date AndAlso
                   (Not bestDate.HasValue OrElse effectiveDate.Value.Date >= bestDate.Value.Date) Then
                    bestDate = effectiveDate.Value.Date
                    bestRate = rate
                End If
            Next

            Return bestRate
        End Function

        Private Function IsStudentDebtDeferred(item As DebtRecord, eventDate As DateTime) As Boolean
            If item.DeferredStatus Then Return True
            If String.Equals(item.StudentRepaymentPlan, "Deferred", StringComparison.OrdinalIgnoreCase) OrElse
               String.Equals(item.StudentRepaymentPlan, "Forbearance", StringComparison.OrdinalIgnoreCase) Then
                Return True
            End If

            Dim deferredUntil = ParseOptionalDate(item.DeferredUntil)
            Return deferredUntil.HasValue AndAlso eventDate.Date <= deferredUntil.Value.Date
        End Function

        Private Function CalculateDebtScheduledPayment(item As DebtRecord, balance As Decimal, interest As Decimal, eventDate As DateTime, cadence As String) As Decimal
            Dim normalizedType = If(item.DebtType, String.Empty).Trim()
            Dim basePayment = Math.Max(0D, item.MinPayment)

            If String.Equals(normalizedType, "Student Loan With Special Repayment Rules", StringComparison.OrdinalIgnoreCase) AndAlso IsStudentDebtDeferred(item, eventDate) Then
                basePayment = 0D
            ElseIf String.Equals(normalizedType, "Fixed-Term Amortizing Loan", StringComparison.OrdinalIgnoreCase) AndAlso basePayment <= 0D Then
                basePayment = CalculateFixedTermDebtPayment(item, cadence)
            ElseIf String.Equals(normalizedType, "Interest-Only Loan", StringComparison.OrdinalIgnoreCase) Then
                basePayment = Math.Max(basePayment, Decimal.Round(interest, 2))
            ElseIf String.Equals(normalizedType, "True 0 Percent Installment Loan", StringComparison.OrdinalIgnoreCase) AndAlso basePayment <= 0D Then
                basePayment = Decimal.Round(Math.Max(0D, balance), 2)
            End If

            Dim payment = basePayment

            If String.Equals(normalizedType, "Balloon Loan", StringComparison.OrdinalIgnoreCase) Then
                Dim balloonDue = ParseOptionalDate(item.BalloonDueDate)
                If balloonDue.HasValue AndAlso eventDate.Date >= balloonDue.Value.Date Then
                    payment += Math.Max(0D, item.BalloonAmount)
                End If
            End If

            Return Math.Max(0D, payment)
        End Function

        Private Function CalculateFixedTermDebtPayment(item As DebtRecord, cadence As String) As Decimal
            Dim principal = Math.Max(0D, If(item.OriginalPrincipal > 0D, item.OriginalPrincipal, item.StartingBalance))
            Dim months = If(item.TermMonths.HasValue AndAlso item.TermMonths.Value > 0, item.TermMonths.Value, 0)
            If principal <= 0D OrElse months <= 0 Then Return 0D

            Dim paymentsPerYear = GetDebtPaymentsPerYear(item, cadence)
            Dim totalPayments = Math.Max(1, CInt(Math.Round(months * (paymentsPerYear / 12D), MidpointRounding.AwayFromZero)))
            Dim periodicRate = If(item.Apr <= 0D, 0D, item.Apr / paymentsPerYear)
            If periodicRate <= 0D Then
                Return Decimal.Round(principal / totalPayments, 2)
            End If

            Dim denominator = 1D - CDec(Math.Pow(CDbl(1D + periodicRate), -totalPayments))
            If denominator <= 0D Then
                Return 0D
            End If

            Return Decimal.Round(principal * periodicRate / denominator, 2)
        End Function

        Private Function SplitManualDebtPayment(total As Decimal, eventCount As Integer, eventIndex As Integer) As Decimal
            If total <= 0D OrElse eventCount <= 0 Then Return 0D
            If eventIndex >= eventCount - 1 Then
                Return total - (Decimal.Round(total / eventCount, 2) * (eventCount - 1))
            End If

            Return Decimal.Round(total / eventCount, 2)
        End Function

        Private Structure DebtPaymentOutcome
            Public Property PaymentOutflow As Decimal
            Public Property NewBalance As Decimal
        End Structure

        Private Function ApplyDebtPayment(item As DebtRecord, balance As Decimal, interest As Decimal, requestedPayment As Decimal, eventDate As DateTime, cadence As String, previousEventDate As DateTime) As DebtPaymentOutcome
            Dim normalizedType = If(item.DebtType, String.Empty).Trim()
            Dim balanceBeforePayment = Math.Max(0D, balance)
            Dim includedEscrow = If(item.EscrowIncluded, Math.Max(0D, item.EscrowMonthly), 0D)
            Dim pmiCharge = Math.Max(0D, item.PmiMonthly)
            Dim requestedOutflow = Math.Max(0D, requestedPayment)
            Dim paymentTowardLoan = Math.Max(0D, requestedOutflow - includedEscrow)

            If String.Equals(normalizedType, "Student Loan With Special Repayment Rules", StringComparison.OrdinalIgnoreCase) Then
                Dim forgivenessDate = ParseOptionalDate(item.ForgivenessDate)
                If forgivenessDate.HasValue AndAlso eventDate.Date >= forgivenessDate.Value.Date Then
                    Return New DebtPaymentOutcome With {
                        .PaymentOutflow = 0D,
                        .NewBalance = 0D
                    }
                End If
            End If

            If String.Equals(normalizedType, "Revolving Credit", StringComparison.OrdinalIgnoreCase) Then
                Dim appliedPayment = Math.Min(paymentTowardLoan, balanceBeforePayment)
                Dim remainingBalance = Math.Max(0D, balanceBeforePayment - appliedPayment)
                Dim remainingInterest = CalculateDebtInterest(item, remainingBalance, eventDate, cadence, previousEventDate)
                Return New DebtPaymentOutcome With {
                    .PaymentOutflow = Decimal.Round(Math.Min(requestedOutflow + pmiCharge, balanceBeforePayment + includedEscrow + pmiCharge), 2),
                    .NewBalance = Decimal.Round(Math.Max(0D, remainingBalance + remainingInterest), 2)
                }
            End If

            Dim balanceWithInterest = balanceBeforePayment + interest
            Dim appliedLoanPayment = paymentTowardLoan
            Dim newBalance As Decimal
            Dim cappedLoanPayment As Decimal

            If String.Equals(normalizedType, "Negative Amortization Loan", StringComparison.OrdinalIgnoreCase) Then
                cappedLoanPayment = appliedLoanPayment
                newBalance = balanceWithInterest - appliedLoanPayment
            Else
                cappedLoanPayment = Math.Min(appliedLoanPayment, balanceWithInterest)
                newBalance = balanceWithInterest - cappedLoanPayment
            End If

            Return New DebtPaymentOutcome With {
                .PaymentOutflow = Decimal.Round(cappedLoanPayment + includedEscrow + pmiCharge, 2),
                .NewBalance = Decimal.Round(Math.Max(0D, newBalance), 2)
            }
        End Function

        Private Function GetDebtPaymentsPerYear(item As DebtRecord, cadence As String) As Integer
            If item.PaymentsPerYear.HasValue AndAlso item.PaymentsPerYear.Value > 0 Then
                Return item.PaymentsPerYear.Value
            End If

            Select Case If(cadence, String.Empty).Trim()
                Case "Weekly"
                    Return 52
                Case "Bi-Weekly"
                    Return 26
                Case "Yearly on Date", "Due Yearly", "Yearly"
                    Return 1
                Case Else
                    Return 12
            End Select
        End Function

        Private Function ParseOptionalDate(value As String) As DateTime?
            Dim parsed As DateTime
            If DateTime.TryParse(value, CultureInfo.CurrentCulture, DateTimeStyles.None, parsed) Then Return parsed.Date
            If DateTime.TryParseExact(value, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, parsed) Then Return parsed.Date
            Return Nothing
        End Function

        Private Function ResolveScheduleStartText(primaryText As String, inheritedText As String) As String
            Dim primaryDate = ParseOptionalDate(primaryText)
            Dim inheritedDate = ParseOptionalDate(inheritedText)

            If primaryDate.HasValue AndAlso inheritedDate.HasValue Then
                Return If(primaryDate.Value.Date >= inheritedDate.Value.Date, primaryDate.Value, inheritedDate.Value).ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
            End If

            If primaryDate.HasValue Then Return primaryDate.Value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
            If inheritedDate.HasValue Then Return inheritedDate.Value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
            Return If(String.IsNullOrWhiteSpace(primaryText), inheritedText, primaryText)
        End Function

        Private Function ResolveScheduleEndText(primaryText As String, inheritedText As String) As String
            Dim primaryDate = ParseOptionalDate(primaryText)
            Dim inheritedDate = ParseOptionalDate(inheritedText)

            If primaryDate.HasValue AndAlso inheritedDate.HasValue Then
                Return If(primaryDate.Value.Date <= inheritedDate.Value.Date, primaryDate.Value, inheritedDate.Value).ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
            End If

            If primaryDate.HasValue Then Return primaryDate.Value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
            If inheritedDate.HasValue Then Return inheritedDate.Value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
            Return If(String.IsNullOrWhiteSpace(primaryText), inheritedText, primaryText)
        End Function

        Private Sub AddScheduledAmount(
            summaries As List(Of BudgetWorkspacePeriodSummary),
            periods As List(Of DateTime),
            budgetPeriod As String,
            amount As Decimal,
            cadence As String,
            dayNumber As Integer?,
            dateText As String,
            startText As String,
            endText As String,
            applyAmount As Action(Of BudgetWorkspacePeriodSummary, Decimal))

            If summaries Is Nothing OrElse periods Is Nothing OrElse summaries.Count = 0 Then Return
            If amount <= 0D AndAlso Not String.Equals(cadence, "Manually Entered", StringComparison.OrdinalIgnoreCase) Then Return

            Dim effectiveCadence = If(String.IsNullOrWhiteSpace(cadence), String.Empty, cadence.Trim())
            If String.Equals(effectiveCadence, "Manually Entered", StringComparison.OrdinalIgnoreCase) Then Return

            Dim itemStart = ParseDateOrDefault(startText, periods(0))
            Dim itemEnd = ParseDateOrDefault(endText, periods(periods.Count - 1).AddYears(50))

            If String.Equals(effectiveCadence, "Per Month", StringComparison.OrdinalIgnoreCase) OrElse
               String.Equals(effectiveCadence, "Per Year", StringComparison.OrdinalIgnoreCase) Then
                Dim perPeriod = GetDistributedPerPeriodAmount(amount, effectiveCadence, budgetPeriod)
                For i = 0 To summaries.Count - 1
                    Dim pStart = periods(i)
                    If pStart.Date < itemStart.Date OrElse pStart.Date > itemEnd.Date Then Continue For
                    applyAmount(summaries(i), perPeriod)
                Next
                Return
            End If

            For Each dueDate In GenerateDueDates(effectiveCadence, If(dayNumber, 1), dateText, itemStart, itemEnd, periods(0), GetPeriodEndExclusive(periods, periods.Count - 1, budgetPeriod))
                Dim idx = FindPeriodIndexForDate(periods, dueDate, budgetPeriod)
                If idx >= 0 AndAlso idx < summaries.Count Then
                    applyAmount(summaries(idx), amount)
                End If
            Next
        End Sub

        Private Function GetIncomeAmountForForecastDate(item As IncomeRecord, targetDate As DateTime, currentBaselineDate As DateTime) As Decimal
            If item Is Nothing Then
                Return 0D
            End If

            Dim baseAmount = item.Amount
            If item.AutoIncrease = 0D Then
                Return baseAmount
            End If

            Dim startDate = ParseDateOrDefault(item.StartDate, targetDate)
            Dim effectiveStartDate = startDate.Date
            If currentBaselineDate.Date > effectiveStartDate AndAlso targetDate.Date >= currentBaselineDate.Date Then
                effectiveStartDate = currentBaselineDate.Date
            End If

            If targetDate.Date < effectiveStartDate Then
                Return baseAmount
            End If

            Dim monthValue = effectiveStartDate.Month
            Dim dayValue = effectiveStartDate.Day
            If Not ParseMonthDay(item.AutoIncreaseOnDate, monthValue, dayValue) Then
                monthValue = effectiveStartDate.Month
                dayValue = effectiveStartDate.Day
            End If

            Dim increaseCount As Integer = 0
            For year = effectiveStartDate.Year + 1 To targetDate.Year
                Dim anniversaryDay = Math.Min(dayValue, DateTime.DaysInMonth(year, monthValue))
                Dim increaseDate As New DateTime(year, monthValue, anniversaryDay)
                If increaseDate <= targetDate.Date Then
                    increaseCount += 1
                End If
            Next

            If increaseCount <= 0 Then
                Return baseAmount
            End If

            Dim factor = CDec(Math.Pow(CDbl(1D + item.AutoIncrease), increaseCount))
            Return Decimal.Round(baseAmount * factor, 2, MidpointRounding.AwayFromZero)
        End Function

        Private Function GeneratePeriods(startDate As DateTime, endExclusive As DateTime, budgetPeriod As String) As List(Of DateTime)
            Dim results As New List(Of DateTime)()
            Dim current = startDate.Date
            While current < endExclusive.Date
                results.Add(current)
                Select Case budgetPeriod
                    Case "Weekly"
                        current = current.AddDays(7)
                    Case "Bi-Weekly"
                        current = current.AddDays(14)
                    Case Else
                        current = current.AddMonths(1)
                End Select
            End While
            Return results
        End Function

        Private Function ParseDateOrDefault(value As String, fallbackValue As DateTime) As DateTime
            Dim parsed As DateTime
            If DateTime.TryParse(value, CultureInfo.CurrentCulture, DateTimeStyles.None, parsed) Then Return parsed.Date
            If DateTime.TryParseExact(value, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, parsed) Then Return parsed.Date
            Return fallbackValue.Date
        End Function

        Private Function ParseMonthDay(textValue As String, ByRef monthValue As Integer, ByRef dayValue As Integer) As Boolean
            monthValue = 0
            dayValue = 0
            If String.IsNullOrWhiteSpace(textValue) Then
                Return False
            End If

            Dim parsed As DateTime
            If DateTime.TryParseExact(textValue.Trim(), New String() {"M/d", "MM/dd", "M/d/yyyy", "MM/dd/yyyy"}, CultureInfo.CurrentCulture, DateTimeStyles.None, parsed) Then
                monthValue = parsed.Month
                dayValue = parsed.Day
                Return True
            End If

            If DateTime.TryParse(textValue.Trim(), CultureInfo.CurrentCulture, DateTimeStyles.None, parsed) Then
                monthValue = parsed.Month
                dayValue = parsed.Day
                Return True
            End If

            Return False
        End Function

        Private Function GetPeriodEndExclusive(periods As List(Of DateTime), index As Integer, budgetPeriod As String) As DateTime
            If index >= 0 AndAlso index < periods.Count - 1 Then Return periods(index + 1)
            Select Case budgetPeriod
                Case "Weekly"
                    Return periods(Math.Max(0, index)).AddDays(7)
                Case "Bi-Weekly"
                    Return periods(Math.Max(0, index)).AddDays(14)
                Case Else
                    Return periods(Math.Max(0, index)).AddMonths(1)
            End Select
        End Function

        Private Function GetDistributedPerPeriodAmount(amount As Decimal, cadence As String, budgetPeriod As String) As Decimal
            Dim periodsPerYear As Decimal
            Select Case budgetPeriod
                Case "Weekly"
                    periodsPerYear = 52D
                Case "Bi-Weekly"
                    periodsPerYear = 26D
                Case Else
                    periodsPerYear = 12D
            End Select

            If String.Equals(cadence, "Per Month", StringComparison.OrdinalIgnoreCase) Then
                Return Decimal.Round((amount * 12D) / periodsPerYear, 2)
            End If

            Return Decimal.Round(amount / periodsPerYear, 2)
        End Function

        Private Iterator Function GenerateDueDates(cadence As String, dayNumber As Integer, dateText As String, itemStart As DateTime, itemEnd As DateTime, horizonStart As DateTime, horizonEndExclusive As DateTime) As IEnumerable(Of DateTime)
            Dim startDate = If(itemStart > horizonStart, itemStart.Date, horizonStart.Date)
            Dim endDate = If(itemEnd < horizonEndExclusive, itemEnd.Date, horizonEndExclusive.AddDays(-1).Date)
            If endDate < startDate Then Return

            Select Case cadence
                Case "Monthly On", "Due Monthly", "Monthly", "Monthly on Day"
                    Dim current = New DateTime(startDate.Year, startDate.Month, Math.Min(Math.Max(1, dayNumber), DateTime.DaysInMonth(startDate.Year, startDate.Month)))
                    If current < startDate Then current = current.AddMonths(1)
                    While current <= endDate
                        Yield current
                        current = New DateTime(current.AddMonths(1).Year, current.AddMonths(1).Month, Math.Min(Math.Max(1, dayNumber), DateTime.DaysInMonth(current.AddMonths(1).Year, current.AddMonths(1).Month)))
                    End While

                Case "Yearly On", "Due Yearly", "Yearly", "Yearly on Date"
                    Dim monthValue = 1
                    Dim dayValue = 1
                    Dim parsed As DateTime
                    If DateTime.TryParse(dateText, CultureInfo.CurrentCulture, DateTimeStyles.None, parsed) Then
                        monthValue = parsed.Month
                        dayValue = parsed.Day
                    ElseIf DateTime.TryParseExact(dateText, {"MM/dd", "M/d", "yyyy-MM-dd"}, CultureInfo.InvariantCulture, DateTimeStyles.None, parsed) Then
                        monthValue = parsed.Month
                        dayValue = parsed.Day
                    End If
                    Dim year = startDate.Year
                    Dim current = New DateTime(year, monthValue, Math.Min(dayValue, DateTime.DaysInMonth(year, monthValue)))
                    If current < startDate Then
                        year += 1
                        current = New DateTime(year, monthValue, Math.Min(dayValue, DateTime.DaysInMonth(year, monthValue)))
                    End If
                    While current <= endDate
                        Yield current
                        year += 1
                        current = New DateTime(year, monthValue, Math.Min(dayValue, DateTime.DaysInMonth(year, monthValue)))
                    End While

                Case "Weekly"
                    Dim current = startDate
                    While current <= endDate
                        Yield current
                        current = current.AddDays(7)
                    End While

                Case "Bi-Weekly"
                    Dim current = startDate
                    While current <= endDate
                        Yield current
                        current = current.AddDays(14)
                    End While

                Case Else
                    Yield startDate
            End Select
        End Function

        Private Function FindPeriodIndexForDate(periods As List(Of DateTime), targetDate As DateTime, budgetPeriod As String) As Integer
            If periods Is Nothing OrElse periods.Count = 0 Then
                Return -1
            End If

            Dim normalizedTarget = targetDate.Date
            Dim firstPeriod = periods(0).Date
            If normalizedTarget < firstPeriod Then
                Return -1
            End If

            Dim lastEndExclusive = GetPeriodEndExclusive(periods, periods.Count - 1, budgetPeriod).Date
            If normalizedTarget >= lastEndExclusive Then
                Return -1
            End If

            Dim fastIndex = TryGetFastPeriodIndex(periods, normalizedTarget, budgetPeriod)
            If fastIndex >= 0 Then
                Return fastIndex
            End If

            For i = 0 To periods.Count - 1
                Dim startDate = periods(i).Date
                Dim endExclusive = GetPeriodEndExclusive(periods, i, budgetPeriod).Date
                If normalizedTarget >= startDate AndAlso normalizedTarget < endExclusive Then
                    Return i
                End If
            Next
            Return -1
        End Function

        Private Function FindCurrentPeriodIndex(periods As List(Of DateTime), budgetPeriod As String) As Integer
            If periods Is Nothing OrElse periods.Count = 0 Then
                Return 0
            End If

            Dim today = DateTime.Today

            Dim fastIndex = TryGetFastPeriodIndex(periods, today, budgetPeriod)
            If fastIndex >= 0 Then
                Return fastIndex
            End If

            For i = 0 To periods.Count - 1
                Dim startDate = periods(i).Date
                Dim endExclusive = GetPeriodEndExclusive(periods, i, budgetPeriod).Date
                If today >= startDate AndAlso today < endExclusive Then
                    Return i
                End If
            Next

            Return Math.Max(0, periods.Count - 1)
        End Function

        Private Function TryGetFastPeriodIndex(periods As List(Of DateTime), targetDate As DateTime, budgetPeriod As String) As Integer
            If periods Is Nothing OrElse periods.Count = 0 Then
                Return -1
            End If

            Dim firstPeriod = periods(0).Date
            Dim candidateIndex As Integer = -1

            Select Case If(budgetPeriod, String.Empty).Trim()
                Case "Weekly"
                    candidateIndex = CInt(Math.Floor((targetDate.Date - firstPeriod).TotalDays / 7D))

                Case "Bi-Weekly"
                    candidateIndex = CInt(Math.Floor((targetDate.Date - firstPeriod).TotalDays / 14D))

                Case Else
                    candidateIndex = ((targetDate.Year - firstPeriod.Year) * 12) + (targetDate.Month - firstPeriod.Month)
            End Select

            If candidateIndex < 0 OrElse candidateIndex >= periods.Count Then
                Return -1
            End If

            Dim candidateStart = periods(candidateIndex).Date
            Dim candidateEndExclusive = GetPeriodEndExclusive(periods, candidateIndex, budgetPeriod).Date
            If targetDate.Date >= candidateStart AndAlso targetDate.Date < candidateEndExclusive Then
                Return candidateIndex
            End If

            If candidateIndex + 1 < periods.Count Then
                Dim nextStart = periods(candidateIndex + 1).Date
                Dim nextEndExclusive = GetPeriodEndExclusive(periods, candidateIndex + 1, budgetPeriod).Date
                If targetDate.Date >= nextStart AndAlso targetDate.Date < nextEndExclusive Then
                    Return candidateIndex + 1
                End If
            End If

            If candidateIndex - 1 >= 0 Then
                Dim previousStart = periods(candidateIndex - 1).Date
                Dim previousEndExclusive = GetPeriodEndExclusive(periods, candidateIndex - 1, budgetPeriod).Date
                If targetDate.Date >= previousStart AndAlso targetDate.Date < previousEndExclusive Then
                    Return candidateIndex - 1
                End If
            End If

            Return -1
        End Function

    End Module

End Namespace

