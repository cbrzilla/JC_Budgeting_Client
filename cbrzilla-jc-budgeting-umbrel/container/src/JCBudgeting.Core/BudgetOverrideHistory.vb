Imports System.Globalization
Imports Microsoft.Data.Sqlite

Namespace Global.JCBudgeting.Core

    Public Class BudgetOverrideHistoryRecord
        Public Property DueDate As DateTime
        Public Property CategoryIndex As Integer
        Public Property ItemKey As String = String.Empty
        Public Property ItemDescription As String = String.Empty
        Public Property Amount As Decimal
        Public Property Additional As Decimal
        Public Property ManualAmount As Decimal
        Public Property Paid As Boolean
        Public Property Notes As String = String.Empty
        Public Property SelectionMode As Integer
        Public Property FromAccountSnapshot As String = String.Empty
        Public Property ToAccountSnapshot As String = String.Empty
        Public Property SameAsSnapshot As String = String.Empty
        Public Property FromAccountSnapshotId As Integer?
        Public Property ToAccountSnapshotId As Integer?
        Public Property SameAsSnapshotId As Integer?
        Public Property FromSavingsSnapshotId As Integer?
        Public Property FromDebtSnapshotId As Integer?

        Public ReadOnly Property IsManual As Boolean
            Get
                Return SelectionMode = 3
            End Get
        End Property
    End Class

    Public Module BudgetOverrideHistoryRepository

        Public Function CaptureLiveOverrideHistory(databasePath As String, categoryIndex As Integer, itemKey As String, itemDescription As String) As List(Of BudgetOverrideHistoryRecord)
            ValidateHistoryArguments(databasePath, categoryIndex, itemKey, itemDescription)

            Dim results As New List(Of BudgetOverrideHistoryRecord)()

            Using conn As New SqliteConnection("Data Source=" & databasePath.Trim())
                conn.Open()
                EnsureBudgetOverrideArchiveTable(conn)

                If Not HasTable(conn, "budgetDateOverrides") Then
                    Return results
                End If

                results = LoadLiveOverrideRecords(conn, categoryIndex, itemKey, itemDescription)
                ArchiveHistoryRecords(conn, results)
            End Using

            Return results
        End Function

        Public Function LoadOverrideHistory(databasePath As String, categoryIndex As Integer, itemKey As String, itemDescription As String) As List(Of BudgetOverrideHistoryRecord)
            Return CaptureLiveOverrideHistory(databasePath, categoryIndex, itemKey, itemDescription)
        End Function

        Public Function HasOverrideHistoryBefore(databasePath As String, categoryIndex As Integer, itemKey As String, itemDescription As String, periodStart As DateTime) As Boolean
            ValidateHistoryArguments(databasePath, categoryIndex, itemKey, itemDescription)

            Using conn As New SqliteConnection("Data Source=" & databasePath.Trim())
                conn.Open()

                If Not HasTable(conn, "budgetDateOverrides") Then
                    Return False
                End If

                Using cmd = conn.CreateCommand()
                    cmd.CommandText = "SELECT 1 FROM budgetDateOverrides WHERE CatIdx=@c AND (ItemDesc=@n OR ItemDesc=@legacy) AND DueDate < @d LIMIT 1"
                    cmd.Parameters.AddWithValue("@c", categoryIndex)
                    cmd.Parameters.AddWithValue("@n", itemKey.Trim())
                    cmd.Parameters.AddWithValue("@legacy", itemDescription.Trim())
                    cmd.Parameters.AddWithValue("@d", periodStart.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture))
                    Dim result = cmd.ExecuteScalar()
                    Return result IsNot Nothing AndAlso result IsNot DBNull.Value
                End Using
            End Using
        End Function

        Public Sub DeleteLiveOverrides(databasePath As String, categoryIndex As Integer, itemKey As String, itemDescription As String)
            ValidateHistoryArguments(databasePath, categoryIndex, itemKey, itemDescription)

            Using conn As New SqliteConnection("Data Source=" & databasePath.Trim())
                conn.Open()

                If Not HasTable(conn, "budgetDateOverrides") Then
                    Return
                End If

                Using cmd = conn.CreateCommand()
                    cmd.CommandText = "DELETE FROM budgetDateOverrides WHERE CatIdx=@c AND (ItemDesc=@n OR ItemDesc=@legacy)"
                    cmd.Parameters.AddWithValue("@c", categoryIndex)
                    cmd.Parameters.AddWithValue("@n", itemKey.Trim())
                    cmd.Parameters.AddWithValue("@legacy", itemDescription.Trim())
                    cmd.ExecuteNonQuery()
                End Using
            End Using
        End Sub

        Public Sub DeleteOverrideHistory(databasePath As String, categoryIndex As Integer, itemKey As String, itemDescription As String)
            Dim existing = CaptureLiveOverrideHistory(databasePath, categoryIndex, itemKey, itemDescription)
            If existing.Count = 0 Then
                Return
            End If

            DeleteLiveOverrides(databasePath, categoryIndex, itemKey, itemDescription)
        End Sub

        Public Sub RestoreOverrideHistory(databasePath As String, records As IEnumerable(Of BudgetOverrideHistoryRecord))
            If String.IsNullOrWhiteSpace(databasePath) Then
                Throw New ArgumentException("A database path is required.", NameOf(databasePath))
            End If

            If records Is Nothing Then
                Return
            End If

            Using conn As New SqliteConnection("Data Source=" & databasePath.Trim())
                conn.Open()
                EnsureLiveBudgetOverridesTable(conn)
                EnsureBudgetOverrideArchiveTable(conn)

                Dim normalizedRecords = records.
                    Where(Function(x) x IsNot Nothing AndAlso x.CategoryIndex >= 0 AndAlso Not String.IsNullOrWhiteSpace(x.ItemKey) AndAlso Not String.IsNullOrWhiteSpace(x.ItemDescription)).
                    ToList()

                For Each record In normalizedRecords
                    Using cmd = conn.CreateCommand()
                        cmd.CommandText = "INSERT OR REPLACE INTO budgetDateOverrides (DueDate, CatIdx, ItemDesc, ItemLabel, Amount, Additional, ManualAmount, Paid, Notes, SelectionMode, FromAccountSnapshot, ToAccountSnapshot, SameAsSnapshot, FromAccountSnapshotId, ToAccountSnapshotId, SameAsSnapshotId, FromSavingsSnapshotId, FromDebtSnapshotId) VALUES (@d, @c, @n, @l, @a, @additional, @manualAmount, @paid, @notes, @selectionMode, @fromAccountSnapshot, @toAccountSnapshot, @sameAsSnapshot, @fromAccountSnapshotId, @toAccountSnapshotId, @sameAsSnapshotId, @fromSavingsSnapshotId, @fromDebtSnapshotId)"
                        cmd.Parameters.AddWithValue("@d", record.DueDate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture))
                        cmd.Parameters.AddWithValue("@c", record.CategoryIndex)
                        cmd.Parameters.AddWithValue("@n", record.ItemKey.Trim())
                        cmd.Parameters.AddWithValue("@l", record.ItemDescription.Trim())
                        cmd.Parameters.AddWithValue("@a", record.Amount)
                        cmd.Parameters.AddWithValue("@additional", record.Additional)
                        cmd.Parameters.AddWithValue("@manualAmount", record.ManualAmount)
                        cmd.Parameters.AddWithValue("@paid", If(record.Paid, 1, 0))
                        cmd.Parameters.AddWithValue("@notes", If(record.Notes, String.Empty).Trim())
                        cmd.Parameters.AddWithValue("@selectionMode", record.SelectionMode)
                        cmd.Parameters.AddWithValue("@fromAccountSnapshot", If(record.FromAccountSnapshot, String.Empty).Trim())
                        cmd.Parameters.AddWithValue("@toAccountSnapshot", If(record.ToAccountSnapshot, String.Empty).Trim())
                        cmd.Parameters.AddWithValue("@sameAsSnapshot", If(record.SameAsSnapshot, String.Empty).Trim())
                        cmd.Parameters.AddWithValue("@fromAccountSnapshotId", If(record.FromAccountSnapshotId.HasValue, CType(record.FromAccountSnapshotId.Value, Object), DBNull.Value))
                        cmd.Parameters.AddWithValue("@toAccountSnapshotId", If(record.ToAccountSnapshotId.HasValue, CType(record.ToAccountSnapshotId.Value, Object), DBNull.Value))
                        cmd.Parameters.AddWithValue("@sameAsSnapshotId", If(record.SameAsSnapshotId.HasValue, CType(record.SameAsSnapshotId.Value, Object), DBNull.Value))
                        cmd.Parameters.AddWithValue("@fromSavingsSnapshotId", If(record.FromSavingsSnapshotId.HasValue, CType(record.FromSavingsSnapshotId.Value, Object), DBNull.Value))
                        cmd.Parameters.AddWithValue("@fromDebtSnapshotId", If(record.FromDebtSnapshotId.HasValue, CType(record.FromDebtSnapshotId.Value, Object), DBNull.Value))
                        cmd.ExecuteNonQuery()
                    End Using
                Next

                ArchiveHistoryRecords(conn, normalizedRecords)
            End Using
        End Sub

        Private Sub ValidateHistoryArguments(databasePath As String, categoryIndex As Integer, itemKey As String, itemDescription As String)
            If String.IsNullOrWhiteSpace(databasePath) Then
                Throw New ArgumentException("A database path is required.", NameOf(databasePath))
            End If

            If categoryIndex < 0 OrElse String.IsNullOrWhiteSpace(itemKey) OrElse String.IsNullOrWhiteSpace(itemDescription) Then
                Throw New ArgumentException("A valid category index, item key, and item description are required.")
            End If
        End Sub

        Private Function LoadLiveOverrideRecords(conn As SqliteConnection, categoryIndex As Integer, itemKey As String, itemDescription As String) As List(Of BudgetOverrideHistoryRecord)
            Dim results As New List(Of BudgetOverrideHistoryRecord)()
            Dim hasItemLabel = HasColumn(conn, "budgetDateOverrides", "ItemLabel")
            Dim hasAdditional = HasColumn(conn, "budgetDateOverrides", "Additional")
            Dim hasManualAmount = HasColumn(conn, "budgetDateOverrides", "ManualAmount")
            Dim hasPaid = HasColumn(conn, "budgetDateOverrides", "Paid")
            Dim hasNotes = HasColumn(conn, "budgetDateOverrides", "Notes")
            Dim hasSelectionMode = HasColumn(conn, "budgetDateOverrides", "SelectionMode")
            Dim hasFromAccountSnapshot = HasColumn(conn, "budgetDateOverrides", "FromAccountSnapshot")
            Dim hasToAccountSnapshot = HasColumn(conn, "budgetDateOverrides", "ToAccountSnapshot")
            Dim hasSameAsSnapshot = HasColumn(conn, "budgetDateOverrides", "SameAsSnapshot")
            Dim hasFromAccountSnapshotId = HasColumn(conn, "budgetDateOverrides", "FromAccountSnapshotId")
            Dim hasToAccountSnapshotId = HasColumn(conn, "budgetDateOverrides", "ToAccountSnapshotId")
            Dim hasSameAsSnapshotId = HasColumn(conn, "budgetDateOverrides", "SameAsSnapshotId")
            Dim hasFromSavingsSnapshotId = HasColumn(conn, "budgetDateOverrides", "FromSavingsSnapshotId")
            Dim hasFromDebtSnapshotId = HasColumn(conn, "budgetDateOverrides", "FromDebtSnapshotId")

            Using cmd = conn.CreateCommand()
                cmd.CommandText =
                    "SELECT DueDate, CatIdx, ItemDesc, " &
                    If(hasItemLabel, "COALESCE(ItemLabel, ItemDesc)", "ItemDesc") & ", " &
                    "Amount, " &
                    If(hasAdditional, "COALESCE(Additional,0)", "0") & ", " &
                    If(hasManualAmount, "COALESCE(ManualAmount,0)", "0") & ", " &
                    If(hasPaid, "COALESCE(Paid,0)", "0") & ", " &
                    If(hasSelectionMode, "COALESCE(SelectionMode,0)", "0") & ", " &
                    If(hasNotes, "COALESCE(Notes,'')", "''") & ", " &
                    If(hasFromAccountSnapshot, "COALESCE(FromAccountSnapshot,'')", "''") & ", " &
                    If(hasToAccountSnapshot, "COALESCE(ToAccountSnapshot,'')", "''") & ", " &
                    If(hasSameAsSnapshot, "COALESCE(SameAsSnapshot,'')", "''") & ", " &
                    If(hasFromAccountSnapshotId, "FromAccountSnapshotId", "NULL") & ", " &
                    If(hasToAccountSnapshotId, "ToAccountSnapshotId", "NULL") & ", " &
                    If(hasSameAsSnapshotId, "SameAsSnapshotId", "NULL") & ", " &
                    If(hasFromSavingsSnapshotId, "FromSavingsSnapshotId", "NULL") & ", " &
                    If(hasFromDebtSnapshotId, "FromDebtSnapshotId", "NULL") &
                    " FROM budgetDateOverrides WHERE CatIdx=@c AND (ItemDesc=@n OR ItemDesc=@legacy) ORDER BY DueDate"
                cmd.Parameters.AddWithValue("@c", categoryIndex)
                cmd.Parameters.AddWithValue("@n", itemKey.Trim())
                cmd.Parameters.AddWithValue("@legacy", itemDescription.Trim())

                Using reader = cmd.ExecuteReader()
                    While reader.Read()
                        Dim dueText = If(reader.IsDBNull(0), String.Empty, Convert.ToString(reader.GetValue(0)))
                        Dim parsedDue As DateTime
                        If Not DateTime.TryParseExact(dueText, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, parsedDue) Then
                            Continue While
                        End If

                        results.Add(New BudgetOverrideHistoryRecord With {
                            .DueDate = parsedDue,
                            .CategoryIndex = If(reader.IsDBNull(1), -1, Convert.ToInt32(reader.GetValue(1), CultureInfo.InvariantCulture)),
                            .ItemKey = If(reader.IsDBNull(2), String.Empty, Convert.ToString(reader.GetValue(2))).Trim(),
                            .ItemDescription = If(reader.IsDBNull(3), String.Empty, Convert.ToString(reader.GetValue(3))).Trim(),
                            .Amount = If(reader.IsDBNull(4), 0D, Convert.ToDecimal(reader.GetValue(4), CultureInfo.InvariantCulture)),
                            .Additional = If(reader.IsDBNull(5), 0D, Convert.ToDecimal(reader.GetValue(5), CultureInfo.InvariantCulture)),
                            .ManualAmount = If(reader.IsDBNull(6), 0D, Convert.ToDecimal(reader.GetValue(6), CultureInfo.InvariantCulture)),
                            .Paid = Not reader.IsDBNull(7) AndAlso Convert.ToInt32(reader.GetValue(7), CultureInfo.InvariantCulture) <> 0,
                            .SelectionMode = If(reader.IsDBNull(8), 0, Convert.ToInt32(reader.GetValue(8), CultureInfo.InvariantCulture)),
                            .Notes = If(reader.IsDBNull(9), String.Empty, Convert.ToString(reader.GetValue(9))).Trim(),
                            .FromAccountSnapshot = If(reader.IsDBNull(10), String.Empty, Convert.ToString(reader.GetValue(10))).Trim(),
                            .ToAccountSnapshot = If(reader.IsDBNull(11), String.Empty, Convert.ToString(reader.GetValue(11))).Trim(),
                            .SameAsSnapshot = If(reader.IsDBNull(12), String.Empty, Convert.ToString(reader.GetValue(12))).Trim(),
                            .FromAccountSnapshotId = If(reader.IsDBNull(13), CType(Nothing, Integer?), Convert.ToInt32(reader.GetValue(13), CultureInfo.InvariantCulture)),
                            .ToAccountSnapshotId = If(reader.IsDBNull(14), CType(Nothing, Integer?), Convert.ToInt32(reader.GetValue(14), CultureInfo.InvariantCulture)),
                            .SameAsSnapshotId = If(reader.IsDBNull(15), CType(Nothing, Integer?), Convert.ToInt32(reader.GetValue(15), CultureInfo.InvariantCulture)),
                            .FromSavingsSnapshotId = If(reader.IsDBNull(16), CType(Nothing, Integer?), Convert.ToInt32(reader.GetValue(16), CultureInfo.InvariantCulture)),
                            .FromDebtSnapshotId = If(reader.IsDBNull(17), CType(Nothing, Integer?), Convert.ToInt32(reader.GetValue(17), CultureInfo.InvariantCulture))
                        })
                    End While
                End Using
            End Using

            Return results
        End Function

        Private Sub ArchiveHistoryRecords(conn As SqliteConnection, records As IEnumerable(Of BudgetOverrideHistoryRecord))
            If conn Is Nothing OrElse records Is Nothing Then
                Return
            End If

            Dim capturedAt = DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture)

            For Each record In records
                If record Is Nothing OrElse record.CategoryIndex < 0 OrElse String.IsNullOrWhiteSpace(record.ItemKey) Then
                    Continue For
                End If

                Using cmd = conn.CreateCommand()
                    cmd.CommandText = "INSERT INTO budgetOverrideHistory (CapturedAtUtc, DueDate, CatIdx, ItemDesc, ItemLabel, Amount, Additional, ManualAmount, Paid, Notes, SelectionMode, FromAccountSnapshot, ToAccountSnapshot, SameAsSnapshot, FromAccountSnapshotId, ToAccountSnapshotId, SameAsSnapshotId, FromSavingsSnapshotId, FromDebtSnapshotId) VALUES (@capturedAt, @d, @c, @n, @l, @a, @additional, @manualAmount, @paid, @notes, @selectionMode, @fromAccountSnapshot, @toAccountSnapshot, @sameAsSnapshot, @fromAccountSnapshotId, @toAccountSnapshotId, @sameAsSnapshotId, @fromSavingsSnapshotId, @fromDebtSnapshotId)"
                    cmd.Parameters.AddWithValue("@capturedAt", capturedAt)
                    cmd.Parameters.AddWithValue("@d", record.DueDate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture))
                    cmd.Parameters.AddWithValue("@c", record.CategoryIndex)
                    cmd.Parameters.AddWithValue("@n", record.ItemKey.Trim())
                    cmd.Parameters.AddWithValue("@l", record.ItemDescription.Trim())
                    cmd.Parameters.AddWithValue("@a", record.Amount)
                    cmd.Parameters.AddWithValue("@additional", record.Additional)
                    cmd.Parameters.AddWithValue("@manualAmount", record.ManualAmount)
                    cmd.Parameters.AddWithValue("@paid", If(record.Paid, 1, 0))
                    cmd.Parameters.AddWithValue("@notes", If(record.Notes, String.Empty).Trim())
                    cmd.Parameters.AddWithValue("@selectionMode", record.SelectionMode)
                    cmd.Parameters.AddWithValue("@fromAccountSnapshot", If(record.FromAccountSnapshot, String.Empty).Trim())
                    cmd.Parameters.AddWithValue("@toAccountSnapshot", If(record.ToAccountSnapshot, String.Empty).Trim())
                    cmd.Parameters.AddWithValue("@sameAsSnapshot", If(record.SameAsSnapshot, String.Empty).Trim())
                    cmd.Parameters.AddWithValue("@fromAccountSnapshotId", If(record.FromAccountSnapshotId.HasValue, CType(record.FromAccountSnapshotId.Value, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@toAccountSnapshotId", If(record.ToAccountSnapshotId.HasValue, CType(record.ToAccountSnapshotId.Value, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@sameAsSnapshotId", If(record.SameAsSnapshotId.HasValue, CType(record.SameAsSnapshotId.Value, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@fromSavingsSnapshotId", If(record.FromSavingsSnapshotId.HasValue, CType(record.FromSavingsSnapshotId.Value, Object), DBNull.Value))
                    cmd.Parameters.AddWithValue("@fromDebtSnapshotId", If(record.FromDebtSnapshotId.HasValue, CType(record.FromDebtSnapshotId.Value, Object), DBNull.Value))
                    cmd.ExecuteNonQuery()
                End Using
            Next
        End Sub

        Private Sub EnsureBudgetOverrideArchiveTable(conn As SqliteConnection)
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "CREATE TABLE IF NOT EXISTS budgetOverrideHistory (Id INTEGER PRIMARY KEY AUTOINCREMENT, CapturedAtUtc TEXT NOT NULL, DueDate TEXT NOT NULL, CatIdx INTEGER NOT NULL, ItemDesc TEXT NOT NULL, ItemLabel TEXT NULL, Amount REAL NOT NULL, Additional REAL NOT NULL DEFAULT 0, ManualAmount REAL NOT NULL DEFAULT 0, Paid INTEGER NOT NULL DEFAULT 0, Notes TEXT NULL, SelectionMode INTEGER NOT NULL DEFAULT 0, FromAccountSnapshot TEXT NULL, ToAccountSnapshot TEXT NULL, SameAsSnapshot TEXT NULL, FromAccountSnapshotId INTEGER NULL, ToAccountSnapshotId INTEGER NULL, SameAsSnapshotId INTEGER NULL, FromSavingsSnapshotId INTEGER NULL, FromDebtSnapshotId INTEGER NULL)"
                cmd.ExecuteNonQuery()
            End Using
            EnsureColumn(conn, "budgetOverrideHistory", "FromAccountSnapshot", "TEXT NULL")
            EnsureColumn(conn, "budgetOverrideHistory", "ToAccountSnapshot", "TEXT NULL")
            EnsureColumn(conn, "budgetOverrideHistory", "SameAsSnapshot", "TEXT NULL")
            EnsureColumn(conn, "budgetOverrideHistory", "FromAccountSnapshotId", "INTEGER NULL")
            EnsureColumn(conn, "budgetOverrideHistory", "ToAccountSnapshotId", "INTEGER NULL")
            EnsureColumn(conn, "budgetOverrideHistory", "SameAsSnapshotId", "INTEGER NULL")
            EnsureColumn(conn, "budgetOverrideHistory", "FromSavingsSnapshotId", "INTEGER NULL")
            EnsureColumn(conn, "budgetOverrideHistory", "FromDebtSnapshotId", "INTEGER NULL")
        End Sub

        Private Sub EnsureLiveBudgetOverridesTable(conn As SqliteConnection)
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "CREATE TABLE IF NOT EXISTS budgetDateOverrides (DueDate TEXT NOT NULL, CatIdx INTEGER NOT NULL, ItemDesc TEXT NOT NULL, ItemLabel TEXT NULL, Amount REAL NOT NULL, Additional REAL NOT NULL DEFAULT 0, ManualAmount REAL NOT NULL DEFAULT 0, Paid INTEGER NOT NULL DEFAULT 0, Notes TEXT NULL, SelectionMode INTEGER NOT NULL DEFAULT 0, FromAccountSnapshot TEXT NULL, ToAccountSnapshot TEXT NULL, SameAsSnapshot TEXT NULL, FromAccountSnapshotId INTEGER NULL, ToAccountSnapshotId INTEGER NULL, SameAsSnapshotId INTEGER NULL, FromSavingsSnapshotId INTEGER NULL, FromDebtSnapshotId INTEGER NULL, PRIMARY KEY(DueDate, CatIdx, ItemDesc))"
                cmd.ExecuteNonQuery()
            End Using

            EnsureColumn(conn, "budgetDateOverrides", "ItemLabel", "TEXT NULL")
            EnsureColumn(conn, "budgetDateOverrides", "Additional", "REAL NOT NULL DEFAULT 0")
            EnsureColumn(conn, "budgetDateOverrides", "ManualAmount", "REAL NOT NULL DEFAULT 0")
            EnsureColumn(conn, "budgetDateOverrides", "Paid", "INTEGER NOT NULL DEFAULT 0")
            EnsureColumn(conn, "budgetDateOverrides", "Notes", "TEXT NULL")
            EnsureColumn(conn, "budgetDateOverrides", "SelectionMode", "INTEGER NOT NULL DEFAULT 0")
            EnsureColumn(conn, "budgetDateOverrides", "FromAccountSnapshot", "TEXT NULL")
            EnsureColumn(conn, "budgetDateOverrides", "ToAccountSnapshot", "TEXT NULL")
            EnsureColumn(conn, "budgetDateOverrides", "SameAsSnapshot", "TEXT NULL")
            EnsureColumn(conn, "budgetDateOverrides", "FromAccountSnapshotId", "INTEGER NULL")
            EnsureColumn(conn, "budgetDateOverrides", "ToAccountSnapshotId", "INTEGER NULL")
            EnsureColumn(conn, "budgetDateOverrides", "SameAsSnapshotId", "INTEGER NULL")
            EnsureColumn(conn, "budgetDateOverrides", "FromSavingsSnapshotId", "INTEGER NULL")
            EnsureColumn(conn, "budgetDateOverrides", "FromDebtSnapshotId", "INTEGER NULL")
        End Sub

        Private Function HasTable(conn As SqliteConnection, tableName As String) As Boolean
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "SELECT 1 FROM sqlite_master WHERE type='table' AND name=@name LIMIT 1"
                cmd.Parameters.AddWithValue("@name", tableName)
                Dim result = cmd.ExecuteScalar()
                Return result IsNot Nothing AndAlso result IsNot DBNull.Value
            End Using
        End Function

        Private Function HasColumn(conn As SqliteConnection, tableName As String, columnName As String) As Boolean
            If Not HasTable(conn, tableName) Then
                Return False
            End If

            Using cmd = conn.CreateCommand()
                cmd.CommandText = "PRAGMA table_info(" & tableName & ")"
                Using reader = cmd.ExecuteReader()
                    While reader.Read()
                        Dim currentName = If(reader.IsDBNull(1), String.Empty, Convert.ToString(reader.GetValue(1)))
                        If String.Equals(currentName, columnName, StringComparison.OrdinalIgnoreCase) Then
                            Return True
                        End If
                    End While
                End Using
            End Using

            Return False
        End Function

        Private Sub EnsureColumn(conn As SqliteConnection, tableName As String, columnName As String, columnDefinition As String)
            If HasColumn(conn, tableName, columnName) Then
                Return
            End If

            Using cmd = conn.CreateCommand()
                cmd.CommandText = "ALTER TABLE " & tableName & " ADD COLUMN " & columnName & " " & columnDefinition
                cmd.ExecuteNonQuery()
            End Using
        End Sub

    End Module

End Namespace
