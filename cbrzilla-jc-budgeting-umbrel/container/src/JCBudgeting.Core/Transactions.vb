Imports Microsoft.Data.Sqlite
Imports System.Text.RegularExpressions

Namespace Global.JCBudgeting.Core

    Public Class TransactionRecord
        Public Property Id As Integer
        Public Property SourceName As String = String.Empty
        Public Property TransactionDate As String = String.Empty
        Public Property Description As String = String.Empty
        Public Property Amount As Decimal
        Public Property Notes As String = String.Empty
    End Class

    Public Class TransactionAssignmentRecord
        Public Property Id As Integer
        Public Property TransactionId As Integer
        Public Property CatIdx As Integer
        Public Property ItemId As Integer
        Public Property Amount As Decimal
        Public Property Notes As String = String.Empty
        Public Property NeedsReview As Boolean
    End Class

    Public Class TransactionDuplicateCandidate
        Public Property ImportIndex As Integer
        Public Property ExistingTransactionId As Integer
        Public Property ExistingDescription As String = String.Empty
        Public Property ExistingSourceName As String = String.Empty
        Public Property ExistingTransactionDate As String = String.Empty
        Public Property ExistingAmount As Decimal
        Public Property Score As Decimal
        Public Property IsExactDuplicate As Boolean
    End Class

    Public Module TransactionRepository

        Public Function NormalizeTransactionDescription(value As String) As String
            Dim trimmed = If(value, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(trimmed) Then
                Return String.Empty
            End If

            Return Regex.Replace(trimmed, "\s{2,}", " ")
        End Function

        Private Function NormalizeTransactionDescriptionForDuplicateMatch(value As String) As String
            Dim normalized = NormalizeTransactionDescription(value)
            If String.IsNullOrWhiteSpace(normalized) Then
                Return String.Empty
            End If

            normalized = normalized.ToUpperInvariant()

            normalized = Regex.Replace(normalized, "\bCHECKING\s+ACH\s+WITHDRAWAL\b", " ")
            normalized = Regex.Replace(normalized, "\bACH\s+WITHDRAWAL\b", " ")
            normalized = Regex.Replace(normalized, "\bPOS\s+(OR|/)\s*VISA\s+PURCHASE\b", " ")
            normalized = Regex.Replace(normalized, "\bPOS\s+(OR|/)\s*VISA\s+PURCH\b", " ")
            normalized = Regex.Replace(normalized, "\bPOS\s+(OR|/)\s*VISA\b", " ")
            normalized = Regex.Replace(normalized, "\bPURCHASED\b", " ")
            normalized = Regex.Replace(normalized, "\bPURCHASE\b", " ")

            normalized = normalized.Replace("/", " ")
            normalized = normalized.Replace("-", " ")
            normalized = normalized.Replace("*", " ")

            normalized = Regex.Replace(normalized, "[^\w\s]", " ")
            normalized = Regex.Replace(normalized, "\s{2,}", " ").Trim()
            Return normalized
        End Function

        Private Function CalculateTransactionDescriptionSimilarity(leftValue As String, rightValue As String) As Decimal
            Dim leftNormalized = NormalizeTransactionDescriptionForDuplicateMatch(leftValue)
            Dim rightNormalized = NormalizeTransactionDescriptionForDuplicateMatch(rightValue)

            If String.IsNullOrWhiteSpace(leftNormalized) OrElse String.IsNullOrWhiteSpace(rightNormalized) Then
                Return 0D
            End If

            If String.Equals(leftNormalized, rightNormalized, StringComparison.OrdinalIgnoreCase) Then
                Return 1D
            End If

            Dim leftTokens = New HashSet(Of String)(
                leftNormalized.Split({" "c}, StringSplitOptions.RemoveEmptyEntries).
                    Where(Function(token) token.Length > 1),
                StringComparer.OrdinalIgnoreCase)
            Dim rightTokens = New HashSet(Of String)(
                rightNormalized.Split({" "c}, StringSplitOptions.RemoveEmptyEntries).
                    Where(Function(token) token.Length > 1),
                StringComparer.OrdinalIgnoreCase)

            If leftTokens.Count = 0 OrElse rightTokens.Count = 0 Then
                Return 0D
            End If

            Dim overlapCount = leftTokens.Intersect(rightTokens, StringComparer.OrdinalIgnoreCase).Count()
            If overlapCount = 0 Then
                Return 0D
            End If

            Dim unionCount = leftTokens.Union(rightTokens, StringComparer.OrdinalIgnoreCase).Count()
            Dim jaccard = CDec(overlapCount) / Math.Max(1D, unionCount)
            Dim containment = CDec(overlapCount) / Math.Max(1D, Math.Min(leftTokens.Count, rightTokens.Count))

            Dim score = Math.Max(jaccard, containment)

            If leftNormalized.IndexOf(rightNormalized, StringComparison.OrdinalIgnoreCase) >= 0 OrElse
               rightNormalized.IndexOf(leftNormalized, StringComparison.OrdinalIgnoreCase) >= 0 Then
                score = Math.Max(score, 0.85D)
            End If

            Return Decimal.Round(score, 4, MidpointRounding.AwayFromZero)
        End Function

        Public Sub EnsureTransactionsSchema(databasePath As String)
            If String.IsNullOrWhiteSpace(databasePath) OrElse Not IO.File.Exists(databasePath) Then
                Return
            End If

            Using conn As New SqliteConnection("Data Source=" & databasePath.Trim())
                conn.Open()
                EnsureTransactionsSchema(conn)
            End Using
        End Sub

        Public Function LoadTransactions(databasePath As String) As List(Of TransactionRecord)
            Dim results As New List(Of TransactionRecord)()
            If String.IsNullOrWhiteSpace(databasePath) OrElse Not IO.File.Exists(databasePath) Then
                Return results
            End If

            Using conn As New SqliteConnection("Data Source=" & databasePath.Trim())
                conn.Open()
                EnsureTransactionsSchema(conn)

                Using cmd = conn.CreateCommand()
                    cmd.CommandText =
                        "SELECT Id, SourceName, TransactionDate, Description, Amount, Notes " &
                        "FROM transactions " &
                        "ORDER BY date(TransactionDate) DESC, Id DESC"

                    Using reader = cmd.ExecuteReader()
                        While reader.Read()
                            results.Add(New TransactionRecord With {
                                .Id = If(reader.IsDBNull(0), 0, reader.GetInt32(0)),
                                .SourceName = If(reader.IsDBNull(1), String.Empty, Convert.ToString(reader.GetValue(1))),
                                .TransactionDate = If(reader.IsDBNull(2), String.Empty, Convert.ToString(reader.GetValue(2))),
                                .Description = NormalizeTransactionDescription(If(reader.IsDBNull(3), String.Empty, Convert.ToString(reader.GetValue(3)))),
                                .Amount = If(reader.IsDBNull(4), 0D, Convert.ToDecimal(reader.GetValue(4))),
                                .Notes = If(reader.IsDBNull(5), String.Empty, Convert.ToString(reader.GetValue(5)))
                            })
                        End While
                    End Using
                End Using
            End Using

            Return results
        End Function

        Public Function LoadTransactionSources(databasePath As String) As List(Of String)
            Dim results As New List(Of String)()
            If String.IsNullOrWhiteSpace(databasePath) OrElse Not IO.File.Exists(databasePath) Then
                Return results
            End If

            Using conn As New SqliteConnection("Data Source=" & databasePath.Trim())
                conn.Open()
                EnsureTransactionsSchema(conn)

                Using cmd = conn.CreateCommand()
                    cmd.CommandText =
                        "SELECT DISTINCT TRIM(SourceName) " &
                        "FROM transactions " &
                        "WHERE TRIM(COALESCE(SourceName,'')) <> '' " &
                        "ORDER BY TRIM(SourceName)"

                    Using reader = cmd.ExecuteReader()
                        While reader.Read()
                            Dim value = If(reader.IsDBNull(0), String.Empty, Convert.ToString(reader.GetValue(0)))
                            If Not String.IsNullOrWhiteSpace(value) Then
                                results.Add(value.Trim())
                            End If
                        End While
                    End Using
                End Using
            End Using

            Return results
        End Function

        Public Function LoadTransactionAssignments(databasePath As String) As List(Of TransactionAssignmentRecord)
            Dim results As New List(Of TransactionAssignmentRecord)()
            If String.IsNullOrWhiteSpace(databasePath) OrElse Not IO.File.Exists(databasePath) Then
                Return results
            End If

            Using conn As New SqliteConnection("Data Source=" & databasePath.Trim())
                conn.Open()
                EnsureTransactionsSchema(conn)

                Using cmd = conn.CreateCommand()
                    cmd.CommandText =
                        "SELECT Id, TransactionId, CatIdx, ItemId, Amount, Notes, COALESCE(NeedsReview,0) " &
                        "FROM transactionAssignments " &
                        "ORDER BY TransactionId, Id"

                    Using reader = cmd.ExecuteReader()
                        While reader.Read()
                            results.Add(New TransactionAssignmentRecord With {
                                .Id = If(reader.IsDBNull(0), 0, reader.GetInt32(0)),
                                .TransactionId = If(reader.IsDBNull(1), 0, Convert.ToInt32(reader.GetValue(1))),
                                .CatIdx = If(reader.IsDBNull(2), -1, Convert.ToInt32(reader.GetValue(2))),
                                .ItemId = If(reader.IsDBNull(3), 0, Convert.ToInt32(reader.GetValue(3))),
                                .Amount = If(reader.IsDBNull(4), 0D, Convert.ToDecimal(reader.GetValue(4))),
                                .Notes = If(reader.IsDBNull(5), String.Empty, Convert.ToString(reader.GetValue(5))),
                                .NeedsReview = Not reader.IsDBNull(6) AndAlso Convert.ToInt32(reader.GetValue(6)) <> 0
                            })
                        End While
                    End Using
                End Using
            End Using

            Return results
        End Function

        Public Sub UpdateTransactionSource(databasePath As String, transactionId As Integer, sourceName As String)
            If String.IsNullOrWhiteSpace(databasePath) OrElse Not IO.File.Exists(databasePath) OrElse transactionId <= 0 Then
                Return
            End If

            Using conn As New SqliteConnection("Data Source=" & databasePath.Trim())
                conn.Open()
                EnsureTransactionsSchema(conn)

                Using cmd = conn.CreateCommand()
                    cmd.CommandText =
                        "UPDATE transactions " &
                        "SET SourceName=@source " &
                        "WHERE Id=@id"
                    cmd.Parameters.AddWithValue("@source", If(sourceName, String.Empty))
                    cmd.Parameters.AddWithValue("@id", transactionId)
                    cmd.ExecuteNonQuery()
                End Using
            End Using
        End Sub

        Public Function CreateTransaction(databasePath As String, record As TransactionRecord, assignments As IEnumerable(Of TransactionAssignmentRecord)) As Integer
            If String.IsNullOrWhiteSpace(databasePath) OrElse Not IO.File.Exists(databasePath) Then
                Throw New IO.FileNotFoundException("Budget database was not found.", databasePath)
            End If

            If record Is Nothing Then
                Throw New ArgumentNullException(NameOf(record))
            End If

            Using conn As New SqliteConnection("Data Source=" & databasePath.Trim())
                conn.Open()
                EnsureTransactionsSchema(conn)

                Using tx = conn.BeginTransaction()
                    Dim createdId As Integer
                    Dim normalizedDescription = NormalizeTransactionDescription(record.Description)
                    Using cmd = conn.CreateCommand()
                        cmd.Transaction = tx
                        cmd.CommandText =
                            "INSERT INTO transactions (SourceName, TransactionDate, Description, Amount, Notes) " &
                            "VALUES (@source, @date, @description, @amount, @notes);" &
                            "SELECT last_insert_rowid();"
                        cmd.Parameters.AddWithValue("@source", If(record.SourceName, String.Empty))
                        cmd.Parameters.AddWithValue("@date", If(record.TransactionDate, String.Empty))
                        cmd.Parameters.AddWithValue("@description", normalizedDescription)
                        cmd.Parameters.AddWithValue("@amount", record.Amount)
                        cmd.Parameters.AddWithValue("@notes", If(record.Notes, String.Empty))
                        createdId = Convert.ToInt32(cmd.ExecuteScalar())
                    End Using

                    SaveAssignments(conn, tx, createdId, assignments)
                    tx.Commit()
                    Return createdId
                End Using
            End Using
        End Function

        Public Sub SaveTransaction(databasePath As String, record As TransactionRecord, assignments As IEnumerable(Of TransactionAssignmentRecord))
            If String.IsNullOrWhiteSpace(databasePath) OrElse Not IO.File.Exists(databasePath) Then
                Throw New IO.FileNotFoundException("Budget database was not found.", databasePath)
            End If

            If record Is Nothing Then
                Throw New ArgumentNullException(NameOf(record))
            End If

            If record.Id <= 0 Then
                Throw New InvalidOperationException("A valid transaction id is required.")
            End If

            Using conn As New SqliteConnection("Data Source=" & databasePath.Trim())
                conn.Open()
                EnsureTransactionsSchema(conn)

                Using tx = conn.BeginTransaction()
                    Dim normalizedDescription = NormalizeTransactionDescription(record.Description)
                    Using cmd = conn.CreateCommand()
                        cmd.Transaction = tx
                        cmd.CommandText =
                            "UPDATE transactions " &
                            "SET SourceName=@source, TransactionDate=@date, Description=@description, Amount=@amount, Notes=@notes " &
                            "WHERE Id=@id"
                        cmd.Parameters.AddWithValue("@source", If(record.SourceName, String.Empty))
                        cmd.Parameters.AddWithValue("@date", If(record.TransactionDate, String.Empty))
                        cmd.Parameters.AddWithValue("@description", normalizedDescription)
                        cmd.Parameters.AddWithValue("@amount", record.Amount)
                        cmd.Parameters.AddWithValue("@notes", If(record.Notes, String.Empty))
                        cmd.Parameters.AddWithValue("@id", record.Id)

                        Dim rowsChanged = cmd.ExecuteNonQuery()
                        If rowsChanged <= 0 Then
                            Throw New InvalidOperationException("The selected transaction could not be saved.")
                        End If
                    End Using

                    SaveAssignments(conn, tx, record.Id, assignments)
                    tx.Commit()
                End Using
            End Using
        End Sub

        Public Sub DeleteTransaction(databasePath As String, transactionId As Integer)
            If String.IsNullOrWhiteSpace(databasePath) OrElse Not IO.File.Exists(databasePath) Then
                Throw New IO.FileNotFoundException("Budget database was not found.", databasePath)
            End If

            If transactionId <= 0 Then
                Throw New ArgumentOutOfRangeException(NameOf(transactionId))
            End If

            Using conn As New SqliteConnection("Data Source=" & databasePath.Trim())
                conn.Open()
                EnsureTransactionsSchema(conn)

                Using tx = conn.BeginTransaction()
                    Using deleteAssignments = conn.CreateCommand()
                        deleteAssignments.Transaction = tx
                        deleteAssignments.CommandText = "DELETE FROM transactionAssignments WHERE TransactionId=@id"
                        deleteAssignments.Parameters.AddWithValue("@id", transactionId)
                        deleteAssignments.ExecuteNonQuery()
                    End Using

                    Using deleteTransaction = conn.CreateCommand()
                        deleteTransaction.Transaction = tx
                        deleteTransaction.CommandText = "DELETE FROM transactions WHERE Id=@id"
                        deleteTransaction.Parameters.AddWithValue("@id", transactionId)
                        deleteTransaction.ExecuteNonQuery()
                    End Using

                    tx.Commit()
                End Using
            End Using
        End Sub

        Public Function FindDuplicateTransactionIds(databasePath As String, records As IEnumerable(Of TransactionRecord)) As HashSet(Of Integer)
            If String.IsNullOrWhiteSpace(databasePath) OrElse Not IO.File.Exists(databasePath) OrElse records Is Nothing Then
                Return New HashSet(Of Integer)()
            End If

            Return FindDuplicateTransactionIds(LoadTransactions(databasePath), records)
        End Function

        Public Function FindDuplicateTransactionIds(existingRecords As IEnumerable(Of TransactionRecord), records As IEnumerable(Of TransactionRecord)) As HashSet(Of Integer)
            Dim duplicates As New HashSet(Of Integer)()
            If existingRecords Is Nothing OrElse records Is Nothing Then
                Return duplicates
            End If

            For Each candidate In FindDuplicateCandidates(existingRecords, records)
                If candidate.ExistingTransactionId > 0 Then
                    duplicates.Add(candidate.ExistingTransactionId)
                End If
            Next

            Return duplicates
        End Function

        Public Function FindDuplicateCandidates(databasePath As String, records As IEnumerable(Of TransactionRecord)) As List(Of TransactionDuplicateCandidate)
            If String.IsNullOrWhiteSpace(databasePath) OrElse Not IO.File.Exists(databasePath) OrElse records Is Nothing Then
                Return New List(Of TransactionDuplicateCandidate)()
            End If

            Return FindDuplicateCandidates(LoadTransactions(databasePath), records)
        End Function

        Public Function FindDuplicateCandidates(existingRecords As IEnumerable(Of TransactionRecord), records As IEnumerable(Of TransactionRecord)) As List(Of TransactionDuplicateCandidate)
            Dim duplicates As New List(Of TransactionDuplicateCandidate)()
            If existingRecords Is Nothing OrElse records Is Nothing Then
                Return duplicates
            End If

            Dim existing = existingRecords.
                Select(Function(item) New TransactionRecord With {
                    .Id = item.Id,
                    .SourceName = If(item.SourceName, String.Empty),
                    .TransactionDate = If(item.TransactionDate, String.Empty),
                    .Description = NormalizeTransactionDescription(If(item.Description, String.Empty)),
                    .Amount = item.Amount,
                    .Notes = If(item.Notes, String.Empty)
                }).
                ToList()

            Dim indexedRecords = records.Select(Function(record, index) New With {
                .Index = index,
                .Record = record
            })

            For Each entry In indexedRecords
                Dim record = entry.Record
                Dim sameDayAmountSource = existing.Where(Function(x)
                                                             Return x.Amount = record.Amount AndAlso
                                                                    String.Equals(x.TransactionDate, If(record.TransactionDate, String.Empty), StringComparison.Ordinal) AndAlso
                                                                    String.Equals(If(x.SourceName, String.Empty), If(record.SourceName, String.Empty), StringComparison.CurrentCultureIgnoreCase)
                                                         End Function).ToList()

                If sameDayAmountSource.Count = 0 Then
                    Continue For
                End If

                Dim bestMatch = sameDayAmountSource.
                    Select(Function(x) New With {
                        .Existing = x,
                        .Score = CalculateTransactionDescriptionSimilarity(record.Description, x.Description),
                        .Exact = String.Equals(
                            NormalizeTransactionDescriptionForDuplicateMatch(record.Description),
                            NormalizeTransactionDescriptionForDuplicateMatch(x.Description),
                            StringComparison.CurrentCultureIgnoreCase)
                    }).
                    OrderByDescending(Function(x) If(x.Exact, 1D, x.Score)).
                    FirstOrDefault()

                If bestMatch IsNot Nothing AndAlso (bestMatch.Exact OrElse bestMatch.Score >= 0.55D) Then
                    duplicates.Add(New TransactionDuplicateCandidate With {
                        .ImportIndex = entry.Index,
                        .ExistingTransactionId = bestMatch.Existing.Id,
                        .ExistingDescription = bestMatch.Existing.Description,
                        .ExistingSourceName = bestMatch.Existing.SourceName,
                        .ExistingTransactionDate = bestMatch.Existing.TransactionDate,
                        .ExistingAmount = bestMatch.Existing.Amount,
                        .Score = If(bestMatch.Exact, 1D, bestMatch.Score),
                        .IsExactDuplicate = bestMatch.Exact
                    })
                End If
            Next

            Return duplicates
        End Function

        Public Sub CleanupOrphanedAssignments(
            databasePath As String,
            incomes As IEnumerable(Of IncomeRecord),
            debts As IEnumerable(Of DebtRecord),
            expenses As IEnumerable(Of ExpenseRecord),
            savings As IEnumerable(Of SavingsRecord))

            If String.IsNullOrWhiteSpace(databasePath) OrElse Not IO.File.Exists(databasePath) Then
                Return
            End If

            Dim validIdsByCategory As New Dictionary(Of Integer, HashSet(Of Integer)) From {
                {0, New HashSet(Of Integer)(incomes.Select(Function(x) x.Id))},
                {1, New HashSet(Of Integer)(debts.Select(Function(x) x.Id))},
                {2, New HashSet(Of Integer)(expenses.Select(Function(x) x.Id))},
                {3, New HashSet(Of Integer)(savings.Select(Function(x) x.Id))}
            }

            Using conn As New SqliteConnection("Data Source=" & databasePath.Trim())
                conn.Open()
                EnsureTransactionsSchema(conn)

                Dim deleteIds As New List(Of Integer)()
                Using cmd = conn.CreateCommand()
                    cmd.CommandText = "SELECT Id, CatIdx, ItemId FROM transactionAssignments"
                    Using reader = cmd.ExecuteReader()
                        While reader.Read()
                            Dim assignmentId = If(reader.IsDBNull(0), 0, reader.GetInt32(0))
                            Dim catIdx = If(reader.IsDBNull(1), Int32.MinValue, Convert.ToInt32(reader.GetValue(1)))
                            Dim itemId = If(reader.IsDBNull(2), 0, Convert.ToInt32(reader.GetValue(2)))
                            If assignmentId <= 0 OrElse itemId <= 0 Then
                                deleteIds.Add(assignmentId)
                                Continue While
                            End If

                            If Not validIdsByCategory.ContainsKey(catIdx) OrElse Not validIdsByCategory(catIdx).Contains(itemId) Then
                                deleteIds.Add(assignmentId)
                            End If
                        End While
                    End Using
                End Using

                For Each assignmentId In deleteIds.Where(Function(x) x > 0)
                    Using deleteCmd = conn.CreateCommand()
                        deleteCmd.CommandText = "DELETE FROM transactionAssignments WHERE Id=@id"
                        deleteCmd.Parameters.AddWithValue("@id", assignmentId)
                        deleteCmd.ExecuteNonQuery()
                    End Using
                Next
            End Using
        End Sub

        Private Sub SaveAssignments(conn As SqliteConnection, tx As SqliteTransaction, transactionId As Integer, assignments As IEnumerable(Of TransactionAssignmentRecord))
            Using deleteCmd = conn.CreateCommand()
                deleteCmd.Transaction = tx
                deleteCmd.CommandText = "DELETE FROM transactionAssignments WHERE TransactionId=@id"
                deleteCmd.Parameters.AddWithValue("@id", transactionId)
                deleteCmd.ExecuteNonQuery()
            End Using

            If assignments Is Nothing Then
                Return
            End If

            For Each assignment In assignments
                If assignment Is Nothing OrElse assignment.ItemId <= 0 OrElse assignment.CatIdx < 0 Then
                    Continue For
                End If

                Using insertCmd = conn.CreateCommand()
                    insertCmd.Transaction = tx
                    insertCmd.CommandText =
                        "INSERT INTO transactionAssignments (TransactionId, CatIdx, ItemId, Amount, Notes, NeedsReview) " &
                        "VALUES (@transactionId, @catIdx, @itemId, @amount, @notes, @needsReview)"
                    insertCmd.Parameters.AddWithValue("@transactionId", transactionId)
                    insertCmd.Parameters.AddWithValue("@catIdx", assignment.CatIdx)
                    insertCmd.Parameters.AddWithValue("@itemId", assignment.ItemId)
                    insertCmd.Parameters.AddWithValue("@amount", assignment.Amount)
                    insertCmd.Parameters.AddWithValue("@notes", If(assignment.Notes, String.Empty))
                    insertCmd.Parameters.AddWithValue("@needsReview", If(assignment.NeedsReview, 1, 0))
                    insertCmd.ExecuteNonQuery()
                End Using
            Next
        End Sub

        Private Sub EnsureTransactionsSchema(conn As SqliteConnection)
            Using createTransactions = conn.CreateCommand()
                createTransactions.CommandText =
                    "CREATE TABLE IF NOT EXISTS transactions (" &
                    "Id INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT, " &
                    "SourceName TEXT NOT NULL DEFAULT '', " &
                    "TransactionDate TEXT NOT NULL DEFAULT '', " &
                    "Description TEXT NOT NULL DEFAULT '', " &
                    "Amount REAL NOT NULL DEFAULT 0, " &
                    "Notes TEXT NOT NULL DEFAULT ''" &
                    ")"
                createTransactions.ExecuteNonQuery()
            End Using

            Using createAssignments = conn.CreateCommand()
                createAssignments.CommandText =
                    "CREATE TABLE IF NOT EXISTS transactionAssignments (" &
                    "Id INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT, " &
                    "TransactionId INTEGER NOT NULL, " &
                    "CatIdx INTEGER NOT NULL, " &
                    "ItemId INTEGER NOT NULL, " &
                    "Amount REAL NOT NULL DEFAULT 0, " &
                    "Notes TEXT NOT NULL DEFAULT '', " &
                    "NeedsReview INTEGER NOT NULL DEFAULT 0" &
                    ")"
                createAssignments.ExecuteNonQuery()
            End Using

            Using indexTransactions = conn.CreateCommand()
                indexTransactions.CommandText = "CREATE INDEX IF NOT EXISTS IX_transactions_DateDescriptionAmount ON transactions (TransactionDate, Description, Amount)"
                indexTransactions.ExecuteNonQuery()
            End Using

            Using indexAssignments = conn.CreateCommand()
                indexAssignments.CommandText = "CREATE INDEX IF NOT EXISTS IX_transactionAssignments_TransactionId ON transactionAssignments (TransactionId)"
                indexAssignments.ExecuteNonQuery()
            End Using
        End Sub
    End Module
End Namespace
