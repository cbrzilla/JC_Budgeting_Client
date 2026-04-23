Imports Microsoft.Data.Sqlite

Namespace Global.JCBudgeting.Core

    Public Class ExpenseRecord
        Public Property Id As Integer
        Public Property Description As String
        Public Property AmountDue As Decimal
        Public Property Cadence As String
        Public Property DueDay As Integer?
        Public Property DueDate As String
        Public Property FromAccount As String
        Public Property SameAs As String
        Public Property Category As String
        Public Property Hidden As Boolean
        Public Property Active As Boolean
        Public Property LoginLink As String
        Public Property Notes As String
    End Class

    Public Module ExpenseRepository
        Private NotInheritable Class SqliteTableColumnDefinition
            Public Property Name As String = String.Empty
            Public Property SqlType As String = String.Empty
            Public Property NotNull As Boolean
            Public Property DefaultValueSql As String = String.Empty
            Public Property PrimaryKeyOrder As Integer
        End Class

        Public Function LoadExpenses(databasePath As String) As List(Of ExpenseRecord)
            Dim results As New List(Of ExpenseRecord)()
            If String.IsNullOrWhiteSpace(databasePath) OrElse Not IO.File.Exists(databasePath) Then
                Return results
            End If

            Using conn As New SqliteConnection("Data Source=" & databasePath.Trim())
                conn.Open()
                EnsureActiveColumn(conn)

                Using cmd = conn.CreateCommand()
                    cmd.CommandText =
                        "SELECT Id, Description, AmtDue, Cadence, DueDay, DueDate, FromAccount, SameAs, Category, Hidden, COALESCE(Active,1), LoginLink, Notes FROM expenses ORDER BY COALESCE(Hidden,0), Description COLLATE NOCASE"
                    Using reader = cmd.ExecuteReader()
                        While reader.Read()
                            results.Add(New ExpenseRecord With {
                                .Id = If(reader.IsDBNull(0), 0, reader.GetInt32(0)),
                                .Description = If(reader.IsDBNull(1), String.Empty, Convert.ToString(reader.GetValue(1))),
                                .AmountDue = If(reader.IsDBNull(2), 0D, Convert.ToDecimal(reader.GetValue(2))),
                                .Cadence = If(reader.IsDBNull(3), String.Empty, Convert.ToString(reader.GetValue(3))),
                                .DueDay = If(reader.IsDBNull(4), CType(Nothing, Integer?), Convert.ToInt32(reader.GetValue(4))),
                                .DueDate = If(reader.IsDBNull(5), String.Empty, Convert.ToString(reader.GetValue(5))),
                                .FromAccount = If(reader.IsDBNull(6), String.Empty, Convert.ToString(reader.GetValue(6))),
                                .SameAs = If(reader.IsDBNull(7), String.Empty, Convert.ToString(reader.GetValue(7))),
                                .Category = If(reader.IsDBNull(8), String.Empty, Convert.ToString(reader.GetValue(8))),
                                .Hidden = Not reader.IsDBNull(9) AndAlso Convert.ToInt32(reader.GetValue(9)) <> 0,
                                .Active = Not reader.IsDBNull(10) AndAlso Convert.ToInt32(reader.GetValue(10)) <> 0,
                                .LoginLink = If(reader.IsDBNull(11), String.Empty, Convert.ToString(reader.GetValue(11))),
                                .Notes = If(reader.IsDBNull(12), String.Empty, Convert.ToString(reader.GetValue(12)))
                            })
                        End While
                    End Using
                End Using
            End Using

            Return results
        End Function

        Public Sub SaveExpense(databasePath As String, expense As ExpenseRecord)
            If String.IsNullOrWhiteSpace(databasePath) OrElse Not IO.File.Exists(databasePath) Then
                Throw New IO.FileNotFoundException("Budget database was not found.", databasePath)
            End If

            If expense Is Nothing Then
                Throw New ArgumentNullException(NameOf(expense))
            End If

            If expense.Id <= 0 Then
                Throw New InvalidOperationException("Only existing expenses can be saved in this migration pass.")
            End If

            Using conn As New SqliteConnection("Data Source=" & databasePath.Trim())
                conn.Open()
                EnsureActiveColumn(conn)

                Using cmd = conn.CreateCommand()
                    cmd.CommandText =
                        "UPDATE expenses " &
                        "SET Description = @Description, " &
                        "AmtDue = @AmtDue, " &
                        "Cadence = @Cadence, " &
                        "DueDay = @DueDay, " &
                        "DueDate = @DueDate, " &
                        "FromAccount = @FromAccount, " &
                        "SameAs = @SameAs, " &
                        "Category = @Category, " &
                        "Hidden = @Hidden, " &
                        "Active = @Active, " &
                        "LoginLink = @LoginLink, " &
                        "Notes = @Notes " &
                        "WHERE Id = @Id"

                    cmd.Parameters.AddWithValue("@Description", If(expense.Description, String.Empty))
                    cmd.Parameters.AddWithValue("@AmtDue", expense.AmountDue)
                    cmd.Parameters.AddWithValue("@Cadence", If(expense.Cadence, String.Empty))

                    Dim dueDayParam = cmd.CreateParameter()
                    dueDayParam.ParameterName = "@DueDay"
                    dueDayParam.Value = If(expense.DueDay.HasValue, expense.DueDay.Value, DBNull.Value)
                    cmd.Parameters.Add(dueDayParam)

                    cmd.Parameters.AddWithValue("@DueDate", If(expense.DueDate, String.Empty))
                    cmd.Parameters.AddWithValue("@FromAccount", If(expense.FromAccount, String.Empty))
                    cmd.Parameters.AddWithValue("@SameAs", If(expense.SameAs, String.Empty))
                    cmd.Parameters.AddWithValue("@Category", If(expense.Category, String.Empty))
                    cmd.Parameters.AddWithValue("@Hidden", If(expense.Hidden, 1, 0))
                    cmd.Parameters.AddWithValue("@Active", If(expense.Active, 1, 0))
                    cmd.Parameters.AddWithValue("@LoginLink", If(expense.LoginLink, String.Empty))
                    cmd.Parameters.AddWithValue("@Notes", If(expense.Notes, String.Empty))
                    cmd.Parameters.AddWithValue("@Id", expense.Id)

                    Dim rowsChanged = cmd.ExecuteNonQuery()
                    If rowsChanged <= 0 Then
                        Throw New InvalidOperationException("The selected expense could not be saved.")
                    End If
                End Using
            End Using
        End Sub

        Public Function CreateExpense(databasePath As String, expense As ExpenseRecord) As Integer
            If String.IsNullOrWhiteSpace(databasePath) OrElse Not IO.File.Exists(databasePath) Then
                Throw New IO.FileNotFoundException("Budget database was not found.", databasePath)
            End If

            If expense Is Nothing Then
                Throw New ArgumentNullException(NameOf(expense))
            End If

            Using conn As New SqliteConnection("Data Source=" & databasePath.Trim())
                conn.Open()
                EnsureActiveColumn(conn)

                Using cmd = conn.CreateCommand()
                    cmd.CommandText =
                        "INSERT INTO expenses (Description, AmtDue, Cadence, DueDay, DueDate, FromAccount, SameAs, Category, Hidden, Active, LoginLink, Notes) " &
                        "VALUES (@Description, @AmtDue, @Cadence, @DueDay, @DueDate, @FromAccount, @SameAs, @Category, @Hidden, @Active, @LoginLink, @Notes);" &
                        "SELECT last_insert_rowid();"

                    cmd.Parameters.AddWithValue("@Description", If(expense.Description, String.Empty))
                    cmd.Parameters.AddWithValue("@AmtDue", expense.AmountDue)
                    cmd.Parameters.AddWithValue("@Cadence", If(expense.Cadence, String.Empty))

                    Dim dueDayParam = cmd.CreateParameter()
                    dueDayParam.ParameterName = "@DueDay"
                    dueDayParam.Value = If(expense.DueDay.HasValue, expense.DueDay.Value, DBNull.Value)
                    cmd.Parameters.Add(dueDayParam)

                    cmd.Parameters.AddWithValue("@DueDate", If(expense.DueDate, String.Empty))
                    cmd.Parameters.AddWithValue("@FromAccount", If(expense.FromAccount, String.Empty))
                    cmd.Parameters.AddWithValue("@SameAs", If(expense.SameAs, String.Empty))
                    cmd.Parameters.AddWithValue("@Category", If(expense.Category, String.Empty))
                    cmd.Parameters.AddWithValue("@Hidden", If(expense.Hidden, 1, 0))
                    cmd.Parameters.AddWithValue("@Active", If(expense.Active, 1, 0))
                    cmd.Parameters.AddWithValue("@LoginLink", If(expense.LoginLink, String.Empty))
                    cmd.Parameters.AddWithValue("@Notes", If(expense.Notes, String.Empty))

                    Return Convert.ToInt32(cmd.ExecuteScalar(), Globalization.CultureInfo.InvariantCulture)
                End Using
            End Using
        End Function

        Public Sub RestoreExpense(databasePath As String, expense As ExpenseRecord)
            If String.IsNullOrWhiteSpace(databasePath) OrElse Not IO.File.Exists(databasePath) Then
                Throw New IO.FileNotFoundException("Budget database was not found.", databasePath)
            End If

            If expense Is Nothing Then
                Throw New ArgumentNullException(NameOf(expense))
            End If

            If expense.Id <= 0 Then
                Throw New InvalidOperationException("A valid expense id is required for restore.")
            End If

            Using conn As New SqliteConnection("Data Source=" & databasePath.Trim())
                conn.Open()
                EnsureActiveColumn(conn)

                Using cmd = conn.CreateCommand()
                    cmd.CommandText =
                        "INSERT OR REPLACE INTO expenses (Id, Description, AmtDue, Cadence, DueDay, DueDate, FromAccount, SameAs, Category, Hidden, Active, LoginLink, Notes) " &
                        "VALUES (@Id, @Description, @AmtDue, @Cadence, @DueDay, @DueDate, @FromAccount, @SameAs, @Category, @Hidden, @Active, @LoginLink, @Notes)"

                    cmd.Parameters.AddWithValue("@Id", expense.Id)
                    cmd.Parameters.AddWithValue("@Description", If(expense.Description, String.Empty))
                    cmd.Parameters.AddWithValue("@AmtDue", expense.AmountDue)
                    cmd.Parameters.AddWithValue("@Cadence", If(expense.Cadence, String.Empty))

                    Dim dueDayParam = cmd.CreateParameter()
                    dueDayParam.ParameterName = "@DueDay"
                    dueDayParam.Value = If(expense.DueDay.HasValue, expense.DueDay.Value, DBNull.Value)
                    cmd.Parameters.Add(dueDayParam)

                    cmd.Parameters.AddWithValue("@DueDate", If(expense.DueDate, String.Empty))
                    cmd.Parameters.AddWithValue("@FromAccount", If(expense.FromAccount, String.Empty))
                    cmd.Parameters.AddWithValue("@SameAs", If(expense.SameAs, String.Empty))
                    cmd.Parameters.AddWithValue("@Category", If(expense.Category, String.Empty))
                    cmd.Parameters.AddWithValue("@Hidden", If(expense.Hidden, 1, 0))
                    cmd.Parameters.AddWithValue("@Active", If(expense.Active, 1, 0))
                    cmd.Parameters.AddWithValue("@LoginLink", If(expense.LoginLink, String.Empty))
                    cmd.Parameters.AddWithValue("@Notes", If(expense.Notes, String.Empty))
                    cmd.ExecuteNonQuery()
                End Using
            End Using
        End Sub

        Public Sub DeleteExpense(databasePath As String, expenseId As Integer)
            If String.IsNullOrWhiteSpace(databasePath) OrElse Not IO.File.Exists(databasePath) Then
                Throw New IO.FileNotFoundException("Budget database was not found.", databasePath)
            End If

            If expenseId <= 0 Then
                Throw New InvalidOperationException("A valid expense id is required for delete.")
            End If

            Using conn As New SqliteConnection("Data Source=" & databasePath.Trim())
                conn.Open()

                Using cmd = conn.CreateCommand()
                    cmd.CommandText = "DELETE FROM expenses WHERE Id = @Id"
                    cmd.Parameters.AddWithValue("@Id", expenseId)

                    Dim rowsChanged = cmd.ExecuteNonQuery()
                    If rowsChanged <= 0 Then
                        Throw New InvalidOperationException("The selected expense could not be deleted.")
                    End If
                End Using
            End Using
        End Sub

        Private Sub EnsureActiveColumn(conn As SqliteConnection)
            If Not HasColumn(conn, "expenses", "Active") Then
                Using addCmd = conn.CreateCommand()
                    addCmd.CommandText = "ALTER TABLE expenses ADD COLUMN Active INTEGER NOT NULL DEFAULT 1"
                    addCmd.ExecuteNonQuery()
                End Using
            End If

            If HasColumn(conn, "expenses", "Inactive") Then
                Using migrateCmd = conn.CreateCommand()
                    migrateCmd.CommandText = "UPDATE expenses SET Active = CASE WHEN COALESCE(Inactive, 0) <> 0 THEN 0 ELSE 1 END"
                    migrateCmd.ExecuteNonQuery()
                End Using

                RebuildTableWithoutColumn(conn, "expenses", "Inactive")
            End If
        End Sub

        Private Sub RebuildTableWithoutColumn(conn As SqliteConnection, tableName As String, removedColumnName As String)
            Dim columns = GetTableColumns(conn, tableName).
                Where(Function(column) Not String.Equals(column.Name, removedColumnName, StringComparison.OrdinalIgnoreCase)).
                ToList()

            If columns.Count = 0 Then
                Return
            End If

            Dim pkColumns = columns.
                Where(Function(column) column.PrimaryKeyOrder > 0).
                OrderBy(Function(column) column.PrimaryKeyOrder).
                ToList()

            Dim columnDefinitions = columns.Select(Function(column)
                                                       Dim parts As New List(Of String) From {
                                                           "[" & column.Name & "]"
                                                       }
                                                       If Not String.IsNullOrWhiteSpace(column.SqlType) Then
                                                           parts.Add(column.SqlType)
                                                       End If
                                                       If pkColumns.Count = 1 AndAlso column.PrimaryKeyOrder = 1 Then
                                                           parts.Add("PRIMARY KEY")
                                                       End If
                                                       If column.NotNull Then
                                                           parts.Add("NOT NULL")
                                                       End If
                                                       If Not String.IsNullOrWhiteSpace(column.DefaultValueSql) Then
                                                           parts.Add("DEFAULT " & column.DefaultValueSql)
                                                       End If
                                                       Return String.Join(" ", parts)
                                                   End Function).ToList()

            If pkColumns.Count > 1 Then
                columnDefinitions.Add("PRIMARY KEY (" & String.Join(", ", pkColumns.Select(Function(column) "[" & column.Name & "]")) & ")")
            End If

            Dim tempTableName = tableName & "__active_migration"
            Dim columnList = String.Join(", ", columns.Select(Function(column) "[" & column.Name & "]"))

            Using transaction = conn.BeginTransaction()
                Using pragmaCmd = conn.CreateCommand()
                    pragmaCmd.Transaction = transaction
                    pragmaCmd.CommandText = "PRAGMA foreign_keys=OFF"
                    pragmaCmd.ExecuteNonQuery()
                End Using

                Using dropTempCmd = conn.CreateCommand()
                    dropTempCmd.Transaction = transaction
                    dropTempCmd.CommandText = "DROP TABLE IF EXISTS [" & tempTableName & "]"
                    dropTempCmd.ExecuteNonQuery()
                End Using

                Using createCmd = conn.CreateCommand()
                    createCmd.Transaction = transaction
                    createCmd.CommandText = "CREATE TABLE [" & tempTableName & "] (" & String.Join(", ", columnDefinitions) & ")"
                    createCmd.ExecuteNonQuery()
                End Using

                Using copyCmd = conn.CreateCommand()
                    copyCmd.Transaction = transaction
                    copyCmd.CommandText = "INSERT INTO [" & tempTableName & "] (" & columnList & ") SELECT " & columnList & " FROM [" & tableName & "]"
                    copyCmd.ExecuteNonQuery()
                End Using

                Using dropOldCmd = conn.CreateCommand()
                    dropOldCmd.Transaction = transaction
                    dropOldCmd.CommandText = "DROP TABLE [" & tableName & "]"
                    dropOldCmd.ExecuteNonQuery()
                End Using

                Using renameCmd = conn.CreateCommand()
                    renameCmd.Transaction = transaction
                    renameCmd.CommandText = "ALTER TABLE [" & tempTableName & "] RENAME TO [" & tableName & "]"
                    renameCmd.ExecuteNonQuery()
                End Using

                Using pragmaCmd = conn.CreateCommand()
                    pragmaCmd.Transaction = transaction
                    pragmaCmd.CommandText = "PRAGMA foreign_keys=ON"
                    pragmaCmd.ExecuteNonQuery()
                End Using

                transaction.Commit()
            End Using
        End Sub

        Private Function GetTableColumns(conn As SqliteConnection, tableName As String) As List(Of SqliteTableColumnDefinition)
            Dim columns As New List(Of SqliteTableColumnDefinition)()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "PRAGMA table_info([" & tableName & "])"
                Using reader = cmd.ExecuteReader()
                    While reader.Read()
                        columns.Add(New SqliteTableColumnDefinition With {
                            .Name = If(reader.IsDBNull(1), String.Empty, Convert.ToString(reader.GetValue(1))),
                            .SqlType = If(reader.IsDBNull(2), String.Empty, Convert.ToString(reader.GetValue(2))),
                            .NotNull = Not reader.IsDBNull(3) AndAlso Convert.ToInt32(reader.GetValue(3)) <> 0,
                            .DefaultValueSql = If(reader.IsDBNull(4), String.Empty, Convert.ToString(reader.GetValue(4))),
                            .PrimaryKeyOrder = If(reader.IsDBNull(5), 0, Convert.ToInt32(reader.GetValue(5)))
                        })
                    End While
                End Using
            End Using

            Return columns
        End Function

        Private Function HasColumn(conn As SqliteConnection, tableName As String, columnName As String) As Boolean
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

    End Module

End Namespace

