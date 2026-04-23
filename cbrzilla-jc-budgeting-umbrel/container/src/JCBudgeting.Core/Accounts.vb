Imports Microsoft.Data.Sqlite

Namespace Global.JCBudgeting.Core

    Public Class AccountRecord
        Public Property Id As Integer
        Public Property Description As String
        Public Property AccountType As String
        Public Property LoginLink As String
        Public Property Notes As String
        Public Property SafetyNet As Decimal
        Public Property Hidden As Boolean
        Public Property Active As Boolean
    End Class

    Public Module AccountRepository
        Private NotInheritable Class SqliteTableColumnDefinition
            Public Property Name As String = String.Empty
            Public Property SqlType As String = String.Empty
            Public Property NotNull As Boolean
            Public Property DefaultValueSql As String = String.Empty
            Public Property PrimaryKeyOrder As Integer
        End Class

        Public Function LoadAccounts(databasePath As String) As List(Of AccountRecord)
            Dim results As New List(Of AccountRecord)()
            If String.IsNullOrWhiteSpace(databasePath) OrElse Not IO.File.Exists(databasePath) Then
                Return results
            End If

            Using conn As New SqliteConnection("Data Source=" & databasePath.Trim())
                conn.Open()
                EnsureActiveColumn(conn)

                Using cmd = conn.CreateCommand()
                    cmd.CommandText =
                        "SELECT Id, Description, Type, LoginLink, Notes, SafetyNet, Hidden, COALESCE(Active,1) FROM accounts ORDER BY COALESCE(Hidden,0), Description COLLATE NOCASE"
                    Using reader = cmd.ExecuteReader()
                        While reader.Read()
                            results.Add(New AccountRecord With {
                                .Id = If(reader.IsDBNull(0), 0, reader.GetInt32(0)),
                                .Description = If(reader.IsDBNull(1), String.Empty, reader.GetString(1)),
                                .AccountType = If(reader.IsDBNull(2), String.Empty, reader.GetString(2)),
                                .LoginLink = If(reader.IsDBNull(3), String.Empty, reader.GetString(3)),
                                .Notes = If(reader.IsDBNull(4), String.Empty, reader.GetString(4)),
                                .SafetyNet = If(reader.IsDBNull(5), 0D, Convert.ToDecimal(reader.GetValue(5))),
                                .Hidden = Not reader.IsDBNull(6) AndAlso Convert.ToInt32(reader.GetValue(6)) <> 0,
                                .Active = Not reader.IsDBNull(7) AndAlso Convert.ToInt32(reader.GetValue(7)) <> 0
                            })
                        End While
                    End Using
                End Using
            End Using

            Return results
        End Function

        Public Sub SaveAccount(databasePath As String, account As AccountRecord)
            If String.IsNullOrWhiteSpace(databasePath) OrElse Not IO.File.Exists(databasePath) Then
                Throw New IO.FileNotFoundException("Budget database was not found.", databasePath)
            End If

            If account Is Nothing Then
                Throw New ArgumentNullException(NameOf(account))
            End If

            If account.Id <= 0 Then
                Throw New InvalidOperationException("Only existing accounts can be saved in this migration pass.")
            End If

            Using conn As New SqliteConnection("Data Source=" & databasePath.Trim())
                conn.Open()
                EnsureActiveColumn(conn)

                Using cmd = conn.CreateCommand()
                    cmd.CommandText = "UPDATE accounts " &
                                      "SET Description = @Description, " &
                                      "Type = @Type, " &
                                      "LoginLink = @LoginLink, " &
                                      "Notes = @Notes, " &
                                      "SafetyNet = @SafetyNet, " &
                                      "Hidden = @Hidden, " &
                                      "Active = @Active " &
                                      "WHERE Id = @Id"

                    cmd.Parameters.AddWithValue("@Description", If(account.Description, String.Empty))
                    cmd.Parameters.AddWithValue("@Type", If(account.AccountType, String.Empty))
                    cmd.Parameters.AddWithValue("@LoginLink", If(account.LoginLink, String.Empty))
                    cmd.Parameters.AddWithValue("@Notes", If(account.Notes, String.Empty))
                    cmd.Parameters.AddWithValue("@SafetyNet", account.SafetyNet)
                    cmd.Parameters.AddWithValue("@Hidden", If(account.Hidden, 1, 0))
                    cmd.Parameters.AddWithValue("@Active", If(account.Active, 1, 0))
                    cmd.Parameters.AddWithValue("@Id", account.Id)

                    Dim rowsChanged = cmd.ExecuteNonQuery()
                    If rowsChanged <= 0 Then
                        Throw New InvalidOperationException("The selected account could not be saved.")
                    End If
                End Using
            End Using
        End Sub

        Public Sub RenameAccountReferences(databasePath As String, previousDescription As String, newDescription As String)
            If String.IsNullOrWhiteSpace(databasePath) OrElse Not IO.File.Exists(databasePath) Then
                Throw New IO.FileNotFoundException("Budget database was not found.", databasePath)
            End If

            Dim previousValue = If(previousDescription, String.Empty).Trim()
            Dim newValue = If(newDescription, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(previousValue) OrElse
               String.IsNullOrWhiteSpace(newValue) OrElse
               String.Equals(previousValue, newValue, StringComparison.CurrentCulture) Then
                Return
            End If

            Using conn As New SqliteConnection("Data Source=" & databasePath.Trim())
                conn.Open()

                Using transaction = conn.BeginTransaction()
                    RenameAccountReferenceColumn(conn, transaction, "expenses", "FromAccount", previousValue, newValue)
                    RenameAccountReferenceColumn(conn, transaction, "savings", "FromAccount", previousValue, newValue)
                    RenameAccountReferenceColumn(conn, transaction, "debts", "FromAccount", previousValue, newValue)
                    RenameAccountReferenceColumn(conn, transaction, "income", "ToAccount", previousValue, newValue)
                    transaction.Commit()
                End Using
            End Using
        End Sub

        Public Function CreateAccount(databasePath As String, account As AccountRecord) As Integer
            If String.IsNullOrWhiteSpace(databasePath) OrElse Not IO.File.Exists(databasePath) Then
                Throw New IO.FileNotFoundException("Budget database was not found.", databasePath)
            End If

            If account Is Nothing Then
                Throw New ArgumentNullException(NameOf(account))
            End If

            Using conn As New SqliteConnection("Data Source=" & databasePath.Trim())
                conn.Open()
                EnsureActiveColumn(conn)

                Using cmd = conn.CreateCommand()
                    cmd.CommandText = "INSERT INTO accounts (Description, Type, LoginLink, Notes, SafetyNet, Hidden, Active) " &
                                      "VALUES (@Description, @Type, @LoginLink, @Notes, @SafetyNet, @Hidden, @Active);" &
                                      "SELECT last_insert_rowid();"

                    cmd.Parameters.AddWithValue("@Description", If(account.Description, String.Empty))
                    cmd.Parameters.AddWithValue("@Type", If(account.AccountType, String.Empty))
                    cmd.Parameters.AddWithValue("@LoginLink", If(account.LoginLink, String.Empty))
                    cmd.Parameters.AddWithValue("@Notes", If(account.Notes, String.Empty))
                    cmd.Parameters.AddWithValue("@SafetyNet", account.SafetyNet)
                    cmd.Parameters.AddWithValue("@Hidden", If(account.Hidden, 1, 0))
                    cmd.Parameters.AddWithValue("@Active", If(account.Active, 1, 0))

                    Dim createdId = cmd.ExecuteScalar()
                    If createdId Is Nothing Then
                        Throw New InvalidOperationException("The new account could not be created.")
                    End If

                    Return Convert.ToInt32(createdId)
                End Using
            End Using
        End Function

        Public Sub DeleteAccount(databasePath As String, accountId As Integer)
            If String.IsNullOrWhiteSpace(databasePath) OrElse Not IO.File.Exists(databasePath) Then
                Throw New IO.FileNotFoundException("Budget database was not found.", databasePath)
            End If

            If accountId <= 0 Then
                Throw New ArgumentOutOfRangeException(NameOf(accountId))
            End If

            Using conn As New SqliteConnection("Data Source=" & databasePath.Trim())
                conn.Open()

                Using cmd = conn.CreateCommand()
                    cmd.CommandText = "DELETE FROM accounts WHERE Id = @Id"
                    cmd.Parameters.AddWithValue("@Id", accountId)

                    Dim rowsChanged = cmd.ExecuteNonQuery()
                    If rowsChanged <= 0 Then
                        Throw New InvalidOperationException("The selected account could not be deleted.")
                    End If
                End Using
            End Using
        End Sub

        Public Sub RestoreAccount(databasePath As String, account As AccountRecord)
            If String.IsNullOrWhiteSpace(databasePath) OrElse Not IO.File.Exists(databasePath) Then
                Throw New IO.FileNotFoundException("Budget database was not found.", databasePath)
            End If

            If account Is Nothing Then
                Throw New ArgumentNullException(NameOf(account))
            End If

            If account.Id <= 0 Then
                Throw New InvalidOperationException("A valid account id is required for restore.")
            End If

            Using conn As New SqliteConnection("Data Source=" & databasePath.Trim())
                conn.Open()
                EnsureActiveColumn(conn)

                Using cmd = conn.CreateCommand()
                    cmd.CommandText = "INSERT OR REPLACE INTO accounts (Id, Description, Type, LoginLink, Notes, SafetyNet, Hidden, Active) " &
                                      "VALUES (@Id, @Description, @Type, @LoginLink, @Notes, @SafetyNet, @Hidden, @Active)"

                    cmd.Parameters.AddWithValue("@Id", account.Id)
                    cmd.Parameters.AddWithValue("@Description", If(account.Description, String.Empty))
                    cmd.Parameters.AddWithValue("@Type", If(account.AccountType, String.Empty))
                    cmd.Parameters.AddWithValue("@LoginLink", If(account.LoginLink, String.Empty))
                    cmd.Parameters.AddWithValue("@Notes", If(account.Notes, String.Empty))
                    cmd.Parameters.AddWithValue("@SafetyNet", account.SafetyNet)
                    cmd.Parameters.AddWithValue("@Hidden", If(account.Hidden, 1, 0))
                    cmd.Parameters.AddWithValue("@Active", If(account.Active, 1, 0))
                    cmd.ExecuteNonQuery()
                End Using
            End Using
        End Sub

        Private Sub EnsureActiveColumn(conn As SqliteConnection)
            If Not HasColumn(conn, "accounts", "Active") Then
                Using addCmd = conn.CreateCommand()
                    addCmd.CommandText = "ALTER TABLE accounts ADD COLUMN Active INTEGER NOT NULL DEFAULT 1"
                    addCmd.ExecuteNonQuery()
                End Using
            End If

            If HasColumn(conn, "accounts", "Inactive") Then
                Using migrateCmd = conn.CreateCommand()
                    migrateCmd.CommandText = "UPDATE accounts SET Active = CASE WHEN COALESCE(Inactive, 0) <> 0 THEN 0 ELSE 1 END"
                    migrateCmd.ExecuteNonQuery()
                End Using

                RebuildTableWithoutColumn(conn, "accounts", "Inactive")
            End If
        End Sub

        Private Sub RenameAccountReferenceColumn(
            conn As SqliteConnection,
            transaction As SqliteTransaction,
            tableName As String,
            columnName As String,
            previousValue As String,
            newValue As String)

            If Not HasColumn(conn, tableName, columnName) Then
                Return
            End If

            Using cmd = conn.CreateCommand()
                cmd.Transaction = transaction
                cmd.CommandText = $"UPDATE [{tableName}] SET [{columnName}] = @NewValue WHERE TRIM(COALESCE([{columnName}], '')) = @PreviousValue"
                cmd.Parameters.AddWithValue("@NewValue", newValue)
                cmd.Parameters.AddWithValue("@PreviousValue", previousValue)
                cmd.ExecuteNonQuery()
            End Using
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

