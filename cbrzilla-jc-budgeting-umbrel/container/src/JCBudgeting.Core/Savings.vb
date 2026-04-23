Imports Microsoft.Data.Sqlite

Namespace Global.JCBudgeting.Core

    Public Class SavingsRecord
        Public Property Id As Integer
        Public Property Description As String
        Public Property GoalAmount As Decimal
        Public Property DepositAmount As Decimal
        Public Property GoalDate As String
        Public Property HasGoal As Boolean
        Public Property Frequency As String
        Public Property OnDay As Integer?
        Public Property OnDate As String
        Public Property StartDate As String
        Public Property EndDate As String
        Public Property FromAccount As String
        Public Property SameAs As String
        Public Property Category As String
        Public Property Hidden As Boolean
        Public Property Active As Boolean
        Public Property LoginLink As String
        Public Property Notes As String
    End Class

    Public Module SavingsRepository

        Public Function LoadSavings(databasePath As String) As List(Of SavingsRecord)
            Dim results As New List(Of SavingsRecord)()
            If String.IsNullOrWhiteSpace(databasePath) OrElse Not IO.File.Exists(databasePath) Then
                Return results
            End If

            Using conn As New SqliteConnection("Data Source=" & databasePath.Trim())
                conn.Open()
                EnsureSavingsSchema(conn)

                Using cmd = conn.CreateCommand()
                    cmd.CommandText = "SELECT Id, Description, GoalAmount, DepositAmount, GoalDate, HasGoal, Frequency, OnDay, OnDate, StartDate, EndDate, FromAccount, SameAs, Category, Hidden, COALESCE(Active,1), LoginLink, Notes FROM savings ORDER BY COALESCE(Hidden,0), COALESCE(Active,1) DESC, Category COLLATE NOCASE, Description COLLATE NOCASE"
                    Using reader = cmd.ExecuteReader()
                        While reader.Read()
                            results.Add(New SavingsRecord With {
                                .Id = If(reader.IsDBNull(0), 0, reader.GetInt32(0)),
                                .Description = If(reader.IsDBNull(1), String.Empty, reader.GetString(1)),
                                .GoalAmount = If(reader.IsDBNull(2), 0D, Convert.ToDecimal(reader.GetValue(2))),
                                .DepositAmount = If(reader.IsDBNull(3), 0D, Convert.ToDecimal(reader.GetValue(3))),
                                .GoalDate = If(reader.IsDBNull(4), String.Empty, Convert.ToString(reader.GetValue(4))),
                                .HasGoal = Not reader.IsDBNull(5) AndAlso Convert.ToInt32(reader.GetValue(5)) <> 0,
                                .Frequency = If(reader.IsDBNull(6), String.Empty, reader.GetString(6)),
                                .OnDay = If(reader.IsDBNull(7), CType(Nothing, Integer?), Convert.ToInt32(reader.GetValue(7))),
                                .OnDate = If(reader.IsDBNull(8), String.Empty, Convert.ToString(reader.GetValue(8))),
                                .StartDate = If(reader.IsDBNull(9), String.Empty, Convert.ToString(reader.GetValue(9))),
                                .EndDate = If(reader.IsDBNull(10), String.Empty, Convert.ToString(reader.GetValue(10))),
                                .FromAccount = If(reader.IsDBNull(11), String.Empty, Convert.ToString(reader.GetValue(11))),
                                .SameAs = If(reader.IsDBNull(12), String.Empty, Convert.ToString(reader.GetValue(12))),
                                .Category = If(reader.IsDBNull(13), String.Empty, Convert.ToString(reader.GetValue(13))),
                                .Hidden = Not reader.IsDBNull(14) AndAlso Convert.ToInt32(reader.GetValue(14)) <> 0,
                                .Active = Not reader.IsDBNull(15) AndAlso Convert.ToInt32(reader.GetValue(15)) <> 0,
                                .LoginLink = If(reader.IsDBNull(16), String.Empty, Convert.ToString(reader.GetValue(16))),
                                .Notes = If(reader.IsDBNull(17), String.Empty, Convert.ToString(reader.GetValue(17)))
                            })
                        End While
                    End Using
                End Using
            End Using

            Return results
        End Function

        Public Sub SaveSavings(databasePath As String, savings As SavingsRecord)
            If String.IsNullOrWhiteSpace(databasePath) OrElse Not IO.File.Exists(databasePath) Then
                Throw New IO.FileNotFoundException("Budget database was not found.", databasePath)
            End If

            If savings Is Nothing Then
                Throw New ArgumentNullException(NameOf(savings))
            End If

            If savings.Id <= 0 Then
                Throw New InvalidOperationException("Only existing savings items can be saved in this migration pass.")
            End If

            Using conn As New SqliteConnection("Data Source=" & databasePath.Trim())
                conn.Open()
                EnsureSavingsSchema(conn)

                Using cmd = conn.CreateCommand()
                    cmd.CommandText =
                        "UPDATE savings " &
                        "SET Description = @Description, " &
                        "GoalAmount = @GoalAmount, " &
                        "DepositAmount = @DepositAmount, " &
                        "GoalDate = @GoalDate, " &
                        "HasGoal = @HasGoal, " &
                        "Frequency = @Frequency, " &
                        "OnDay = @OnDay, " &
                        "OnDate = @OnDate, " &
                        "StartDate = @StartDate, " &
                        "EndDate = @EndDate, " &
                        "FromAccount = @FromAccount, " &
                        "SameAs = @SameAs, " &
                        "Category = @Category, " &
                        "Hidden = @Hidden, " &
                        "Active = @Active, " &
                        "LoginLink = @LoginLink, " &
                        "Notes = @Notes " &
                        "WHERE Id = @Id"

                    cmd.Parameters.AddWithValue("@Description", If(savings.Description, String.Empty))
                    cmd.Parameters.AddWithValue("@GoalAmount", savings.GoalAmount)
                    cmd.Parameters.AddWithValue("@DepositAmount", savings.DepositAmount)
                    cmd.Parameters.AddWithValue("@GoalDate", If(savings.GoalDate, String.Empty))
                    cmd.Parameters.AddWithValue("@HasGoal", If(savings.HasGoal, 1, 0))
                    cmd.Parameters.AddWithValue("@Frequency", If(savings.Frequency, String.Empty))
                    cmd.Parameters.AddWithValue("@OnDay", If(savings.OnDay.HasValue, savings.OnDay.Value, CType(DBNull.Value, Object)))
                    cmd.Parameters.AddWithValue("@OnDate", If(savings.OnDate, String.Empty))
                    cmd.Parameters.AddWithValue("@StartDate", If(savings.StartDate, String.Empty))
                    cmd.Parameters.AddWithValue("@EndDate", If(savings.EndDate, String.Empty))
                    cmd.Parameters.AddWithValue("@FromAccount", If(savings.FromAccount, String.Empty))
                    cmd.Parameters.AddWithValue("@SameAs", If(savings.SameAs, String.Empty))
                    cmd.Parameters.AddWithValue("@Category", If(savings.Category, String.Empty))
                    cmd.Parameters.AddWithValue("@Hidden", If(savings.Hidden, 1, 0))
                    cmd.Parameters.AddWithValue("@Active", If(savings.Active, 1, 0))
                    cmd.Parameters.AddWithValue("@LoginLink", If(savings.LoginLink, String.Empty))
                    cmd.Parameters.AddWithValue("@Notes", If(savings.Notes, String.Empty))
                    cmd.Parameters.AddWithValue("@Id", savings.Id)

                    Dim rowsChanged = cmd.ExecuteNonQuery()
                    If rowsChanged <= 0 Then
                        Throw New InvalidOperationException("The selected savings item could not be saved.")
                    End If
                End Using
            End Using
        End Sub

        Public Function CreateSavings(databasePath As String, savings As SavingsRecord) As Integer
            If String.IsNullOrWhiteSpace(databasePath) OrElse Not IO.File.Exists(databasePath) Then
                Throw New IO.FileNotFoundException("Budget database was not found.", databasePath)
            End If

            If savings Is Nothing Then
                Throw New ArgumentNullException(NameOf(savings))
            End If

            Using conn As New SqliteConnection("Data Source=" & databasePath.Trim())
                conn.Open()
                EnsureSavingsSchema(conn)

                Using cmd = conn.CreateCommand()
                    cmd.CommandText =
                        "INSERT INTO savings " &
                        "(Description, GoalAmount, DepositAmount, GoalDate, HasGoal, Frequency, OnDay, OnDate, StartDate, EndDate, FromAccount, SameAs, Category, Hidden, Active, LoginLink, Notes) " &
                        "VALUES (@Description, @GoalAmount, @DepositAmount, @GoalDate, @HasGoal, @Frequency, @OnDay, @OnDate, @StartDate, @EndDate, @FromAccount, @SameAs, @Category, @Hidden, @Active, @LoginLink, @Notes);" &
                        "SELECT last_insert_rowid();"

                    cmd.Parameters.AddWithValue("@Description", If(savings.Description, String.Empty))
                    cmd.Parameters.AddWithValue("@GoalAmount", savings.GoalAmount)
                    cmd.Parameters.AddWithValue("@DepositAmount", savings.DepositAmount)
                    cmd.Parameters.AddWithValue("@GoalDate", If(savings.GoalDate, String.Empty))
                    cmd.Parameters.AddWithValue("@HasGoal", If(savings.HasGoal, 1, 0))
                    cmd.Parameters.AddWithValue("@Frequency", If(savings.Frequency, String.Empty))
                    cmd.Parameters.AddWithValue("@OnDay", If(savings.OnDay.HasValue, savings.OnDay.Value, CType(DBNull.Value, Object)))
                    cmd.Parameters.AddWithValue("@OnDate", If(savings.OnDate, String.Empty))
                    cmd.Parameters.AddWithValue("@StartDate", If(savings.StartDate, String.Empty))
                    cmd.Parameters.AddWithValue("@EndDate", If(savings.EndDate, String.Empty))
                    cmd.Parameters.AddWithValue("@FromAccount", If(savings.FromAccount, String.Empty))
                    cmd.Parameters.AddWithValue("@SameAs", If(savings.SameAs, String.Empty))
                    cmd.Parameters.AddWithValue("@Category", If(savings.Category, String.Empty))
                    cmd.Parameters.AddWithValue("@Hidden", If(savings.Hidden, 1, 0))
                    cmd.Parameters.AddWithValue("@Active", If(savings.Active, 1, 0))
                    cmd.Parameters.AddWithValue("@LoginLink", If(savings.LoginLink, String.Empty))
                    cmd.Parameters.AddWithValue("@Notes", If(savings.Notes, String.Empty))

                    Return Convert.ToInt32(cmd.ExecuteScalar())
                End Using
            End Using
        End Function

        Public Sub DeleteSavings(databasePath As String, savingsId As Integer)
            If String.IsNullOrWhiteSpace(databasePath) OrElse Not IO.File.Exists(databasePath) Then
                Throw New IO.FileNotFoundException("Budget database was not found.", databasePath)
            End If

            If savingsId <= 0 Then
                Throw New ArgumentOutOfRangeException(NameOf(savingsId))
            End If

            Using conn As New SqliteConnection("Data Source=" & databasePath.Trim())
                conn.Open()

                Using cmd = conn.CreateCommand()
                    cmd.CommandText = "DELETE FROM savings WHERE Id = @Id"
                    cmd.Parameters.AddWithValue("@Id", savingsId)

                    Dim rowsChanged = cmd.ExecuteNonQuery()
                    If rowsChanged <= 0 Then
                        Throw New InvalidOperationException("The selected savings item could not be deleted.")
                    End If
                End Using
            End Using
        End Sub

        Public Sub RestoreSavings(databasePath As String, savings As SavingsRecord)
            If String.IsNullOrWhiteSpace(databasePath) OrElse Not IO.File.Exists(databasePath) Then
                Throw New IO.FileNotFoundException("Budget database was not found.", databasePath)
            End If

            If savings Is Nothing Then
                Throw New ArgumentNullException(NameOf(savings))
            End If

            If savings.Id <= 0 Then
                Throw New InvalidOperationException("A valid savings id is required for restore.")
            End If

            Using conn As New SqliteConnection("Data Source=" & databasePath.Trim())
                conn.Open()
                EnsureSavingsSchema(conn)

                Using cmd = conn.CreateCommand()
                    cmd.CommandText =
                        "INSERT OR REPLACE INTO savings " &
                        "(Id, Description, GoalAmount, DepositAmount, GoalDate, HasGoal, Frequency, OnDay, OnDate, StartDate, EndDate, FromAccount, SameAs, Category, Hidden, Active, LoginLink, Notes) " &
                        "VALUES (@Id, @Description, @GoalAmount, @DepositAmount, @GoalDate, @HasGoal, @Frequency, @OnDay, @OnDate, @StartDate, @EndDate, @FromAccount, @SameAs, @Category, @Hidden, @Active, @LoginLink, @Notes)"

                    cmd.Parameters.AddWithValue("@Id", savings.Id)
                    cmd.Parameters.AddWithValue("@Description", If(savings.Description, String.Empty))
                    cmd.Parameters.AddWithValue("@GoalAmount", savings.GoalAmount)
                    cmd.Parameters.AddWithValue("@DepositAmount", savings.DepositAmount)
                    cmd.Parameters.AddWithValue("@GoalDate", If(savings.GoalDate, String.Empty))
                    cmd.Parameters.AddWithValue("@HasGoal", If(savings.HasGoal, 1, 0))
                    cmd.Parameters.AddWithValue("@Frequency", If(savings.Frequency, String.Empty))
                    cmd.Parameters.AddWithValue("@OnDay", If(savings.OnDay.HasValue, savings.OnDay.Value, CType(DBNull.Value, Object)))
                    cmd.Parameters.AddWithValue("@OnDate", If(savings.OnDate, String.Empty))
                    cmd.Parameters.AddWithValue("@StartDate", If(savings.StartDate, String.Empty))
                    cmd.Parameters.AddWithValue("@EndDate", If(savings.EndDate, String.Empty))
                    cmd.Parameters.AddWithValue("@FromAccount", If(savings.FromAccount, String.Empty))
                    cmd.Parameters.AddWithValue("@SameAs", If(savings.SameAs, String.Empty))
                    cmd.Parameters.AddWithValue("@Category", If(savings.Category, String.Empty))
                    cmd.Parameters.AddWithValue("@Hidden", If(savings.Hidden, 1, 0))
                    cmd.Parameters.AddWithValue("@Active", If(savings.Active, 1, 0))
                    cmd.Parameters.AddWithValue("@LoginLink", If(savings.LoginLink, String.Empty))
                    cmd.Parameters.AddWithValue("@Notes", If(savings.Notes, String.Empty))
                    cmd.ExecuteNonQuery()
                End Using
            End Using
        End Sub

        Private Sub EnsureSavingsSchema(conn As SqliteConnection)
            Dim columnNames As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            Using pragma = conn.CreateCommand()
                pragma.CommandText = "PRAGMA table_info(savings)"
                Using reader = pragma.ExecuteReader()
                    While reader.Read()
                        columnNames.Add(Convert.ToString(reader("name")))
                    End While
                End Using
            End Using

            Dim requiresRebuild =
                columnNames.Contains("Type") OrElse
                Not columnNames.Contains("Category") OrElse
                Not columnNames.Contains("HasGoal") OrElse
                Not columnNames.Contains("Active") OrElse
                columnNames.Contains("Inactive")

            If Not requiresRebuild Then
                Return
            End If

            Dim categorySource = If(columnNames.Contains("Category"), "COALESCE(Category, '')", If(columnNames.Contains("Type"), "COALESCE(Type, '')", "''"))
            Dim activeSource = If(columnNames.Contains("Inactive"),
                                  "CASE WHEN COALESCE(Inactive, 0) <> 0 THEN 0 ELSE 1 END",
                                  If(columnNames.Contains("Active"), "COALESCE(Active, 1)", "1"))

            Using tx = conn.BeginTransaction()
                Using create = conn.CreateCommand()
                    create.Transaction = tx
                    create.CommandText =
                        "CREATE TABLE savings_new (" &
                        "Id INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT, " &
                        "Description TEXT NULL, " &
                        "GoalAmount REAL NOT NULL DEFAULT 0, " &
                        "DepositAmount REAL NOT NULL DEFAULT 0, " &
                        "GoalDate TEXT NULL, " &
                        "HasGoal INTEGER NOT NULL DEFAULT 0, " &
                        "Frequency TEXT NULL, " &
                        "OnDay INTEGER NULL, " &
                        "OnDate TEXT NULL, " &
                        "StartDate TEXT NULL, " &
                        "EndDate TEXT NULL, " &
                        "FromAccount TEXT NULL, " &
                        "SameAs TEXT NULL, " &
                        "Category TEXT NULL, " &
                        "Hidden INTEGER NOT NULL DEFAULT 0, " &
                        "Active INTEGER NOT NULL DEFAULT 1, " &
                        "LoginLink TEXT NULL, " &
                        "Notes TEXT NULL" &
                        ")"
                    create.ExecuteNonQuery()
                End Using

                Using copy = conn.CreateCommand()
                    copy.Transaction = tx
                    copy.CommandText =
                        "INSERT INTO savings_new (Id, Description, GoalAmount, DepositAmount, GoalDate, HasGoal, Frequency, OnDay, OnDate, StartDate, EndDate, FromAccount, SameAs, Category, Hidden, Active, LoginLink, Notes) " &
                        "SELECT Id, COALESCE(Description, ''), COALESCE(GoalAmount, 0), COALESCE(DepositAmount, 0), COALESCE(GoalDate, ''), " &
                        "CASE WHEN COALESCE(GoalAmount, 0) <> 0 OR COALESCE(GoalDate, '') <> '' THEN 1 ELSE 0 END, " &
                        "COALESCE(Frequency, ''), OnDay, COALESCE(OnDate, ''), COALESCE(StartDate, ''), COALESCE(EndDate, ''), COALESCE(FromAccount, ''), COALESCE(SameAs, ''), " & categorySource & ", COALESCE(Hidden, 0), " & activeSource & ", COALESCE(LoginLink, ''), COALESCE(Notes, '') " &
                        "FROM savings"
                    copy.ExecuteNonQuery()
                End Using

                Using dropOld = conn.CreateCommand()
                    dropOld.Transaction = tx
                    dropOld.CommandText = "DROP TABLE savings"
                    dropOld.ExecuteNonQuery()
                End Using

                Using rename = conn.CreateCommand()
                    rename.Transaction = tx
                    rename.CommandText = "ALTER TABLE savings_new RENAME TO savings"
                    rename.ExecuteNonQuery()
                End Using

                tx.Commit()
            End Using
        End Sub

    End Module

End Namespace

