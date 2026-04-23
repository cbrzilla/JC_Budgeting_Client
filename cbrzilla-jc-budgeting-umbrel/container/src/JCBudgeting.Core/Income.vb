Imports System
Imports System.Collections.Generic
Imports Microsoft.Data.Sqlite

Namespace Global.JCBudgeting.Core
    Public Class IncomeRecord
        Public Property Id As Integer
        Public Property Description As String = String.Empty
        Public Property Amount As Decimal
        Public Property Cadence As String = String.Empty
        Public Property OnDay As Integer?
        Public Property OnDate As String = String.Empty
        Public Property AutoIncrease As Decimal
        Public Property AutoIncreaseOnDate As String = String.Empty
        Public Property StartDate As String = String.Empty
        Public Property EndDate As String = String.Empty
        Public Property ToAccount As String = String.Empty
        Public Property SameAs As String = String.Empty
        Public Property Hidden As Boolean
        Public Property Active As Boolean = True
        Public Property LoginLink As String = String.Empty
        Public Property Notes As String = String.Empty
    End Class

    Public NotInheritable Class IncomeRepository
        Private Sub New()
        End Sub

        Public Shared Function LoadIncome(databasePath As String) As IReadOnlyList(Of IncomeRecord)
            Dim items As New List(Of IncomeRecord)()

            Using conn As New SqliteConnection($"Data Source={databasePath}")
                conn.Open()
                EnsureActiveColumn(conn)
                EnsureSameAsColumn(conn)
                EnsureSameAsColumn(conn)
                EnsureSameAsColumn(conn)
                EnsureSameAsColumn(conn)

                Using cmd = conn.CreateCommand()
                    cmd.CommandText = "SELECT Id, Description, Amount, Cadence, OnDay, OnDate, autoincrease, AutoIncreaseOnDate, StartDate, EndDate, ToAccount, SameAs, Hidden, COALESCE(Active,1), LoginLink, Notes FROM income ORDER BY COALESCE(Hidden,0), Description COLLATE NOCASE"

                    Using reader = cmd.ExecuteReader()
                        While reader.Read()
                            Dim item As New IncomeRecord() With {
                                .Id = If(reader.IsDBNull(0), 0, reader.GetInt32(0)),
                                .Description = If(reader.IsDBNull(1), String.Empty, reader.GetString(1)),
                                .Amount = If(reader.IsDBNull(2), 0D, reader.GetDecimal(2)),
                                .Cadence = If(reader.IsDBNull(3), String.Empty, reader.GetString(3)),
                                .OnDate = If(reader.IsDBNull(5), String.Empty, reader.GetString(5)),
                                .AutoIncrease = If(reader.IsDBNull(6), 0D, reader.GetDecimal(6)),
                                .AutoIncreaseOnDate = If(reader.IsDBNull(7), String.Empty, reader.GetString(7)),
                                .StartDate = If(reader.IsDBNull(8), String.Empty, reader.GetString(8)),
                                .EndDate = If(reader.IsDBNull(9), String.Empty, reader.GetString(9)),
                                .ToAccount = If(reader.IsDBNull(10), String.Empty, reader.GetString(10)),
                                .SameAs = If(reader.IsDBNull(11), String.Empty, reader.GetString(11)),
                                .Hidden = Not reader.IsDBNull(12) AndAlso reader.GetInt64(12) <> 0,
                                .Active = Not reader.IsDBNull(13) AndAlso reader.GetInt64(13) <> 0,
                                .LoginLink = If(reader.IsDBNull(14), String.Empty, reader.GetString(14)),
                                .Notes = If(reader.IsDBNull(15), String.Empty, reader.GetString(15))
                            }

                            If Not reader.IsDBNull(4) Then
                                item.OnDay = reader.GetInt32(4)
                            End If

                            items.Add(item)
                        End While
                    End Using
                End Using
            End Using

            Return items
        End Function

        Public Shared Sub SaveIncome(databasePath As String, item As IncomeRecord)
            If String.IsNullOrWhiteSpace(databasePath) OrElse Not IO.File.Exists(databasePath) Then
                Throw New IO.FileNotFoundException("Budget database was not found.", databasePath)
            End If

            If item Is Nothing Then
                Throw New ArgumentNullException(NameOf(item))
            End If

            If item.Id <= 0 Then
                Throw New InvalidOperationException("Only existing income items can be saved in this migration pass.")
            End If

            Using conn As New SqliteConnection($"Data Source={databasePath}")
                conn.Open()
                EnsureActiveColumn(conn)

                Using cmd = conn.CreateCommand()
                    cmd.CommandText =
                        "UPDATE income " &
                        "SET Description = @Description, " &
                        "Amount = @Amount, " &
                        "Cadence = @Cadence, " &
                        "OnDay = @OnDay, " &
                        "OnDate = @OnDate, " &
                        "AutoIncrease = @AutoIncrease, " &
                        "AutoIncreaseOnDate = @AutoIncreaseOnDate, " &
                        "StartDate = @StartDate, " &
                        "EndDate = @EndDate, " &
                        "ToAccount = @ToAccount, " &
                        "SameAs = @SameAs, " &
                        "Hidden = @Hidden, " &
                        "Active = @Active, " &
                        "LoginLink = @LoginLink, " &
                        "Notes = @Notes " &
                        "WHERE Id = @Id"

                    cmd.Parameters.AddWithValue("@Description", If(item.Description, String.Empty))
                    cmd.Parameters.AddWithValue("@Amount", item.Amount)
                    cmd.Parameters.AddWithValue("@Cadence", If(item.Cadence, String.Empty))

                    Dim onDayParam = cmd.CreateParameter()
                    onDayParam.ParameterName = "@OnDay"
                    onDayParam.Value = If(item.OnDay.HasValue, item.OnDay.Value, DBNull.Value)
                    cmd.Parameters.Add(onDayParam)

                    cmd.Parameters.AddWithValue("@OnDate", If(item.OnDate, String.Empty))
                    cmd.Parameters.AddWithValue("@AutoIncrease", item.AutoIncrease)
                    cmd.Parameters.AddWithValue("@AutoIncreaseOnDate", If(item.AutoIncreaseOnDate, String.Empty))
                    cmd.Parameters.AddWithValue("@StartDate", If(item.StartDate, String.Empty))
                    cmd.Parameters.AddWithValue("@EndDate", If(item.EndDate, String.Empty))
                    cmd.Parameters.AddWithValue("@ToAccount", If(item.ToAccount, String.Empty))
                    cmd.Parameters.AddWithValue("@SameAs", If(item.SameAs, String.Empty))
                    cmd.Parameters.AddWithValue("@Hidden", If(item.Hidden, 1, 0))
                    cmd.Parameters.AddWithValue("@Active", If(item.Active, 1, 0))
                    cmd.Parameters.AddWithValue("@LoginLink", If(item.LoginLink, String.Empty))
                    cmd.Parameters.AddWithValue("@Notes", If(item.Notes, String.Empty))
                    cmd.Parameters.AddWithValue("@Id", item.Id)

                    Dim rowsChanged = cmd.ExecuteNonQuery()
                    If rowsChanged <= 0 Then
                        Throw New InvalidOperationException("The selected income item could not be saved.")
                    End If
                End Using
            End Using
        End Sub

        Public Shared Function CreateIncome(databasePath As String, item As IncomeRecord) As Integer
            If String.IsNullOrWhiteSpace(databasePath) OrElse Not IO.File.Exists(databasePath) Then
                Throw New IO.FileNotFoundException("Budget database was not found.", databasePath)
            End If

            If item Is Nothing Then
                Throw New ArgumentNullException(NameOf(item))
            End If

            Using conn As New SqliteConnection($"Data Source={databasePath}")
                conn.Open()
                EnsureActiveColumn(conn)
                EnsureSameAsColumn(conn)

                Using cmd = conn.CreateCommand()
                    cmd.CommandText =
                        "INSERT INTO income (Description, Amount, Cadence, OnDay, OnDate, AutoIncrease, AutoIncreaseOnDate, StartDate, EndDate, ToAccount, SameAs, Hidden, Active, LoginLink, Notes) " &
                        "VALUES (@Description, @Amount, @Cadence, @OnDay, @OnDate, @AutoIncrease, @AutoIncreaseOnDate, @StartDate, @EndDate, @ToAccount, @SameAs, @Hidden, @Active, @LoginLink, @Notes);" &
                        "SELECT last_insert_rowid();"

                    cmd.Parameters.AddWithValue("@Description", If(item.Description, String.Empty))
                    cmd.Parameters.AddWithValue("@Amount", item.Amount)
                    cmd.Parameters.AddWithValue("@Cadence", If(item.Cadence, String.Empty))

                    Dim onDayParam = cmd.CreateParameter()
                    onDayParam.ParameterName = "@OnDay"
                    onDayParam.Value = If(item.OnDay.HasValue, item.OnDay.Value, DBNull.Value)
                    cmd.Parameters.Add(onDayParam)

                    cmd.Parameters.AddWithValue("@OnDate", If(item.OnDate, String.Empty))
                    cmd.Parameters.AddWithValue("@AutoIncrease", item.AutoIncrease)
                    cmd.Parameters.AddWithValue("@AutoIncreaseOnDate", If(item.AutoIncreaseOnDate, String.Empty))
                    cmd.Parameters.AddWithValue("@StartDate", If(item.StartDate, String.Empty))
                    cmd.Parameters.AddWithValue("@EndDate", If(item.EndDate, String.Empty))
                    cmd.Parameters.AddWithValue("@ToAccount", If(item.ToAccount, String.Empty))
                    cmd.Parameters.AddWithValue("@SameAs", If(item.SameAs, String.Empty))
                    cmd.Parameters.AddWithValue("@Hidden", If(item.Hidden, 1, 0))
                    cmd.Parameters.AddWithValue("@Active", If(item.Active, 1, 0))
                    cmd.Parameters.AddWithValue("@LoginLink", If(item.LoginLink, String.Empty))
                    cmd.Parameters.AddWithValue("@Notes", If(item.Notes, String.Empty))

                    Dim createdId = cmd.ExecuteScalar()
                    If createdId Is Nothing Then
                        Throw New InvalidOperationException("The new income item could not be created.")
                    End If

                    Return Convert.ToInt32(createdId)
                End Using
            End Using
        End Function

        Public Shared Sub DeleteIncome(databasePath As String, incomeId As Integer)
            If String.IsNullOrWhiteSpace(databasePath) OrElse Not IO.File.Exists(databasePath) Then
                Throw New IO.FileNotFoundException("Budget database was not found.", databasePath)
            End If

            If incomeId <= 0 Then
                Throw New ArgumentOutOfRangeException(NameOf(incomeId))
            End If

            Using conn As New SqliteConnection($"Data Source={databasePath}")
                conn.Open()
                EnsureActiveColumn(conn)

                Using cmd = conn.CreateCommand()
                    cmd.CommandText = "DELETE FROM income WHERE Id = @Id"
                    cmd.Parameters.AddWithValue("@Id", incomeId)

                    Dim rowsChanged = cmd.ExecuteNonQuery()
                    If rowsChanged <= 0 Then
                        Throw New InvalidOperationException("The selected income item could not be deleted.")
                    End If
                End Using
            End Using
        End Sub

        Public Shared Sub RestoreIncome(databasePath As String, item As IncomeRecord)
            If String.IsNullOrWhiteSpace(databasePath) OrElse Not IO.File.Exists(databasePath) Then
                Throw New IO.FileNotFoundException("Budget database was not found.", databasePath)
            End If

            If item Is Nothing Then
                Throw New ArgumentNullException(NameOf(item))
            End If

            If item.Id <= 0 Then
                Throw New InvalidOperationException("A valid income id is required for restore.")
            End If

            Using conn As New SqliteConnection($"Data Source={databasePath}")
                conn.Open()
                EnsureActiveColumn(conn)
                EnsureSameAsColumn(conn)

                Using cmd = conn.CreateCommand()
                    cmd.CommandText =
                        "INSERT OR REPLACE INTO income (Id, Description, Amount, Cadence, OnDay, OnDate, AutoIncrease, AutoIncreaseOnDate, StartDate, EndDate, ToAccount, SameAs, Hidden, Active, LoginLink, Notes) " &
                        "VALUES (@Id, @Description, @Amount, @Cadence, @OnDay, @OnDate, @AutoIncrease, @AutoIncreaseOnDate, @StartDate, @EndDate, @ToAccount, @SameAs, @Hidden, @Active, @LoginLink, @Notes)"

                    cmd.Parameters.AddWithValue("@Id", item.Id)
                    cmd.Parameters.AddWithValue("@Description", If(item.Description, String.Empty))
                    cmd.Parameters.AddWithValue("@Amount", item.Amount)
                    cmd.Parameters.AddWithValue("@Cadence", If(item.Cadence, String.Empty))

                    Dim onDayParam = cmd.CreateParameter()
                    onDayParam.ParameterName = "@OnDay"
                    onDayParam.Value = If(item.OnDay.HasValue, item.OnDay.Value, DBNull.Value)
                    cmd.Parameters.Add(onDayParam)

                    cmd.Parameters.AddWithValue("@OnDate", If(item.OnDate, String.Empty))
                    cmd.Parameters.AddWithValue("@AutoIncrease", item.AutoIncrease)
                    cmd.Parameters.AddWithValue("@AutoIncreaseOnDate", If(item.AutoIncreaseOnDate, String.Empty))
                    cmd.Parameters.AddWithValue("@StartDate", If(item.StartDate, String.Empty))
                    cmd.Parameters.AddWithValue("@EndDate", If(item.EndDate, String.Empty))
                    cmd.Parameters.AddWithValue("@ToAccount", If(item.ToAccount, String.Empty))
                    cmd.Parameters.AddWithValue("@SameAs", If(item.SameAs, String.Empty))
                    cmd.Parameters.AddWithValue("@Hidden", If(item.Hidden, 1, 0))
                    cmd.Parameters.AddWithValue("@Active", If(item.Active, 1, 0))
                    cmd.Parameters.AddWithValue("@LoginLink", If(item.LoginLink, String.Empty))
                    cmd.Parameters.AddWithValue("@Notes", If(item.Notes, String.Empty))
                    cmd.ExecuteNonQuery()
                End Using
            End Using
        End Sub

    Private Shared Sub EnsureActiveColumn(conn As SqliteConnection)
        If Not HasColumn(conn, "income", "Active") Then
            Using addCmd = conn.CreateCommand()
                addCmd.CommandText = "ALTER TABLE income ADD COLUMN Active INTEGER NOT NULL DEFAULT 1"
                addCmd.ExecuteNonQuery()
            End Using
        End If
    End Sub

    Private Shared Sub EnsureSameAsColumn(conn As SqliteConnection)
        If Not HasColumn(conn, "income", "SameAs") Then
            Using addCmd = conn.CreateCommand()
                addCmd.CommandText = "ALTER TABLE income ADD COLUMN SameAs TEXT NULL"
                addCmd.ExecuteNonQuery()
            End Using
        End If

        Using migrateCmd = conn.CreateCommand()
            migrateCmd.CommandText = "UPDATE income SET SameAs = COALESCE(ToAccount, ''), ToAccount = '' WHERE COALESCE(Cadence, '') = 'Same As' AND COALESCE(SameAs, '') = '' AND COALESCE(ToAccount, '') <> ''"
            migrateCmd.ExecuteNonQuery()
        End Using
    End Sub

        Private Shared Function HasColumn(conn As SqliteConnection, tableName As String, columnName As String) As Boolean
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
    End Class
End Namespace
