Public Class QueryHelper

    Private str As String = "Data Source=localhost;Initial Catalog=mTracker;User ID=sa;Password=gilagila ;Persist Security Info=True;MultipleActiveResultSets=true;"
    Private sqlCon


    Private Sub OpenConnection()
        sqlCon = New SqlClient.SqlConnection()
        If sqlCon.State = ConnectionState.Closed Then
            Try
                sqlCon.ConnectionString = str
                sqlCon.Open()
            Catch ex As Exception
                Console.WriteLine(ex.ToString)
                LogHelper.InsertLog(ex.Message)
            End Try

        End If
    End Sub

    Private Sub CloseConnection()
        If sqlCon.State = ConnectionState.Open Then
            Try
                sqlCon.Close()
            Catch ex As Exception
                Console.WriteLine(ex.ToString)
                LogHelper.InsertLog(ex.Message)
            End Try
        End If
    End Sub


    Public Function GetDataByStoredProcedure(ByVal SPName As String) As DataTable
        Dim sqlCommand As New SqlClient.SqlCommand()
        Dim dt As DataTable = New DataTable()
        Dim dateNow As DateTime = DateTime.Now
        Dim listLogTrack As New ArrayList
        Dim dataAdapt As New SqlClient.SqlDataAdapter()

        Try
            OpenConnection()
            sqlCommand.CommandText = SPName
            sqlCommand.CommandType = CommandType.StoredProcedure
            sqlCommand.Connection = sqlCon
            sqlCommand.CommandTimeout = 0

            dataAdapt.SelectCommand = sqlCommand
            dataAdapt.Fill(dt)

        Catch ex As Exception
            LogHelper.InsertLog(ex.Message)
            Console.WriteLine(ex.Message)
        Finally
            CloseConnection()
            sqlCon = Nothing
            sqlCommand = Nothing
            dataAdapt = Nothing
        End Try
        Return dt
    End Function

    Public Function GetDataByQuery(ByVal query As String) As DataTable
        Dim sqlCommand As New SqlClient.SqlCommand()
        Dim dt As DataTable = New DataTable()
        Dim dateNow As DateTime = DateTime.Now
        Dim listLogTrack As New ArrayList
        Dim dataAdapt As New SqlClient.SqlDataAdapter()

        Try
            OpenConnection()
            sqlCommand.CommandText = query
            sqlCommand.Connection = sqlCon
            sqlCommand.CommandTimeout = 0

            dataAdapt.SelectCommand = sqlCommand
            dataAdapt.Fill(dt)
        Catch ex As Exception
            LogHelper.InsertLog(ex.Message)
            Console.WriteLine(ex.Message)
        Finally
            CloseConnection()
            sqlCon = Nothing
            sqlCommand = Nothing
            dataAdapt = Nothing
        End Try
        Return dt
    End Function


    Public Function GetScalarDataByQuery(ByVal query As String) As String
        Dim sqlCommand As New SqlClient.SqlCommand()
        Dim result As String = ""
        Try
            OpenConnection()
            sqlCommand = New SqlClient.SqlCommand()
            sqlCommand.CommandText = query
            sqlCommand.Connection = sqlCon
            result = sqlCommand.ExecuteScalar()
        Catch ex As Exception
            LogHelper.InsertLog(ex.Message)
            Console.WriteLine(ex.Message)
        Finally
            CloseConnection()
            sqlCon = Nothing
            sqlCommand = Nothing
        End Try
        Return result

    End Function


    Public Sub ExecuteNonQuery(ByVal query As String)
        Dim sqlCommand As New SqlClient.SqlCommand()

        Try
            OpenConnection()
            sqlCommand = New SqlClient.SqlCommand()
            sqlCommand.CommandText = query
            sqlCommand.Connection = sqlCon
            sqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            LogHelper.InsertLog(ex.Message)
            Console.WriteLine(ex.Message)
        Finally
            CloseConnection()
            sqlCon = Nothing
            sqlCommand = Nothing
        End Try

    End Sub

    Public Sub ExecuteNonQueryByStoredProcedure(ByVal SPName As String)
        Dim sqlCommand As New SqlClient.SqlCommand()

        Try
            OpenConnection()
            sqlCommand = New SqlClient.SqlCommand()
            sqlCommand.CommandText = SPName
            sqlCommand.CommandType = CommandType.StoredProcedure
            sqlCommand.Connection = sqlCon
            sqlCommand.CommandTimeout = 0            
            sqlCommand.ExecuteNonQuery()
        Catch ex As Exception
            LogHelper.InsertLog(ex.Message)
            Console.WriteLine(ex.Message)
        Finally
            CloseConnection()
            sqlCon = Nothing
            sqlCommand = Nothing
        End Try

    End Sub
End Class
