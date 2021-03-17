Imports System.Data
Imports System.Data.SqlClient
Imports System.Configuration
Imports System.Collections.Generic
'Imports MySql.Data.MySqlClient
Imports Npgsql

Public Class GeneralDAO
    Dim oGeneralDA As New AppService.OdooServiceSoapClient

    Friend Function ExecuteSQL(ByVal QueryString As String) As DataTable
        Dim dtResult As New DataTable
        Dim DBCnxStr As String = My.Settings.PostgreSQLCnxString
        dtResult = oGeneralDA.ExecuteSQL(DBCnxStr, QueryString)
        Return dtResult
    End Function

    Friend Function ExecuteMSSQL(ByVal QueryString As String) As DataSet
        Dim dsResult As New DataSet
        Using connection As New SqlConnection(My.Settings.DBConnectionString)
            Dim Command As New SqlCommand(QueryString, connection)
            Try
                Command.Connection.Open()
                Dim reader As SqlDataReader = Command.ExecuteReader()
                dsResult.Tables.Add()
                dsResult.Tables(0).Load(reader)
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            Finally
                Command.Connection.Close()
            End Try
            Return dsResult
        End Using
    End Function

    Friend Function SQLTestConnection() As Boolean
        Dim bResult As Boolean = True
        Using connection As New SqlConnection(My.Settings.DBConnectionString)
            Try
                connection.Open()
                If Not connection.State = ConnectionState.Open Then
                    bResult = False
                End If
            Catch ex As Exception
                bResult = False
            Finally
                connection.Close()
            End Try
            Return bResult
        End Using
    End Function

    'Friend Function ExecuteMySQL(ByVal QueryString As String) As DataSet
    '    Dim dsResult As New DataSet
    '    Using connection As New MySqlConnection(My.Settings.MySQLCnxString)

    '        Dim Command As New MySqlCommand(QueryString, connection)
    '        Try
    '            Command.Connection.Open()
    '            Dim reader As MySqlDataReader = Command.ExecuteReader()
    '            dsResult.Tables.Add()
    '            dsResult.Tables(0).Load(reader)
    '        Catch ex As Exception
    '            MessageBox.Show(ex.Message)
    '        Finally
    '            Command.Connection.Close()
    '        End Try
    '        Return dsResult
    '    End Using
    'End Function

    Friend Function ExecutePostgreSQL(ByVal QueryString As String) As DataTable
        Dim dsResult As New DataSet
        Using connection As New NpgsqlConnection(My.Settings.PostgreSQLCnxString & ";timeout=1024;")
            Dim Command As New NpgsqlCommand(QueryString, connection)
            Try
                Command.Connection.Open()
                Dim reader As NpgsqlDataReader = Command.ExecuteReader()
                dsResult.Tables.Add()
                dsResult.Tables(0).Load(reader)
            Catch ex As Exception
                MessageBox.Show(ex.Message)
                Return Nothing
            Finally
                Command.Connection.Close()
            End Try
            Return dsResult.Tables(0)
        End Using
    End Function

End Class
