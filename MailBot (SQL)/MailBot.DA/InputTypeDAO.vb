Imports System.Data
Imports System.Configuration
Imports System.Data.SqlClient
Imports MailBot.Entidades
Public Class InputTypeDAO

    Dim constr As String = ConfigurationManager.AppSettings("dbSolution").ToString()
    Dim Schema As String = ConfigurationManager.AppSettings("Schema").ToString()

    Public Sub Insert(oInputType As InputType)

        Dim query As String = Schema & "upInputTypeInsert"
        Dim con As SqlConnection = New SqlConnection(constr)
        Dim cmd As SqlCommand = New SqlCommand(query)
        Try
            cmd.Parameters.AddWithValue("@IdInputType", oInputType.IdInputType)
            cmd.Parameters.AddWithValue("@InputTypeDescription", oInputType.InputTypeDescription)
            cmd.Parameters.AddWithValue("@UserCreate", oInputType.UserCreate)
            cmd.Parameters.AddWithValue("@DateCreate", oInputType.DateCreate)
            cmd.Parameters.AddWithValue("@UserUpdate", oInputType.UserUpdate)
            cmd.Parameters.AddWithValue("@DateUpdate", oInputType.DateUpdate)
            cmd.Connection = con
            con.Open()
            cmd.ExecuteNonQuery()

        Catch ex As Exception
        Finally
            con.Close()
        End Try


    End Sub


    Public Sub Update(oInputType As InputType)

        Dim query As String = Schema & "upInputTypeUpdate"
        Dim con As SqlConnection = New SqlConnection(constr)
        Dim cmd As SqlCommand = New SqlCommand(query)

        Try
            cmd.Parameters.AddWithValue("@IdInputType", oInputType.IdInputType)
            cmd.Parameters.AddWithValue("@InputTypeDescription", oInputType.InputTypeDescription)
            cmd.Parameters.AddWithValue("@UserCreate", oInputType.UserCreate)
            cmd.Parameters.AddWithValue("@DateCreate", oInputType.DateCreate)
            cmd.Parameters.AddWithValue("@UserUpdate", oInputType.UserUpdate)
            cmd.Parameters.AddWithValue("@DateUpdate", oInputType.DateUpdate)
            cmd.Connection = con
            con.Open()
            cmd.ExecuteNonQuery()

        Catch ex As Exception

        Finally
            con.Close()
        End Try


    End Sub

    Public Sub Delete(oInputType As InputType)

        Dim query As String = Schema & "upInputTypeDelete"
        Dim con As SqlConnection = New SqlConnection(constr)
        Dim cmd As SqlCommand = New SqlCommand(query)
        Try
            cmd.Parameters.AddWithValue("@IdInputType", oInputType.IdInputType)
            cmd.Connection = con
            con.Open()
            cmd.ExecuteNonQuery()
        Catch ex As Exception

        Finally
            con.Close()
        End Try

    End Sub


    Public Function GetAll(oInputType As InputType) As List(Of InputType)
        Dim lstInputType As List(Of InputType) = New List(Of InputType)
        Dim query As String = Schema & "upGetInputType"
        Dim con As SqlConnection = New SqlConnection(constr)
        Dim cmd As SqlCommand = New SqlCommand(query)
        Dim rdr As SqlDataReader

        Try

            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@IdInputType", oInputType.IdInputType)
            cmd.Parameters.AddWithValue("@InputTypeDescription", oInputType.InputTypeDescription)
            cmd.Parameters.AddWithValue("@UserCreate", oInputType.UserCreate)
            cmd.Parameters.AddWithValue("@DateCreate", oInputType.DateCreate)
            cmd.Parameters.AddWithValue("@UserUpdate", oInputType.UserUpdate)
            cmd.Parameters.AddWithValue("@DateUpdate", oInputType.DateUpdate)
            con.Open()
            rdr = cmd.ExecuteReader()

            While rdr.Read()
                Dim bInputType As InputType = New InputType()

                bInputType.IdInputType = Convert.ToInt32(rdr("IdInputType").ToString())
                bInputType.InputTypeDescription = rdr("InputTypeDescription").ToString()
                bInputType.UserCreate = rdr("UserCreate").ToString()
                bInputType.DateCreate = Convert.ToDateTime(rdr("DateCreate").ToString())
                bInputType.UserUpdate = rdr("UserUpdate").ToString()
                bInputType.DateUpdate = Convert.ToDateTime(rdr("DateUpdate").ToString())



                lstInputType.Add(bInputType)
            End While


            Return lstInputType

        Catch ex As Exception
        Finally
            con.Close()

        End Try


    End Function


    Public Function GetAllDataTable(oInputType As InputType) As DataTable
        Dim dt As DataTable = New DataTable()
        Dim query As String = Schema & "upGetInputType"
        Dim con As SqlConnection = New SqlConnection(constr)
        Dim cmd As SqlCommand = New SqlCommand(query)
        Dim da As SqlDataAdapter

        Try

            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@IdInputType", oInputType.IdInputType)
            cmd.Parameters.AddWithValue("@InputTypeDescription", oInputType.InputTypeDescription)
            cmd.Parameters.AddWithValue("@UserCreate", oInputType.UserCreate)
            cmd.Parameters.AddWithValue("@DateCreate", oInputType.DateCreate)
            cmd.Parameters.AddWithValue("@UserUpdate", oInputType.UserUpdate)
            cmd.Parameters.AddWithValue("@DateUpdate", oInputType.DateUpdate)
            con.Open()
            da = New SqlDataAdapter(cmd)
            da.Fill(dt)

            Return dt
        Catch ex As Exception
        Finally
            con.Close()

        End Try


    End Function


End Class
