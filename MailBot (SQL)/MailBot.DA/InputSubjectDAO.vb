Imports System.Data
Imports System.Configuration
Imports System.Data.SqlClient
Imports MailBot.Entidades
Public Class InputSubjectDAO

    Dim constr As String = ConfigurationManager.AppSettings("dbSolution").ToString()
    Dim Schema As String = ConfigurationManager.AppSettings("Schema").ToString()

    Public Sub Insert(oInputSubject As InputSubject)

        Dim query As String = Schema & "upInputSubjectInsert"
        Dim con As SqlConnection = New SqlConnection(constr)
        Dim cmd As SqlCommand = New SqlCommand(query)
        Try
            cmd.Parameters.AddWithValue("@IdMessageSubject", oInputSubject.IdMessageSubject)
            cmd.Parameters.AddWithValue("@IdMessageConfiguration", oInputSubject.IdMessageConfiguration)
            cmd.Parameters.AddWithValue("@SubjectIdentifier", oInputSubject.SubjectIdentifier)
            cmd.Parameters.AddWithValue("@UserCreate", oInputSubject.UserCreate)
            cmd.Parameters.AddWithValue("@DateCreate", oInputSubject.DateCreate)
            cmd.Parameters.AddWithValue("@UserUpdate", oInputSubject.UserUpdate)
            cmd.Parameters.AddWithValue("@DateUpdate", oInputSubject.DateUpdate)
            cmd.Connection = con
            con.Open()
            cmd.ExecuteNonQuery()

        Catch ex As Exception
        Finally
            con.Close()
        End Try


    End Sub


    Public Sub Update(oInputSubject As InputSubject)

        Dim query As String = Schema & "upInputSubjectUpdate"
        Dim con As SqlConnection = New SqlConnection(constr)
        Dim cmd As SqlCommand = New SqlCommand(query)

        Try
            cmd.Parameters.AddWithValue("@IdMessageSubject", oInputSubject.IdMessageSubject)
            cmd.Parameters.AddWithValue("@IdMessageConfiguration", oInputSubject.IdMessageConfiguration)
            cmd.Parameters.AddWithValue("@SubjectIdentifier", oInputSubject.SubjectIdentifier)
            cmd.Parameters.AddWithValue("@UserCreate", oInputSubject.UserCreate)
            cmd.Parameters.AddWithValue("@DateCreate", oInputSubject.DateCreate)
            cmd.Parameters.AddWithValue("@UserUpdate", oInputSubject.UserUpdate)
            cmd.Parameters.AddWithValue("@DateUpdate", oInputSubject.DateUpdate)
            cmd.Connection = con
            con.Open()
            cmd.ExecuteNonQuery()

        Catch ex As Exception

        Finally
            con.Close()
        End Try


    End Sub

    Public Sub Delete(oInputSubject As InputSubject)

        Dim query As String = Schema & "upInputSubjectDelete"
        Dim con As SqlConnection = New SqlConnection(constr)
        Dim cmd As SqlCommand = New SqlCommand(query)
        Try
            cmd.Parameters.AddWithValue("@IdMessageSubject", oInputSubject.IdMessageSubject)
            cmd.Connection = con
            con.Open()
            cmd.ExecuteNonQuery()
        Catch ex As Exception

        Finally
            con.Close()
        End Try

    End Sub


    Public Function GetAll(oInputSubject As InputSubject) As List(Of InputSubject)
        Dim lstInputSubject As List(Of InputSubject) = New List(Of InputSubject)
        Dim query As String = Schema & "upGetInputSubject"
        Dim con As SqlConnection = New SqlConnection(constr)
        Dim cmd As SqlCommand = New SqlCommand(query)
        Dim rdr As SqlDataReader

        Try

            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@IdMessageSubject", oInputSubject.IdMessageSubject)
            cmd.Parameters.AddWithValue("@IdMessageConfiguration", oInputSubject.IdMessageConfiguration)
            cmd.Parameters.AddWithValue("@SubjectIdentifier", oInputSubject.SubjectIdentifier)
            cmd.Parameters.AddWithValue("@UserCreate", oInputSubject.UserCreate)
            cmd.Parameters.AddWithValue("@DateCreate", oInputSubject.DateCreate)
            cmd.Parameters.AddWithValue("@UserUpdate", oInputSubject.UserUpdate)
            cmd.Parameters.AddWithValue("@DateUpdate", oInputSubject.DateUpdate)
            con.Open()
            rdr = cmd.ExecuteReader()

            While rdr.Read()
                Dim bInputSubject As InputSubject = New InputSubject()

                bInputSubject.IdMessageSubject = Convert.ToInt32(rdr("IdMessageSubject").ToString())
                bInputSubject.IdMessageConfiguration = Convert.ToInt32(rdr("IdMessageConfiguration").ToString())
                bInputSubject.SubjectIdentifier = rdr("IdMessageConfiguration").ToString()
                bInputSubject.UserCreate = rdr("UserCreate").ToString()
                bInputSubject.DateCreate = Convert.ToDateTime(rdr("DateCreate").ToString())
                bInputSubject.UserUpdate = rdr("UserUpdate").ToString()
                bInputSubject.DateUpdate = Convert.ToDateTime(rdr("DateUpdate").ToString())



                lstInputSubject.Add(bInputSubject)
            End While


            Return lstInputSubject

        Catch ex As Exception
        Finally
            con.Close()

        End Try


    End Function


    Public Function GetAllDataTable(oInputSubject As InputSubject) As DataTable
        Dim dt As DataTable = New DataTable()
        Dim query As String = Schema & "upGetInputSubject"
        Dim con As SqlConnection = New SqlConnection(constr)
        Dim cmd As SqlCommand = New SqlCommand(query)
        Dim da As SqlDataAdapter

        Try

            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@IdMessageSubject", oInputSubject.IdMessageSubject)
            cmd.Parameters.AddWithValue("@IdMessageConfiguration", oInputSubject.IdMessageConfiguration)
            cmd.Parameters.AddWithValue("@SubjectIdentifier", oInputSubject.SubjectIdentifier)
            cmd.Parameters.AddWithValue("@UserCreate", oInputSubject.UserCreate)
            cmd.Parameters.AddWithValue("@DateCreate", oInputSubject.DateCreate)
            cmd.Parameters.AddWithValue("@UserUpdate", oInputSubject.UserUpdate)
            cmd.Parameters.AddWithValue("@DateUpdate", oInputSubject.DateUpdate)
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
