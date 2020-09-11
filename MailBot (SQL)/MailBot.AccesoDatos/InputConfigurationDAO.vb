Imports System.Data
Imports System.Configuration
Imports System.Data.SqlClient
Imports MailBot.Entidades

Public Class InputConfigurationDAO
    'Dim constr As String = ConfigurationManager.ConnectionStrings("constr").ConnectionString
    Dim constr As String = ConfigurationManager.AppSettings("dbSolution").ToString()
    Dim Schema As String = ConfigurationManager.AppSettings("Schema").ToString()

    Public Sub Insert(oInputConfiguration As InputConfiguration)

        Dim query As String = Schema & "upInputConfigurationInsert"
        Dim con As SqlConnection = New SqlConnection(constr)
        Dim cmd As SqlCommand = New SqlCommand(query)
        Try
            cmd.Parameters.AddWithValue("@IdInputConfiguration", oInputConfiguration.IdInputConfiguration)
            cmd.Parameters.AddWithValue("@InputConfigurationDescription", oInputConfiguration.InputConfigurationDescription)
            cmd.Parameters.AddWithValue("@IdMessageConfiguration", oInputConfiguration.IdMessageConfiguration)
            cmd.Parameters.AddWithValue("@IdInputType", oInputConfiguration.IdInputType)
            cmd.Parameters.AddWithValue("@ProcessFile", oInputConfiguration.ProcessFile)
            cmd.Parameters.AddWithValue("@UserUpdate", oInputConfiguration.UserUpdate)
            cmd.Parameters.AddWithValue("@DateUpdate", oInputConfiguration.DateUpdate)
            cmd.Connection = con
            con.Open()
            cmd.ExecuteNonQuery()

        Catch ex As Exception
        Finally
            con.Close()
        End Try


    End Sub


    Public Sub Update(oInputConfiguration As InputConfiguration)

        Dim query As String = Schema & "upInputConfigurationUpdate"
        Dim con As SqlConnection = New SqlConnection(constr)
        Dim cmd As SqlCommand = New SqlCommand(query)

        Try
            cmd.Parameters.AddWithValue("@IdInputConfiguration", oInputConfiguration.IdInputConfiguration)
            cmd.Parameters.AddWithValue("@InputConfigurationDescription", oInputConfiguration.InputConfigurationDescription)
            cmd.Parameters.AddWithValue("@IdMessageConfiguration", oInputConfiguration.IdMessageConfiguration)
            cmd.Parameters.AddWithValue("@IdInputType", oInputConfiguration.IdInputType)
            cmd.Parameters.AddWithValue("@ProcessFile", oInputConfiguration.ProcessFile)
            cmd.Parameters.AddWithValue("@UserCreate", oInputConfiguration.UserCreate)
            cmd.Parameters.AddWithValue("@DateCreate", oInputConfiguration.DateCreate)
            cmd.Parameters.AddWithValue("@UserUpdate", oInputConfiguration.UserUpdate)
            cmd.Parameters.AddWithValue("@DateUpdate", oInputConfiguration.DateUpdate)
            cmd.Connection = con
            con.Open()
            cmd.ExecuteNonQuery()

        Catch ex As Exception

        Finally
            con.Close()
        End Try


    End Sub

    Public Sub Delete(oInputConfiguration As InputConfiguration)

        Dim query As String = Schema & "upInputConfigurationDelete"
        Dim con As SqlConnection = New SqlConnection(constr)
        Dim cmd As SqlCommand = New SqlCommand(query)
        Try
            cmd.Parameters.AddWithValue("@IdInputConfiguration", oInputConfiguration.IdInputConfiguration)
            cmd.Connection = con
            con.Open()
            cmd.ExecuteNonQuery()
        Catch ex As Exception

        Finally
            con.Close()
        End Try

    End Sub


    Public Function GetAll(oInputConfiguration As InputConfiguration) As List(Of InputConfiguration)
        Dim lstInputConfiguration As List(Of InputConfiguration) = New List(Of InputConfiguration)
        Dim query As String = Schema & "upGetInputConfiguration"
        Dim con As SqlConnection = New SqlConnection(constr)
        Dim cmd As SqlCommand = New SqlCommand(query)
        Dim rdr As SqlDataReader

        Try

            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@IdInputConfiguration", oInputConfiguration.IdInputConfiguration)
            cmd.Parameters.AddWithValue("@InputConfigurationDescription", oInputConfiguration.InputConfigurationDescription)
            cmd.Parameters.AddWithValue("@IdMessageConfiguration", oInputConfiguration.IdMessageConfiguration)
            cmd.Parameters.AddWithValue("@IdInputType", oInputConfiguration.IdInputType)
            cmd.Parameters.AddWithValue("@ProcessFile", oInputConfiguration.ProcessFile)
            cmd.Parameters.AddWithValue("@UserCreate", oInputConfiguration.UserCreate)
            cmd.Parameters.AddWithValue("@DateCreate", oInputConfiguration.DateCreate)
            cmd.Parameters.AddWithValue("@UserUpdate", oInputConfiguration.UserUpdate)
            cmd.Parameters.AddWithValue("@DateUpdate", oInputConfiguration.DateUpdate)
            con.Open()
            rdr = cmd.ExecuteReader()

            While rdr.Read()
                Dim bInputConfiguration As InputConfiguration = New InputConfiguration()

                bInputConfiguration.IdInputConfiguration = Convert.ToInt32(rdr("IdInputConfiguration").ToString())
                bInputConfiguration.InputConfigurationDescription = rdr("InputConfigurationDescription").ToString()
                bInputConfiguration.IdMessageConfiguration = Convert.ToInt32(rdr("IdMessageConfiguration").ToString())
                bInputConfiguration.IdInputType = Convert.ToInt32(rdr("IdInputType").ToString())
                bInputConfiguration.ProcessFile = rdr("ProcessFile").ToString()
                bInputConfiguration.UserCreate = rdr("UserCreate").ToString()
                bInputConfiguration.DateCreate = Convert.ToDateTime(rdr("DateCreate").ToString())
                bInputConfiguration.UserUpdate = rdr("UserUpdate").ToString()
                bInputConfiguration.DateUpdate = Convert.ToDateTime(rdr("DateUpdate").ToString())



                lstInputConfiguration.Add(bInputConfiguration)
            End While


            Return lstInputConfiguration

        Catch ex As Exception
        Finally
            con.Close()

        End Try


    End Function


    Public Function GetAllDataTable(oInputConfiguration As InputConfiguration) As DataTable
        Dim dt As DataTable = New DataTable()
        Dim query As String = Schema & "upGetInputConfiguration"
        Dim con As SqlConnection = New SqlConnection(constr)
        Dim cmd As SqlCommand = New SqlCommand(query)
        Dim da As SqlDataAdapter

        Try

            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@IdInputConfiguration", oInputConfiguration.IdInputConfiguration)
            cmd.Parameters.AddWithValue("@InputConfigurationDescription", oInputConfiguration.InputConfigurationDescription)
            cmd.Parameters.AddWithValue("@IdMessageConfiguration", oInputConfiguration.IdMessageConfiguration)
            cmd.Parameters.AddWithValue("@IdInputType", oInputConfiguration.IdInputType)
            cmd.Parameters.AddWithValue("@ProcessFile", oInputConfiguration.ProcessFile)
            cmd.Parameters.AddWithValue("@UserCreate", oInputConfiguration.UserCreate)
            cmd.Parameters.AddWithValue("@DateCreate", oInputConfiguration.DateCreate)
            cmd.Parameters.AddWithValue("@UserUpdate", oInputConfiguration.UserUpdate)
            cmd.Parameters.AddWithValue("@DateUpdate", oInputConfiguration.DateUpdate)
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
