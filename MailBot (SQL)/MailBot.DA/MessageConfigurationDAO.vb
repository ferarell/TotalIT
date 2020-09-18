Imports System.Data
Imports System.Configuration
Imports System.Data.SqlClient
Imports MailBot.Entidades
Public Class MessageConfigurationDAO
    Dim constr As String = ConfigurationManager.AppSettings("dbSolution").ToString()
    Dim Schema As String = ConfigurationManager.AppSettings("Schema").ToString()

    Public Sub Insert(oMessageConfiguration As MessageConfiguration)

        Dim query As String = Schema & "upMessageConfigurationInsert"
        Dim con As SqlConnection = New SqlConnection(constr)
        Dim cmd As SqlCommand = New SqlCommand(query)
        Try
            cmd.Parameters.AddWithValue("@IdMessageConfiguration", oMessageConfiguration.IdMessageConfiguration)
            cmd.Parameters.AddWithValue("@IdInputType", oMessageConfiguration.IdInputType)

            cmd.Parameters.AddWithValue("@Recipients", oMessageConfiguration.Recipients)
            cmd.Parameters.AddWithValue("@CopyRecipients", oMessageConfiguration.CopyRecipients)
            cmd.Parameters.AddWithValue("@BlindCopyRecipients", oMessageConfiguration.BlindCopyRecipients)
            cmd.Parameters.AddWithValue("@CustomizedQuery", oMessageConfiguration.CustomizedQuery)
            cmd.Parameters.AddWithValue("@AttachedFiles", oMessageConfiguration.AttachedFiles)
            cmd.Parameters.AddWithValue("@Image", oMessageConfiguration.Image)
            cmd.Parameters.AddWithValue("@ImageFile", oMessageConfiguration.ImageFile)

            cmd.Parameters.AddWithValue("@UserCreate", oMessageConfiguration.UserCreate)
            cmd.Parameters.AddWithValue("@DateCreate", oMessageConfiguration.DateCreate)
            cmd.Parameters.AddWithValue("@UserUpdate", oMessageConfiguration.UserUpdate)
            cmd.Parameters.AddWithValue("@DateUpdate", oMessageConfiguration.DateUpdate)
            cmd.Connection = con
            con.Open()
            cmd.ExecuteNonQuery()

        Catch ex As Exception
        Finally
            con.Close()
        End Try


    End Sub


    Public Sub Update(oMessageConfiguration As MessageConfiguration)

        Dim query As String = Schema & "upMessageConfigurationUpdate"
        Dim con As SqlConnection = New SqlConnection(constr)
        Dim cmd As SqlCommand = New SqlCommand(query)

        Try
            cmd.Parameters.AddWithValue("@IdMessageConfiguration", oMessageConfiguration.IdMessageConfiguration)
            cmd.Parameters.AddWithValue("@IdInputType", oMessageConfiguration.IdInputType)

            cmd.Parameters.AddWithValue("@Recipients", oMessageConfiguration.Recipients)
            cmd.Parameters.AddWithValue("@CopyRecipients", oMessageConfiguration.CopyRecipients)
            cmd.Parameters.AddWithValue("@BlindCopyRecipients", oMessageConfiguration.BlindCopyRecipients)
            cmd.Parameters.AddWithValue("@CustomizedQuery", oMessageConfiguration.CustomizedQuery)
            cmd.Parameters.AddWithValue("@AttachedFiles", oMessageConfiguration.AttachedFiles)
            cmd.Parameters.AddWithValue("@Image", oMessageConfiguration.Image)
            cmd.Parameters.AddWithValue("@ImageFile", oMessageConfiguration.ImageFile)

            cmd.Parameters.AddWithValue("@UserCreate", oMessageConfiguration.UserCreate)
            cmd.Parameters.AddWithValue("@DateCreate", oMessageConfiguration.DateCreate)
            cmd.Parameters.AddWithValue("@UserUpdate", oMessageConfiguration.UserUpdate)
            cmd.Parameters.AddWithValue("@DateUpdate", oMessageConfiguration.DateUpdate)
            cmd.Connection = con
            con.Open()
            cmd.ExecuteNonQuery()

        Catch ex As Exception

        Finally
            con.Close()
        End Try


    End Sub

    Public Sub Delete(oMessageConfiguration As MessageConfiguration)

        Dim query As String = Schema & "upMessageConfigurationDelete"
        Dim con As SqlConnection = New SqlConnection(constr)
        Dim cmd As SqlCommand = New SqlCommand(query)
        Try
            cmd.Parameters.AddWithValue("@IdMessageConfiguration", oMessageConfiguration.IdMessageConfiguration)
            cmd.Connection = con
            con.Open()
            cmd.ExecuteNonQuery()
        Catch ex As Exception

        Finally
            con.Close()
        End Try

    End Sub


    Public Function GetAll(oMessageConfiguration As MessageConfiguration) As List(Of MessageConfiguration)
        Dim lstMessageConfiguration As List(Of MessageConfiguration) = New List(Of MessageConfiguration)
        Dim query As String = Schema & "upGetMessageConfiguration"
        Dim con As SqlConnection = New SqlConnection(constr)
        Dim cmd As SqlCommand = New SqlCommand(query)
        Dim rdr As SqlDataReader

        Try

            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@IdMessageConfiguration", oMessageConfiguration.IdMessageConfiguration)
            cmd.Parameters.AddWithValue("@IdInputType", oMessageConfiguration.IdInputType)

            cmd.Parameters.AddWithValue("@Recipients", oMessageConfiguration.Recipients)
            cmd.Parameters.AddWithValue("@CopyRecipients", oMessageConfiguration.CopyRecipients)
            cmd.Parameters.AddWithValue("@BlindCopyRecipients", oMessageConfiguration.BlindCopyRecipients)
            cmd.Parameters.AddWithValue("@CustomizedQuery", oMessageConfiguration.CustomizedQuery)
            cmd.Parameters.AddWithValue("@AttachedFiles", oMessageConfiguration.AttachedFiles)
            cmd.Parameters.AddWithValue("@Image", oMessageConfiguration.Image)
            cmd.Parameters.AddWithValue("@ImageFile", oMessageConfiguration.ImageFile)

            cmd.Parameters.AddWithValue("@UserCreate", oMessageConfiguration.UserCreate)
            cmd.Parameters.AddWithValue("@DateCreate", oMessageConfiguration.DateCreate)
            cmd.Parameters.AddWithValue("@UserUpdate", oMessageConfiguration.UserUpdate)
            cmd.Parameters.AddWithValue("@DateUpdate", oMessageConfiguration.DateUpdate)
            con.Open()
            rdr = cmd.ExecuteReader()

            While rdr.Read()
                Dim bMessageConfiguration As MessageConfiguration = New MessageConfiguration()

                bMessageConfiguration.IdMessageConfiguration = Convert.ToInt32(rdr("IdMessageConfiguration").ToString())
                bMessageConfiguration.IdInputType = Convert.ToInt32(rdr("IdInputType").ToString())


                bMessageConfiguration.Recipients = rdr("Recipients").ToString()
                bMessageConfiguration.CopyRecipients = rdr("CopyRecipients").ToString()
                bMessageConfiguration.BlindCopyRecipients = rdr("BlindCopyRecipients").ToString()
                bMessageConfiguration.CustomizedQuery = rdr("CustomizedQuery").ToString()
                bMessageConfiguration.AttachedFiles = System.Text.Encoding.Default.GetBytes(rdr("AttachedFiles").ToString())
                bMessageConfiguration.Image = System.Text.Encoding.Default.GetBytes(rdr("Image").ToString())
                bMessageConfiguration.ImageFile = System.Text.Encoding.Default.GetBytes(rdr("ImageFile").ToString())



                bMessageConfiguration.UserCreate = rdr("UserCreate").ToString()
                bMessageConfiguration.DateCreate = Convert.ToDateTime(rdr("DateCreate").ToString())
                bMessageConfiguration.UserUpdate = rdr("UserUpdate").ToString()
                bMessageConfiguration.DateUpdate = Convert.ToDateTime(rdr("DateUpdate").ToString())



                lstMessageConfiguration.Add(bMessageConfiguration)
            End While


            Return lstMessageConfiguration

        Catch ex As Exception
        Finally
            con.Close()

        End Try


    End Function


    Public Function GetAllDataTable(oMessageConfiguration As MessageConfiguration) As DataTable
        Dim dt As DataTable = New DataTable()
        Dim query As String = Schema & "upGetMessageConfiguration"
        Dim con As SqlConnection = New SqlConnection(constr)
        Dim cmd As SqlCommand = New SqlCommand(query)
        Dim da As SqlDataAdapter

        Try

            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@IdMessageConfiguration", oMessageConfiguration.IdMessageConfiguration)
            cmd.Parameters.AddWithValue("@IdInputType", oMessageConfiguration.IdInputType)

            cmd.Parameters.AddWithValue("@Recipients", oMessageConfiguration.Recipients)
            cmd.Parameters.AddWithValue("@CopyRecipients", oMessageConfiguration.CopyRecipients)
            cmd.Parameters.AddWithValue("@BlindCopyRecipients", oMessageConfiguration.BlindCopyRecipients)
            cmd.Parameters.AddWithValue("@CustomizedQuery", oMessageConfiguration.CustomizedQuery)
            cmd.Parameters.AddWithValue("@AttachedFiles", oMessageConfiguration.AttachedFiles)
            cmd.Parameters.AddWithValue("@Image", oMessageConfiguration.Image)
            cmd.Parameters.AddWithValue("@ImageFile", oMessageConfiguration.ImageFile)

            cmd.Parameters.AddWithValue("@UserCreate", oMessageConfiguration.UserCreate)
            cmd.Parameters.AddWithValue("@DateCreate", oMessageConfiguration.DateCreate)
            cmd.Parameters.AddWithValue("@UserUpdate", oMessageConfiguration.UserUpdate)
            cmd.Parameters.AddWithValue("@DateUpdate", oMessageConfiguration.DateUpdate)
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
