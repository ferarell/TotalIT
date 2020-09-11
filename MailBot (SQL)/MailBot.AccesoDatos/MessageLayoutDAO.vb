Imports System.Data
Imports System.Configuration
Imports System.Data.SqlClient
Imports MailBot.Entidades

Public Class MessageLayoutDAO
    Dim constr As String = ConfigurationManager.AppSettings("dbSolution").ToString()
    Dim Schema As String = ConfigurationManager.AppSettings("Schema").ToString()

    Public Sub Insert(oMessageLayout As MessageLayout)

        Dim query As String = Schema & "upMessageLayoutInsert"
        Dim con As SqlConnection = New SqlConnection(constr)
        Dim cmd As SqlCommand = New SqlCommand(query)
        Try
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@IdMessageLayout", oMessageLayout.IdMessageLayout)
            cmd.Parameters.AddWithValue("@IdMessageStructure", oMessageLayout.IdMessageStructure)
            cmd.Parameters.AddWithValue("@IdInputConfiguration", oMessageLayout.IdInputConfiguration)
            cmd.Parameters.AddWithValue("@MessageText", System.Text.Encoding.Unicode.GetBytes(oMessageLayout.MessageText))
            'cmd.Parameters.AddWithValue("@ValidFrom", vbNull)
            'If Not oMessageLayout.ValidFrom Is Nothing Then
            cmd.Parameters.AddWithValue("@ValidFrom", oMessageLayout.ValidFrom)
            'End If
            'cmd.Parameters.AddWithValue("@ValidTo", vbNull)
            'If Not oMessageLayout.ValidTo Is Nothing Then
            cmd.Parameters.AddWithValue("@ValidTo", oMessageLayout.ValidTo)
            'End If
            cmd.Parameters.AddWithValue("@UserCreate", oMessageLayout.UserCreate)
            cmd.Parameters.AddWithValue("@DateCreate", oMessageLayout.DateCreate)
            cmd.Connection = con
            con.Open()
            cmd.ExecuteNonQuery()


        Catch ex As Exception
            Dim Mensaje As String = ex.Message
        Finally
            con.Close()
        End Try


    End Sub


    Public Sub Update(oMessageLayout As MessageLayout)

        Dim query As String = Schema & "upMessageLayoutUpdate"
        Dim con As SqlConnection = New SqlConnection(constr)
        Dim cmd As SqlCommand = New SqlCommand(query)

        Try
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@IdMessageLayout", oMessageLayout.IdMessageLayout)
            cmd.Parameters.AddWithValue("@IdMessageStructure", oMessageLayout.IdMessageStructure)
            cmd.Parameters.AddWithValue("@IdInputConfiguration", oMessageLayout.IdInputConfiguration)
            'cmd.Parameters.AddWithValue("@MessageText", System.Text.Encoding.ASCII.GetBytes(oMessageLayout.MessageText))
            cmd.Parameters.AddWithValue("@MessageText", System.Text.Encoding.Unicode.GetBytes(oMessageLayout.MessageText))
            cmd.Parameters.AddWithValue("@ValidFrom", oMessageLayout.ValidFrom)
            cmd.Parameters.AddWithValue("@ValidTo", oMessageLayout.ValidTo)
            cmd.Parameters.AddWithValue("@UserUpdate", oMessageLayout.UserUpdate)
            cmd.Parameters.AddWithValue("@DateUpdate", oMessageLayout.DateUpdate)
            cmd.Connection = con
            con.Open()
            cmd.ExecuteNonQuery()

        Catch ex As Exception
            Dim Mensaje As String = ex.Message
        Finally
            con.Close()
        End Try


    End Sub

    Public Sub Delete(oMessageLayout As MessageLayout)

        Dim query As String = Schema & "upMessageLayoutDelete"
        Dim con As SqlConnection = New SqlConnection(constr)
        Dim cmd As SqlCommand = New SqlCommand(query)
        Try
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@IdMessageLayout", oMessageLayout.IdMessageLayout)
            cmd.Connection = con
            con.Open()
            cmd.ExecuteNonQuery()
        Catch ex As Exception

        Finally
            con.Close()
        End Try

    End Sub


    Public Function GetAll(oMessageLayout As MessageLayout) As List(Of MessageLayout)
        Dim lstMessageLayout As List(Of MessageLayout) = New List(Of MessageLayout)
        Dim query As String = Schema & "upGetMessageLayout"
        Dim con As SqlConnection = New SqlConnection(constr)
        Dim cmd As SqlCommand = New SqlCommand(query)
        Dim rdr As SqlDataReader

        Try

            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@IdMessageLayout", oMessageLayout.IdMessageLayout)
            cmd.Parameters.AddWithValue("@IdMessageStructure", oMessageLayout.IdMessageStructure)
            cmd.Parameters.AddWithValue("@IdInputConfiguration", oMessageLayout.IdInputConfiguration)
            cmd.Parameters.AddWithValue("@MessageText", oMessageLayout.MessageText)
            cmd.Parameters.AddWithValue("@ValidFrom", oMessageLayout.ValidFrom)
            cmd.Parameters.AddWithValue("@ValidTo", oMessageLayout.ValidTo)
            cmd.Parameters.AddWithValue("@UserCreate", oMessageLayout.UserCreate)
            cmd.Parameters.AddWithValue("@DateCreate", oMessageLayout.DateCreate)
            cmd.Parameters.AddWithValue("@UserUpdate", oMessageLayout.UserUpdate)
            cmd.Parameters.AddWithValue("@DateUpdate", oMessageLayout.DateUpdate)
            con.Open()
            rdr = cmd.ExecuteReader()

            While rdr.Read()
                Dim bMessageLayout As MessageLayout = New MessageLayout()

                bMessageLayout.IdMessageLayout = Convert.ToInt32(rdr("IdMessageLayout").ToString())
                bMessageLayout.IdMessageStructure = Convert.ToInt32(rdr("IdMessageStructure").ToString())


                bMessageLayout.IdInputConfiguration = Convert.ToInt32(rdr("IdInputConfiguration").ToString())
                bMessageLayout.MessageText = rdr("MessageText").ToString()
                bMessageLayout.ValidFrom = Convert.ToDateTime(rdr("ValidFrom").ToString())
                bMessageLayout.ValidTo = Convert.ToDateTime(rdr("ValidTo").ToString())

                bMessageLayout.UserCreate = rdr("UserCreate").ToString()
                bMessageLayout.DateCreate = Convert.ToDateTime(rdr("DateCreate").ToString())
                bMessageLayout.UserUpdate = rdr("UserUpdate").ToString()
                bMessageLayout.DateUpdate = Convert.ToDateTime(rdr("DateUpdate").ToString())

                lstMessageLayout.Add(bMessageLayout)
            End While


            Return lstMessageLayout

        Catch ex As Exception
        Finally
            con.Close()

        End Try


    End Function


    Public Function GetAllDataTable(oMessageLayout As MessageLayout) As DataTable
        Dim dt As DataTable = New DataTable()
        Dim query As String = Schema & "upGetMessageConfiguration"
        Dim con As SqlConnection = New SqlConnection(constr)
        Dim cmd As SqlCommand = New SqlCommand(query)
        Dim da As SqlDataAdapter

        Try

            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@IdMessageLayout", oMessageLayout.IdMessageLayout)
            cmd.Parameters.AddWithValue("@IdMessageStructure", oMessageLayout.IdMessageStructure)
            cmd.Parameters.AddWithValue("@IdInputConfiguration", oMessageLayout.IdInputConfiguration)
            cmd.Parameters.AddWithValue("@MessageText", oMessageLayout.MessageText)
            cmd.Parameters.AddWithValue("@ValidFrom", oMessageLayout.ValidFrom)
            cmd.Parameters.AddWithValue("@ValidTo", oMessageLayout.ValidTo)
            cmd.Parameters.AddWithValue("@UserCreate", oMessageLayout.UserCreate)
            cmd.Parameters.AddWithValue("@DateCreate", oMessageLayout.DateCreate)
            cmd.Parameters.AddWithValue("@UserUpdate", oMessageLayout.UserUpdate)
            cmd.Parameters.AddWithValue("@DateUpdate", oMessageLayout.DateUpdate)
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
