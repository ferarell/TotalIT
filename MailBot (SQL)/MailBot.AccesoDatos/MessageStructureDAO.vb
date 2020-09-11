Imports System.Data
Imports System.Configuration
Imports System.Data.SqlClient
Imports MailBot.Entidades
Public Class MessageStructureDAO

    Dim constr As String = ConfigurationManager.AppSettings("dbSolution").ToString()
    Dim Schema As String = ConfigurationManager.AppSettings("Schema").ToString()

    Public Sub Insert(oMessageStructure As MessageStructure)

        Dim query As String = Schema & "upMessageStructureInsert"
        Dim con As SqlConnection = New SqlConnection(constr)
        Dim cmd As SqlCommand = New SqlCommand(query)
        Try
            cmd.Parameters.AddWithValue("@IdMessageStructure", oMessageStructure.IdMessageStructure)
            cmd.Parameters.AddWithValue("@MessageStructureDescription", oMessageStructure.MessageStructureDescription)
            cmd.Parameters.AddWithValue("@IndexLocation", oMessageStructure.IndexLocation)
            cmd.Parameters.AddWithValue("@Signature", oMessageStructure.Signature)
            cmd.Parameters.AddWithValue("@UserCreate", oMessageStructure.UserCreate)
            cmd.Parameters.AddWithValue("@DateCreate", oMessageStructure.DateCreate)
            cmd.Parameters.AddWithValue("@UserUpdate", oMessageStructure.UserUpdate)
            cmd.Parameters.AddWithValue("@DateUpdate", oMessageStructure.DateUpdate)
            cmd.Connection = con
            con.Open()
            cmd.ExecuteNonQuery()


        Catch ex As Exception
        Finally
            con.Close()
        End Try


    End Sub


    Public Sub Update(oMessageStructure As MessageStructure)

        Dim query As String = Schema & "upMessageStructureUpdate"
        Dim con As SqlConnection = New SqlConnection(constr)
        Dim cmd As SqlCommand = New SqlCommand(query)

        Try
            cmd.Parameters.AddWithValue("@IdMessageStructure", oMessageStructure.IdMessageStructure)
            cmd.Parameters.AddWithValue("@MessageStructureDescription", oMessageStructure.MessageStructureDescription)
            cmd.Parameters.AddWithValue("@IndexLocation", oMessageStructure.IndexLocation)
            cmd.Parameters.AddWithValue("@Signature", oMessageStructure.Signature)
            cmd.Parameters.AddWithValue("@UserCreate", oMessageStructure.UserCreate)
            cmd.Parameters.AddWithValue("@DateCreate", oMessageStructure.DateCreate)
            cmd.Parameters.AddWithValue("@UserUpdate", oMessageStructure.UserUpdate)
            cmd.Parameters.AddWithValue("@DateUpdate", oMessageStructure.DateUpdate)
            cmd.Connection = con
            con.Open()
            cmd.ExecuteNonQuery()

        Catch ex As Exception

        Finally
            con.Close()
        End Try


    End Sub

    Public Sub Delete(oMessageStructure As MessageStructure)

        Dim query As String = Schema & "upMessageStructureDelete"
        Dim con As SqlConnection = New SqlConnection(constr)
        Dim cmd As SqlCommand = New SqlCommand(query)
        Try
            cmd.Parameters.AddWithValue("@IdMessageStructure", oMessageStructure.IdMessageStructure)
            cmd.Connection = con
            con.Open()
            cmd.ExecuteNonQuery()
        Catch ex As Exception

        Finally
            con.Close()
        End Try

    End Sub


    Public Function GetAll(oMessageStructure As MessageStructure) As List(Of MessageStructure)
        Dim lstMessageStructure As List(Of MessageStructure) = New List(Of MessageStructure)
        Dim query As String = Schema & "upGetMessageStructure"
        Dim con As SqlConnection = New SqlConnection(constr)
        Dim cmd As SqlCommand = New SqlCommand(query)
        Dim rdr As SqlDataReader

        Try

            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@IdMessageStructure", oMessageStructure.IdMessageStructure)
            cmd.Parameters.AddWithValue("@MessageStructureDescription", oMessageStructure.MessageStructureDescription)
            cmd.Parameters.AddWithValue("@IndexLocation", oMessageStructure.IndexLocation)
            cmd.Parameters.AddWithValue("@Signature", oMessageStructure.Signature)
            cmd.Parameters.AddWithValue("@UserCreate", oMessageStructure.UserCreate)
            cmd.Parameters.AddWithValue("@DateCreate", oMessageStructure.DateCreate)
            cmd.Parameters.AddWithValue("@UserUpdate", oMessageStructure.UserUpdate)
            cmd.Parameters.AddWithValue("@DateUpdate", oMessageStructure.DateUpdate)
            con.Open()
            rdr = cmd.ExecuteReader()





            While rdr.Read()
                Dim bMessageStructure As MessageStructure = New MessageStructure()

                bMessageStructure.IdMessageStructure = Convert.ToInt32(rdr("IdMessageLayout").ToString())
                bMessageStructure.MessageStructureDescription = rdr("IdMessageStructure").ToString()
                bMessageStructure.IndexLocation = Convert.ToInt32(rdr("IdInputConfiguration").ToString())
                bMessageStructure.Signature = Convert.ToBoolean(rdr("MessageText").ToString())
                bMessageStructure.UserCreate = rdr("UserCreate").ToString()
                bMessageStructure.DateCreate = Convert.ToDateTime(rdr("DateCreate").ToString())
                bMessageStructure.UserUpdate = rdr("UserUpdate").ToString()
                bMessageStructure.DateUpdate = Convert.ToDateTime(rdr("DateUpdate").ToString())

                lstMessageStructure.Add(bMessageStructure)
            End While


            Return lstMessageStructure

        Catch ex As Exception
        Finally
            con.Close()

        End Try


    End Function


    Public Function GetAllDataTable(oMessageStructure As MessageStructure) As DataTable
        Dim dt As DataTable = New DataTable()
        Dim query As String = Schema & "upGetMessageStructure"
        Dim con As SqlConnection = New SqlConnection(constr)
        Dim cmd As SqlCommand = New SqlCommand(query)
        Dim da As SqlDataAdapter

        Try

            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@IdMessageStructure", oMessageStructure.IdMessageStructure)
            cmd.Parameters.AddWithValue("@MessageStructureDescription", oMessageStructure.MessageStructureDescription)
            cmd.Parameters.AddWithValue("@IndexLocation", oMessageStructure.IndexLocation)
            cmd.Parameters.AddWithValue("@Signature", oMessageStructure.Signature)
            cmd.Parameters.AddWithValue("@UserCreate", oMessageStructure.UserCreate)
            cmd.Parameters.AddWithValue("@DateCreate", oMessageStructure.DateCreate)
            cmd.Parameters.AddWithValue("@UserUpdate", oMessageStructure.UserUpdate)
            cmd.Parameters.AddWithValue("@DateUpdate", oMessageStructure.DateUpdate)
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
