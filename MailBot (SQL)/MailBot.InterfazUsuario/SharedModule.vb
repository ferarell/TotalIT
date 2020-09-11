Imports DevExpress.XtraRichEdit
Imports System.Windows.Forms
Imports System.Data
Imports System.Collections
Imports System.IO


Module SharedModule
    Friend BlList As New ArrayList
    Friend dtConfig As New DataTable
    Friend drConfig As DataRow
    Friend Identifier As String = ""
    Friend AppPath As String = IO.Path.GetDirectoryName(New Uri(System.Reflection.Assembly.GetExecutingAssembly().CodeBase).LocalPath)
    Friend oGeneralDAO As New AppService.MailBotServicesWCFClient

    Friend Function ExecuteSQL(QueryString As String) As DataTable
        Dim dsResult As New DataSet
        Try
            dsResult = oGeneralDAO.NewExecuteSQL(QueryString)
            If dsResult.Tables.Contains("Error") Then
                If dsResult.Tables("Error").Rows.Count > 0 Then
                    Throw New ApplicationException(dsResult.Tables("Error")(0)(0))
                End If
            End If
        Catch ex As Exception
            Throw
            DevExpress.XtraEditors.XtraMessageBox.Show(dsResult.Tables("Error")(0)(0), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Return dsResult.Tables(0)
    End Function

    Friend Function ExecuteSQLNonQuery(QueryString As String) As Boolean
        Dim bResult As Boolean = True
        Dim aResult As New ArrayList
        aResult.AddRange(oGeneralDAO.NewExecuteSQLNonQuery(QueryString))
        If Not aResult(0) Then
            DevExpress.XtraEditors.XtraMessageBox.Show(aResult(1), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            bResult = False
        End If
        Return bResult
    End Function

    Friend Function LoadExcel(ByVal FileName As String, ByRef Hoja As String) As DataSet
        Dim dsResult As New DataSet
        Dim ExcelConnectionString As String = "provider=Microsoft.ACE.OLEDB.12.0; Data Source='" & FileName & "'; Extended Properties=Excel 8.0;"
        Using connection As New System.Data.OleDb.OleDbConnection(ExcelConnectionString)
            Try
                connection.Open()
                If Hoja = "{0}" Then
                    For r = 0 To connection.GetSchema("Tables").Rows.Count - 1
                        If Not connection.GetSchema("Tables").Rows(r)("TABLE_NAME").toupper.contains("FILTER") Then
                            Hoja = connection.GetSchema("Tables").Rows(r)("TABLE_NAME")
                            Exit For
                        End If
                    Next
                End If
                Dim Command As New System.Data.OleDb.OleDbDataAdapter("select * from [" & Hoja & "]", connection)
                Command.Fill(dsResult)
            Catch ex As Exception
                DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Finally
                connection.Close()
            End Try
            Return dsResult
        End Using
    End Function

    Friend Function QueryExcel(FileName As String, Query As String) As DataSet
        Dim dsResult As New DataSet
        Dim ExcelConnectionString As String = "provider=Microsoft.ACE.OLEDB.12.0; Data Source='" & FileName & "'; Extended Properties=Excel 8.0;"
        Using connection As New System.Data.OleDb.OleDbConnection(ExcelConnectionString)
            Try
                connection.Open()
                Dim Command As New System.Data.OleDb.OleDbDataAdapter(Query, connection)
                Command.Fill(dsResult)
            Catch ex As Exception
                'DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Finally
                connection.Close()
            End Try
            Return dsResult
        End Using
    End Function

    'Friend Function SendNewMessage(TypMsg As String, oMailItem As Outlook.MailItem, Identifier As String, Message As String, Attachments As ArrayList) As Boolean
    '    Dim NewMessage As Outlook.MailItem
    '    Dim AppOutlook As New Outlook.Application
    '    Dim bResult As Boolean = True
    '    NewMessage = AppOutlook.CreateItem(Outlook.OlItemType.olMailItem)
    '    Dim Recipents As Outlook.Recipients = NewMessage.Recipients
    '    Recipents.Add(oMailItem.SenderEmailAddress)
    '    NewMessage.BodyFormat = Outlook.OlBodyFormat.olFormatHTML
    '    If TypMsg = "OK" Then
    '        'Valid Subject
    '        NewMessage.Subject = oMailItem.Subject
    '        NewMessage.HTMLBody = "" 'GetValidMessageBody(oMailItem.SenderName)
    '    ElseIf TypMsg = "MSG_ERROR" Then
    '        'Invalid Subject
    '        NewMessage.Subject = "Asunto de mensaje inválido"
    '        NewMessage.HTMLBody = "" 'GetInvalidMessageBody(oMailItem.SenderName)
    '    ElseIf TypMsg = "PRC_OK" Then
    '        If Attachments.Count > 0 Then
    '            For i = 0 To Attachments.Count - 1
    '                NewMessage.Attachments.Add(Attachments(i))
    '            Next
    '        End If
    '        NewMessage.Subject = oMailItem.Subject
    '        NewMessage.BCC = "aremonfe@gmail.com"
    '        If oMailItem.HTMLBody.Trim = "" Then
    '            NewMessage.HTMLBody = GetProcessResponseMessage(True, oMailItem.SenderName, Message, 0)
    '        Else
    '            NewMessage.HTMLBody = GetCustomProcessResponseMessage(oMailItem)
    '        End If
    '    ElseIf TypMsg = "PRC_ERROR" Then
    '        NewMessage.Subject = "Error al procesar " & Identifier
    '        'NewMessage.CC = "itsupport@delfingroupco.com.pe"
    '        NewMessage.BCC = "aremonfe@gmail.com"
    '        NewMessage.HTMLBody = GetProcessResponseMessage(False, oMailItem.SenderName, Message, 0)
    '    End If
    '    NewMessage.Send()
    '    oMailItem.Close(Microsoft.Office.Interop.Outlook.OlInspectorClose.olDiscard)
    '    Return bResult
    'End Function

    Friend Function GetProcessResponseMessage(bProcess As Boolean, SenderName As String, Message As String, iProcessType As Integer) As String
        Dim oText As New RichTextBox
        oText.AppendText("<html><body lang=ES style='tab-interval:35.4pt;font-size:10.0pt;font-family:""Tahoma"",sans-serif'>")
        oText.AppendText("Estimado(a): " & SenderName & "<br><br>")
        If bProcess Then
            oText.AppendText("Los datos asociados al proceso " & Identifier & " se han procesado satisfactoriamente." & "<br><br>")
        Else
            oText.AppendText("El proceso asociado al identificador " & Identifier & " ha generado el siguiente error:" & "<br><br>")
            oText.AppendText(Message & "<br><br>")
            oText.AppendText("Los datos no han sido actualizados." & "<br><br>")
        End If
        oText.AppendText(drConfig("Signature"))
        Return oText.Text
    End Function

    'Friend Function GetCustomProcessResponseMessage(oMailItem As Outlook.MailItem) As String
    '    Dim oText As New RichTextBox
    '    oText.AppendText("<html><body lang=ES style='tab-interval:35.4pt;font-size:10.0pt;font-family:""Tahoma"",sans-serif'>")
    '    oText.AppendText("Estimado(a): " & oMailItem.SenderName & "<br><br>")
    '    oText.AppendText(oMailItem.HTMLBody & "<br><br>")
    '    Return oText.Text
    'End Function

    <System.Runtime.CompilerServices.Extension> _
    Public Function Contains(ByVal str As String, ByVal ParamArray values As String()) As Boolean
        For Each value In values
            If str.Contains(value) Then
                Return True
            End If
        Next
        Return False
    End Function

    Friend Function SelectDistinct(ByVal SourceTable As System.Data.DataTable, ByVal Condition As String, ByVal ParamArray FieldNames() As String) As System.Data.DataTable
        Dim lastValues() As Object
        Dim newTable As System.Data.DataTable

        If FieldNames Is Nothing OrElse FieldNames.Length = 0 Then
            Throw New ArgumentNullException("FieldNames")
        End If

        lastValues = New Object(FieldNames.Length - 1) {}
        newTable = New System.Data.DataTable

        For Each field As String In FieldNames
            newTable.Columns.Add(field, SourceTable.Columns(field).DataType)
        Next

        For Each Row As DataRow In SourceTable.Select(Condition, String.Join(", ", FieldNames))
            If Not fieldValuesAreEqual(lastValues, Row, FieldNames) Then
                newTable.Rows.Add(createRowClone(Row, newTable.NewRow(), FieldNames))

                setLastValues(lastValues, Row, FieldNames)
            End If
        Next

        Return newTable
    End Function

    Friend Function fieldValuesAreEqual(ByVal lastValues() As Object, ByVal currentRow As DataRow, ByVal fieldNames() As String) As Boolean
        Dim areEqual As Boolean = True

        For i As Integer = 0 To fieldNames.Length - 1
            If lastValues(i) Is Nothing OrElse Not lastValues(i).Equals(currentRow(fieldNames(i))) Then
                areEqual = False
                Exit For
            End If
        Next

        Return areEqual
    End Function

    Private Function createRowClone(ByVal sourceRow As DataRow, ByVal newRow As DataRow, ByVal fieldNames() As String) As DataRow
        For Each field As String In fieldNames
            newRow(field) = sourceRow(field)
        Next

        Return newRow
    End Function

    Private Sub setLastValues(ByVal lastValues() As Object, ByVal sourceRow As DataRow, ByVal fieldNames() As String)
        For i As Integer = 0 To fieldNames.Length - 1
            lastValues(i) = sourceRow(fieldNames(i))
        Next
    End Sub

    Friend Function UnicodeStringToBytes(ByVal str As String) As Byte()
        Return System.Text.Encoding.Unicode.GetBytes(str)
    End Function

    Friend Function UnicodeBytesToString(ByVal bytes() As Byte) As String
        Return System.Text.Encoding.Unicode.GetString(bytes)
    End Function

End Module
