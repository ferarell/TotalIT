Imports System.Collections
Imports System.Data
Imports System.Windows.Forms
Imports System.Threading
Imports DevExpress.XtraEditors

Public Class ThisAddIn
    Dim outlookNameSpace As Outlook.NameSpace
    Dim inbox As Outlook.MAPIFolder
    Dim WithEvents items As Outlook.Items
    Dim oDataAccess As New DataAccess

    Private Sub ThisAddIn_Startup() Handles Me.Startup
        Try
            If Not My.Settings.GetPreviousVersion("DBFileName") Is Nothing Then
                'If My.Computer.Name <> "FARELLANO" Then
                My.Settings.Upgrade()
            End If
            'End If
            If Not IO.File.Exists(My.Settings.DBFileName) Then
                DevExpress.XtraEditors.XtraMessageBox.Show("Cannot find the database, please assign it from the configuration option. ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If
            outlookNameSpace = Me.Application.GetNamespace("MAPI")
            inbox = outlookNameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox)
            items = inbox.Items
            'MDBFileName = My.Settings.DBFileName
        Catch ex As Exception
            DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub

    Private Sub Items_ItemAdd(ByVal item As Object) Handles items.ItemAdd
        Dim mailItem As Outlook.MailItem = item
        Filter = GetFilter(mailItem)
        BlList.Clear()
        If Filter = "" Then
            dtConfig = oDataAccess.ExecuteAccessQuery("SELECT TOP 1 * FROM " & My.Settings.ConfigTableName & " WHERE ResponseType=0").Tables(0)
            SendNewMessage("MSG_ERROR", item, "", "")
            Return
        End If
        If mailItem.Subject Is Nothing Then
            mailItem.Subject = " "
        End If
        If mailItem.Body Is Nothing Then
            mailItem.Body = " "
        End If
        If TypeOf (item) Is Outlook.MailItem Then
            If mailItem.Sender Is Nothing Then
                Return
            End If
            If mailItem.Recipients.Item(1).Address <> mailItem.Sender.Address Then
                If mailItem.Subject.ToUpper.Contains(Filter) Or mailItem.Body.ToUpper.Contains(Filter) Then
                    dtConfig = oDataAccess.ExecuteAccessQuery("SELECT * FROM " & My.Settings.ConfigTableName & " WHERE Identifier = '" & Filter & "'").Tables(0)
                    'If Filter = "OBLI" Then
                    '    DataProcess1(item)
                    'End If
                    If dtConfig.Select("ResponseType=4").Length > 0 Then
                        dtConfig = dtConfig.Select("ResponseType=4").CopyToDataTable
                        DataProcess4(item, dtConfig.Rows(0)("Identificador"))
                    End If
                    If dtConfig.Select("ResponseType=3").Length > 0 Then
                        dtConfig = dtConfig.Select("ResponseType=3").CopyToDataTable
                        If dtConfig.Rows(0)("Identifier").ToString.ToUpper.Contains("BL") Or dtConfig.Rows(0)("Identifier").ToString.ToUpper.Contains("SWB") Then
                            dtCnfgLayout = oDataAccess.ExecuteAccessQuery("SELECT Label, Result1, Result2 FROM RobotConfigurationLayout WHERE Position=" & dtConfig.Rows(0)("Position") & " and ColumnType='R'").Tables(0)
                            DataProcess3(item, dtConfig.Rows(0)("Identifier"))
                        End If
                    End If
                    If dtConfig.Select("ResponseType=2").Length > 0 Then
                        dtConfig = dtConfig.Select("ResponseType=2").CopyToDataTable
                        'DataProcess2(item, dtConfig.Select("TipoRespuesta=2")(0)("Identificador"))
                    Else
                        SendNewMessage("OK", item, dtConfig.Rows(0)("Identifier"), "")
                    End If

                End If
            End If
        End If
    End Sub

    Private Sub DataProcess1(item As Object)
        Dim mailItem As Outlook.MailItem = item
        'Find in Subject
        For s = 1 To mailItem.Subject.Length - 1
            If Mid(mailItem.Subject.ToUpper, s, 4) = "HLCU" Then
                BlList.Add(Mid(mailItem.Subject.ToUpper, s, 16))
            End If
        Next
        'Find in Body
        For b = 1 To mailItem.Body.Length - 1
            If Mid(mailItem.Body.ToUpper, b, 4) = "HLCU" Then
                BlList.Add(Mid(mailItem.Body.ToUpper, b, 16))
            End If
        Next
        'FillDataQry1()
    End Sub

    'Private Sub DataProcess2(item As Object, identifier As String)
    '    Dim mailItem As Outlook.MailItem = item
    '    Dim sFileName = FileIO.FileSystem.GetTempFileName
    '    Dim oLogFileUpdate As New LogFileGenerate

    '    For a = 1 To mailItem.Attachments.Count
    '        If mailItem.Attachments(a).FileName.ToUpper.Contains({"XLS", "TXT"}) Then
    '            sFileName = My.Settings.AttachedFilePath & "\" & Format(Now, "ddMMyyyy HHmmss") & " - " & mailItem.Attachments(a).FileName
    '            mailItem.Attachments(a).SaveAsFile(sFileName)
    '            If Not IO.File.Exists(sFileName) Then
    '                DevExpress.XtraEditors.XtraMessageBox.Show("No se descargó el archivo adjunto.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '                Return
    '            End If
    '        End If
    '    Next
    '    'CheckForIllegalCrossThreadCalls = False
    '    Dim tProc As Thread = Nothing
    '    If identifier = "PTWO" And item.Attachments.Count > 0 Then
    '        Dim oPurchaseOrderControl As New PurchaseOrderControl
    '        tProc = New Thread(Sub() oPurchaseOrderControl.StartProcess(item))
    '    ElseIf item.Subject.ToUpper.Contains("HL-") And item.Body.ToUpper.Contains("BOOKING CONFIRMATION CANCELLATION") And item.Attachments.Count > 0 Then
    '        Dim oBookingCancellation As New BookingCancellation
    '        tProc = New Thread(Sub() oBookingCancellation.StartProcess(item))
    '    ElseIf identifier = "EQEO0801" Then
    '        Dim oReeferDataMasterUpdate As New ReeferDataMasterUpdate
    '        tProc = New Thread(Sub() oReeferDataMasterUpdate.DataProcess(sFileName))
    '    ElseIf identifier = "VOYC2502" Then
    '        Dim oSheduleLocalVoyageUpdate As New ScheduleLocalVoyageUpdate
    '        tProc = New Thread(Sub() oSheduleLocalVoyageUpdate.DataProcess(sFileName))
    '    ElseIf identifier = "VOYC3001" Then
    '        Dim oScheduleTranshipmentVoyageUpdate As New ScheduleTranshipmentVoyageUpdate
    '        tProc = New Thread(Sub() oScheduleTranshipmentVoyageUpdate.DataProcess(sFileName))
    '    ElseIf identifier = "RETORNO GENSET" Then
    '        Dim oReturnRequestGenset As New ReturnRequestGenset
    '        tProc = New Thread(Sub() oReturnRequestGenset.DataProcess(item, sFileName))
    '    ElseIf identifier = "UPDATE-GENSET" Then
    '        Dim oGensetDataTableUpdate As New GensetDataTableUpdate
    '        tProc = New Thread(Sub() oGensetDataTableUpdate.DataProcess(sFileName, item))
    '    End If
    '    Try
    '        tProc.Start()
    '    Catch ex As Exception
    '        oLogFileUpdate.TextFileUpdate(identifier, ex.Message)
    '        SendNewMessage("PRC_ERROR", item, identifier, ex.Message)
    '    End Try
    'End Sub

    Private Sub DataProcess3(item As Object, identifier As String)
        Dim mailItem As Outlook.MailItem = item
        Dim bIssued As Boolean
        Dim sFecha As String = ""
        Dim sQuery As String = ""
        Dim dtDataQry As New DataTable
        'Find in Subject
        For s = 1 To mailItem.Subject.Length - 1
            If Mid(mailItem.Subject.ToUpper, s, 4) = "HLCU" Then
                BlList.Add(Mid(mailItem.Subject.ToUpper, s, 16))
            End If
        Next
        'Find in Body
        For b = 1 To mailItem.Body.Length - 1
            If Mid(mailItem.Body.ToUpper, b, 4) = "HLCU" Then
                BlList.Add(Mid(mailItem.Body.ToUpper, b, 16))
            End If
        Next
        sQuery = dtConfig.Rows(0)("QuerySQL")
        If identifier.Contains("BL") Or identifier.Contains("SWB") Then
            dtQuery = oDataAccess.ExecuteAccessQuery(Replace(sQuery, "[BLNO]", "'#'")).Tables(0)
            For i = 0 To BlList.Count - 1
                If BlList(i).Trim <> "" Then
                    dtDataQry = oDataAccess.ExecuteAccessQuery(Replace(sQuery, "[BLNO]", "'" & BlList(i) & "'")).Tables(0)
                    bIssued = False
                    If dtDataQry.Rows.Count > 0 Then
                        bIssued = True
                        sFecha = dtDataQry.Rows(0)(2).ToString
                        dtQuery.Rows.Add(BlList(i), dtCnfgLayout.Rows(0)("Result1"), sFecha, dtDataQry(0)(3))
                    Else
                        dtQuery.Rows.Add(BlList(i), dtCnfgLayout.Rows(0)("Result2"), dtDataQry(3))
                    End If
                End If
            Next
        End If

    End Sub

    Private Sub DataProcess4(item As Object, identifier As String)
        Dim mailItem As Outlook.MailItem = item
        Dim tProc As Thread = Nothing
        If Not mailItem.Sender.Address.ToUpper.Contains("HLAG.COM") Then
            'Return
        End If
        Dim sPrm1, sPrm2, sPrm3, sPrm4, sPrm5, sPrm6 As String
        Dim sPrm7 As DateTime
        sPrm1 = Replace(Mid(mailItem.Subject, InStr(mailItem.Subject, "HLCU"), 16), "'", "")
        sPrm2 = mailItem.ReceivedTime.ToShortDateString
        sPrm3 = mailItem.SenderName
        sPrm4 = mailItem.To
        sPrm5 = Replace(mailItem.Subject, "'", "")
        sPrm6 = Environment.UserName
        sPrm7 = DateTime.Now
        Dim oUpdateBlIssued As New BlIssuedUpdate
        tProc = New Thread(Sub() oUpdateBlIssued.DataProcess(sPrm1, sPrm2, sPrm3, sPrm4, sPrm5, sPrm6, sPrm7))
        Try
            tProc.Start()
        Catch ex As Exception
            SendNewMessage("PRC_ERROR", item, identifier, ex.Message)
        End Try

    End Sub

    Friend Function GetFilter(mailItem As Outlook.MailItem) As String
        Dim sResult As String = ""
        If mailItem.Subject Is Nothing Then
            mailItem.Subject = " "
        End If
        dtSubjects.Rows.Clear()
        dtSubjects = oDataAccess.ExecuteAccessQuery("SELECT * FROM " & My.Settings.ConfigTableName).Tables(0)
        For r = 0 To dtSubjects.Rows.Count - 1
            If IsDBNull(dtSubjects.Rows(r)("Identifier")) Then
                Continue For
            End If
            If mailItem.Subject.ToUpper.Contains(dtSubjects.Rows(r)("Identifier")) Then
                sResult = dtSubjects.Rows(r)("Identifier")
                drConfig = dtSubjects.Rows(r)
                Exit For
            End If
        Next
        Return sResult
    End Function

    Public Sub SearchInBox()
        Dim inbox As Outlook.MAPIFolder = outlookNameSpace.ActiveExplorer.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox)
        Dim items As Outlook.Items = inbox.Items
        Dim mailItem As Outlook.MailItem = Nothing
        Dim folderItem As Object
        Dim subjectName As String = String.Empty
        Dim filter As String = "[Subject] > 's' And [Subject] <'u'"
        folderItem = items.Find(filter)
        While folderItem IsNot Nothing
            mailItem = TryCast(folderItem, Outlook.MailItem)
            If mailItem IsNot Nothing Then
                subjectName += vbLf + mailItem.Subject
            End If
            folderItem = items.FindNext()
        End While
        subjectName = " The following e-mail messages were found: " + subjectName
        MessageBox.Show(subjectName)
    End Sub

    Private Sub Application_Startup() Handles Application.Startup

    End Sub

End Class
