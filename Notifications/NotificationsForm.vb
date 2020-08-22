Imports DevExpress.XtraEditors.DXErrorProvider
Imports DevExpress.XtraGrid.Views.Grid
Imports System.Drawing
Imports System.IO
Imports System.Globalization
Imports DevExpress.XtraSplashScreen
Imports System.Threading
Imports DevExpress.XtraGrid.Views.Grid.ViewInfo
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices
Imports System.Net.Mail
Imports DevExpress.XtraRichEdit
Imports System.Net.Mime
Imports System.Text.RegularExpressions
Imports DevExpress.Export
Imports DevExpress.XtraPrinting

Public Class NotificationsForm
    Dim eMailTo As String = ""
    Dim dtContacts, dtBookings, dtBlackList As New DataTable
    Dim oAppService As New AppService.GlobalServiceClient

    Private Sub NotificationsForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim iWidth As Integer = 400
        SplitContainerControl1.SplitterPosition = Me.Size.Height - 350
        SplitContainerControl1.Panel2.Visible = False
        SplitContainerControl2.SplitterPosition = iWidth
        SplitContainerControl3.SplitterPosition = Me.Size.Width - iWidth
        ImageListBoxControl1.Items.Clear()
        ImageListBoxControl1.MultiColumn = True
        DevExpress.LookAndFeel.UserLookAndFeel.Default.SetSkinStyle(SkinName)
        bsiVersion.Caption = My.User.Name
    End Sub

    Private Sub bbiOpenFile_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiOpenFile.ItemClick
        OpenFileDialog2.Filter = "Source Files (*.doc*;*.htm*;*.mht;*.rtf;*.odt)|*.doc*;*.htm*;*.mht;*.rtf;*.odt"
        If Not OpenFileDialog2.ShowDialog = DialogResult.OK Then
            Return
        End If
        If OpenFileDialog2.FileName.ToUpper.Contains(".HTM") Then
            richEditControl.LoadDocument(OpenFileDialog2.FileName, DevExpress.XtraRichEdit.DocumentFormat.Html)
        ElseIf OpenFileDialog2.FileName.ToUpper.Contains(".DOC") Then
            richEditControl.LoadDocument(OpenFileDialog2.FileName)
        ElseIf OpenFileDialog2.FileName.ToUpper.Contains(".MHT") Then
            richEditControl.LoadDocument(OpenFileDialog2.FileName, DevExpress.XtraRichEdit.DocumentFormat.Mht)
        ElseIf OpenFileDialog2.FileName.ToUpper.Contains(".RTF") Then
            richEditControl.LoadDocument(OpenFileDialog2.FileName, DevExpress.XtraRichEdit.DocumentFormat.Rtf)
        ElseIf OpenFileDialog2.FileName.ToUpper.Contains(".ODT") Then
            richEditControl.LoadDocument(OpenFileDialog2.FileName, DevExpress.XtraRichEdit.DocumentFormat.OpenDocument)
        End If
    End Sub

    Private Sub bbiLoadContacts_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiLoadContacts.ItemClick

    End Sub

    Private Sub bbiAttachFiles_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiAttachFiles.ItemClick
        OpenFileDialog1.ShowDialog()
        ImageListBoxControl1.Items.Clear()
        For r = 0 To OpenFileDialog1.FileNames.Count - 1
            ImageListBoxControl1.Items.Add(OpenFileDialog1.FileNames(r))
        Next

    End Sub

    Private Sub bbiMessagePreview_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiMessagePreview.ItemClick
        Validate()
        Dim aFiles As New ArrayList
        Dim oRow As DataRow
        For i = 0 To ImageListBoxControl1.Items.Count - 1
            aFiles.Add(ImageListBoxControl1.Items(i).Value)
        Next
        eMailTo = edtTO.Text.Trim
        SplashScreenManager.ShowForm(Me, GetType(WaitForm), True, True, False)
        SplashScreenManager.Default.SetWaitFormDescription("Loading eMail Adresses List")
        For r = 0 To GridView1.RowCount - 1
            oRow = GridView1.GetDataRow(r)
            If IsDBNull(oRow("Checked")) Then
                Continue For
            End If
            If oRow("Checked") Then
                eMailTo += IIf(eMailTo.Trim.Length > 0, ";", "") & oRow("ENTC_Email").ToString.Trim
            End If
        Next
        SplashScreenManager.CloseForm(False)
        If eMailTo = "" Then
            DevExpress.XtraEditors.XtraMessageBox.Show(Me.LookAndFeel, "You must select at least one email", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        CreateSendItem(edtSubject.Text, richEditControl.HtmlText, aFiles)
    End Sub

    Private Sub bbiSendMessage_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiSendMessage.ItemClick
        'If GridView2.RowCount = 0 Then
        '    DevExpress.XtraEditors.XtraMessageBox.Show("You must load bookings list.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        '    Return
        'End If
        If DevExpress.XtraEditors.XtraMessageBox.Show(Me.LookAndFeel, "Are you sure to send?", "Confirmation", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No Then
            Return
        End If
        Dim dtBookingByMatchCode As New DataTable
        Dim eMail As String = ""
        Dim oMatchCode1 As String = ""
        Dim oMatchCode2 As String = ""
        Validate()
        SplashScreenManager.ShowForm(Me, GetType(WaitForm), True, True, False)
        SplashScreenManager.Default.SetWaitFormDescription("Sending Messages")
        For r = 0 To GridView1.RowCount - 1
            Try
                Dim oRow As DataRow = GridView1.GetDataRow(r)
                If IsDBNull(oRow("Checked")) Then
                    Continue For
                End If
                If Not oRow("Checked") Then
                    Continue For
                End If
                SendMessageByMail(oRow("ENTC_RazonSocial"), oRow("ENTC_Email"))
            Catch ex As Exception

            End Try
        Next
        SplashScreenManager.CloseForm(False)
    End Sub

    Private Sub SendMessageByMail(Company As String, eMail As String)
        Dim Application As New Outlook.Application
        Dim mail As Outlook.MailItem = Nothing
        Dim mailRecipients As Outlook.Recipients = Nothing
        Dim mailRecipient As Outlook.Recipient = Nothing
        'Dim mailSender As Outlook.AddressEntry
        Try
            mail = Application.CreateItem(Outlook.OlItemType.olMailItem)
            mail.Subject = edtSubject.Text
            'mail.HTMLBody = " <img src =""~\delfin_logo_notificaciones.png"">"
            mail.HTMLBody = Replace(richEditControl.HtmlText, "[COMPANY]", Company)
            For i = 0 To ImageListBoxControl1.Items.Count - 1
                mail.Attachments.Add(ImageListBoxControl1.Items(i).Value)
            Next
            'mail.Sender.Name = "Delfin Group Co. SAC"
            mail.To = eMail & ";" & edtTO.Text
            mail.CC = edtCC.Text
            mail.BCC = edtBCC.Text
            mail.Send()
        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message,
                "An exception is occured in the code of add-in.")
        Finally
            If Not IsNothing(mailRecipient) Then Marshal.ReleaseComObject(mailRecipient)
            If Not IsNothing(mailRecipients) Then Marshal.ReleaseComObject(mailRecipients)
            If Not IsNothing(mail) Then Marshal.ReleaseComObject(mail)
        End Try
    End Sub

    Private Sub SendMessageByNetMail(Company As String, eMail As String)
        Dim oSendMail As New MessageSender.SendMessage
        Dim _message As New MailMessage
        Dim _BodyHtml As New RichEditControl
        _BodyHtml.LoadDocument(OpenFileDialog2.FileName, DevExpress.XtraRichEdit.DocumentFormat.Html)
        _message.Subject = edtSubject.Text
        _message.From = New MailAddress("delfin@delfingroupco.com.pe", "Delfin Group Co. SAC")
        _message.IsBodyHtml = True
        _message.Body = Replace(_BodyHtml.HtmlText, "[COMPANY]", Company)
        _message.Priority = MailPriority.Normal
        _message.To.Add(eMail)
        If edtCC.Text.Trim <> "" Then
            _message.CC.Add(edtCC.Text)
        End If
        If edtBCC.Text.Trim <> "" Then
            _message.Bcc.Add(edtBCC.Text)
        End If
        For i = 0 To ImageListBoxControl1.Items.Count - 1
            Dim sAttach As String = ImageListBoxControl1.Items(i).Value
            Dim imageBody As String = " <img src =""cid:" & sAttach & """>"
            Dim htmlView As System.Net.Mail.AlternateView = System.Net.Mail.AlternateView.CreateAlternateViewFromString(_BodyHtml.HtmlText, Nothing, "text/html")
            Dim plainTextView As System.Net.Mail.AlternateView = System.Net.Mail.AlternateView.CreateAlternateViewFromString(_BodyHtml.HtmlText, Nothing, "text/plain")
            Dim imageResource As New System.Net.Mail.LinkedResource(sAttach, MediaTypeNames.Image.Gif)
            Dim attachment As System.Net.Mail.Attachment
            attachment = New System.Net.Mail.Attachment(sAttach)
            _message.Attachments.Add(attachment)
            'imageResource.ContentId = "HDIImage"
            htmlView.LinkedResources.Add(imageResource)
            htmlView.TransferEncoding = Net.Mime.TransferEncoding.Base64
            htmlView.ContentType = New System.Net.Mime.ContentType("text/html")
            _message.AlternateViews.Add(htmlView)
        Next

        oSendMail.SendMail(_message, False, Nothing)
    End Sub

    'Private Sub SendBookingsByMail(dtSource As DataTable, eMail As String)
    '    Dim Application As New Outlook.Application
    '    Dim mail As Outlook.MailItem = Nothing
    '    Dim mailRecipients As Outlook.Recipients = Nothing
    '    Dim mailRecipient As Outlook.Recipient = Nothing
    '    Dim BookingsList As New RichTextBox
    '    Dim sMatchCode As String = ""
    '    Try
    '        For r = 0 To dtSource.Rows.Count - 1
    '            BookingsList.AppendText(dtSource.Rows(r)("F1") & "<br>")
    '            sMatchCode = dtSource.Rows(r)("F2")
    '        Next
    '        mail = Application.CreateItem(Outlook.OlItemType.olMailItem)
    '        mail.Subject = edtSubject.Text
    '        mail.HTMLBody = Replace(richEditControl.HtmlText, "[BookingsList]", BookingsList.Text)
    '        mail.HTMLBody = Replace(mail.HTMLBody, "[MatchCode]", sMatchCode)
    '        mail.HTMLBody += "<br><br> " & eMail
    '        For i = 0 To ImageListBoxControl1.Items.Count - 1
    '            mail.Attachments.Add(ImageListBoxControl1.Items(i).Value)
    '        Next
    '        mail.To = eMail & ";" & edtTO.Text
    '        mail.CC = edtCC.Text
    '        mail.BCC = edtBCC.Text
    '        mail.Send()
    '    Catch ex As Exception
    '        System.Windows.Forms.MessageBox.Show(ex.Message,
    '            "An exception is occured in the code of add-in.")
    '    Finally
    '        If Not IsNothing(mailRecipient) Then Marshal.ReleaseComObject(mailRecipient)
    '        If Not IsNothing(mailRecipients) Then Marshal.ReleaseComObject(mailRecipients)
    '        If Not IsNothing(mail) Then Marshal.ReleaseComObject(mail)
    '    End Try
    'End Sub

    Friend Sub CreateSendItem(Subject As String, Body As String, AttachFile As ArrayList)
        Dim Application As New Outlook.Application
        Dim mail As Outlook.MailItem = Nothing
        Dim mailRecipients As Outlook.Recipients = Nothing
        Dim mailRecipient As Outlook.Recipient = Nothing
        Try
            mail = Application.CreateItem(Outlook.OlItemType.olMailItem)
            mail.Subject = Subject
            mail.HTMLBody = Body
            If AttachFile.Count > 0 Then
                For f = 0 To AttachFile.Count - 1
                    mail.Attachments.Add(AttachFile(f))
                Next
            End If
            mail.To = edtTO.Text
            mail.BCC = eMailTo
            mail.CC = edtCC.Text
            SplashScreenManager.ShowForm(Me, GetType(WaitForm), True, True, False)
            SplashScreenManager.Default.SetWaitFormDescription("Creating a New Message")
            mail.Display()
            SplashScreenManager.CloseForm(False)
        Catch ex As Exception
            SplashScreenManager.CloseForm(False)
            System.Windows.Forms.MessageBox.Show(ex.Message,
                "An exception is occured in the code of add-in.")
        Finally
            If Not IsNothing(mailRecipient) Then Marshal.ReleaseComObject(mailRecipient)
            If Not IsNothing(mailRecipients) Then Marshal.ReleaseComObject(mailRecipients)
            If Not IsNothing(mail) Then Marshal.ReleaseComObject(mail)
        End Try

    End Sub

    Private Sub bbiClose_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiClose.ItemClick
        Close()
    End Sub

    Private Sub ImageListBoxControl1_KeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyValue = Keys.Delete Then
            ImageListBoxControl1.Items.RemoveAt(ImageListBoxControl1.SelectedIndex)
        End If
    End Sub

    Private Sub rgMatchCode_SelectedIndexChanged(sender As Object, e As EventArgs)
        FilterContactsByMatchCode()
    End Sub

    Private Sub FilterContactsByMatchCode()
        Dim iPos As Integer = 0
        If MemoEdit1.Text.Length = 0 Then
            Return
        End If
        Dim sFilter As String = ""
        If rgMatchCode.SelectedIndex = 0 Then
            GridView1.ActiveFilterString = ""
        End If
        If rgMatchCode.SelectedIndex = 1 Then
            GridView1.ActiveFilterString = ""
            For l = 0 To MemoEdit1.Lines.Count - 1
                If MemoEdit1.Lines(l).Length > 0 Then
                    iPos = InStr(1, MemoEdit1.Lines(l), " ") - 1
                    sFilter += IIf(l > 0, " OR ", "") & "[MatchCode] = '" & Mid(MemoEdit1.Lines(l), 1, iPos) & Space(1) & Trim(Mid(MemoEdit1.Lines(l), iPos + 1, Len(MemoEdit1.Lines(l)))) & "'"
                End If
            Next
            GridView1.ActiveFilterCriteria = DevExpress.Data.Filtering.CriteriaOperator.Or
            GridView1.ActiveFilterString = "(" & sFilter & ")"
        End If
    End Sub

    Private Sub SeleccionaFilas(caso As Integer)
        Dim i As Integer = 0
        Do While i < GridView1.RowCount
            Dim row As DataRow = GridView1.GetDataRow(i)
            If caso = 0 Then
                row("Checked") = True
            End If
            If caso = 1 Then
                row("Checked") = False
            End If
            If caso = 2 Then
                If row("Checked") Then
                    row("Checked") = False
                Else
                    row("Checked") = True
                End If
            End If
            i += 1
        Loop
    End Sub

    Private Sub SeleccionaTodosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SeleccionaTodosToolStripMenuItem.Click
        SeleccionaFilas(0)
    End Sub

    Private Sub DeseleccionaTodosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DeseleccionaTodosToolStripMenuItem.Click
        SeleccionaFilas(1)
    End Sub

    Private Sub InvertirSelecciónToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles InvertirSelecciónToolStripMenuItem.Click
        SeleccionaFilas(2)
    End Sub

    Private Sub RepositoryItemCheckEdit1_CheckedChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub bbiConfiguration_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiConfiguration.ItemClick
        Dim oForm As New PreferencesForm
        oForm.ShowDialog()
    End Sub

    Private Sub NotificationsForm_TextChanged(sender As Object, e As EventArgs) Handles MyBase.TextChanged
        Me.Text = My.Application.Info.ProductName + " [" + My.Application.Info.Version.ToString + "]"
    End Sub

    Public Sub New()

        ' Llamada necesaria para el diseñador.
        InitializeComponent()

        DevExpress.Skins.SkinManager.EnableFormSkins()
        DevExpress.UserSkins.BonusSkins.Register()

        SkinName = My.Settings.LookAndFeel

    End Sub

    Private Sub bbiImportContacts_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiImportContacts.ItemClick
        Dim dtSource As New DataTable
        OpenFileDialog3.Filter = "Import File (*.xls*)|*.xls*"
        If Not OpenFileDialog3.ShowDialog = DialogResult.OK Then
            Return
        End If
        SplashScreenManager.ShowForm(Me, GetType(WaitForm), True, True, False)
        SplashScreenManager.Default.SetWaitFormDescription("Load Data Source of Contacts")
        dtSource = LoadExcelWC(OpenFileDialog3.FileName, "Data-MSG-Defaults per MC filter$", "F10='EMAIL'").Tables(0)
        If dtSource.Rows.Count = 0 Then
            Return
        End If
        Dim dtResult As New DataTable
        Dim sMatchCode As String = ""
        dtResult.Columns.Add("MatchCode", GetType(String)).DefaultValue = ""
        dtResult.Columns.Add("eMail", GetType(String)).DefaultValue = ""
        dtResult.Columns.Add("Status", GetType(String)).DefaultValue = "A"
        dtResult.Columns.Add("CreatedBy", GetType(String)).DefaultValue = ""
        dtResult.Columns.Add("CreatedDate", GetType(Date)).DefaultValue = Now
        SplashScreenManager.Default.SetWaitFormDescription("Delete Current Contacts")
        ExecuteAccessQuery("delete from Contacts")
        For r = 0 To dtSource.Rows.Count - 1
            Dim oRow As DataRow = dtSource.Rows(r)
            sMatchCode = oRow("F1").ToString.Trim & Space(1) & CInt(oRow("F2")).ToString
            oRow("F11") = Replace(oRow("F11"), "'", "").Trim
            dtResult.Rows.Add(sMatchCode, oRow("F11"), "A", My.User.Name, Now)
            SplashScreenManager.Default.SetWaitFormDescription("Insert eMail Contacts (" & (r + 1).ToString & " of " & dtSource.Rows.Count.ToString & ")")
            InsertIntoAccess("Contacts", dtResult.Rows(r))
        Next
        SplashScreenManager.CloseForm(False)
        'bbiLoadContacts.PerformClick()
    End Sub

    Private Sub gcBookingFilter_EmbeddedNavigator_ButtonClick(sender As Object, e As DevExpress.XtraEditors.NavigatorButtonClickEventArgs) Handles gcBookingFilter.EmbeddedNavigator.ButtonClick
        If e.Button.Tag = "LoadExcel" Then
            OpenFileDialog3.Filter = "Import File (*.xls*)|*.xls*"
            If Not OpenFileDialog3.ShowDialog = DialogResult.OK Then
                Return
            End If
            SplashScreenManager.ShowForm(Me, GetType(WaitForm), True, True, False)
            SplashScreenManager.Default.SetWaitFormDescription("Load Excel Data File")
            dtBookings.Rows.Clear()
            dtBookings = LoadExcelWH(OpenFileDialog3.FileName, "{0}").Tables(0)
            SplashScreenManager.CloseForm(False)
            If dtBookings.Rows.Count = 0 Then
                Return
            End If

            dtBookings = UpdateDataTable(SelectDistinct(dtBookings, "", "F1", "F2"))
            gcBookingFilter.DataSource = dtBookings
            FilterContactsByBookings(SelectDistinct(dtBookings, "", "F2"))
        End If
    End Sub

    Private Sub gcBlackList_EmbeddedNavigator_ButtonClick(sender As Object, e As DevExpress.XtraEditors.NavigatorButtonClickEventArgs) Handles gcBlackList.EmbeddedNavigator.ButtonClick
        If e.Button.Tag = "LoadText" Then
            OpenFileDialog3.Filter = "Import File (*.txt)|*.txt"
            If Not OpenFileDialog3.ShowDialog = DialogResult.OK Then
                Return
            End If
            SplashScreenManager.ShowForm(Me, GetType(WaitForm), True, True, False)
            SplashScreenManager.Default.SetWaitFormDescription("Load Text File")
            dtBlackList.Rows.Clear()
            dtBlackList = LoadBlackListTextFile(OpenFileDialog3.FileName)
            SplashScreenManager.CloseForm(False)
            If dtBookings.Rows.Count = 0 Then
                Return
            End If
        End If
    End Sub

    Friend Function LoadBlackListTextFile(FileName As String) As DataTable
        Dim iPosition As Integer = 0
        Dim sEqpType, sMatchCode As String
        Dim dtResult As New DataTable
        dtResult.Columns.Add("MatchCode", GetType(String)).DefaultValue = ""

        Try
            SplashScreenManager.ShowForm(Me, GetType(WaitForm), True, True, False)
            SplashScreenManager.Default.SetWaitFormDescription("Clean Black List")
            ExecuteAccessNonQuery("delete from BlackList")
            SplashScreenManager.Default.SetWaitFormDescription("Update Black List")
            Using sr As New StreamReader(FileName)
                Dim lines As List(Of String) = New List(Of String)
                Dim line As String = ""
                Dim bExit As Boolean = False
                Do While Not sr.EndOfStream
                    line = sr.ReadLine()
                    If Mid(line, 105, 5) = "INACT" Or Mid(line, 105, 5) = "INVAL" Then
                        lines.Add(line)
                    End If
                Loop
                Dim bSkip As Boolean = True
                For i As Integer = 0 To lines.Count - 1
                    SplashScreenManager.Default.SetWaitFormDescription("Update Black List (Row: " & (i + 1).ToString & " of " & (lines.Count - 1).ToString & ")")
                    If Mid(lines(i), 23, 10).Trim = "" Then
                        Continue For
                    End If
                    'If Mid(lines(i), 105, 5).Trim = "ACTIV" Then
                    '    Continue For
                    'End If
                    sMatchCode = Mid(lines(i), 23, 7).Trim & Space(1) & CInt(Mid(lines(i), 30, 3)).ToString
                    If ExecuteAccessQuery("select * from Contacts where MatchCode = '" & sMatchCode & "'").Tables(0).Rows.Count = 0 Then
                        Continue For
                    End If
                    dtResult.Rows.Add()
                    iPosition = dtResult.Rows.Count - 1
                    dtResult.Rows(iPosition).Item(0) = sMatchCode
                    InsertIntoAccess("BlackList", dtResult.Rows(iPosition))
                Next
            End Using
            gcBlackList.DataSource = dtResult
            SplashScreenManager.CloseForm(False)
            DevExpress.XtraEditors.XtraMessageBox.Show(Me.LookAndFeel, "The process has been completed successfully", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            SplashScreenManager.CloseForm(False)
            DevExpress.XtraEditors.XtraMessageBox.Show(Me.LookAndFeel, ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Return dtResult
    End Function

    Friend Function UpdateDataTable(dtSource As DataTable) As DataTable
        Dim IPos As Integer = 0
        For r = 0 To dtSource.Rows.Count - 1
            Dim oRow As DataRow = dtSource.Rows(r)
            IPos = InStr(oRow("F2").ToString.Trim, " ") - 1
            oRow("F2") = Mid(oRow("F2"), 1, IPos) & Space(1) & Trim(CInt(Mid(oRow("F2"), IPos + 1, Len(oRow("F2")))).ToString)
        Next
        Return dtSource
    End Function

    Private Sub bbiCustomers_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiCustomers.ItemClick
        Dim dtQuery As New DataTable
        SplashScreenManager.ShowForm(Me, GetType(WaitForm), True, True, False)
        SplashScreenManager.Default.SetWaitFormDescription("Load Contacts")
        dtQuery = oAppService.ExecuteSQL("EXEC TotalIT.dbo.upGetMailCustomer").Tables(0)
        If Not dtQuery.Columns.Contains("Checked") Then
            dtQuery.Columns.Add("Checked", GetType(Boolean)).DefaultValue = False
        End If
        'SplashScreenManager.Default.SetWaitFormDescription("Load Black List")
        'dtBlackList = ExecuteAccessQuery("select distinct * from BlackList").Tables(0)
        SplashScreenManager.CloseForm(False)
        gcBlackList.DataSource = dtBlackList
        gcMainData.DataSource = dtQuery
    End Sub

    Private Sub bbiSuppliers_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiSuppliers.ItemClick
        Dim dtQuery As New DataTable
        SplashScreenManager.ShowForm(Me, GetType(WaitForm), True, True, False)
        SplashScreenManager.Default.SetWaitFormDescription("Load Contacts")
        dtQuery = oAppService.ExecuteSQL("EXEC TotalIT.ntf.paObtieneListaProveedores").Tables(0)
        If Not dtQuery.Columns.Contains("Checked") Then
            dtQuery.Columns.Add("Checked", GetType(Boolean)).DefaultValue = False
        End If
        'SplashScreenManager.Default.SetWaitFormDescription("Load Black List")
        'dtBlackList = ExecuteAccessQuery("select distinct * from BlackList").Tables(0)
        SplashScreenManager.CloseForm(False)
        gcBlackList.DataSource = dtBlackList
        gcMainData.DataSource = dtQuery
    End Sub

    Private Sub NotificationsForm_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If DevExpress.XtraEditors.XtraMessageBox.Show(Me.LookAndFeel, "Are you sure to exit?", "Confirmation", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
            Try
                For f = 1 To Application.OpenForms.Count - 1
                    Application.OpenForms.Item(f).Close()
                Next
            Catch ex As Exception
            End Try
            TextToSpeak("Hasta luego, ")
        Else
            e.Cancel = True
        End If
    End Sub

    Private Sub FilterContactsByBookings(dtBookings As DataTable)
        If GridView1.RowCount = 0 Then
            bbiLoadContacts.PerformClick()
        End If
        Dim sFilter As String = ""
        GridView1.ActiveFilterString = ""
        For r = 0 To dtBookings.Rows.Count - 1
            Dim oRow As DataRow = dtBookings.Rows(r)
            sFilter += IIf(r > 0, " OR ", "") & "[MatchCode] = '" & oRow("F2") & "'"
        Next
        GridView1.ActiveFilterCriteria = DevExpress.Data.Filtering.CriteriaOperator.Or
        GridView1.ActiveFilterString = "(" & sFilter & ")"
    End Sub

    Private Sub GridView2_RowStyle(sender As Object, e As RowStyleEventArgs) Handles GridView2.RowStyle
        Dim View As GridView = sender
        If (e.RowHandle >= 0) Then
            If Not View.Columns(1) Is Nothing Then
                Dim C1 As String = View.GetRowCellDisplayText(e.RowHandle, View.Columns(1))
                If dtContacts.Select("MatchCode='" & C1 & "'").Length = 0 Then
                    e.Appearance.BackColor = Color.Salmon
                    e.Appearance.BackColor2 = Color.SeaShell
                End If
            End If
        End If
    End Sub

    Private Sub bbiExportToExcel_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiExportToExcel.ItemClick
        Dim sPath As String = Path.GetTempPath
        Dim sFileName = (FileIO.FileSystem.GetTempFileName).Replace(".tmp", ".xlsx")
        GridView1.OptionsPrint.ExpandAllDetails = True
        GridView1.OptionsPrint.AutoWidth = False
        GridView1.OptionsPrint.UsePrintStyles = True
        GridView1.BestFitMaxRowCount = GridView1.RowCount
        Dim Options As New XlsxExportOptionsEx
        Options.ExportType = ExportType.WYSIWYG
        GridView1.ExportToXlsx(sFileName, Options)
        System.Diagnostics.Process.Start(sFileName)
        'If IO.File.Exists(sFileName) Then
        '    Dim oXls As New Excel.Application 'Crea el objeto excel 
        '    oXls.Workbooks.Open(sFileName, , False) 'El true es para abrir el archivo en modo Solo lectura (False si lo quieres de otro modo)
        '    oXls.Visible = True
        '    oXls.WindowState = Excel.XlWindowState.xlMaximized 'Para que la ventana aparezca maximizada.
        'End If
    End Sub

    Private Sub GridView1_RowCellStyle(sender As Object, e As RowCellStyleEventArgs) Handles GridView1.RowCellStyle
        Dim View As GridView = sender
        If (e.RowHandle >= 0) Then
            If e.Column.FieldName <> "ENTC_Email" Then
                Return
            End If
            If Not View.Columns("ENTC_Email") Is Nothing Then
                Dim C1 As String = View.GetRowCellDisplayText(e.RowHandle, View.Columns("ENTC_Email"))
                For Each mail As String In C1.ToString.Split(";")
                    If IsValidEmail(LCase(mail.Trim)) = False Then
                        e.Appearance.BackColor = Color.Salmon
                        e.Appearance.BackColor2 = Color.SeaShell
                    End If
                Next

            End If
        End If
    End Sub

    Public Function IsValidEmail(ByVal email As String) As Boolean
        Return Regex.IsMatch(email, "^([0-9a-zA-Z]([-\.\w]*[0-9a-zA-Z])*@([0-9a-zA-Z][-\w]*[0-9a-zA-Z]\.)+[a-zA-Z]{2,9})$") '"^[a-zA-Z][\w\._-]*[a-zA-Z0-9]@[a-zA-Z0-9][\w\.-]*[a-zA-Z0-9]\.[a-zA-Z][a-zA-Z\.]*[a-zA-Z]{2,4}$")
    End Function

End Class