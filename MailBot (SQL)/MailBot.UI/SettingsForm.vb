Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraRichEdit.Export
Imports DevExpress.XtraRichEdit.Services
Imports System.Windows.Forms
Imports System.Data
Imports System.IO
Imports System.Globalization
Imports MailBot.InterfazUsuario.MailBotServiceWCF
Imports DevExpress.XtraBars

Public Class SettingsForm
    Dim dtInputType, dtStructure, dtInputList, dtIdentifier, dtConfiguration As New DataTable
    Dim oProxyService As MailBot.InterfazUsuario.MailBotServiceWCF.MailBotServicesWCFClient = New MailBot.InterfazUsuario.MailBotServiceWCF.MailBotServicesWCFClient()

    Private Sub SettingsForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        grInputType.Width = 250
        grInputType.Width = 350
        grIdentifiers.Width = 250

        dtInputType = ExecuteSQL("EXEC bot.upGetInputTypeById 0")
        If dtInputType.Rows.Count > 0 Then
            gcInputType.DataSource = dtInputType
        End If
        dtStructure = ExecuteSQL("EXEC bot.upGetMessageStructure")
        If dtStructure.Rows.Count > 0 Then
            gcStructure.DataSource = dtStructure
        End If
        GetConfiguration("All")
    End Sub

    Private Sub GridView3_FocusedRowChanged(sender As Object, e As DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs) Handles GridView3.FocusedRowChanged, GridView1.FocusedRowChanged, GridView2.FocusedRowChanged
        GetConfiguration(sender.Name)
    End Sub

    Private Sub bbiSave_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiSave.ItemClick
        For r = 1 To dtStructure.Rows.Count
            If r = GridView2.GetFocusedRowCellValue("IdMessageStructure") Then
                'If recText.Text.Trim <> "" Then
                If dtConfiguration.Rows.Count = 0 Then
                    UpdateConfiguration(0)
                Else
                    UpdateConfiguration(dtConfiguration.Rows(0)("IdMessageLayout"))
                End If

                'End If
            End If
        Next
        GetConfiguration("All")
    End Sub

    Friend Function UpdateConfiguration(iLayout As Integer) As Boolean
        Dim bResult As Boolean = True
        Dim sQuery As String = ""
        Try
            If iLayout = 0 Then
                'Insert
                'sQuery = "EXEC bot.upMessageLayoutInsert 0, " & GridView2.GetFocusedRowCellValue("IdMessageStructure").ToString & ", " & GridView1.GetFocusedRowCellValue("IdInputConfiguration").ToString & ", '" & recText.HtmlText & "', NULL, NULL, '" & My.User.Name & "', '" & Now.ToString & "'"
                Dim oMessageLayout As MessageLayout = New MessageLayout()
                oMessageLayout.IdMessageLayout = Convert.ToInt32(iLayout.ToString())
                oMessageLayout.IdMessageStructure = Convert.ToInt32(GridView2.GetFocusedRowCellValue("IdMessageStructure").ToString())
                oMessageLayout.IdInputConfiguration = Convert.ToInt32(GridView1.GetFocusedRowCellValue("IdInputConfiguration").ToString())
                oMessageLayout.MessageText = recText.HtmlText
                oMessageLayout.ValidFrom = Nothing
                oMessageLayout.ValidTo = Nothing
                oMessageLayout.UserCreate = My.User.Name
                oMessageLayout.DateCreate = Now.ToString("dd/MM/yyyy")
                oProxyService.InsertMessageLayout(oMessageLayout)
            Else
                'Update
                'Dim usDtfi As DateTimeFormatInfo = New CultureInfo("es-ES", False).DateTimeFormat

                'var result = Convert.ToDateTime("12/01/2011", usDtfi)
                Dim oMessageLayout As MessageLayout = New MessageLayout()
                oMessageLayout.IdMessageLayout = Convert.ToInt32(iLayout.ToString())
                oMessageLayout.IdMessageStructure = Convert.ToInt32(GridView2.GetFocusedRowCellValue("IdMessageStructure").ToString())
                oMessageLayout.IdInputConfiguration = Convert.ToInt32(GridView1.GetFocusedRowCellValue("IdInputConfiguration").ToString())
                oMessageLayout.MessageText = recText.HtmlText
                oMessageLayout.ValidFrom = Nothing
                oMessageLayout.ValidTo = Nothing
                oMessageLayout.UserUpdate = My.User.Name
                oMessageLayout.DateUpdate = Now.ToString("dd/MM/yyyy")
                oProxyService.UpdateMessageLayout(oMessageLayout)
                'sQuery = "EXEC bot.upMessageLayoutUpdate " & iLayout.ToString & ", " & GridView2.GetFocusedRowCellValue("IdMessageStructure").ToString & ", " & GridView1.GetFocusedRowCellValue("IdInputConfiguration").ToString & ", '" & recText.HtmlText & "', NULL, NULL, '" & My.User.Name & "', '" & Now.ToString & "'"
            End If
            'ExecuteSQLNonQuery(sQuery)
        Catch ex As Exception
            DevExpress.XtraEditors.XtraMessageBox.Show(Me.LookAndFeel, ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            bResult = False
        End Try
        Return bResult
    End Function

    Private Sub beDatabase_Properties_ButtonClick(sender As Object, e As DevExpress.XtraEditors.Controls.ButtonPressedEventArgs) Handles beDatabase.Properties.ButtonClick
        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
            beDatabase.Text = OpenFileDialog1.FileName
        End If
    End Sub

    Private Sub GetConfiguration(GridName As String)
        recText.HtmlText = Nothing
        If GridName = "GridView3" Then
            gcInputList.DataSource = Nothing
            If Not GridView3.GetFocusedRowCellValue("IdInputType") Then
                dtInputList.Rows.Clear()
                dtInputList = ExecuteSQL("EXEC bot.upGetInputConfigurationByType " & GridView3.GetFocusedRowCellValue("IdInputType"))
                If dtInputList.Rows.Count > 0 Then
                    gcInputList.DataSource = dtInputList
                End If
            End If
        End If
        If GridName = "GridView1" Then
            gcIdentifiers.DataSource = Nothing
            If Not GridView1.GetFocusedRowCellValue("IdInputConfiguration") Is Nothing Then
                dtIdentifier.Rows.Clear()
                dtIdentifier = ExecuteSQL("EXEC bot.upGetIdentifiersById " & GridView1.GetFocusedRowCellValue("IdInputConfiguration"))
                If dtIdentifier.Rows.Count > 0 Then
                    gcIdentifiers.DataSource = dtIdentifier
                End If
            End If
        End If
        If dtInputList.Rows.Count = 0 Or GridView1.GetFocusedRowCellValue("IdInputConfiguration") Is Nothing Then
            Return
        End If
        dtConfiguration = ExecuteSQL("EXEC bot.upGetConfigurationById " & GridView1.GetFocusedRowCellValue("IdInputConfiguration") & ", " & GridView2.GetFocusedRowCellValue("IdMessageStructure"))
        recText.RtfText = Nothing
        If dtConfiguration.Rows.Count = 0 Then
            Return
        End If
        If Not IsDBNull(dtConfiguration.Rows(0)("MessageText")) Then
            recText.HtmlText = UnicodeBytesToString(dtConfiguration.Rows(0)("MessageText"))
        End If
    End Sub

    Private Sub bbiClose_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiClose.ItemClick
        Close()
    End Sub

    Private Sub bbiRefresh_ItemClick(sender As Object, e As ItemClickEventArgs) Handles bbiRefresh.ItemClick
        GetConfiguration("All")
    End Sub

    'Private Sub SettingsForm_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
    '    If Not IO.File.Exists(beDatabase.Text) Then
    '        DialogResult = Windows.Forms.DialogResult.No
    '    End If
    'End Sub
End Class