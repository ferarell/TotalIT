Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraRichEdit.Export
Imports DevExpress.XtraRichEdit.Services
Imports System.Windows.Forms
Imports System.Data
Imports System.IO
Imports DevExpress.XtraEditors

Public Class LayoutForm
    Dim oDataAccess As New DataAccess
    Dim dtConfiguration, dtMessageConfig, dtChild As New System.Data.DataTable
    Dim FileName As String = ""
    Dim oRichTextEdit As New DevExpress.XtraRichEdit.RichEditControl

    Private Sub SettingsForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        rgResponse.SelectedIndex = 0
        rgFieldName.SelectedIndex = 0
        SplitContainerControl2.Panel2.Visible = False
        If Not IO.File.Exists(My.Settings.DBFileName) Then
            DevExpress.XtraEditors.XtraMessageBox.Show(Me.LookAndFeel, "The database was not found, please check the setting option", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If
        LoadConfiguration()
        GetMessageConfig()
    End Sub

    Private Sub FillRichTextControl(drConfig As DataRow)
        If GridView1.FocusedRowHandle < 0 Then
            gcMainData = Nothing
            Return
        End If
        recText.RtfText = Nothing
        If Not IsDBNull(drConfig(rgFieldName.EditValue)) Then
            recText.HtmlText = drConfig(rgFieldName.EditValue)
        End If
        SplitContainerControl2.PanelVisibility = DevExpress.XtraEditors.SplitPanelVisibility.Panel1
    End Sub

    Private Sub LoadConfiguration()
        dtConfiguration.Rows.Clear()
        dtConfiguration = oDataAccess.ExecuteAccessQuery("SELECT * FROM " & My.Settings.ConfigTableName).Tables(0)
        If dtConfiguration.Rows.Count > 0 Then
            gcMainData.DataSource = dtConfiguration
            FillRichTextControl(dtConfiguration.Rows(GridView1.FocusedRowHandle))
        End If
    End Sub

    Private Sub GetMessageConfig()
        If dtConfiguration.Rows.Count = 0 Then
            Return
        End If
        dtMessageConfig = dtConfiguration.Select("ResponseType=" & rgResponse.SelectedIndex.ToString).CopyToDataTable
        If dtMessageConfig.Rows.Count > 0 Then
            gcMainData.DataSource = dtMessageConfig
            FillRichTextControl(dtMessageConfig.Rows(GridView1.FocusedRowHandle))
        End If
    End Sub

    Private Sub bbiSave_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiSave.ItemClick
        Validate()
        gcMainData.EmbeddedNavigator.Buttons.CustomButtons.Item(0).Enabled = True
        gcMainData.EmbeddedNavigator.Buttons.CustomButtons.Item(0).Enabled = True
        Try
            If Not UpdateConfiguration() Then
                XtraMessageBox.Show(Me.LookAndFeel, "An error occurred while saving changes. ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If
            XtraMessageBox.Show(Me.LookAndFeel, "Changes were saved successfully", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            XtraMessageBox.Show(Me.LookAndFeel, "An error occurred while saving changes. " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        bbiRefresh.PerformClick()
    End Sub

    Friend Function UpdateConfiguration() As Boolean
        Dim bResult As Boolean = True
        Dim oCondition, oValues As String
        Try
            Dim oRow As DataRow = GridView1.GetFocusedDataRow

            oCondition = "Position=" & oRow("Position").ToString & " AND ResponseType=" & oRow("ResponseType").ToString

            If Not IsDBNull(oRow("Identifier")) Then
                oValues = "[Identifier]='" & oRow("Identifier").ToString & "'"
            End If
            If Not IsDBNull(oRow("Description")) Then
                oValues += ", " & "[Description]='" & oRow("Description").ToString & "'"
            End If
            If Not IsDBNull(oRow("Header")) Then
                oValues += ", " & "Header='" & Replace(oRow("Header"), "'", "") & "'"
            End If
            If Not IsDBNull(oRow("Body")) Then
                oValues += ", " & "[Body]='" & Replace(oRow("Body"), "'", "") & "'"
            End If
            If Not IsDBNull(oRow("Signature")) Then
                oValues += ", " & "Signature='" & Replace(oRow("Signature"), "'", "") & "'"
            End If
            If Not IsDBNull(oRow("News")) Then
                oValues += ", " & "News='" & Replace(oRow("News"), "'", "") & "'"
            End If
            If IsDBNull(oRow("NewsValidityFrom")) Then
                oValues += ", " & "[NewsValidityFrom]=NULL"
            Else
                oValues += ", " & "[NewsValidityFrom]='" & Format(oRow("NewsValidityFrom"), "yyyy-MM-dd") & "'"
            End If
            If IsDBNull(oRow("NewsValidityTo")) Then
                oValues += ", " & "[NewsValidityTo]=NULL"
            Else
                oValues += ", " & "[NewsValidityTo]='" & Format(oRow("NewsValidityTo"), "yyyy-MM-dd") & "'"
            End If

            If oRow.RowState = DataRowState.Added Then
                bResult = oDataAccess.InsertIntoAccess(My.Settings.ConfigTableName, oRow)
            Else
                bResult = oDataAccess.UpdateAccess(My.Settings.ConfigTableName, oCondition, oValues)
            End If
        Catch ex As Exception
            bResult = False
        End Try
        Return bResult
    End Function

    Private Sub bbiRefresh_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiRefresh.ItemClick
        LoadConfiguration()
        GetMessageConfig()
    End Sub

    Private Sub bbiClose_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiClose.ItemClick
        Close()
    End Sub

    Private Sub rgResponse_SelectedIndexChanged(sender As Object, e As EventArgs) Handles rgResponse.SelectedIndexChanged
        rgFieldName.SelectedIndex = 0
        rgFieldName.Enabled = True
        recText.Enabled = True
        rgFieldName.Properties.Items(4).Enabled = False

        If rgResponse.EditValue = "I" Then
            'rgFieldName.Enabled = False
            'recText.Enabled = False
        End If
        If rgResponse.EditValue = "C" Then
            rgFieldName.Properties.Items(4).Enabled = True
            rgFieldName.EditValue = "Query"
        End If
        GetMessageConfig()
    End Sub

    Private Sub rgFieldName_Properties_EditValueChanging(sender As Object, e As Controls.ChangingEventArgs) Handles rgFieldName.Properties.EditValueChanging
        GridView1.SetFocusedRowCellValue(rgFieldName.EditValue, recText.HtmlText)

    End Sub

    Private Sub bbiMessagePreview_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiMessagePreview.ItemClick
        Dim oMailItem As New CreateMailItem
        oMailItem.subject = GridView1.GetFocusedRowCellValue("Identifier")
        oMailItem.htmlBody.AppendText(GridView1.GetFocusedRowCellValue("Header") & "<br>")
        oMailItem.htmlBody.AppendText(GridView1.GetFocusedRowCellValue("Body") & "<br>")
        oMailItem.htmlBody.AppendText(GridView1.GetFocusedRowCellValue("Signature") & "<br>")
        oMailItem.htmlBody.AppendText(GridView1.GetFocusedRowCellValue("News") & "<br>")
        oMailItem.CreateCustomMessage("Display")
    End Sub

    Private Sub GridView1_CellValueChanged(sender As Object, e As DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs) Handles GridView1.CellValueChanged
        GridView1.UpdateCurrentRow()
        If GridView1.FocusedRowHandle >= 0 Then
            Dim oRow As DataRow = GridView1.GetDataRow(GridView1.FocusedRowHandle)
            oRow("Position") = IIf(oRow("Position").ToString = "", dtConfiguration.Compute("MAX(Position)", "") + 1, oRow("Position"))
            oRow("ResponseType") = rgResponse.SelectedIndex
        End If
    End Sub

    Private Sub gcMainData_EmbeddedNavigator_ButtonClick(sender As Object, e As NavigatorButtonClickEventArgs) Handles gcMainData.EmbeddedNavigator.ButtonClick
        If e.Button.Hint = "New Row" Then
            GridView1.AddNewRow()
            e.Button.Enabled = False
        End If
        If e.Button.Hint = "Remove Row" Then
            If XtraMessageBox.Show("Are you sure you want to delete this row?", "Confirmation", MessageBoxButtons.YesNo) = DialogResult.No Then
                Return
            End If
            If GridView1.FocusedRowHandle >= 0 Then
                Dim oRow As DataRow = GridView1.GetDataRow(GridView1.FocusedRowHandle)
                If Not oDataAccess.ExecuteAccessNonQuery("DELETE FROM " & My.Settings.ConfigTableName & " WHERE Position=" & oRow("Position").ToString & " AND ResponseType=" & oRow("ResponseType").ToString) Then
                    XtraMessageBox.Show(Me.LookAndFeel, "An error occurred while saving changes. ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Else
                    XtraMessageBox.Show(Me.LookAndFeel, "Changes were saved successfully", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End If
            e.Button.Enabled = True
            bbiRefresh.PerformClick()
        End If

    End Sub

    Private Sub recText_TextChanged(sender As Object, e As EventArgs) Handles recText.TextChanged
        GridView1.SetFocusedRowCellValue(rgFieldName.EditValue, recText.HtmlText)
    End Sub

    Private Sub rgOleFieldName_SelectedIndexChanged(sender As Object, e As EventArgs) Handles rgFieldName.SelectedIndexChanged
        GridView1.Columns("NewsValidityFrom").Visible = False
        GridView1.Columns("NewsValidityTo").Visible = False
        If dtMessageConfig.Rows.Count > 0 Then
            FillRichTextControl(GridView1.GetFocusedDataRow)
        End If
        If rgFieldName.SelectedIndex = 3 Then
            GridView1.Columns("NewsValidityFrom").Visible = True
            GridView1.Columns("NewsValidityTo").Visible = True
        End If
    End Sub

    Private Sub GridView1_FocusedRowChanged(sender As Object, e As DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs) Handles GridView1.FocusedRowChanged
        If e.FocusedRowHandle < 0 Or dtMessageConfig.Rows.Count = 0 Then
            Return
        End If
        If dtConfiguration.Rows.Count > 0 Then
            FillRichTextControl(dtMessageConfig.Rows(GridView1.FocusedRowHandle))
        End If
    End Sub

End Class