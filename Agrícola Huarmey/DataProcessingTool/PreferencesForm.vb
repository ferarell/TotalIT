Imports DevExpress.XtraGrid
Imports DevExpress.XtraGrid.Views.Grid
Imports DevExpress.XtraGrid.Views.Grid.ViewInfo
Imports System.Configuration

Public Class PreferencesForm
    Dim dtCompany As New DataTable
    Dim drCompany As DataRow = Nothing
    Dim teWebServiceURL As New TextBox

    Private Sub beArchivoOrigen_Properties_ButtonClick(sender As Object, e As DevExpress.XtraEditors.Controls.ButtonPressedEventArgs)
        If FolderBrowserDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            sender.EditValue = FolderBrowserDialog1.SelectedPath
        End If
    End Sub

    Private Sub bbiCerrar_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiCerrar.ItemClick
        Close()
    End Sub

    Private Sub bbiGuardar_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiGuardar.ItemClick
        Try
            Validate()
            My.Settings.DBHostServer = teDBHostServer.Text
            My.Settings.DBPortServer = teDBPortServer.Text
            My.Settings.DBName = teDatabase.Text
            My.Settings.DBSslConnection = ceSSLDBConnection.Checked
            My.Settings.DBUser = teDBUser.Text
            My.Settings.DBPassword = teDBPassword.Text
            My.Settings.PostgreSQLConStr = "Server=" & teDBHostServer.Text & ";Port=" & teDBPortServer.Text & ";Database=" & teDatabase.Text & ";User Id=" & teDBUser.Text & ";Password=" & teDBPassword.Text & ";"
            SaveCompanies()
            My.Settings.Save()
        Catch ex As Exception
            DevExpress.XtraEditors.XtraMessageBox.Show(Me.LookAndFeel, ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            DevExpress.XtraEditors.XtraMessageBox.Show(Me.LookAndFeel, "Los cambios fueron aplicados satisfactoriamente.", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Try

    End Sub

    Private Sub PreferencesForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadCompanies()
        LoadSettings()
    End Sub

    Private Sub LoadCompanies()
        dtCompany = GetCompaniesList()
        gcCompanyList.DataSource = dtCompany
    End Sub

    Private Sub LoadSettings()
        teDBHostServer.EditValue = My.Settings.DBHostServer
        teDBPortServer.EditValue = My.Settings.DBPortServer
        teDatabase.EditValue = My.Settings.DBName
        ceSSLDBConnection.Checked = My.Settings.DBSslConnection
        teDBUser.EditValue = My.Settings.DBUser
        teDBPassword.EditValue = My.Settings.DBPassword
    End Sub

    Private Sub bbiReset_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiReset.ItemClick
        If DevExpress.XtraEditors.XtraMessageBox.Show(Me.LookAndFeel, "Se procederá a reiniciar la configuración de la aplicación, desea continuar?", "Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
            My.Settings.Reset()
            Me.PreferencesForm_Load(sender, e)
        End If
    End Sub

    Private Sub gcExternalData_EmbeddedNavigator_ButtonClick(sender As Object, e As DevExpress.XtraEditors.NavigatorButtonClickEventArgs) Handles gcCompanyList.EmbeddedNavigator.ButtonClick
        If e.Button.Tag = 0 Then
            RepositoryItemTextEdit1.AllowFocused = True
            GridView1.AddNewRow()
        ElseIf e.Button.Tag = 1 Then
            GridView1.DeleteSelectedRows()
        End If
    End Sub

    Private Sub RepositoryItemCheckEdit1_Click(sender As Object, e As EventArgs) Handles RepositoryItemCheckEdit1.Click
        If GridView1.GetFocusedRowCellValue("Active") Then
            teWebServiceURL.Text = GridView1.GetFocusedRowCellValue("URL")
        End If
    End Sub

    Private Sub SaveCompanies()
        If GridView1.RowCount = 0 Then
            Return
        End If
        For r = 0 To dtCompany.Rows.Count - 1
            If GridView1.GetRowCellValue(r, "company_ruc").ToString.Trim <> "" Then
                If dtCompany.Select("company_ruc='" & My.Settings.CompanyTable.Rows(r)("company_ruc") & "'").Length = 0 Then
                    My.Settings.CompanyTable.Rows.Add(dtCompany.Rows(r).ItemArray)
                End If
            Else
                GridView1.DeleteRow(r)
            End If
        Next
        gcCompanyList.RefreshDataSource()
    End Sub

    Private Sub RepositoryItemCheckEdit1_CheckStateChanged(sender As Object, e As EventArgs) Handles RepositoryItemCheckEdit1.CheckStateChanged
        GridView1.CloseEditor()
    End Sub

    Private Sub RepositoryItemTextEdit1_Leave(sender As Object, e As EventArgs) Handles RepositoryItemTextEdit1.Leave
        Dim dtQuery As New DataTable
        Dim protocol As String = "http://"
        If GridView1.GetFocusedRowCellValue("service_url") = "" Then
            Return
        End If
        If Not GridView1.GetFocusedRowCellValue("service_url").ToString.Contains(protocol) Then
            GridView1.SetFocusedRowCellValue("service_url", protocol & GridView1.GetFocusedRowCellValue("service_url"))
        End If
        Try
            dtQuery = WebServiceValidate(GridView1.GetFocusedRowCellValue("service_url"), "select t0.id, t1.name, t1.ref, t1.street from res_company t0 join res_partner t1 on t1.id = t0.partner_id")
            If dtQuery.Rows.Count > 0 Then
                For r = 0 To dtQuery.Rows.Count - 1
                    If r > 0 Then
                        GridView1.AddNewRow()
                        GridView1.SetFocusedRowCellValue("service_url", GridView1.GetRowCellValue(GridView1.RowCount - 2, "service_url"))
                    End If
                    GridView1.SetFocusedRowCellValue("company_id", dtQuery.Rows(r)(0))
                    GridView1.SetFocusedRowCellValue("company_name", dtQuery.Rows(r)(1))
                    GridView1.SetFocusedRowCellValue("company_ruc", dtQuery.Rows(r)(2))
                    GridView1.SetFocusedRowCellValue("company_address", dtQuery.Rows(r)(3))
                Next
            Else
                GridView1.FocusedColumn.ColumnHandle = 0
            End If
        Catch ex As Exception
            DevExpress.XtraEditors.XtraMessageBox.Show(Me.LookAndFeel, ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub GridView1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GridView1.Click
        If GridView1.GetFocusedRowCellValue("Company") = "" Then
            Return
        End If
        Dim hit As Integer = GridView1.FocusedRowHandle
        For r = 0 To GridView1.RowCount - 1
            GridView1.SetRowCellValue(r, "Active", False)
            If hit = r Then
                GridView1.SetRowCellValue(r, "Active", True)
            End If
        Next
    End Sub

End Class