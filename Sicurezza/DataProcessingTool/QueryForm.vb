Imports DevExpress.XtraEditors.DXErrorProvider
Imports DevExpress.XtraGrid.Views.Grid
Imports System.Drawing
Imports System.IO
Imports System.Globalization
Imports DevExpress.XtraSplashScreen
Imports System.Threading

Public Class QueryForm
    Dim dtQuery As New DataTable
    Dim dsMain As New dsMain

    Private Sub bbiExecute_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiExecute.ItemClick
        Try
            SplashScreenManager.ShowForm(Me, GetType(WaitForm), True, True, False)
            dtQuery.Rows.Clear()
            dtQuery = ExecuteSQL(lueCompany.GetColumnValue("service_url"), meQuery.Text)
            gcExternalData.DataSource = dtQuery
            GridView1.PopulateColumns()
            If dtQuery.Rows.Count > 0 Then
                bbiExport.Enabled = True
                bbiPrint.Enabled = True
            End If
        Catch ex As Exception
            SplashScreenManager.CloseForm(False)
            DevExpress.XtraEditors.XtraMessageBox.Show(Me.LookAndFeel, ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        SplashScreenManager.CloseForm(False)
    End Sub

    Private Sub BarButtonItem2_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles BarButtonItem2.ItemClick
        Dim myForm As New PreferencesForm
        myForm.ShowDialog()
    End Sub

    Private Sub bbiClose_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiClose.ItemClick
        Close()
    End Sub

    Private Sub QueryForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        bbiExport.Enabled = False
        bbiPrint.Enabled = False
        LoadCompany()
    End Sub

    Private Sub LoadCompany()
        dtCompany = GetCompaniesList()
        lueCompany.Properties.DataSource = dtCompany
        lueCompany.Properties.DisplayMember = "company_name"
        lueCompany.Properties.ValueMember = "company_ruc"
    End Sub

    Private Sub bbiExport_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiExport.ItemClick
        ExportarExcel(GridView1, Nothing)
    End Sub

    Private Sub bbiPrint_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiPrint.ItemClick
        Dim dtPrint As New DataTable
        dtPrint = dsMain.Tables("Test")
        If dtQuery.Columns.Count < 5 Then
            DevExpress.XtraEditors.XtraMessageBox.Show(Me.LookAndFeel, "The query should return at least 5 columns..", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If
        For Each row As DataRow In dtQuery.Rows
            dtPrint.Rows.Add(row(1), row(2), row(3), row(4), row(5))
        Next
        Dim pForm As New PrintForm
        pForm.dtPrint = dtPrint
        pForm.ShowDialog()

    End Sub

End Class