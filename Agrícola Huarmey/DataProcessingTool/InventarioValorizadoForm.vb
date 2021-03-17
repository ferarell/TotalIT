Imports DevExpress.XtraEditors.DXErrorProvider
Imports DevExpress.XtraGrid.Views.Grid
Imports System.Drawing
Imports System.IO
Imports System.Globalization
Imports DevExpress.XtraSplashScreen
Imports System.Threading

Public Class InventarioValorizadoForm
    Dim dtQuery, dtPrint As New DataTable
    Dim dsMain As New dsMain

    Private Sub InventarioValorizadoForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        bbiExport.Enabled = False
        bbiPrint.Enabled = False
        LoadCompany()
        LoadInputValidations()
        dtPrint = dsMain.Tables("InventarioValorizado")
    End Sub

    Private Sub LoadCompany()
        dtCompany = GetCompaniesList()
        lueCompany.Properties.DataSource = dtCompany
        lueCompany.Properties.DisplayMember = "company_name"
        lueCompany.Properties.ValueMember = "company_id"
    End Sub

    Private Sub bbiClose_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiClose.ItemClick
        Close()
    End Sub

    Private Sub bbiPrint_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiPrint.ItemClick
        Dim pForm As New PrintForm
        pForm.dtPrint = dtPrint
        pForm.aParams = GetParamValues()
        pForm.RptFile = "InventarioValorizado.rpt"
        pForm.ShowDialog()
    End Sub

    Friend Function GetParamValues() As ArrayList
        Dim aParams As New ArrayList
        aParams.Add("Formato 13.1 - REGISTRO DE INVENTARIO PERMANENTE VALORIZADO - DETALLE DEL INVENTARIO VALORIZADO")
        aParams.Add(deDateFrom.Text)
        aParams.Add(deDateTo.Text)
        aParams.Add(lueCompany.GetColumnValue("company_id"))
        aParams.Add(lueCompany.GetColumnValue("company_name"))
        aParams.Add(lueCompany.GetColumnValue("street"))
        Return aParams
    End Function

    Private Sub bbiProcess_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiProcess.ItemClick
        If Not vpInputs.Validate Then
            Return
        End If
        Try
            SplashScreenManager.ShowForm(Me, GetType(WaitForm), True, True, False)
            DataProcess()
        Catch ex As Exception
            SplashScreenManager.CloseForm(False)
            DevExpress.XtraEditors.XtraMessageBox.Show(Me.LookAndFeel, ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        SplashScreenManager.CloseForm(False)
    End Sub

    Private Sub DataProcess()
        Dim sQuery As String = ""
        'dtQuery.Rows.Clear()
        'dtQuery = oGeneralDao.ExecuteSQL("select id from public.account_period where date_start = '" & FecIni & "' and date_stop = '" & FecFin & "'")
        'If dtQuery.Rows.Count = 0 Then
        '    SplashScreenManager.CloseForm(False)
        '    DevExpress.XtraEditors.XtraMessageBox.Show(Me.LookAndFeel, "No existe el periodo fiscal asignado.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        '    Return
        'End If
        'Dim sPeriodId As String = dtQuery.Rows(0)("id")
        sQuery += "SELECT pp.id as product_id, pp.default_code as default_code, pt.name as product_name, sv.picking_id as picking_id, "
        sQuery += "sv.number as number, sv.date as date, sv.location_id as location_id, sv.location_dest_id as location_dest_id, "
        sQuery += "sv.type as type, sum(sv.entry_qty) as entry_qty, avg(sv.entry_price) as entry_price, "
        sQuery += "sum(sv.exit_qty) as exit_qty, avg(sv.exit_price) as exit_price, "
        sQuery += "min(pib.saldo_qty) as inibal_qty, min(pib.saldo_price) as inibal_price, min(pib.saldo_total) as inibal_total "
        sQuery += "FROM stock_valued sv "
        sQuery += "LEFT JOIN product_product pp ON pp.id = sv.product_id "
        sQuery += "LEFT JOIN product_initial_balance pib ON pib.product_id = sv.product_id "
        sQuery += "LEFT JOIN product_template pt ON pt.id = pp.product_tmpl_id "
        sQuery += "WHERE sv.company_id = " & lueCompany.GetColumnValue("id").ToString & " AND sv.date>= '" & deDateFrom.Text & "' AND sv.date <= '" & deDateTo.Text & "' AND sv.location_id != 5 "
        sQuery += "GROUP BY pp.id, pp.default_code, pt.name, sv.picking_id, sv.number, sv.date, sv.location_id, sv.location_dest_id, sv.type "
        sQuery += "ORDER BY pp.default_code, sv.date, sv.number; "
        dtQuery.Rows.Clear()
        dtQuery = ExecuteSQL(lueCompany.GetColumnValue("service_url"), sQuery)
        If dtQuery.Rows.Count = 0 Then
            SplashScreenManager.CloseForm(False)
            DevExpress.XtraEditors.XtraMessageBox.Show(Me.LookAndFeel, "La consulta no retornó datos.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        If dtQuery.Rows.Count > 0 Then
            bbiExport.Enabled = True
            bbiPrint.Enabled = True
        End If
        Dim sCurrency As String = ""
        dtPrint.Rows.Clear()
        If dtQuery.Columns.Count = 0 Then
            DevExpress.XtraEditors.XtraMessageBox.Show(Me.LookAndFeel, "La consulta no retornó información.", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If
        'For Each row As DataRow In dtQuery.Rows
        '    dtPrint.Rows.Add(row("date"), row("move"), row("glosa"), row("t8_code"), row("move"), row("account_code"), row("account_name"), sCurrency, row("debit"), row("credit"), IIf(row("debit") > 0, row("amount_currency"), 0), IIf(row("credit") > 0, row("amount_currency"), 0), row("account_code"))
        'Next
        gcExternalData.DataSource = dtQuery
        'LayoutApply()
    End Sub

    Private Sub LayoutApply()
        GridView1.PopulateColumns()
        GridView1.Columns(3).Visible = False
        GridView1.Columns(12).Visible = False
        For i = 8 To 9
            GridView1.Columns(i).SummaryItem.SetSummary(DevExpress.Data.SummaryItemType.Sum, "{0:n2}")
            GridView1.Columns(i).DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
            GridView1.Columns(i).DisplayFormat.FormatString = "n2"
        Next

    End Sub
    Private Sub bbiExport_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiExport.ItemClick
        ExportarExcel(gcExternalData)
    End Sub

    Private Sub LoadInputValidations()
        Validate()
        Dim containsValidationRule As New DevExpress.XtraEditors.DXErrorProvider.ConditionValidationRule()
        containsValidationRule.ConditionOperator = ConditionOperator.IsNotBlank
        containsValidationRule.ErrorText = "Asigne un valor."
        containsValidationRule.ErrorType = ErrorType.Critical
        Dim customValidationRule As New CustomValidationRule()
        customValidationRule.ErrorText = "Valor obligatorio."
        customValidationRule.ErrorType = ErrorType.Critical
        vpInputs.SetValidationRule(Me.lueCompany, customValidationRule)
    End Sub
End Class