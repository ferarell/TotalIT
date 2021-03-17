Imports DevExpress.XtraEditors.DXErrorProvider
Imports DevExpress.XtraGrid.Views.Grid
Imports System.Drawing
Imports System.IO
Imports System.Globalization
Imports DevExpress.XtraSplashScreen
Imports System.Threading

Public Class MovimientoContableForm
    Dim dsMain As New dsMain
    Dim dtQuery, dtPrint As New DataTable
    Dim bError As Boolean = False

    Private Sub MovimientoContableForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        GridView1.RestoreLayoutFromRegistry(regKey)
        GridView1.ActiveFilterEnabled = False
        GridView1.OptionsFilter.Reset()
        EnableButtons(False)
        dtPrint = dsMain.Tables("MovimientoContable")
        Try
            LoadCompany()
            If dtCompany.Rows.Count = 0 Then
                bError = True
                DevExpress.XtraEditors.XtraMessageBox.Show(Me.LookAndFeel, "No existe ningún servicio web disponible, por favor asignarlo en preferencias de la aplicación.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If
        Catch ex As Exception
            bError = True
            DevExpress.XtraEditors.XtraMessageBox.Show(Me.LookAndFeel, ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub LoadCompany()
        dtCompany = GetCompaniesList()
        lueCompany.Properties.DataSource = dtCompany
        lueCompany.Properties.DisplayMember = "company_name"
        lueCompany.Properties.ValueMember = "company_ruc"
    End Sub

    Private Sub LoadAccount()
        dtAccount = GetAccountsList(lueCompany.GetColumnValue("service_url"))
        lueAccount.Properties.DataSource = dtAccount
        lueAccount.Properties.DisplayMember = "code"
        lueAccount.Properties.ValueMember = "code"
    End Sub

    Private Sub LoadAPartner()
        dtPartner = GetPartnersList(lueCompany.GetColumnValue("service_url"))
        luePartner.Properties.DataSource = dtPartner
        luePartner.Properties.DisplayMember = "name"
        luePartner.Properties.ValueMember = "id"
    End Sub

    Private Sub bbiClose_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiClose.ItemClick
        Close()
    End Sub

    Private Sub bbiPrint_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiPrint.ItemClick
        Dim pForm As New PrintForm
        pForm.dtPrint = dtPrint
        pForm.aParams = GetParamValues()
        pForm.RptFile = "MovimientoContable.rpt"
        pForm.ShowDialog()
    End Sub

    Friend Function GetParamValues() As ArrayList
        Dim aParams As New ArrayList
        aParams.Add("ANALÍTICO CONTABLE")
        aParams.Add(GetConditions())
        aParams.Add(lueCompany.GetColumnValue("company_ruc"))
        aParams.Add(lueCompany.GetColumnValue("company_name"))
        aParams.Add(lueCompany.GetColumnValue("company_address"))
        Return aParams
    End Function

    Private Sub bbiProcess_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiProcess.ItemClick
        LoadInputValidations()
        If Not vpInputs.Validate Then
            Return
        End If
        InitializeGrid()
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
        gcMainData.DataSource = Nothing
        'dtQuery = ExecuteSQL("select id from public.account_period")
        sQuery = " SELECT am.id as move_id, am.name as move, l.id as line_id, aj.sequence_report as sequence_report, aj.id as journal_id, aj.name as journal, "
        sQuery += " am.date as date, l.name as glosa, l.debit as debit, l.credit as credit, l.amount_currency as amount_currency, "
        'sQuery += " case left(ac.code,2) when '12' then ai.nro_invoice when '42' then ai.supplier_invoice_number else l.ref end as ref, "
        sQuery += " case when ai.type LIKE 'out%' then ai.nro_invoice when ai.type LIKE 'in%' then ai.supplier_invoice_number else case when l.ref_doc is NULL then '' else l.ref_doc end end as ref, "
        sQuery += " ac.code as account_code, ac.name as account_name, t10.name as name_receipt_type, "
        sQuery += " l.partner_id, rp.name as partner_name, rp.street, rp.city, rp.street2, rp.supplier, rp.customer, case when rp.ref is NULL then '' else rp.ref end as ruc, "
        sQuery += " am.state, t1.code as code_pay_method, t1.name as name_pay_method, t10.code as code_receipt_type, "
        sQuery += " case when l.currency_id is NULL then 'PEN' else 'USD' end as code_currency, "
        sQuery += " rc.name as code_currency2, rc.nombre as name_currency, "
        sQuery += " case when ai.nro_invoice is NULL then '' else ai.nro_invoice end as nro_invoice,  "
        sQuery += " aa.code as py_code "
        sQuery += " FROM account_move_line as l "
        sQuery += " LEFT JOIN account_move am ON l.move_id=am.id "
        sQuery += " LEFT JOIN account_journal aj ON l.journal_id = aj.id "
        sQuery += " LEFT JOIN account_account ac ON ac.id = l.account_id "
        sQuery += " LEFT JOIN account_period ap ON ap.id = l.period_id "
        sQuery += " LEFT JOIN res_partner rp ON rp.id = l.partner_id "
        sQuery += " LEFT JOIN account_invoice ai ON ai.id = l.invoice_id "
        sQuery += " LEFT JOIN account_invoice_line il ON il.id = ai.id "
        sQuery += " LEFT JOIN account_analytic_account aa on aa.id = il.account_analytic_id "
        sQuery += " LEFT JOIN account_voucher av ON av.move_id = am.id "
        sQuery += " LEFT JOIN sunat_table_01 t1 ON t1.id = l.table_01_id "
        sQuery += " LEFT JOIN sunat_table_02 t2 ON t2.id = rp.table_02_id "
        sQuery += " LEFT JOIN sunat_table_10 t10 ON t10.id = aj.table_10_id "
        sQuery += " LEFT JOIN res_currency rc ON rc.id = ai.currency_id "
        sQuery += " LEFT JOIN res_partner_bank pb ON pb.journal_id = l.journal_id "
        sQuery += " LEFT JOIN res_bank bn ON bn.id = pb.bank "
        sQuery += " WHERE l.company_id = " & lueCompany.GetColumnValue("company_id").ToString
        sQuery += " AND am.state='posted' "
        If teDocument.Text <> "" Then
            If teDocument.Text.Contains("*", "%") Then
                sQuery += " AND am.name LIKE '" & Replace(teDocument.Text, "*", "%") & "' "
            Else
                sQuery += " AND am.name='" & teDocument.Text & "' "
            End If
        End If
        If deDateFrom.Text <> "" Then
            sQuery += " AND am.date>='" & Format(deDateFrom.EditValue, "yyyy-MM-dd") & "' "
        End If
        If deDateTo.Text <> "" Then
            sQuery += " AND am.date<='" & Format(deDateTo.EditValue, "yyyy-MM-dd") & "' "
        End If
        If lueAccount.Text <> "" Then
            sQuery += " AND LEFT(ac.code," & Len(lueAccount.Text.Trim) & ") ='" & lueAccount.Text & "'"
        End If
        If luePartner.Text <> "" Then
            sQuery += " AND am.partner_id ='" & luePartner.EditValue & "'"
        End If
        sQuery += " AND left(aj.code,2) not in ('13', 'CI') "
        sQuery += " ORDER BY am.date, am.name, am.id;"
        'dtQuery.Rows.Clear()
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
        If rgReportType.SelectedIndex = 1 Then
            dtQuery = DataFilter()
        End If
        For Each row As DataRow In dtQuery.Rows
            'dtPrint.Rows.Add(row("account_code"), row("account_name"), row("move_id"), row("partner_id"), row("partner_name"), row("date"), "", row("ref"), sCurrency, row("debit"), row("credit"), IIf(row("debit") > 0, row("amount_currency"), 0), IIf(row("credit") > 0, row("amount_currency"), 0), row("account_code"))
            dtPrint.Rows.Add(row("account_code"), row("account_name"), row("move"), row("partner_id"), row("partner_name"), row("date"), row("code_receipt_type"), row("ref"), row("code_currency"), row("debit"), row("credit"), row("amount_currency"), row("code_pay_method"), row("name_pay_method"), row("date"), row("move"), row("move"), 0, row("name_currency"), row("ruc"))
        Next
        gcMainData.DataSource = dtQuery
        'GridView1.PopulateColumns()
        LayoutApply()
    End Sub

    Friend Function DataFilter() As DataTable
        Dim dtFilter As New DataTable
        Dim dtResult As New DataTable
        Dim sQuery As String = ""
        dtFilter = SelectDistinct(dtQuery, "account_code", "ruc", "ref")
        dtResult = dtQuery.Clone
        For Each row As DataRow In dtFilter.Rows
            sQuery = "account_code='" & row(0) & "' and ruc='" & row(1) & "' and ref='" & row(2) & "'"
            If dtQuery.Compute("SUM(Debit)-SUM(Credit)", sQuery) <> 0 Then
                For Each dtrow As DataRow In dtQuery.Select(sQuery).CopyToDataTable.Rows
                    dtResult.ImportRow(dtrow)
                Next
            End If
        Next
        Return dtResult
    End Function

    Friend Function GetConditions() As String
        Dim RptCnd As String = "Condiciones:"
        If teDocument.EditValue <> "" Then
            RptCnd += IIf(Len(RptCnd) > 12, ", ", "") & " Documento=" & teDocument.EditValue
        End If
        If deDateFrom.Text <> "" Then
            RptCnd += IIf(Len(RptCnd) > 12, ", ", "") & " Fecha Desde=" & deDateFrom.Text
        End If
        If deDateTo.Text <> "" Then
            RptCnd += IIf(Len(RptCnd) > 12, ", ", "") & " Fecha Hasta=" & deDateTo.Text
        End If
        If lueAccount.Text.Trim <> "" Then
            RptCnd += IIf(Len(RptCnd) > 12, ", ", "") & " Cuenta=" & lueAccount.Text.Trim
        End If
        If luePartner.Text.Trim <> "" Then
            RptCnd += IIf(Len(RptCnd) > 12, ", ", "") & " Socio=" & luePartner.Text.Trim
        End If
        RptCnd += IIf(Len(RptCnd) > 12, ", ", "") & " Resultado=" & IIf(rgReportType.SelectedIndex = 0, "Todo", "Pendiente")
        Return RptCnd
    End Function

    Private Sub LayoutApply()
        'GridView1.PopulateColumns()
        'GridView1.Columns(3).Visible = False
        'GridView1.Columns(12).Visible = False
        For i = 7 To 8
            GridView1.Columns(i).SummaryItem.SetSummary(DevExpress.Data.SummaryItemType.Sum, "{0:n2}")
            GridView1.Columns(i).DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
            GridView1.Columns(i).DisplayFormat.FormatString = "n2"
        Next

    End Sub
    Private Sub bbiExport_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiExport.ItemClick
        ExportarExcel(GridView1)
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
        vpInputs.SetValidationRule(Me.deDateFrom, customValidationRule)
    End Sub

 
    Private Sub GridView1_FocusedRowChanged(sender As Object, e As DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs) Handles GridView1.FocusedRowChanged
        If e.FocusedRowHandle >= 0 Then
            vgcAdditionalData.DataSource = dtQuery.Select("move_id=" & GridView1.GetFocusedRowCellValue("move_id") & " and line_id=" & GridView1.GetFocusedRowCellValue("line_id")).CopyToDataTable
        End If
    End Sub

    Private Sub lueCompany_EditValueChanged(sender As Object, e As EventArgs) Handles lueCompany.EditValueChanged
        SplashScreenManager.ShowForm(Me, GetType(WaitForm), True, True, False)
        SplashScreenManager.Default.SetWaitFormDescription("Cargando tablas de cuentas y socios...")
        LoadAccount()
        LoadAPartner()
        SplashScreenManager.CloseForm(False)
    End Sub

    Private Sub InitializeGrid()
        gcMainData.DataSource = Nothing
        'GridView1.PopulateColumns()
    End Sub

    Private Sub gcMainData_DataSourceChanged(sender As Object, e As EventArgs) Handles gcMainData.DataSourceChanged
        If GridView1.RowCount = 0 Then
            EnableButtons(False)
        End If
    End Sub

    Private Sub EnableButtons(bEnabled As Boolean)
        bbiExport.Enabled = bEnabled
        bbiPrint.Enabled = bEnabled
    End Sub

    Private Sub lueCompany_TextChanged(sender As Object, e As EventArgs) Handles lueCompany.TextChanged, teDocument.TextChanged, deDateFrom.TextChanged, deDateTo.TextChanged, lueAccount.TextChanged, luePartner.TextChanged
        InitializeGrid()
    End Sub

    Private Sub MovimientoContableForm_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        SplitContainerControl2.Collapsed = True
        If bError Then
            Close()
        End If
    End Sub

    Private Sub MovimientoContableForm_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        GridView1.ActiveFilterEnabled = False
        GridView1.OptionsFilter.Reset()
        GridView1.SaveLayoutToRegistry(regKey)
    End Sub
End Class