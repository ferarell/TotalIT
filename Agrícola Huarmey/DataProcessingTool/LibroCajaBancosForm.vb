Imports DevExpress.XtraEditors.DXErrorProvider
Imports DevExpress.XtraGrid.Views.Grid
Imports System.Drawing
Imports System.IO
Imports System.Globalization
Imports DevExpress.XtraSplashScreen
Imports System.Threading

Public Class LibroCajaBancosForm
    Dim dtQuery, dtPrint As New DataTable
    Dim dsMain As New dsMain
    Dim sPeriod, FecIni, FecFin As String
    Dim bError As Boolean = False

    Private Sub LibroCajaBancosForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        EnableButtons(False)
        rgFormat.SelectedIndex = 0
        seFiscalYear.Text = Year(Now).ToString
        seFiscalPeriod.Text = Month(Now).ToString
        Try
            LoadCompany()
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

    Private Sub bbiClose_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiClose.ItemClick
        Close()
    End Sub

    Private Sub bbiPrint_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiPrint.ItemClick
        Dim pForm As New PrintForm
        pForm.dtPrint = dtPrint
        pForm.aParams = GetParamValues()
        pForm.RptFile = "LibroCajaBancos" & rgFormat.EditValue & ".rpt"
        pForm.ShowDialog()
    End Sub

    Friend Function GetParamValues() As ArrayList
        Dim aParams As New ArrayList
        Dim sFormat As String = "LIBRO CAJA Y BANCOS - DETALLE DE LOS MOVIMIENTOS "
        If rgFormat.EditValue = "1" Then
            sFormat += "DEL EFECTIVO (Formato 1.1)"
        Else
            sFormat += "DE LA CUENTA CORRIENTE (Formato 1.2)"
        End If
        aParams.Add(sFormat)
        aParams.Add(MonthName(seFiscalPeriod.EditValue).ToUpper & " " & seFiscalYear.Text)
        aParams.Add(lueCompany.GetColumnValue("company_ruc"))
        aParams.Add(lueCompany.GetColumnValue("company_name"))
        aParams.Add("Soles")
        Return aParams
    End Function

    Private Sub bbiProcess_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiProcess.ItemClick
        LoadInputValidations()
        If Not vpInputs.Validate Then
            Return
        End If
        InitializeGrid()
        dtPrint = dsMain.Tables("LibroCajaBanco" & rgFormat.EditValue)
        sPeriod = seFiscalYear.Text & Format(seFiscalPeriod.EditValue, "00") & "00"
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
        Dim dtBalance As New DataTable
        Dim sQuery As String = ""
        FecIni = seFiscalYear.Text & "-" & "01-01"
        FecFin = Format(LastDayOfMonth(seFiscalYear.Text & "-" & Format(seFiscalPeriod.EditValue - 1, "00") & "-01"), "yyyy-MM-dd")
        If rgFormat.SelectedIndex = 0 Then
            sQuery = " SELECT 1 as RegType, ac.code as account_code, ac.name as account_name, "
        Else
            sQuery = " SELECT 1 as RegType, ac.code as account_code, ac.name as account_name, pb.acc_number, pb.bank_name, "
        End If
        sQuery += " SUM(l.debit) as debit, SUM(l.credit) AS credit "
        sQuery += " FROM account_move_line as l "
        sQuery += " LEFT JOIN account_move am ON l.move_id=am.id "
        sQuery += " LEFT JOIN account_account ac ON ac.id = l.account_id "
        If rgFormat.SelectedIndex = 1 Then
            sQuery += " LEFT JOIN res_partner_bank pb ON pb.journal_id = l.journal_id "
            sQuery += " LEFT JOIN res_bank bn ON bn.id = pb.bank "
        End If
        sQuery += " WHERE l.company_id = " & lueCompany.GetColumnValue("company_id").ToString
        sQuery += " AND  am.state='posted' AND l.date BETWEEN '" & FecIni & "' AND '" & FecFin & "' "
        If rgFormat.SelectedIndex = 0 Then
            sQuery += " AND LEFT(ac.code,3) BETWEEN '101' AND '103' "
        Else
            sQuery += " AND LEFT(ac.code,3) BETWEEN '104' AND '109' "
        End If
        If rgFormat.SelectedIndex = 0 Then
            sQuery += " GROUP BY ac.code, ac.name "
        Else
            sQuery += " GROUP BY ac.code, ac.name, pb.acc_number, pb.bank_name "
        End If
        dtBalance = ExecuteSQL(lueCompany.GetColumnValue("service_url"), sQuery)
        dtQuery.Rows.Clear()
        FecIni = seFiscalYear.Text & "-" & Format(seFiscalPeriod.EditValue, "00") & "-01"
        FecFin = Format(LastDayOfMonth(FecIni), "yyyy-MM-dd")
        sQuery = "select id from account_period where date_start = '" & FecIni & "' and date_stop = '" & FecFin & "' and company_id = " & lueCompany.GetColumnValue("company_id").ToString
        dtQuery = ExecuteSQL(lueCompany.GetColumnValue("service_url"), sQuery)
        If dtQuery.Rows.Count = 0 Then
            SplashScreenManager.CloseForm(False)
            DevExpress.XtraEditors.XtraMessageBox.Show(Me.LookAndFeel, "No existe el periodo fiscal asignado.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        Dim sPeriodId As String = dtQuery.Rows(0)("id")
        sQuery = " SELECT 2 as RegType, am.id as move_id, am.name as move, l.id as line_id, aj.sequence_report as sequence_report, aj.id as journal_id, aj.name as journal, "
        sQuery += " am.date as date, l.name as glosa, l.debit as debit, l.credit as credit, l.amount_currency as amount_currency, "
        sQuery += " case when ai.type LIKE 'out%' then ai.nro_invoice when ai.type LIKE 'in%' then ai.supplier_invoice_number else case when l.ref is NULL then am.ref else l.ref end end as ref, "
        'sQuery += "l.currency_id as currency_id, t8.code as t8_code, ac.code as account_code, ac.name as account_name, "
        sQuery += " l.currency_id as currency_id, ac.code as account_code, ac.name as account_name, "
        sQuery += " am.partner_id, rp.name as partner_name, rp.street, rp.city, rp.street2, rp.supplier, rp.customer, rp.ref as ruc, "
        sQuery += " am.state, t1.code as code_pay_method, t1.name as name_pay_method, t10.code as code_receipt_type, "
        sQuery += " t10.name as name_receipt_type, rc.name as code_currency, rc.nombre as name_currency, pb.acc_number, "
        sQuery += " pb.bank_name, bn.code_sunat as code_bank_sunat, t2.code as identity_id, "
        If rgFormat.SelectedIndex = 0 Then
            sQuery += " CASE WHEN LEFT(ac.code,3) BETWEEN '101' AND '103' THEN 'G' ELSE ' ' END AS account_type "
        Else
            sQuery += " CASE WHEN LEFT(ac.code,3) BETWEEN '104' AND '109' THEN 'G' ELSE ' ' END AS account_type "
        End If
        sQuery += " FROM account_move_line as l "
        sQuery += " LEFT JOIN account_move am ON l.move_id=am.id "
        sQuery += " LEFT JOIN account_journal aj ON l.journal_id = aj.id "
        sQuery += " LEFT JOIN account_account ac ON ac.id = l.account_id "
        sQuery += " LEFT JOIN account_period ap ON ap.id = l.period_id "
        sQuery += " LEFT JOIN res_partner rp ON rp.id = am.partner_id "
        sQuery += " LEFT JOIN account_invoice ai ON ai.move_id = l.move_id "
        sQuery += " LEFT JOIN sunat_table_01 t1 ON t1.id = l.table_01_id "
        sQuery += " LEFT JOIN sunat_table_02 t2 ON t2.id = rp.table_02_id "
        sQuery += " LEFT JOIN sunat_table_10 t10 ON t10.id = aj.table_10_id "
        sQuery += " LEFT JOIN res_currency rc ON rc.id = ai.currency_id "
        sQuery += " LEFT JOIN res_partner_bank pb ON pb.journal_id = l.journal_id "
        sQuery += " LEFT JOIN res_bank bn ON bn.id = pb.bank "
        sQuery += " LEFT JOIN account_voucher av ON av.move_id = am.id "

        'sQuery += " LEFT JOIN account_transfer at ON av.move_id = am.id "
        sQuery += " WHERE l.company_id = " & lueCompany.GetColumnValue("company_id").ToString
        sQuery += " AND am.state='posted' "
        sQuery += " AND l.period_id=" & sPeriodId
        If rgFormat.SelectedIndex = 0 Then
            sQuery += " AND EXISTS(SELECT * FROM account_move_line aml JOIN account_account aa ON aa.id = aml.account_id"
            sQuery += "             WHERE aml.move_id = l.move_id and LEFT(aa.code,3) BETWEEN '101' AND '103')"
            'sQuery += " AND LEFT(ac.code,3) NOT BETWEEN '101' AND '103' "
        Else
            sQuery += " AND EXISTS(SELECT * FROM account_move_line aml JOIN account_account aa ON aa.id = aml.account_id"
            sQuery += "             WHERE aml.move_id = l.move_id and LEFT(aa.code,3) BETWEEN '104' AND '109')"
            'sQuery += " AND LEFT(ac.code,3) NOT BETWEEN '104' AND '109' "
        End If
        sQuery += "ORDER BY date, l.id "
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
        dtPrint.Rows.Clear()
        If dtQuery.Columns.Count = 0 Then
            DevExpress.XtraEditors.XtraMessageBox.Show(Me.LookAndFeel, "La consulta no retornó información.", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If
        dtBalance.Columns.Add("account_group", GetType(String)).DefaultValue = ""
        For Each bRow As DataRow In dtBalance.Rows
            If rgFormat.SelectedIndex = 0 Then
                bRow("account_group") = bRow("account_code") & " - " & bRow("account_name")
                dtPrint.Rows.Add(bRow("RegType"), sPeriod, "", "", bRow("account_code"), "", "", "", "", "", "", "", "", "", "", "", bRow("debit"), bRow("credit"), "", "", bRow("account_group"), bRow("account_name"))
            Else
                bRow("account_group") = bRow("acc_number") & " - " & bRow("bank_name")
                dtPrint.Rows.Add(bRow("RegType"), sPeriod, "", "", "", "", "", "", "", "", "", "", "", bRow("debit"), bRow("credit"), "", bRow("account_group"), bRow("account_code"), "", bRow("account_name"))
            End If
        Next
        Dim dtDetail As New DataTable
        dtDetail = dtQuery.Copy
        dtQuery.Columns.Add("account_group", GetType(String)).DefaultValue = ""
        For Each dRow As DataRow In dtQuery.Rows
            For c = 0 To dRow.ItemArray.Count - 1
                If IsDBNull(dRow(c)) Then
                    If dRow.Table.Columns(c).DataType.Name = "String" Then
                        dRow(c) = ""
                    End If
                    If dRow.Table.Columns(c).DataType.Name = "Decimal" Then
                        dRow(c) = 0
                    End If
                End If
            Next
            Dim oRow As DataRow = dtDetail.Select("move_id=" & dRow("move_id") & " and account_type='G'").CopyToDataTable.Rows(0)
            If rgFormat.SelectedIndex = 0 Then
                dRow("account_group") = oRow("account_code") & " - " & oRow("account_name")
            Else
                dRow("account_group") = oRow("acc_number") & " - " & oRow("bank_name")
            End If
            If dRow("account_type") <> "G" Then
                If rgFormat.SelectedIndex = 0 Then
                    dtPrint.Rows.Add(dRow("RegType"), sPeriod, dRow("move"), "M" & dRow("line_id"), dRow("account_code"), "", "", dRow("code_currency"), dRow("code_receipt_type"), "", Mid(dRow("move"), 1, 20), Format(dRow("date"), "dd/MM/yyyy"), "", Format(dRow("date"), "dd/MM/yyyy"), dRow("glosa"), "", dRow("credit"), dRow("debit"), "", "", dRow("account_group"), dRow("account_name"))
                Else
                    dtPrint.Rows.Add(dRow("RegType"), sPeriod, dRow("move"), "M" & dRow("line_id"), Strings.Right("00" & dRow("code_bank_sunat"), 2), dRow("acc_number"), Format(dRow("date"), "dd/MM/yyyy"), dRow("code_pay_method"), dRow("glosa"), dRow("identity_id"), dRow("ruc"), dRow("partner_name"), dRow("ref"), dRow("credit"), dRow("debit"), "", dRow("account_group"), dRow("account_code"), dRow("account_name"), dRow("bank_name"))
                End If
            End If
        Next
        gcMainData.DataSource = dtPrint
        LayoutApply()
    End Sub

    Private Sub LayoutApply()
        GridView1.PopulateColumns()
        For i = 13 To 17
            If (rgFormat.SelectedIndex = 0 And (i = 16 Or i = 17)) Or (rgFormat.SelectedIndex = 1 And (i = 13 Or i = 14)) Then
                GridView1.Columns(i).SummaryItem.SetSummary(DevExpress.Data.SummaryItemType.Sum, "{0:n2}")
                GridView1.Columns(i).DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
                GridView1.Columns(i).DisplayFormat.FormatString = "n2"
            End If
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
    End Sub

    Private Sub LibroDiarioForm_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        If bError Then
            Close()
        End If
    End Sub

    Private Sub rgFormat_SelectedIndexChanged(sender As Object, e As EventArgs) Handles rgFormat.SelectedIndexChanged
        EnableButtons(False)
        InitializeGrid()
    End Sub

    Private Sub InitializeGrid()
        gcMainData.DataSource = Nothing
        GridView1.PopulateColumns()
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

    Private Sub lueCompany_TextChanged(sender As Object, e As EventArgs) Handles lueCompany.TextChanged, seFiscalPeriod.TextChanged, seFiscalYear.TextChanged
        InitializeGrid()
    End Sub

End Class