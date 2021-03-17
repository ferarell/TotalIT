Imports DevExpress.XtraEditors.DXErrorProvider
Imports DevExpress.XtraGrid.Views.Grid
Imports System.Drawing
Imports System.IO
Imports System.Globalization
Imports DevExpress.XtraSplashScreen
Imports System.Threading

Public Class LibroDiarioForm
    Dim dtQuery, dtPrint, dtFlatFile1, dtFlatFile2 As New DataTable
    Dim dsMain As New dsMain
    Dim FecIni, FecFin As String
    Dim bError As Boolean = False

    Private Sub LibroDiarioForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        EnableButtons(False)
        seFiscalYear.Text = Year(Now).ToString
        seFiscalPeriod.Text = Month(Now).ToString
        dtPrint = dsMain.Tables("LibroDiario")
        dtFlatFile1 = dsMain.Tables("LibroDiarioPle1")
        dtFlatFile2 = dsMain.Tables("LibroDiarioPle2")
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

    Private Sub bbiClose_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiClose.ItemClick
        Close()
    End Sub

    Private Sub bbiPrint_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiPrint.ItemClick
        Dim pForm As New PrintForm
        pForm.dtPrint = dtPrint
        pForm.aParams = GetParamValues()
        pForm.RptFile = "LibroDiario.rpt"
        pForm.ShowDialog()
    End Sub

    Friend Function GetParamValues() As ArrayList
        Dim aParams As New ArrayList
        Dim sPeriod As String = MonthName(seFiscalPeriod.EditValue).ToUpper & " " & seFiscalYear.Text
        aParams.Add("FORMATO 5.1 - LIBRO DIARIO")
        aParams.Add(sPeriod)
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
        FecIni = seFiscalYear.Text & "-" & Format(seFiscalPeriod.EditValue, "00") & "-01"
        FecFin = Format(LastDayOfMonth(FecIni), "yyyy-MM-dd")
        If rgLedgerType.SelectedIndex = 0 Then
            FecIni = seFiscalYear.Text & "-" & "01-01"
            FecFin = seFiscalYear.Text & "-" & "01-01"
        ElseIf rgLedgerType.SelectedIndex = 2 Then
            FecIni = seFiscalYear.Text & "-" & "12-31"
            FecFin = seFiscalYear.Text & "-" & "12-31"
        End If
        dtQuery.Rows.Clear()
        sQuery = "select id from account_period where date_start = '" & FecIni & "' and date_stop = '" & FecFin & "' and company_id = " & lueCompany.GetColumnValue("company_id").ToString
        dtQuery = ExecuteSQL(lueCompany.GetColumnValue("service_url"), sQuery)
        If dtQuery.Rows.Count = 0 Then
            SplashScreenManager.CloseForm(False)
            DevExpress.XtraEditors.XtraMessageBox.Show(Me.LookAndFeel, "No existe el periodo fiscal asignado.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        Dim sPeriodId As String = dtQuery.Rows(0)("id")
        sQuery = " SELECT am.id as move_id, am.name as move, l.id as line_id, aj.sequence_report as sequence_report, aj.id as journal_id, aj.name as journal, am.date as date, l.name as glosa, l.debit as debit, l.credit as credit, l.amount_currency as amount_currency, l.currency_id as currency_id, t8.code as t8_code, ac.code as account_code, ac.name as account_name, "
        sQuery += " case when ai.type LIKE 'out%' then ai.nro_invoice when ai.type LIKE 'in%' then ai.supplier_invoice_number else case when l.ref is NULL then am.ref else l.ref end end as ref "
        sQuery += " FROM account_move_line as l "
        sQuery += " LEFT JOIN account_move am ON l.move_id=am.id "
        sQuery += " LEFT JOIN account_journal aj ON l.journal_id = aj.id "
        sQuery += " LEFT JOIN sunat_table_08 t8 ON t8.id = aj.table_08_id "
        sQuery += " LEFT JOIN account_account ac ON ac.id = l.account_id "
        sQuery += " LEFT JOIN account_invoice ai ON ai.id = l.invoice_id "
        sQuery += " WHERE l.company_id = " & lueCompany.GetColumnValue("company_id").ToString
        sQuery += " AND am.state='posted' "
        sQuery += " AND l.period_id=" & sPeriodId
        sQuery += " ORDER BY aj.sequence_report , am.name, l.id"
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
        For Each row As DataRow In dtQuery.Rows
            sCurrency = "PEN"
            If Not IsDBNull(row("currency_id")) Then
                sCurrency = "USD"
            End If
            dtPrint.Rows.Add(row("move"), row("date"), row("glosa"), row("t8_code"), row("ref"), row("account_code"), row("account_name"), sCurrency, row("debit"), row("credit"), IIf(row("debit") > 0, row("amount_currency"), 0), IIf(row("credit") > 0, row("amount_currency"), 0), row("sequence_report") & row("move") & row("move_id"))
        Next
        gcMainData.DataSource = dtPrint
        LayoutApply()
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
        ExportarExcel(GridView1)
    End Sub

    Private Sub rgTipoAsiento_SelectedIndexChanged(sender As Object, e As EventArgs) Handles rgLedgerType.SelectedIndexChanged
        seFiscalPeriod.Enabled = False
        EnableButtons(False)
        InitializeGrid()
        If sender.selectedindex = 0 Then
            seFiscalPeriod.EditValue = 1
        ElseIf sender.selectedindex = 2 Then
            seFiscalPeriod.EditValue = 12
        ElseIf sender.selectedindex = 1 Then
            seFiscalPeriod.Enabled = True
        End If
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

    Private Sub LibroDiarioForm_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        If bError Then
            Close()
        End If
    End Sub

    Private Sub bbiSunatPle_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiSunatPle.ItemClick

    End Sub

End Class