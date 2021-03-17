Imports DevExpress.XtraEditors.DXErrorProvider
Imports DevExpress.XtraGrid.Views.Grid
Imports System.Drawing
Imports System.IO
Imports System.Globalization
Imports DevExpress.XtraSplashScreen
Imports System.Threading

Public Class SaldoCuentaAnualForm
    Dim dtQuery, dtPrint As New DataTable
    Dim dsMain As New dsMain
    Dim FecIni, FecFin As String
    Dim bError As Boolean = False

    Private Sub GastosIngresosForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        EnableButtons(False)
        seFiscalYear.Text = Year(Now).ToString
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
        aParams.Add("ANALÍTICO CONTABLE")
        aParams.Add(GetConditions())
        aParams.Add(lueCompany.GetColumnValue("company_ruc"))
        aParams.Add(lueCompany.GetColumnValue("company_name"))
        aParams.Add(lueCompany.GetColumnValue("company_address"))
        Return aParams
    End Function

    Friend Function GetConditions() As String
        Dim RptCnd As String = "Condiciones:"
        If seFiscalYear.Text <> "" Then
            RptCnd += "Año=" & seFiscalYear.Text
        End If
        Return RptCnd
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
        sQuery = " SELECT LEFT(ac.code,2) as account_group, ac.code as account_code, ac.name as account_name, "
        sQuery += " SUM(CASE WHEN to_char(l.date, 'MM') = '01' THEN l.debit - l.credit ELSE 0 END) AS M01, "
        sQuery += " SUM(CASE WHEN to_char(l.date, 'MM') = '02' THEN l.debit - l.credit ELSE 0 END) AS M02, "
        sQuery += " SUM(CASE WHEN to_char(l.date, 'MM') = '03' THEN l.debit - l.credit ELSE 0 END) AS M03, "
        sQuery += " SUM(CASE WHEN to_char(l.date, 'MM') = '04' THEN l.debit - l.credit ELSE 0 END) AS M04, "
        sQuery += " SUM(CASE WHEN to_char(l.date, 'MM') = '05' THEN l.debit - l.credit ELSE 0 END) AS M05, "
        sQuery += " SUM(CASE WHEN to_char(l.date, 'MM') = '06' THEN l.debit - l.credit ELSE 0 END) AS M06, "
        sQuery += " SUM(CASE WHEN to_char(l.date, 'MM') = '07' THEN l.debit - l.credit ELSE 0 END) AS M07, "
        sQuery += " SUM(CASE WHEN to_char(l.date, 'MM') = '08' THEN l.debit - l.credit ELSE 0 END) AS M08, "
        sQuery += " SUM(CASE WHEN to_char(l.date, 'MM') = '09' THEN l.debit - l.credit ELSE 0 END) AS M09, "
        sQuery += " SUM(CASE WHEN to_char(l.date, 'MM') = '10' THEN l.debit - l.credit ELSE 0 END) AS M10, "
        sQuery += " SUM(CASE WHEN to_char(l.date, 'MM') = '11' THEN l.debit - l.credit ELSE 0 END) AS M11, "
        sQuery += " SUM(CASE WHEN to_char(l.date, 'MM') = '12' THEN l.debit - l.credit ELSE 0 END) AS M12, "
        sQuery += " SUM(l.debit - l.credit) AS total "
        sQuery += " FROM account_move_line as l "
        sQuery += " LEFT JOIN account_move am ON l.move_id=am.id "
        sQuery += " LEFT JOIN account_account ac ON ac.id = l.account_id "
        sQuery += " WHERE l.company_id = " & lueCompany.GetColumnValue("company_id").ToString
        sQuery += " AND  am.state='posted' AND l.date BETWEEN '" & seFiscalYear.Text & "-01-01' AND '" & seFiscalYear.Text & "-12-31' "
        sQuery += " GROUP BY LEFT(ac.code,2), ac.code, ac.name "
        sQuery += " ORDER BY LEFT(ac.code,2);"
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
        'Dim sCurrency As String = ""
        'dtPrint.Rows.Clear()
        If dtQuery.Columns.Count = 0 Then
            DevExpress.XtraEditors.XtraMessageBox.Show(Me.LookAndFeel, "La consulta no retornó información.", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If
        'For Each row As DataRow In dtQuery.Rows
        '    dtPrint.Rows.Add(row("move"), row("date"), row("glosa"), row("t8_code"), row("move"), row("account_code"), row("account_name"), sCurrency, row("debit"), row("credit"), IIf(row("debit") > 0, row("amount_currency"), 0), IIf(row("credit") > 0, row("amount_currency"), 0), row("sequence_report") & row("move") & row("move_id"))
        'Next
        gcMainData.DataSource = dtQuery
        LayoutApply()
    End Sub

    Private Sub LayoutApply()
        For i = 2 To GridView1.Columns.Count - 1
            GridView1.Columns(i).DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
            GridView1.Columns(i).DisplayFormat.FormatString = "###,###,##0.00"
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
        vpInputs.SetValidationRule(Me.seFiscalYear, customValidationRule)
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

    Private Sub lueCompany_TextChanged(sender As Object, e As EventArgs) Handles lueCompany.TextChanged, seFiscalYear.TextChanged
        InitializeGrid()
    End Sub

    Private Sub GastosIngresosForm_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        If bError Then
            Close()
        End If
    End Sub
End Class