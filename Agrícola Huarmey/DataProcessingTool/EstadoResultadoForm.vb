Imports DevExpress.XtraEditors.DXErrorProvider
Imports DevExpress.XtraGrid.Views.Grid
Imports System.Drawing
Imports System.IO
Imports System.Globalization
Imports DevExpress.XtraSplashScreen
Imports System.Threading

Public Class EstadoResultadoForm
    Dim dtQuery, dtPrint As New DataTable
    Dim dsMain As New dsMain
    Dim FecIni1, FecFin1, FecIni2, FecFin2 As String
    Dim bError As Boolean = False

    Private Sub EstadoResultadoForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        EnableButtons(False)
        seFiscalYear.Text = Year(Now).ToString
        seFiscalPeriod.Text = Month(Now).ToString
        dtPrint = dsMain.Tables("ReportLayout")
        LoadInputValidations()
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
        pForm.RptFile = "EstadoResultado.rpt"
        pForm.ShowDialog()
    End Sub

    Friend Function GetParamValues() As ArrayList
        Dim aParams As New ArrayList
        aParams.Add("ESTADO DE GANANCIAS Y PÉRDIDAS")
        aParams.Add(seFiscalYear.Text & "-" & Format(seFiscalPeriod.EditValue, "00"))
        aParams.Add(lueCompany.GetColumnValue("company_ruc"))
        aParams.Add(lueCompany.GetColumnValue("company_name"))
        aParams.Add("Soles")
        Return aParams
    End Function

    Private Sub bbiProcess_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiProcess.ItemClick
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
        Dim dAmount1 As Double = 0
        Dim dAmount2 As Double = 0
        FecIni1 = seFiscalYear.Text & "-" & "01-01"
        FecIni2 = (seFiscalYear.EditValue - 1).ToString & "-" & "01-01"
        FecFin1 = Format(LastDayOfMonth(seFiscalYear.Text & "-" & Format(seFiscalPeriod.EditValue, "00") & "-01"), "yyyy-MM-dd")
        FecFin2 = Format(LastDayOfMonth((seFiscalYear.EditValue - 1).ToString & "-" & Format(seFiscalPeriod.EditValue, "00") & "-01"), "yyyy-MM-dd")
        sQuery = " SELECT am.id as move_id, am.name as move, l.id as line_id, aj.sequence_report as sequence_report, aj.id as journal_id, aj.name as journal, am.date as date, l.name as glosa, l.ref as ref, l.debit as debit, l.credit as credit, l.amount_currency as amount_currency, l.currency_id as currency_id, t8.code as t8_code, ac.code as account_code, ac.name as account_name "
        sQuery += " FROM account_move_line as l "
        sQuery += " LEFT JOIN account_move am ON l.move_id=am.id "
        sQuery += " LEFT JOIN account_journal aj ON l.journal_id = aj.id "
        sQuery += " LEFT JOIN sunat_table_08 t8 ON t8.id = aj.table_08_id "
        sQuery += " LEFT JOIN account_account ac ON ac.id = l.account_id "
        sQuery += " LEFT JOIN account_period ap ON ap.id = l.period_id "
        sQuery += " WHERE l.company_id = " & lueCompany.GetColumnValue("company_id").ToString
        sQuery += " AND am.state='posted' "
        sQuery += " AND ((ap.date_start>='" & FecIni1 & "'and ap.date_stop<='" & FecFin1 & "') "
        sQuery += "  OR  (ap.date_start>='" & FecIni2 & "'and ap.date_stop<='" & FecFin2 & "')) "
        sQuery += " AND left(aj.code,2) not in ('13', 'CI') "  'AND l.journal_id= "
        sQuery += " ORDER BY aj.sequence_report, am.name, l.id;"
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
        If dtQuery.Columns.Count = 0 Then
            DevExpress.XtraEditors.XtraMessageBox.Show(Me.LookAndFeel, "La consulta no retornó información.", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If
        dtPrint = ExecuteAccessQuery("select * from ReportLayoutQry where report_id = 1").Select("", "group_id, line_no, code_id").CopyToDataTable
        For Each row As DataRow In dtPrint.Rows
            row("fiscal_year") = seFiscalYear.EditValue
            row("account_id") = IIf(IsDBNull(row("account_id")), "", row("account_id"))
            If row("account_id") <> "" Then
                row("amount1") = GetAmountByAccount(row("account_id"), row("factor"), seFiscalYear.EditValue)
                row("amount2") = GetAmountByAccount(row("account_id"), row("factor"), seFiscalYear.EditValue - 1)
                dAmount1 += row("amount1")
                dAmount2 += row("amount2")
            End If
            If row("account_id") = "CALCULATED" Then
                row("amount1") = Math.Abs(dAmount1) * row("factor")
                row("amount2") = Math.Abs(dAmount2) * row("factor")
                dAmount1 += row("amount1")
                dAmount2 += row("amount2")
            End If
            If row("line_no") = 99 Then
                row("amount1") = dAmount1
                row("amount2") = dAmount2
            End If
        Next
        gcMainData.DataSource = dtPrint
        LayoutApply()
    End Sub

    Friend Function GetAmountByAccount(chart As String, factor As Integer, fiscal_year As Integer) As Double
        Dim dResult As Double = 0
        For Each row As DataRow In dtQuery.Rows
            If Year(row("date")) = fiscal_year Then
                If Strings.Left(row("account_code"), Len(chart)) = chart Then
                    dResult += (row("credit") - row("debit")) * factor
                End If
            End If
        Next
        Return dResult
    End Function

    Private Sub LayoutApply()
        'GridView1.PopulateColumns()
        For i = 4 To 5
            GridView1.Columns(i).SummaryItem.SetSummary(DevExpress.Data.SummaryItemType.Sum, "{0:n2}")
            GridView1.Columns(i).DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
            GridView1.Columns(i).DisplayFormat.FormatString = "n2"
        Next

    End Sub
    Private Sub bbiExport_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiExport.ItemClick
        ExportarExcel(GridView1)
    End Sub

    Private Sub rgTipoAsiento_SelectedIndexChanged(sender As Object, e As EventArgs)
        seFiscalPeriod.Enabled = False
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

    Private Sub EstadoResultadoForm_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        If bError Then
            Close()
        End If
    End Sub
End Class