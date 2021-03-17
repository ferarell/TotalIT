Imports DevExpress.XtraEditors
Imports DevExpress.XtraEditors.DXErrorProvider
Imports DevExpress.XtraSplashScreen
Imports System.Threading
Public Class DeclaracionImpuestoForm
    Dim oSharePointTransactions As New SharePointListTransactions
    Dim dtSunat, dtPeriod, dtMovimientos, dtImpuesto As New DataTable
    Public Sub New()

        ' Esta llamada es exigida por el diseñador.
        InitializeComponent()

        ' Agregue cualquier inicialización después de la llamada a InitializeComponent().
        DevExpress.Skins.SkinManager.EnableFormSkins()
        DevExpress.UserSkins.BonusSkins.Register()
        oSharePointTransactions.SharePointUrl = My.Settings.SharePoint_Url
    End Sub

    Private Sub bbiQuery_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiConsultar.ItemClick
        Try
            SplashScreenManager.ShowForm(Me, GetType(WaitForm), True, True, False)
            SplashScreenManager.Default.SetWaitFormDescription("Obteniendo Declaraciones de Impuestos")
            gcDeclaraciones.DataSource = Nothing
            gcMovimientos.DataSource = Nothing
            LoadTaxes()
            SplashScreenManager.Default.SetWaitFormDescription("Obteniendo Movimientos")
            LoadMoviments(luePeriodos.Text)
            SplashScreenManager.CloseForm(False)
            GridView1.BestFitColumns()
            GridView2.ExpandAllGroups()
        Catch ex As Exception
            SplashScreenManager.CloseForm(False)
            DevExpress.XtraEditors.XtraMessageBox.Show(Me.LookAndFeel, ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub
    Private Sub LoadTaxes()
        oSharePointTransactions.SharePointList = "DeclaracionImpuesto"
        oSharePointTransactions.FieldsList.Clear()
        oSharePointTransactions.FieldsList.Add({"ID"})
        oSharePointTransactions.FieldsList.Add({"PeriodoContable"})
        oSharePointTransactions.FieldsList.Add({"BaseImponibleIngreso"})
        oSharePointTransactions.FieldsList.Add({"ImpuestoIngreso"})
        oSharePointTransactions.FieldsList.Add({"BaseImponibleEgreso"})
        oSharePointTransactions.FieldsList.Add({"ImpuestoEgreso"})
        oSharePointTransactions.FieldsList.Add({"ValorImpuesto"})
        oSharePointTransactions.FieldsList.Add({"ValorRenta"})
        oSharePointTransactions.FieldsList.Add({"ImporteCalculadoImpuesto"})
        oSharePointTransactions.FieldsList.Add({"ImporteCalculadoRenta"})
        oSharePointTransactions.FieldsList.Add({"SaldoImpuesto"})
        oSharePointTransactions.FieldsList.Add({"SaldoRenta"})
        oSharePointTransactions.FieldsList.Add({"ImporteImpuesto"})
        oSharePointTransactions.FieldsList.Add({"ImporteRenta"})

        dtSunat.Clear()
        dtSunat = oSharePointTransactions.GetItems()
        gcDeclaraciones.DataSource = dtSunat
    End Sub
    Private Sub DeclaracionImpuestoForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        DevExpress.LookAndFeel.UserLookAndFeel.Default.SetSkinStyle(My.Settings.LookAndFeel)
        LoadPeriod()
    End Sub

    Private Sub GridView1_SelectionChanged(sender As Object, e As DevExpress.Data.SelectionChangedEventArgs) Handles GridView1.SelectionChanged
        LoadMoviments(luePeriodos.EditValue)
    End Sub

    Private Sub bbiCalcular_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiCalcular.ItemClick
        LoadInputValidations()
        If Not vpInputs.Validate Then
            Return
        End If
        If GridView1.RowCount = 0 Then
            bbiConsultar.PerformClick()
        End If
        GridView1.ActiveFilterString = "([PeriodoContable] = '" & Convert.ToString(CInt(luePeriodos.Text)) & "')"
        If GridView1.RowCount > 0 Then
            If GridView1.GetFocusedRowCellValue("ID").ToString <> "" Then
                XtraMessageBox.Show("Ya existe la declaración de impuesto del periodo que ha seleecionado para el cáculo.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            Else
                GridView1.GetFocusedDataRow.AcceptChanges()
                GridView1.GetFocusedDataRow.Delete()
            End If

        End If
        If Mid(luePeriodos.Text, 5, 2) = "01" Then
            GridView1.ActiveFilterString = "([PeriodoContable] = '" & Convert.ToString(CInt(Mid(luePeriodos.Text, 1, 4) - 1) & "12") & "')"
        Else
            GridView1.ActiveFilterString = "([PeriodoContable] = '" & Convert.ToString(CInt(luePeriodos.Text) - 1) & "')"
        End If

        If GridView1.RowCount = 0 Then
            XtraMessageBox.Show("No existe declaración de impuesto del periodo anterior.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        Dim drImpuestoPeriodoAnterior As DataRow = GridView1.GetFocusedDataRow
        GridView1.ActiveFilterString = Nothing
        Dim dtMovPeriodo As New DataTable
        dtMovPeriodo = dtMovimientos.Select("[Periodo] = '" & luePeriodos.Text & "'").CopyToDataTable
        GridView2.ActiveFilterString = "([Periodo] = '" & luePeriodos.Text & "')"
        Dim BaseImponibleIngreso As Double = SumFieldByType(dtMovPeriodo, "ImporteML", "Ingreso")
        Dim ImpuestoIngreso As Double = SumFieldByType(dtMovPeriodo, "ImporteImpuestoML", "Ingreso")
        Dim BaseImponibleEgreso As Double = SumFieldByType(dtMovPeriodo, "ImporteML", "Egreso")
        Dim ImpuestoEgreso As Double = SumFieldByType(dtMovPeriodo, "ImporteImpuestoML", "Egreso")
        Dim iPos As Integer = 0
        dtSunat.Rows.Add()
        iPos = dtSunat.Rows.Count - 1
        dtSunat.Rows(iPos)("PeriodoContable") = luePeriodos.Text
        dtSunat.Rows(iPos)("BaseImponibleIngreso") = Math.Round(BaseImponibleIngreso)
        dtSunat.Rows(iPos)("BaseImponibleEgreso") = Math.Round(BaseImponibleEgreso)
        dtSunat.Rows(iPos)("ImpuestoIngreso") = Math.Round(ImpuestoIngreso)
        dtSunat.Rows(iPos)("ImpuestoEgreso") = Math.Round(ImpuestoEgreso)
        dtSunat.Rows(iPos)("ImporteCalculadoImpuesto") = Math.Round(ImpuestoIngreso) - Math.Round(ImpuestoEgreso)
        dtSunat.Rows(iPos)("ImporteCalculadoRenta") = Math.Round(BaseImponibleIngreso * 0.01)
        dtSunat.Rows(iPos)("ImporteImpuesto") = dtSunat.Rows(iPos)("ImporteCalculadoImpuesto") - drImpuestoPeriodoAnterior("SaldoImpuesto")
        dtSunat.Rows(iPos)("ImporteRenta") = dtSunat.Rows(iPos)("ImporteCalculadoRenta") - drImpuestoPeriodoAnterior("SaldoRenta")
        dtSunat.Rows(iPos)("SaldoImpuesto") = IIf(CDbl(dtSunat.Rows(iPos)("ImporteImpuesto")) >= 0, 0, Math.Abs(CDbl(dtSunat.Rows(iPos)("ImporteImpuesto"))))
        dtSunat.Rows(iPos)("SaldoRenta") = IIf(CDbl(dtSunat.Rows(iPos)("ImporteRenta")) >= 0, 0, Math.Abs(CDbl(dtSunat.Rows(iPos)("ImporteRenta"))))
        dtSunat.Rows(iPos)("ImporteImpuesto") = IIf(dtSunat.Rows(iPos)("ImporteImpuesto") < 0, 0, dtSunat.Rows(iPos)("ImporteImpuesto"))
        dtSunat.Rows(iPos)("ImporteRenta") = IIf(dtSunat.Rows(iPos)("ImporteRenta") < 0, 0, dtSunat.Rows(iPos)("ImporteRenta"))
    End Sub

    Function SumFieldByType(dtSource As DataTable, sField As String, sType As String) As Double
        Dim dResult As Double = 0
        For r = 0 To dtSource.Rows.Count - 1
            Dim oRow As DataRow = dtSource.Rows(r)
            If oRow("TipoMovimiento") = sType Then
                dResult += CDbl(oRow(sField))
            End If
        Next
        Return dResult
    End Function

    Private Sub bbiExportar_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiExportar.ItemClick
        If GridView1.IsFocusedView Then
            ExportToExcel(gcDeclaraciones)
        ElseIf GridView2.IsFocusedView Then
            ExportToExcel(gcMovimientos)
        End If
    End Sub

    Private Sub bbiPreferencias_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiPreferencias.ItemClick
        Dim oForm As New PreferencesForm
        oForm.ShowDialog()
    End Sub

    Private Sub bbiGuardar_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiGuardar.ItemClick
        Validate()
        Dim bSave As Boolean = False
        Dim oRow As DataRow
        For r = 0 To GridView1.RowCount - 1
            oRow = GridView1.GetDataRow(r)
            If oRow("ID").ToString = "" Then
                bSave = True
                Exit For
            End If
        Next
        If Not bSave Then
            XtraMessageBox.Show("No existe información pendiente por guardar.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        If XtraMessageBox.Show("Está seguro de guardar la información calculada?", "Confirmación", MessageBoxButtons.YesNo) = DialogResult.No Then
            Return
        End If
        Try
            oSharePointTransactions.SharePointList = "DeclaracionImpuesto"
            oSharePointTransactions.ValuesList.Clear()
            oSharePointTransactions.ValuesList.Add({"PeriodoContable", luePeriodos.EditValue})
            oSharePointTransactions.ValuesList.Add({"BaseImponibleIngreso", oRow("BaseImponibleIngreso")})
            oSharePointTransactions.ValuesList.Add({"ImpuestoIngreso", oRow("ImpuestoIngreso")})
            oSharePointTransactions.ValuesList.Add({"BaseImponibleEgreso", oRow("BaseImponibleEgreso")})
            oSharePointTransactions.ValuesList.Add({"ImpuestoEgreso", oRow("ImpuestoEgreso")})
            oSharePointTransactions.ValuesList.Add({"ValorImpuesto", 1})
            oSharePointTransactions.ValuesList.Add({"ValorRenta", 2})
            oSharePointTransactions.ValuesList.Add({"ImporteCalculadoImpuesto", oRow("ImporteCalculadoImpuesto")})
            oSharePointTransactions.ValuesList.Add({"ImporteCalculadoRenta", oRow("ImporteCalculadoRenta")})
            oSharePointTransactions.ValuesList.Add({"SaldoImpuesto", oRow("SaldoImpuesto")})
            oSharePointTransactions.ValuesList.Add({"SaldoRenta", oRow("SaldoRenta")})
            oSharePointTransactions.ValuesList.Add({"ImporteImpuesto", oRow("ImporteImpuesto")})
            oSharePointTransactions.ValuesList.Add({"ImporteRenta", oRow("ImporteRenta")})
            oSharePointTransactions.InsertItem()
        Catch ex As Exception
            XtraMessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub LoadPeriod()
        oSharePointTransactions.SharePointList = "Período Contable"
        oSharePointTransactions.FieldsList.Clear()
        oSharePointTransactions.FieldsList.Add({"ID"})
        oSharePointTransactions.FieldsList.Add({"Periodo"})
        dtPeriod = oSharePointTransactions.GetItems()
        luePeriodos.Properties.DataSource = dtPeriod
        luePeriodos.Properties.DisplayMember = "Periodo"
        luePeriodos.Properties.ValueMember = "ID"
        luePeriodos.Properties.KeyMember = "Periodo"

    End Sub

    Private Sub LoadMoviments(sPeriod As String)
        oSharePointTransactions.SharePointList = "Movimientos"
        oSharePointTransactions.FieldsList.Clear()
        oSharePointTransactions.FieldsList.Add({"ID", "Integer"})
        oSharePointTransactions.FieldsList.Add({"Periodo", "String"})
        oSharePointTransactions.FieldsList.Add({"TipoMovimiento", "String"})
        oSharePointTransactions.FieldsList.Add({"Tipo_x0020_de_x0020_Comprobante", "String"})
        oSharePointTransactions.FieldsList.Add({"SerieComprobante", "String"})
        oSharePointTransactions.FieldsList.Add({"NumeroComprobante", "String"})
        oSharePointTransactions.FieldsList.Add({"FechaComprobante", "Date"})
        oSharePointTransactions.FieldsList.Add({"RazonSocial", "String"})
        oSharePointTransactions.FieldsList.Add({"DescripcionMovimiento", "String"})
        oSharePointTransactions.FieldsList.Add({"ImporteMD", "Decimal"})
        oSharePointTransactions.FieldsList.Add({"CodigoMoneda", "String"})
        oSharePointTransactions.FieldsList.Add({"FechaComprobante_x003a_Importe_x", "Decimal"})
        oSharePointTransactions.FieldsList.Add({"ImporteML", "Decimal"})
        oSharePointTransactions.FieldsList.Add({"Impuesto_x003a_FactorImpuesto", "Decimal"})
        oSharePointTransactions.FieldsList.Add({"ImporteImpuestoML", "Decimal"})
        dtMovimientos.Clear()
        dtMovimientos = oSharePointTransactions.GetItems()
        gcMovimientos.DataSource = dtMovimientos
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
        vpInputs.SetValidationRule(Me.luePeriodos, customValidationRule)
    End Sub
End Class