Imports System.IO
Imports DevExpress.XtraEditors.DXErrorProvider
Imports DevExpress.XtraGrid.Views.Grid
Imports DevExpress.XtraSplashScreen

Public Class PurchasesAndInventoryForm

    Dim dtPresupuesto, dtInventario, dtControlCalidad, dtResult, dtDifference As New DataTable
    Private Sub bePresupuesto_Properties_ButtonClick(sender As Object, e As DevExpress.XtraEditors.Controls.ButtonPressedEventArgs) Handles bePresupuesto.Properties.ButtonClick
        'XtraOpenFileDialog1.Filter = "Source Files (*.xls*;*.csv)|*.xls*;*.csv"
        XtraOpenFileDialog1.Filter = "Source Files (*.xls*)|*.xls*"
        XtraOpenFileDialog1.FileName = ""
        'XtraOpenFileDialog1.InitialDirectory = ""
        If XtraOpenFileDialog1.ShowDialog() = DialogResult.OK Then
            bePresupuesto.Text = XtraOpenFileDialog1.FileName
        End If
    End Sub

    Private Sub beInventario_Properties_ButtonClick(sender As Object, e As DevExpress.XtraEditors.Controls.ButtonPressedEventArgs) Handles beInventario.Properties.ButtonClick
        'XtraOpenFileDialog1.Filter = "Source Files (*.xls*;*.csv)|*.xls*;*.csv"
        XtraOpenFileDialog1.Filter = "Source Files (*.xls*)|*.xls*"
        XtraOpenFileDialog1.FileName = ""
        'XtraOpenFileDialog1.InitialDirectory = ""
        If XtraOpenFileDialog1.ShowDialog() = DialogResult.OK Then
            beInventario.Text = XtraOpenFileDialog1.FileName
        End If
    End Sub

    Private Sub beControlCalidad_Properties_ButtonClick(sender As Object, e As DevExpress.XtraEditors.Controls.ButtonPressedEventArgs) Handles beControlCalidad.Properties.ButtonClick
        'XtraOpenFileDialog1.Filter = "Source Files (*.xls*;*.csv)|*.xls*;*.csv"
        XtraOpenFileDialog1.Filter = "Source Files (*.xls*)|*.xls*"
        XtraOpenFileDialog1.FileName = ""
        'XtraOpenFileDialog1.InitialDirectory = ""
        If XtraOpenFileDialog1.ShowDialog() = DialogResult.OK Then
            beControlCalidad.Text = XtraOpenFileDialog1.FileName
        End If
    End Sub

    Private Sub ResultTableCreate()
        dtResult.Columns.Add("Proveedor", GetType(String))
        dtResult.Columns.Add("Producto", GetType(String))
        dtResult.Columns.Add("UnidadMedida", GetType(String))
        dtResult.Columns.Add("PptoCantidad", GetType(Decimal))
        dtResult.Columns.Add("PptoImporte", GetType(Decimal))
        dtResult.Columns.Add("SaldoIniCantidad", GetType(Decimal))
        dtResult.Columns.Add("SaldoIniImporte", GetType(Decimal))
        dtResult.Columns.Add("PptoIngresoCantidad", GetType(Decimal))
        dtResult.Columns.Add("PptoIngresoImporte", GetType(Decimal))
        dtResult.Columns.Add("IngresoCantidad", GetType(Decimal))
        dtResult.Columns.Add("IngresoImporte", GetType(Decimal))
        dtResult.Columns.Add("PptoSalidaCantidad", GetType(Decimal))
        dtResult.Columns.Add("PptoSalidaImporte", GetType(Decimal))
        dtResult.Columns.Add("SalidaCantidad", GetType(Decimal))
        dtResult.Columns.Add("SalidaImporte", GetType(Decimal))
        dtResult.Columns.Add("PptoSaldoCantidad", GetType(Decimal))
        dtResult.Columns.Add("PptoSaldoImporte", GetType(Decimal))
        'dtResult.Columns.Add("SaldoCalcCantidad", GetType(Decimal))
        'dtResult.Columns.Add("SaldoCalcImporte", GetType(Decimal))
        dtResult.Columns.Add("SaldoSistCantidad", GetType(Decimal))
        dtResult.Columns.Add("SaldoSistImporte", GetType(Decimal))
        dtResult.Columns.Add("VarPtoConsumoCantidad", GetType(Decimal))
        dtResult.Columns.Add("VarPtoSaldoCantidad", GetType(Decimal))
        'dtResult.Columns.Add("DifSaldoPptoCalcCantidad", GetType(Decimal))

    End Sub

    Private Sub DifferenceTableCreate()
        dtDifference.Columns.Add("Producto", GetType(String))
        dtDifference.Columns.Add("IngresoCantidad", GetType(Decimal))
        dtDifference.Columns.Add("IngresoImporte", GetType(Decimal))
        dtDifference.Columns.Add("SalidaCantidad", GetType(Decimal))
        dtDifference.Columns.Add("SalidaImporte", GetType(Decimal))
        dtDifference.Columns.Add("SaldoCantidad", GetType(Decimal))
        dtDifference.Columns.Add("SaldoImporte", GetType(Decimal))
    End Sub
    Private Sub SaleByWebForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        seFiscalYear.Text = Year(Now).ToString
        seFiscalPeriod.Text = Month(Now).ToString
        GridView1.RestoreLayoutFromRegistry(Directory.GetCurrentDirectory)
        SplitContainerControl2.Collapsed = True
        ResultTableCreate()
        DifferenceTableCreate()
    End Sub

    Private Sub bbiProcesar_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiProcesar.ItemClick
        LoadInputValidations()
        If Not vpInputs.Validate Then
            Return
        End If
        Try
            SplashScreenManager.ShowForm(Me, GetType(WaitForm), True, True, False)
            SplashScreenManager.Default.SetWaitFormDescription("Cargando archivos seleccionados")
            If bePresupuesto.Text.ToUpper.Contains("XLS") Then
                dtPresupuesto.Rows.Clear()
                dtPresupuesto = LoadExcel(bePresupuesto.Text, "INSUMOS QUIMICOS FINAL$").Tables(0)
                If dtPresupuesto.Rows.Count = 0 Then
                    DevExpress.XtraEditors.XtraMessageBox.Show("No se cargó el archivo de Presupuesto, verifique el formato e intente nuevamente.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End If
            End If
            If beInventario.Text.ToUpper.Contains("XLS") Then
                dtInventario.Rows.Clear()
                dtInventario = LoadExcel(beInventario.Text, "{0}").Tables(0)
                If dtInventario.Rows.Count = 0 Then
                    DevExpress.XtraEditors.XtraMessageBox.Show("No se cargó el archivo de Inventario, verifique el formato e intente nuevamente.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End If
            End If
            If beControlCalidad.Text.ToUpper.Contains("XLS") Then
                dtControlCalidad.Rows.Clear()
                dtControlCalidad = LoadExcel(beControlCalidad.Text, "{0}").Tables(0)
                If dtControlCalidad.Rows.Count = 0 Then
                    DevExpress.XtraEditors.XtraMessageBox.Show("No se cargó el archivo de Control de Calidad, verifique el formato e intente nuevamente.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End If
            End If
            'Next
            SplashScreenManager.Default.SetWaitFormDescription("Procesando datos")
            ProcessAllData(dtPresupuesto, dtInventario, dtControlCalidad)
            gcInputsControl.DataSource = dtResult
            FormatGridView(GridView1)
            gcValidations.DataSource = dtDifference
            FormatGridView(GridView2)
            SplashScreenManager.CloseForm(False)
        Catch ex As Exception
            SplashScreenManager.CloseForm(False)
            DevExpress.XtraEditors.XtraMessageBox.Show(Me.LookAndFeel, ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub FormatGridView(oGridView As GridView)
        oGridView.BestFitColumns()
        For c = 0 To oGridView.Columns.Count - 1
            '    If oGridView.Columns(c).Tag = 1 Then
            '        oGridView.Columns(c).AppearanceCell.BackColor = Color.WhiteSmoke
            '    ElseIf oGridView.Columns(c).Tag = 2 Then
            '        oGridView.Columns(c).AppearanceCell.BackColor = Color.LightSeaGreen
            '    ElseIf oGridView.Columns(c).Tag = 3 Then
            '        oGridView.Columns(c).AppearanceCell.BackColor = Color.LightBlue
            '    End If
            If oGridView.Columns(c).ColumnType = GetType(Decimal) Then
                oGridView.Columns(c).DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
                oGridView.Columns(c).DisplayFormat.FormatString = "###,###,##0.00"
                oGridView.Columns(c).SummaryItem.SetSummary(DevExpress.Data.SummaryItemType.Sum, "{0:n2}")
                oGridView.Columns(c).DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
                oGridView.Columns(c).DisplayFormat.FormatString = "n2"
            End If
        Next

    End Sub

    Private Sub ProcessAllData(dtSource1 As DataTable, dtSource2 As DataTable, dtSource3 As DataTable)
        SplitContainerControl2.Collapsed = True
        If Not dtResult.Columns.Contains("Errores") Then
            dtResult.Columns.Add("Errores", GetType(String)).DefaultValue = Nothing
        End If
        dtResult.Rows.Clear()
        dtDifference.Rows.Clear()
        Dim iPos As Integer = 0
        Dim sPeriod As String = seFiscalYear.EditValue.ToString & Format(seFiscalPeriod.EditValue, "00")
        For r = 0 To dtPresupuesto.Rows.Count - 1
            Dim oRow As DataRow = dtPresupuesto.Rows(r)
            SplashScreenManager.Default.SetWaitFormDescription("Procesando Presupuesto...")
            Try
                SplashScreenManager.Default.SetWaitFormDescription("Procesando Presupuesto (" & (r + 1).ToString & " de " & dtPresupuesto.Rows.Count.ToString & ")")
                If dtResult.Select("PRODUCTO LIKE '%" & oRow("PRODUCTO") & "'").Length > 0 Or oRow("PRODUCTO").ToString.ToUpper = "PRODUCTO" Then
                    Continue For
                End If
                dtResult.Rows.Add()
                iPos = dtResult.Rows.Count - 1
                dtResult.Rows(iPos)("Producto") = oRow("PRODUCTO").ToString.ToUpper
                dtResult.Rows(iPos)("PptoCantidad") = dtPresupuesto.Compute("SUM([TOTAL CANT])", "PRODUCTO LIKE '%" & oRow("PRODUCTO") & "'")
                dtResult.Rows(iPos)("PptoImporte") = dtPresupuesto.Compute("SUM([TOTAL SOLES])", "PRODUCTO LIKE '%" & oRow("PRODUCTO") & "'")
                dtResult.Rows(iPos)("SaldoIniCantidad") = SumFieldByType(oRow, dtInventario, "Cantidad Saldo", "Inventario", False)
                dtResult.Rows(iPos)("SaldoIniImporte") = SumFieldByType(oRow, dtInventario, "Saldo Soles", "Inventario", False)
                dtResult.Rows(iPos)("IngresoCantidad") = SumFieldByType(oRow, dtInventario, "Cantidad Ingreso", "Inventario", False)
                dtResult.Rows(iPos)("IngresoImporte") = SumFieldByType(oRow, dtInventario, "Ingreso Soles", "Inventario", False)
                dtResult.Rows(iPos)("PptoIngresoCantidad") = dtResult.Rows(iPos)("IngresoCantidad") 'SumInputByPeriod(oRow, dtPresupuesto, "Cantidad", "Presupuesto", sPeriod)
                dtResult.Rows(iPos)("PptoIngresoImporte") = dtResult.Rows(iPos)("IngresoImporte") 'SumInputByPeriod(oRow, dtPresupuesto, "Importe", "Presupuesto", sPeriod)
                dtResult.Rows(iPos)("PptoSalidaCantidad") = 0 'SumFieldByPeriod(oRow, dtPresupuesto, "Cantidad", "Presupuesto", sPeriod)
                dtResult.Rows(iPos)("PptoSalidaImporte") = 0 'SumFieldByPeriod(oRow, dtPresupuesto, "Importe", "Presupuesto", sPeriod)
                dtResult.Rows(iPos)("SalidaCantidad") = SumFieldByType(oRow, dtInventario, "Cantidad Salida", "Inventario", False)
                dtResult.Rows(iPos)("SalidaImporte") = SumFieldByType(oRow, dtInventario, "Salida Soles", "Inventario", False)
                dtResult.Rows(iPos)("PptoSaldoCantidad") = dtResult.Rows(iPos)("PptoIngresoCantidad") - dtResult.Rows(iPos)("PptoSalidaCantidad")
                dtResult.Rows(iPos)("PptoSaldoImporte") = dtResult.Rows(iPos)("PptoIngresoImporte") - dtResult.Rows(iPos)("PptoSalidaImporte")
                dtResult.Rows(iPos)("SaldoSistCantidad") = dtResult.Rows(iPos)("SaldoIniCantidad") + dtResult.Rows(iPos)("IngresoCantidad") - dtResult.Rows(iPos)("SalidaCantidad")
                dtResult.Rows(iPos)("SaldoSistImporte") = dtResult.Rows(iPos)("SaldoIniImporte") + dtResult.Rows(iPos)("IngresoImporte") - dtResult.Rows(iPos)("SalidaImporte")
                dtResult.Rows(iPos)("VarPtoConsumoCantidad") = dtResult.Rows(iPos)("PptoCantidad") - dtResult.Rows(iPos)("SalidaCantidad")
                '    dtResult.Rows(iPos)("VarPtoSaldoCantidad") =
                '    dtResult.Rows(iPos)("DifSaldoPptoCalcCantidad") =

            Catch ex As Exception

            End Try
        Next
        For r = 0 To dtInventario.Rows.Count - 1
            Dim oRow As DataRow = dtInventario.Rows(r)
            'SplashScreenManager.Default.SetWaitFormDescription("Procesando Inventario...")
            Try
                SplashScreenManager.Default.SetWaitFormDescription("Procesando Inventario (" & (r + 1).ToString & " de " & dtInventario.Rows.Count.ToString & ")")
                If dtResult.Select("PRODUCTO LIKE '%" & oRow("Descripción").ToString.Trim & "%'").Length > 0 Then
                    Continue For
                End If
                If dtDifference.Select("Producto LIKE '%" & oRow("Descripción").ToString.Trim & "%'").Length > 0 Then
                    Continue For
                End If
                dtDifference.Rows.Add()
                iPos = dtDifference.Rows.Count - 1
                dtDifference.Rows(iPos)("Producto") = oRow("Descripción").ToString.Trim.ToUpper
                dtDifference.Rows(iPos)("IngresoCantidad") = SumFieldByType(oRow, dtInventario, "Cantidad Ingreso", "Diferencia", True) + SumFieldByType(oRow, dtInventario, "Cantidad Saldo", "Diferencia", True)
                dtDifference.Rows(iPos)("IngresoImporte") = SumFieldByType(oRow, dtInventario, "Ingreso Soles", "Diferencia", True) + SumFieldByType(oRow, dtInventario, "Saldo Soles", "Diferencia", True)
                dtDifference.Rows(iPos)("SalidaCantidad") = SumFieldByType(oRow, dtInventario, "Cantidad Salida", "Diferencia", False)
                dtDifference.Rows(iPos)("SalidaImporte") = SumFieldByType(oRow, dtInventario, "Salida Soles", "Diferencia", False)
                dtDifference.Rows(iPos)("SaldoCantidad") = dtDifference.Rows(iPos)("IngresoCantidad") - dtDifference.Rows(iPos)("SalidaCantidad")
                dtDifference.Rows(iPos)("SaldoImporte") = dtDifference.Rows(iPos)("IngresoImporte") - dtDifference.Rows(iPos)("SalidaImporte")
            Catch ex As Exception

            End Try
        Next
        If dtDifference.Rows.Count > 0 Then
            SplitContainerControl2.Collapsed = False
        End If

    End Sub

    Function SumInputByPeriod(drSource As DataRow, dtQuery As DataTable, sField As String, SourceName As String, sPeriod As String) As Double
        Dim dResult As Double = 0
        For r = 0 To dtQuery.Rows.Count - 1
            Dim oRow As DataRow = dtQuery.Rows(r)
            Try
                If SourceName = "Presupuesto" Then
                    If oRow("PRODUCTO").ToString.Trim.ToUpper.Contains(drSource("PRODUCTO").ToString.Trim.ToUpper) Then
                        For c = 0 To dtQuery.Columns.Count - 1
                            If dtQuery.Columns(c).ColumnName.Contains(Mid(sPeriod, 1, 4)) Then
                                If Mid(dtQuery.Columns(c).ColumnName, 1, 6) <= sPeriod Then
                                    If dtQuery.Columns(c).ColumnName.Contains(sField) Then
                                        dResult += CDbl(oRow(dtQuery.Columns(c).ColumnName))
                                    End If
                                    'If sField.Trim.ToUpper.Contains("SALDO") And oRow("Documento").ToString.Trim = "S.INICIAL" Then
                                    '    dResult += CDbl(oRow(sField))
                                    'End If
                                    'If sField.Trim.ToUpper.Contains("SALDO") And oRow("Documento").ToString.Trim <> "S.INICIAL" Then
                                    '        dResult += CDbl(oRow(sField))
                                    '    End If
                                    'End If
                                End If
                            End If
                        Next
                    End If
                End If
            Catch ex As Exception

            End Try
        Next
        Return dResult
    End Function

    Function SumFieldByType(drSource As DataRow, dtQuery As DataTable, sField As String, SourceName As String, OpeningBalance As Boolean) As Double
        Dim dResult As Double = 0
        For r = 0 To dtQuery.Rows.Count - 1
            Dim oRow As DataRow = dtQuery.Rows(r)
            Try
                If SourceName = "Inventario" Then
                    If oRow("Descripción").ToString.Trim.ToUpper.Contains(drSource("PRODUCTO").ToString.Trim.ToUpper) Then
                        If sField.Trim.ToUpper.Contains("SALDO") And oRow("Documento").ToString.Trim = "S.INICIAL" Then
                            dResult += CDbl(oRow(sField))
                        End If
                        If Not sField.Trim.ToUpper.Contains("SALDO") And oRow("Documento").ToString.Trim <> "S.INICIAL" Then
                            dResult += CDbl(oRow(sField))
                        End If
                    End If
                End If
                If SourceName = "Diferencia" Then
                    If IsDBNull(oRow(sField)) Or oRow(sField).ToString.Trim = "" Then
                        oRow(sField) = 0
                    End If
                    If oRow("Descripción").ToString.Trim.ToUpper.Contains(drSource("Descripción").ToString.Trim.ToUpper) Then
                        If sField.Trim.ToUpper.Contains("SALDO") And oRow("Documento").ToString.Trim = "S.INICIAL" Then
                            'If OpeningBalance And oRow("Documento").ToString.Trim = "S.INICIAL" Then
                            dResult += CDbl(oRow(sField))
                        End If
                        If Not sField.Trim.ToUpper.Contains("SALDO") Then
                            'If OpeningBalance = True And oRow("Documento").ToString.Trim = "S.INICIAL" Then
                            '    dResult += CDbl(oRow(sField))
                            'End If
                            If oRow("Documento").ToString.Trim <> "S.INICIAL" Then
                                dResult += CDbl(oRow(sField))
                            End If
                        End If
                    End If
                End If
            Catch ex As Exception

            End Try
        Next
        Return dResult
    End Function

    Private Sub bbiExportar_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiExportar.ItemClick
        If GridView1.IsFocusedView Then
            ExportarExcel(gcInputsControl, Nothing)
        Else
            ExportarExcel(gcValidations, Nothing)
        End If

    End Sub

    Private Sub bbiCerrar_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiCerrar.ItemClick
        Close()
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

        vpInputs.SetValidationRule(Me.bePresupuesto, customValidationRule)
        vpInputs.SetValidationRule(Me.beInventario, customValidationRule)
        vpInputs.SetValidationRule(Me.beControlCalidad, customValidationRule)

    End Sub

    Private Sub GridView1_RowStyle(ByVal sender As Object, ByVal e As DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs) Handles GridView1.RowStyle
        Dim View As GridView = sender
        'If (e.RowHandle >= 0) Then
        '    Dim _Errores As String = View.GetRowCellDisplayText(e.RowHandle, View.Columns("Errores"))

        '    If _Errores <> "" Then
        '        e.Appearance.BackColor = Color.Salmon
        '        e.Appearance.BackColor2 = Color.SeaShell
        '    End If

        'End If
    End Sub

    Private Sub SaleByWebForm_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        GridView1.ActiveFilterEnabled = False
        GridView1.ClearSorting()
        gcInputsControl.MainView.SaveLayoutToRegistry(IO.Directory.GetCurrentDirectory)
    End Sub
End Class