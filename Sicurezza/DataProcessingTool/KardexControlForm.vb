Imports System.IO
Imports DevExpress.Utils.Drawing.Helpers
Imports DevExpress.XtraEditors.DXErrorProvider
Imports DevExpress.XtraGrid.Views.Grid
Imports DevExpress.XtraSplashScreen
'Imports Microsoft.Office.Interop.Excel

Public Class KardexControlForm
    Dim dtSaldoInicial, dtImportaciones, dtSalidas, dtResult As New DataTable
    Private Sub SaleByWebForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        seFiscalYear.Text = Year(Now).ToString
        seFiscalPeriod.Text = Month(DateAdd(DateInterval.Month, -1, Now)).ToString
        GridView1.RestoreLayoutFromRegistry(Directory.GetCurrentDirectory)
        SplitContainerControl2.Collapsed = True
        ResultTableCreate()
    End Sub

    Private Sub ResultTableCreate()
        dtResult.Columns.Add("Tipo", GetType(Integer))
        dtResult.Columns.Add("Fecha", GetType(DateTime))
        dtResult.Columns.Add("Movimiento", GetType(String))
        dtResult.Columns.Add("Categoria", GetType(String))
        dtResult.Columns.Add("CodigoInventario", GetType(String))
        dtResult.Columns.Add("Descripcion", GetType(String))
        dtResult.Columns.Add("Talla", GetType(String))
        dtResult.Columns.Add("Color", GetType(String))
        dtResult.Columns.Add("CantidadIngreso", GetType(Integer))
        dtResult.Columns.Add("CostoUnitarioIngreso", GetType(Decimal))
        dtResult.Columns.Add("TotalIngreso", GetType(Decimal))
        dtResult.Columns.Add("CantidadSalida", GetType(Integer))
        dtResult.Columns.Add("CostoUnitarioSalida", GetType(Decimal))
        dtResult.Columns.Add("TotalSalida", GetType(Decimal))
        dtResult.Columns.Add("SaldoTotal", GetType(Decimal))
        dtResult.Columns.Add("CantidadTotal", GetType(Integer))
        dtResult.Columns.Add("CostoUnitarioCalculado", GetType(Decimal))
        dtResult.Columns.Add("ImporteCalculado", GetType(Decimal))
    End Sub
    Private Sub beImportaciones_Properties_ButtonClick(sender As Object, e As DevExpress.XtraEditors.Controls.ButtonPressedEventArgs) Handles beImportaciones.Properties.ButtonClick
        xofdImportaciones.Filter = "Source Files (*.xls*)|*.xls*"
        xofdImportaciones.FileName = ""
        'XtraOpenFileDialog1.InitialDirectory = ""
        If xofdImportaciones.ShowDialog() = DialogResult.OK Then
            beImportaciones.Text = xofdImportaciones.FileName
        End If
    End Sub

    Private Sub beSalidas_Properties_ButtonClick(sender As Object, e As DevExpress.XtraEditors.Controls.ButtonPressedEventArgs) Handles beSalidas.Properties.ButtonClick
        xofdSalidas.Filter = "Source Files (*.xls*)|*.xls*"
        xofdSalidas.FileName = ""
        'XtraOpenFileDialog1.InitialDirectory = ""
        If xofdSalidas.ShowDialog() = DialogResult.OK Then
            beSalidas.Text = xofdSalidas.FileName
        End If
    End Sub
    Private Sub beResultado_Properties_ButtonClick(sender As Object, e As DevExpress.XtraEditors.Controls.ButtonPressedEventArgs) Handles beResultado.Properties.ButtonClick
        If xfbdResultado.ShowDialog = DialogResult.OK Then
            beResultado.Text = xfbdResultado.SelectedPath & "\Saldo Inicial Inventario " & seFiscalYear.Text & "-" & Format(seFiscalPeriod.EditValue, "00") & ".xlsx"
        End If
    End Sub
    Private Sub beSaldoInicial_Properties_ButtonClick(sender As Object, e As DevExpress.XtraEditors.Controls.ButtonPressedEventArgs) Handles beSaldoInicial.Properties.ButtonClick
        xofdSaldoInicial.Filter = "Source Files (*.xls*)|*.xls*"
        xofdSaldoInicial.FileName = ""
        'XtraOpenFileDialog1.InitialDirectory = ""
        If xofdSaldoInicial.ShowDialog() = DialogResult.OK Then
            beSaldoInicial.Text = xofdSaldoInicial.FileName
        End If
    End Sub
    Private Sub bbiProcesar_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiProcesar.ItemClick
        LoadInputValidations()
        If Not vpInputs.Validate Then
            Return
        End If
        Try
            SplashScreenManager.ShowForm(Me, GetType(WaitForm), True, True, False)
            SplashScreenManager.Default.SetWaitFormDescription("Cargando archivos seleccionados")
            'Saldo Inicial
            dtSaldoInicial = LoadExcel(xofdSaldoInicial.FileName, "{0}").Tables(0)
            'Importaciones
            Dim dtTemp As New DataTable
            dtImportaciones.Rows.Clear()
            For f = 0 To xofdImportaciones.FileNames.Count - 1
                dtTemp = LoadExcel(xofdImportaciones.FileNames(f), "Calculo de Costo Unitario$").Tables(0)
                If dtImportaciones.Rows.Count = 0 Then
                    dtImportaciones = dtTemp.Copy
                Else
                    dtImportaciones.Merge(dtTemp)
                End If
            Next
            'Salidas
            dtSalidas = LoadExcel(xofdSalidas.FileName, "{0}").Tables(0).Select("", "CodigoInt,Fecha de Venta").CopyToDataTable
            SplashScreenManager.Default.SetWaitFormDescription("Procesando datos")
            ProcessAllData()
            dtResult = dtResult.Select("", "CodigoInventario, Tipo, Fecha, CantidadTotal DESC").CopyToDataTable
            gcKardex.DataSource = dtResult
            FormatGridView(GridView1)
            SplashScreenManager.CloseForm(False)
        Catch ex As Exception
            SplashScreenManager.CloseForm(False)
            DevExpress.XtraEditors.XtraMessageBox.Show(Me.LookAndFeel, ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub FormatGridView(oGridView As GridView)
        oGridView.Columns("CodigoInventario").GroupIndex = 0
        oGridView.Columns("CodigoInventario").Caption = "Código de Inventario (SKU)"
        oGridView.ExpandAllGroups()
        oGridView.BestFitColumns()
        For c = 0 To oGridView.Columns.Count - 1
            If oGridView.Columns(c).Tag = 1 Then
                oGridView.Columns(c).AppearanceCell.BackColor = Color.WhiteSmoke
            ElseIf oGridView.Columns(c).Tag = 2 Then
                oGridView.Columns(c).AppearanceCell.BackColor = Color.LightGray
            ElseIf oGridView.Columns(c).Tag = 3 Then
                oGridView.Columns(c).AppearanceCell.BackColor = Color.LightGreen
            End If
        Next
    End Sub

    Private Sub ProcessAllData()
        dtResult.Rows.Clear()
        Dim oRow As DataRow
        Dim iPos As Integer = 0
        Try
            'Saldo Inicial
            For r = 0 To dtSaldoInicial.Rows.Count - 1
                oRow = dtSaldoInicial.Rows(r)
                SplashScreenManager.ShowForm(Me, GetType(WaitForm), True, True, False)
                SplashScreenManager.Default.SetWaitFormDescription("Saldo Inicial (" & (r + 1).ToString & " de " & dtSaldoInicial.Rows.Count.ToString & ")")
                If IsDBNull(oRow(1)) Then
                    Continue For
                End If
                If oRow(1) = "Codigo de Inventario" Then
                    Continue For
                End If
                'Dim iPos As Integer = 0
                dtResult.Rows.Add()
                iPos = dtResult.Rows.Count - 1
                dtResult.Rows(iPos)("Tipo") = 0
                dtResult.Rows(iPos)("Fecha") = seFiscalYear.Text & "-" & seFiscalPeriod.Text & "-01"
                dtResult.Rows(iPos)("Movimiento") = "SALDO INICIAL"
                dtResult.Rows(iPos)("Categoria") = oRow("Categoria")
                dtResult.Rows(iPos)("CodigoInventario") = oRow("Codigo de Inventario")
                dtResult.Rows(iPos)("Descripcion") = oRow("Descripcion")
                dtResult.Rows(iPos)("Talla") = oRow("Talla")
                dtResult.Rows(iPos)("Color") = oRow("Color")
                dtResult.Rows(iPos)("CantidadIngreso") = oRow("Stock Sistema")
                dtResult.Rows(iPos)("CostoUnitarioIngreso") = Math.Round(oRow("Costo Un# Ref# PEN") + oRow("Costo Empaque"), 2)
                dtResult.Rows(iPos)("TotalIngreso") = Math.Round(dtResult.Rows(iPos)("CantidadIngreso") * dtResult.Rows(iPos)("CostoUnitarioIngreso"), 2)
                dtResult.Rows(iPos)("CantidadSalida") = 0
                dtResult.Rows(iPos)("CostoUnitarioSalida") = 0
                dtResult.Rows(iPos)("TotalSalida") = 0
                dtResult.Rows(iPos)("CantidadTotal") = oRow("Stock Sistema")
                dtResult.Rows(iPos)("CostoUnitarioCalculado") = Math.Round(oRow("Costo Un# Ref# PEN") + oRow("Costo Empaque"), 2)
                dtResult.Rows(iPos)("SaldoTotal") = Math.Round(dtResult.Rows(iPos)("CantidadTotal") * dtResult.Rows(iPos)("CostoUnitarioCalculado"), 2)
                dtResult.Rows(iPos)("ImporteCalculado") = 0
            Next
        Catch ex As Exception

        End Try

        Try
            'Importaciones (Ingresos)
            For r = 0 To dtImportaciones.Rows.Count - 1
                oRow = dtImportaciones.Rows(r)
                SplashScreenManager.ShowForm(Me, GetType(WaitForm), True, True, False)
                SplashScreenManager.Default.SetWaitFormDescription("Importaciones (" & (r + 1).ToString & " de " & dtImportaciones.Rows.Count.ToString & ")")
                If IsDBNull(oRow(0)) Then
                    Continue For
                End If

                dtResult.Rows.Add()
                iPos = dtResult.Rows.Count - 1
                dtResult.Rows(iPos)("Tipo") = 1
                dtResult.Rows(iPos)("Fecha") = oRow("Fecha")
                dtResult.Rows(iPos)("Movimiento") = oRow("Número correlativo de importación")
                dtResult.Rows(iPos)("Categoria") = "" 'oRow("Categoria")
                dtResult.Rows(iPos)("CodigoInventario") = oRow("SKU")
                dtResult.Rows(iPos)("Descripcion") = oRow("Descripción de producto")
                dtResult.Rows(iPos)("Talla") = oRow("Talla")
                dtResult.Rows(iPos)("Color") = oRow("Color")
                dtResult.Rows(iPos)("CantidadIngreso") = oRow("Cantidad")
                dtResult.Rows(iPos)("CostoUnitarioIngreso") = Math.Round(oRow("Costo Unitario"), 2)
                dtResult.Rows(iPos)("TotalIngreso") = Math.Round(dtResult.Rows(iPos)("CantidadIngreso") * dtResult.Rows(iPos)("CostoUnitarioIngreso"), 2)
                dtResult.Rows(iPos)("CantidadSalida") = 0
                dtResult.Rows(iPos)("CostoUnitarioSalida") = 0
                dtResult.Rows(iPos)("TotalSalida") = 0
                dtResult.Rows(iPos)("CantidadTotal") = 0
                If dtResult.Select("CodigoInventario='" & oRow("SKU") & "'").Length > 0 Then
                    dtResult.Rows(iPos)("CantidadTotal") = dtResult.Compute("SUM(CantidadTotal)", "CodigoInventario='" & oRow("SKU") & "'") + oRow("Cantidad")
                End If
                dtResult.Rows(iPos)("SaldoTotal") = 0
                If dtResult.Select("CodigoInventario='" & oRow("SKU") & "'").Length > 0 Then
                    dtResult.Rows(iPos)("SaldoTotal") = Math.Round(dtResult.Compute("SUM(SaldoTotal)", "CodigoInventario='" & oRow("SKU") & "'") + dtResult.Rows(iPos)("TotalIngreso"), 2)
                End If
                dtResult.Rows(iPos)("CostoUnitarioCalculado") = Math.Round(dtResult.Rows(iPos)("SaldoTotal") / dtResult.Rows(iPos)("CantidadTotal"), 2)
                dtResult.Rows(iPos)("ImporteCalculado") = 0
            Next
        Catch ex As Exception

        End Try
        Try
            'Ventas (Salidas)
            For r = 0 To dtSalidas.Rows.Count - 1
                oRow = dtSalidas.Rows(r)
                SplashScreenManager.ShowForm(Me, GetType(WaitForm), True, True, False)
                SplashScreenManager.Default.SetWaitFormDescription("Ventas (" & (r + 1).ToString & " de " & dtSalidas.Rows.Count.ToString & ")")
                If IsDBNull(oRow(0)) Then
                    Continue For
                End If
                If Format(oRow("Fecha de Venta"), "yyyyMM") <> seFiscalYear.Text & Strings.Left("0" & seFiscalPeriod.Text, 2) Then
                    Continue For
                End If
                'Dim iPos As Integer = 0
                Dim dCostoUnitarioCalculado As Decimal = 0
                Dim dSaldoTotal As Decimal = 0
                Dim iSaldoCantidadTotal As Integer = 0
                Dim iCantidadSalida As Integer = 0
                If dtResult.Select("CodigoInventario='" & oRow("CodigoInt") & "' AND Tipo<>2").Length > 0 Then
                    dCostoUnitarioCalculado = dtResult.Compute("MAX(CostoUnitarioCalculado)", "CodigoInventario='" & oRow("CodigoInt") & "' AND Tipo<>2")
                End If
                If dtResult.Select("CodigoInventario='" & oRow("CodigoInt") & "' AND Tipo=2").Length > 0 Then
                    iCantidadSalida = dtResult.Compute("SUM(CantidadSalida)", "CodigoInventario='" & oRow("CodigoInt") & "' AND Tipo=2")
                End If
                If dtResult.Select("CodigoInventario='" & oRow("CodigoInt") & "' AND Tipo<>2").Length > 0 Then
                    iSaldoCantidadTotal = dtResult.Compute("MAX(CantidadTotal)", "CodigoInventario='" & oRow("CodigoInt") & "' AND Tipo<>2")
                End If
                If dtResult.Select("CodigoInventario='" & oRow("CodigoInt") & "' AND Tipo<>2").Length > 0 Then
                    dSaldoTotal = dtResult.Compute("MAX(SaldoTotal)", "CodigoInventario='" & oRow("CodigoInt") & "' AND Tipo<>2")
                End If
                dtResult.Rows.Add()
                iPos = dtResult.Rows.Count - 1
                dtResult.Rows(iPos)("Tipo") = 2
                dtResult.Rows(iPos)("Fecha") = oRow("Fecha de Venta")
                dtResult.Rows(iPos)("Movimiento") = oRow("Serie") & "-" & oRow("Correlativo")
                dtResult.Rows(iPos)("Categoria") = oRow("Categoria")
                dtResult.Rows(iPos)("CodigoInventario") = oRow("CodigoInt")
                dtResult.Rows(iPos)("Descripcion") = oRow("Nombre")
                dtResult.Rows(iPos)("Talla") = oRow("Talla")
                dtResult.Rows(iPos)("Color") = oRow("Color")
                dtResult.Rows(iPos)("CantidadIngreso") = 0
                dtResult.Rows(iPos)("CostoUnitarioIngreso") = 0
                dtResult.Rows(iPos)("TotalIngreso") = 0
                dtResult.Rows(iPos)("CantidadSalida") = oRow("Cantidad")
                dtResult.Rows(iPos)("CostoUnitarioSalida") = dCostoUnitarioCalculado
                dtResult.Rows(iPos)("TotalSalida") = dtResult.Rows(iPos)("CantidadSalida") * dtResult.Rows(iPos)("CostoUnitarioSalida")
                dtResult.Rows(iPos)("CantidadTotal") = iSaldoCantidadTotal - iCantidadSalida - oRow("Cantidad")
                dtResult.Rows(iPos)("SaldoTotal") = Math.Round(dSaldoTotal - dtResult.Compute("SUM(TotalSalida)", "CodigoInventario='" & oRow("CodigoInt") & "' AND Tipo=2"), 2)
                dtResult.Rows(iPos)("CostoUnitarioCalculado") = dCostoUnitarioCalculado
                dtResult.Rows(iPos)("ImporteCalculado") = 0
            Next
        Catch ex As Exception

        End Try
    End Sub

    Private Sub bbiExportar_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiExportar.ItemClick
        If GridView1.IsFocusedView Then
            ExportarExcel(gcKardex, beResultado.Text)
            'Else
            '    ExportarExcel(gcValidations)
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

        vpInputs.SetValidationRule(Me.beSaldoInicial, customValidationRule)
        vpInputs.SetValidationRule(Me.beImportaciones, customValidationRule)
        vpInputs.SetValidationRule(Me.beSalidas, customValidationRule)
        vpInputs.SetValidationRule(Me.beResultado, customValidationRule)
    End Sub

    'Private Sub GridView1_RowStyle(ByVal sender As Object, ByVal e As DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs) Handles GridView1.RowStyle
    '    Dim View As GridView = sender
    '    If (e.RowHandle >= 0) Then
    '        Dim _Errores As String = View.GetRowCellDisplayText(e.RowHandle, View.Columns("Errores"))

    '        If _Errores <> "" Then
    '            e.Appearance.BackColor = Color.Salmon
    '            e.Appearance.BackColor2 = Color.SeaShell
    '        End If

    '    End If
    'End Sub

    Private Sub SaleByWebForm_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        GridView1.ActiveFilterEnabled = False
        GridView1.ClearSorting()
        gcKardex.MainView.SaveLayoutToRegistry(IO.Directory.GetCurrentDirectory)
    End Sub

    Private Sub beResultado_ButtonClick(sender As Object, e As DevExpress.XtraEditors.Controls.ButtonPressedEventArgs) Handles beResultado.ButtonClick

    End Sub
End Class