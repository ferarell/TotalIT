Imports System.IO
Imports DevExpress.XtraEditors.DXErrorProvider
Imports DevExpress.XtraGrid.Views.Grid
Imports DevExpress.XtraSplashScreen

Public Class SaleByWebForm
    Private Sub beArchivoOrigen_Properties_ButtonClick(sender As Object, e As DevExpress.XtraEditors.Controls.ButtonPressedEventArgs) Handles beArchivoOrigen.Properties.ButtonClick
        'XtraOpenFileDialog1.Filter = "Source Files (*.xls*;*.csv)|*.xls*;*.csv"
        XtraOpenFileDialog1.Filter = "Source Files (*.xls*)|*.xls*"
        XtraOpenFileDialog1.FileName = ""
        'XtraOpenFileDialog1.InitialDirectory = ""
        If XtraOpenFileDialog1.ShowDialog() = DialogResult.OK Then
            beArchivoOrigen.Text = XtraOpenFileDialog1.FileName
        End If
    End Sub

    Private Sub SaleByWebForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        GridView1.RestoreLayoutFromRegistry(Directory.GetCurrentDirectory)
        SplitContainerControl2.Collapsed = True
    End Sub

    Private Sub bbiProcesar_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiProcesar.ItemClick
        LoadInputValidations()
        If Not vpInputs.Validate Then
            Return
        End If
        Try
            SplashScreenManager.ShowForm(Me, GetType(WaitForm), True, True, False)
            SplashScreenManager.Default.SetWaitFormDescription("Cargando archivos seleccionados")
            Dim dtSource1, dtSource2, dtSource3 As New DataTable
            For f = 0 To XtraOpenFileDialog1.FileNames.Count - 1
                If XtraOpenFileDialog1.FileNames(f).ToString.ToUpper.Contains("CSV") Then
                    dtSource1 = LoadCSV(XtraOpenFileDialog1.FileNames(f), True, ",")
                End If
                If XtraOpenFileDialog1.FileNames(f).ToString.ToUpper.Contains("XLS") Then
                    Dim dtTemp As DataTable = LoadExcel(XtraOpenFileDialog1.FileNames(f), "{0}").Tables(0)
                    dtTemp.Rows(0)(0) = IIf(IsDBNull(dtTemp.Rows(2)(0)), "", dtTemp.Rows(2)(0))
                    If dtTemp.Columns(0).ColumnName = "Name" Then
                        dtSource1 = LoadExcel(XtraOpenFileDialog1.FileNames(f), "{0}").Tables(0).Select("", "Name").CopyToDataTable
                    ElseIf dtTemp.Columns.Contains("Fecha de Venta") And dtTemp.Columns.Contains("CodigoInt") Then
                        dtSource2 = LoadExcel(XtraOpenFileDialog1.FileNames(f), "{0}").Tables(0)
                    Else
                        dtSource3 = LoadExcel(XtraOpenFileDialog1.FileNames(f), "{0}").Tables(0)
                    End If
                End If
            Next
            If dtSource1.Rows.Count = 0 Then
                DevExpress.XtraEditors.XtraMessageBox.Show("No se cargó el archivo de Shopify, verifique el formato e intente nuevamente.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If
            If dtSource2.Rows.Count = 0 Then
                DevExpress.XtraEditors.XtraMessageBox.Show("No se cargó el archivo de PayU, verifique el formato e intente nuevamente.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If
            If dtSource3.Rows.Count = 0 Then
                DevExpress.XtraEditors.XtraMessageBox.Show("No se cargó el archivo de Contabilidad, verifique el formato e intente nuevamente.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If
            SplashScreenManager.Default.SetWaitFormDescription("Procesando datos")
            ProcessAllData(dtSource1, dtSource2, dtSource3)
            gcSaleByWeb.DataSource = dtSource1
            FormatGridView(GridView1)
            SplashScreenManager.CloseForm(False)
        Catch ex As Exception
            SplashScreenManager.CloseForm(False)
            DevExpress.XtraEditors.XtraMessageBox.Show(Me.LookAndFeel, ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub FormatGridView(oGridView As GridView)
        GridView1.BestFitColumns()
        For c = 0 To oGridView.Columns.Count - 1
            If oGridView.Columns(c).Tag = 1 Then
                oGridView.Columns(c).AppearanceCell.BackColor = Color.WhiteSmoke
            ElseIf oGridView.Columns(c).Tag = 2 Then
                oGridView.Columns(c).AppearanceCell.BackColor = Color.LightSeaGreen
            ElseIf oGridView.Columns(c).Tag = 3 Then
                oGridView.Columns(c).AppearanceCell.BackColor = Color.LightBlue
            End If
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
        If Not dtSource1.Columns.Contains("Errores") Then
            dtSource1.Columns.Add("Errores", GetType(String)).DefaultValue = Nothing
            dtSource1.Columns.Add("ValorVenta", GetType(Decimal)).DefaultValue = 0
            dtSource1.Columns.Add("ValorImpuesto", GetType(Decimal)).DefaultValue = 0
        End If
        'PAYU (dtSource3)
        If Not dtSource1.Columns.Contains("PayU") Then
            dtSource1.Columns.Add("PayU Fecha", GetType(DateTime)).DefaultValue = Nothing
            dtSource1.Columns.Add("PayU Descripcion", GetType(String)).DefaultValue = Nothing
            dtSource1.Columns.Add("PayU Importe Venta", GetType(Decimal)).DefaultValue = 0
            dtSource1.Columns.Add("PayU Comision", GetType(Decimal)).DefaultValue = 0
            dtSource1.Columns.Add("PayU IGV Comision", GetType(Decimal)).DefaultValue = 0
            dtSource1.Columns.Add("PayU ComiCalc", GetType(Decimal)).DefaultValue = 0
            dtSource1.Columns.Add("PayU CargoFijo", GetType(Decimal)).DefaultValue = 0
        End If
        For r = 0 To dtSource1.Rows.Count - 1
            Dim oRow As DataRow = dtSource1.Rows(r)
            Dim dtSourceTmp3 As New DataTable
            Dim SalePayU As String = Nothing
            oRow("Errores") = ""
            SplashScreenManager.Default.SetWaitFormDescription("Validando Shopify - PayU")
            Try
                If Not IsDBNull(oRow("Name")) Then
                    oRow("ValorVenta") = Math.Round((oRow("Subtotal") / 1.18), 2)
                    oRow("ValorImpuesto") = Math.Round((oRow("Subtotal") * 0.18), 2)
                End If
                oRow("Payment Reference") = IIf(IsDBNull(oRow("Payment Reference")), "", oRow("Payment Reference"))
                If oRow("Payment Reference") = "" Then
                    Continue For
                End If
                If Not IsNumeric(ExtractOnlyNumbers(oRow("Payment Reference"))) Or Not oRow("Payment Reference").ToString.Contains(".") Then
                    Continue For
                End If
                SalePayU = Mid(oRow("Payment Reference"), 2, InStr(oRow("Payment Reference"), ".") - 2)
                    If dtSource3.Select("DESCRIPCION LIKE '%" & SalePayU & "%'").Length = 0 Then
                        If Not IsDBNull(oRow("Financial Status")) Then
                            oRow("Errores") += IIf(oRow("Errores") = "", "", " | ") & "No se encontró la relación Shopify-PayU"
                        End If
                        Continue For
                    End If
                    dtSourceTmp3 = dtSource3.Select("DESCRIPCION LIKE '%" & SalePayU & "%'").CopyToDataTable
                For p = 0 To dtSourceTmp3.Rows.Count - 1
                    Dim oRow3 As DataRow = dtSourceTmp3.Rows(p)
                    'For c = 0 To dtSource3.Columns.Count - 1
                    '    Dim ColName As String = dtSource3.Columns(c).ColumnName
                    '    oRow("PAYU_" & ColName) = oRow3(ColName)
                    'Next

                    oRow("PayU Fecha") = oRow3("FECHA")
                    If oRow3("DESCRIPCION").ToString.StartsWith("SALES") Then
                        oRow("PayU Importe Venta") = oRow3("CREDITOS")
                        oRow("PayU ComiCalc") = Math.Round(oRow3("CREDITOS") * (My.Settings.ComPayU / 100), 2)
                        oRow("PayU CargoFijo") = My.Settings.CargoFijoPayU
                    End If
                    If oRow3("DESCRIPCION").ToString.StartsWith("POL_COMMISSION") Then
                        oRow("PayU Comision") = oRow3("DEBITOS")
                    End If
                    If oRow3("DESCRIPCION").ToString.StartsWith("TAX_POL_COMMISSION") Then
                        oRow("PayU IGV Comision") = oRow3("DEBITOS")
                    End If
                Next
                'VALIDACIONES

                'If Not IsDBNull(oRow("Financial Status")) And IsDBNull(oRow("PayU Fecha")) Then
                '    oRow("Errores") += IIf(oRow("Errores") = "", "", " | ") & "No se encontró la relación Shopify-PayU"
                'End If

            Catch ex As Exception

            End Try
        Next

        'CONTABILIDAD (dtSource2)
        If Not dtSource1.Columns.Contains("FechaVenta") Then
            For c = 0 To dtSource2.Columns.Count - 1
                dtSource1.Columns.Add(dtSource2.Columns(c).ColumnName, dtSource2.Columns(c).DataType)
            Next
            dtSource1.Columns.Add("SubTotal Contable", GetType(Decimal)).DefaultValue = 0
            dtSource1.Columns.Add("Total Contable", GetType(Decimal)).DefaultValue = 0
            dtSource1.Columns.Add("Flete", GetType(Decimal)).DefaultValue = 0
            dtSource1.Columns.Add("Embajadora", GetType(String)).DefaultValue = Nothing
        End If
        SplashScreenManager.Default.SetWaitFormDescription("Validando Shopify - Contabilidad")
        For r = 0 To dtSource1.Rows.Count - 1
            Dim oRow As DataRow = dtSource1.Rows(r)
            Dim dtSourceTmp2 As New DataTable
            Dim PedidoShopify, PedidoConta As String
            Try
                oRow("Name") = ExtractOnlyNumbers(oRow("Name"))
                PedidoShopify = oRow("Name")
                If dtSource2.Select("Descripcion LIKE '%" & PedidoShopify & "%'").Length = 0 Then
                    If Not IsDBNull(oRow("Financial Status")) Then
                        oRow("Errores") += IIf(oRow("Errores") = "", "", " | ") & "No se encontró la relación Shopify-Contabilidad"
                    End If
                    Continue For
                End If
                dtSourceTmp2 = dtSource2.Select("Descripcion LIKE '%" & PedidoShopify & "%'").CopyToDataTable
                For p = 0 To dtSourceTmp2.Rows.Count - 1
                    Dim oRow2 As DataRow = dtSourceTmp2.Rows(p)
                    oRow2("Descripcion") = IIf(IsDBNull(oRow2("Descripcion")), "", oRow2("Descripcion"))
                    If oRow2("Descripcion") = "" Then
                        Continue For
                    End If
                    PedidoConta = ExtractOnlyNumbers(oRow2("Descripcion"))
                    If PedidoConta.Length = 0 Then
                        Continue For
                    End If
                    'oRow("Fecha de Venta") = oRow2("Fecha de Venta")
                    'oRow("Anulado?") = oRow2("Anulado?")
                    'oRow("Sunat Tipo Comp#") = oRow2("Sunat Tipo Comp#")
                    'oRow("Serie") = oRow2("Serie")
                    'oRow("Correlativo") = oRow2("Correlativo")
                    'oRow("Sunat Tipo Doc#") = oRow2("Sunat Tipo Doc#")
                    'oRow("RUC/Documento") = oRow2("RUC/Documento")
                    'oRow("Nombre Cliente") = oRow2("Nombre Cliente")
                    'oRow("ValorVentaTotal") = 0 'oRow2("Fecha de Venta")
                    For c = 0 To dtSourceTmp2.Columns.Count - 1
                        Dim ColName As String = dtSourceTmp2.Columns(c).ColumnName
                        If oRow("Lineitem sku").ToString = oRow2("CodigoInt").ToString Then
                            oRow(ColName) = oRow2(ColName)
                        End If
                        If Not IsDBNull(oRow("Financial Status")) Then
                            If oRow2("Categoria").ToString.Contains("FLETE") Then
                                oRow("Flete") = oRow2("Precio Vta Tot")
                            End If
                            oRow("SubTotal Contable") = dtSourceTmp2.Compute("SUM([Valor Vta Tot])", "")
                            oRow("Total Contable") = dtSourceTmp2.Compute("SUM([Precio Vta Tot])", "")
                            If oRow2("Descripcion").ToString.Contains("-") Then
                                oRow("Embajadora") = Mid(oRow2("Descripcion"), InStr(oRow2("Descripcion"), "-") + 1, Len(oRow2("Descripcion"))).Trim
                            End If
                        End If
                    Next
                Next
                ' VALIDACIONES
                If Not IsDBNull(oRow("Anulado?")) Then
                    Continue For
                End If
                'oRow("Errores") = ""
                oRow("Shipping") = IIf(IsDBNull(oRow("Shipping")), 0, oRow("Shipping"))
                oRow("Flete") = IIf(IsDBNull(oRow("Flete")), 0, oRow("Flete"))
                oRow("Total") = IIf(IsDBNull(oRow("Total")), 0, oRow("Total"))
                oRow("Lineitem quantity") = IIf(IsDBNull(oRow("Lineitem quantity")), 0, oRow("Lineitem quantity"))
                oRow("Cantidad") = IIf(IsDBNull(oRow("Cantidad")), 0, oRow("Cantidad"))
                oRow("SubTotal Contable") = IIf(IsDBNull(oRow("SubTotal Contable")), 0, oRow("SubTotal Contable"))
                oRow("Total Contable") = IIf(IsDBNull(oRow("Total Contable")), 0, oRow("Total Contable"))

                If IsDBNull(oRow("Fecha de Venta")) And oRow("Total Contable") > 0 Then
                    oRow("Errores") += IIf(oRow("Errores") = "", "", " | ") & "No se encontró la relación SKU (Shopify-Contabilidad)"
                    Continue For
                End If
                If oRow("Shipping") <> 0 And oRow("Shipping") <> oRow("Flete") Then
                    oRow("Errores") += "Existe una diferencia en el importe del shipping (Shopify-Contabilidad)"
                End If
                If oRow("Total") <> 0 Then
                    If oRow("ValorVenta") <> oRow("SubTotal Contable") Then
                        oRow("Errores") += IIf(oRow("Errores") = "", "", " | ") & "Existe una diferencia en el valor de venta (Shopify-Contabilidad)"
                    End If
                    If oRow("Total") <> oRow("Total Contable") Then
                        oRow("Errores") += IIf(oRow("Errores") = "", "", " | ") & "Existe una diferencia en el precio de venta (Shopify-Contabilidad)"
                    End If
                End If
                If Not IsDBNull(oRow("Financial Status")) Then
                    If oRow("Discount Code").ToString <> oRow("Embajadora").ToString Then
                        oRow("Errores") += IIf(oRow("Errores") = "", "", " | ") & "Existe una diferencia en Embajadora (Shopify-Contabilidad)"
                    End If
                    If oRow("Lineitem quantity") <> 0 And oRow("Lineitem quantity") <> oRow("Cantidad") Then
                        oRow("Errores") += IIf(oRow("Errores") = "", "", " | ") & "Existe una diferencia en la cantidad (Shopify-Contabilidad)"
                    End If
                End If
            Catch ex As Exception

            End Try
        Next

        Dim dtValidations As New DataTable
        Try

            dtValidations = dtSource1.Clone
        If Not dtValidations.Columns.Contains("Origen") Then
            dtValidations.Columns.Add("Origen", GetType(String)).DefaultValue = Nothing
        End If
        Dim iPos As Integer = 0
        'Verifica si existen movimientos en PayU, pero que no tengan relación con Shopify
        SplashScreenManager.Default.SetWaitFormDescription("Validando PayU - Shopify")
        For r = 0 To dtSource3.Rows.Count - 1
            Dim oRow3 As DataRow = dtSource3.Rows(r)
                Dim ReferenciaPayU As String = ""
                Try
                    If IsDBNull(oRow3(0)) Then
                        Continue For
                    End If
                    If Not oRow3("DESCRIPCION").ToString.StartsWith("SALES") Then
                        Continue For
                    End If
                    ReferenciaPayU = ExtractOnlyNumbers(oRow3("DESCRIPCION"))
                    If dtSource1.Select("[Payment Reference] LIKE '%" & ReferenciaPayU & "%'").Length = 0 Then
                        dtValidations.Rows.Add()
                        iPos = dtValidations.Rows.Count - 1
                        dtValidations.Rows(iPos)("Origen") = "PayU"
                        dtValidations.Rows(iPos)("Errores") = "No se encontró la referencia de pago en Shopify"
                        dtValidations.Rows(iPos)("PayU Fecha") = oRow3("FECHA")
                        dtValidations.Rows(iPos)("PayU Descripcion") = ReferenciaPayU
                    End If
                Catch ex As Exception

                End Try
            Next

            'Verifica si existen movimientos en PayU, pero que no tengan relación con Shopify
            SplashScreenManager.Default.SetWaitFormDescription("Validando Contabilidad - Shopify")
            For r = 0 To dtSource2.Rows.Count - 1
                Dim oRow2 As DataRow = dtSource2.Rows(r)
                If IsDBNull(oRow2(0)) Then
                    Continue For
                End If
                Dim PedidoContable As String = ExtractOnlyNumbers(oRow2("Descripcion"))
                Try
                    If dtSource1.Select("Name LIKE '%" & PedidoContable & "%'").Length = 0 Then
                        dtValidations.Rows.Add()
                        iPos = dtValidations.Rows.Count - 1
                        dtValidations.Rows(iPos)("Origen") = "Contabilidad"
                        dtValidations.Rows(iPos)("Errores") = "No se encontró el movimiento contable en Shopify"
                        dtValidations.Rows(iPos)("Fecha de Venta") = oRow2("Fecha de Venta")
                        dtValidations.Rows(iPos)("Anulado?") = oRow2("Anulado?")
                        dtValidations.Rows(iPos)("Serie") = oRow2("Serie")
                        dtValidations.Rows(iPos)("Correlativo") = oRow2("Correlativo")
                        dtValidations.Rows(iPos)("Sunat Tipo Doc#") = oRow2("Sunat Tipo Doc#")
                        dtValidations.Rows(iPos)("RUC/Documento") = oRow2("RUC/Documento")
                        dtValidations.Rows(iPos)("Nombre Cliente") = oRow2("Nombre Cliente")
                        dtValidations.Rows(iPos)("Descripcion") = oRow2("Descripcion")
                        dtValidations.Rows(iPos)("CodigoInt") = oRow2("CodigoInt")
                    End If

                Catch ex As Exception

                End Try

            Next

        Catch ex As Exception

        End Try
        If dtValidations.Rows.Count > 0 Then
            SplitContainerControl2.Collapsed = False
        End If
        gcValidations.DataSource = dtValidations
        GridView2.BestFitColumns()
    End Sub

    Private Sub bbiExportar_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiExportar.ItemClick
        If GridView1.IsFocusedView Then
            ExportarExcel(gcSaleByWeb, Nothing)
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

        vpInputs.SetValidationRule(Me.beArchivoOrigen, customValidationRule)
    End Sub

    Private Sub GridView1_RowStyle(ByVal sender As Object, ByVal e As DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs) Handles GridView1.RowStyle
        Dim View As GridView = sender
        If (e.RowHandle >= 0) Then
            Dim _Errores As String = View.GetRowCellDisplayText(e.RowHandle, View.Columns("Errores"))

            If _Errores <> "" Then
                e.Appearance.BackColor = Color.Salmon
                e.Appearance.BackColor2 = Color.SeaShell
            End If

        End If
    End Sub

    Private Sub SaleByWebForm_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        GridView1.ActiveFilterEnabled = False
        GridView1.ClearSorting()
        gcSaleByWeb.MainView.SaveLayoutToRegistry(IO.Directory.GetCurrentDirectory)
    End Sub
End Class