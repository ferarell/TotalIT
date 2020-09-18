Imports System.IO
Imports DevExpress.XtraEditors.DXErrorProvider
Imports DevExpress.XtraSplashScreen

Public Class SaleByWebForm
    Private Sub beArchivoOrigen_Properties_ButtonClick(sender As Object, e As DevExpress.XtraEditors.Controls.ButtonPressedEventArgs) Handles beArchivoOrigen.Properties.ButtonClick
        XtraOpenFileDialog1.Filter = "Source Files (*.xls*;*.csv)|*.xls*;*.csv"
        XtraOpenFileDialog1.FileName = ""
        'XtraOpenFileDialog1.InitialDirectory = ""
        If XtraOpenFileDialog1.ShowDialog() = DialogResult.OK Then
            beArchivoOrigen.Text = XtraOpenFileDialog1.FileName
        End If
    End Sub

    Private Sub SaleByWebForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        GridView1.RestoreLayoutFromRegistry(Directory.GetCurrentDirectory)
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
                    If dtTemp.Columns(0).ColumnName = "Fecha de Venta" Then
                        dtSource2 = LoadExcel(XtraOpenFileDialog1.FileNames(f), "{0}").Tables(0)
                    Else
                        dtSource3 = LoadExcel(XtraOpenFileDialog1.FileNames(f), "{0}").Tables(0)
                    End If
                End If
            Next
            SplashScreenManager.Default.SetWaitFormDescription("Procesando datos")
            ProcessAllData(dtSource1, dtSource2, dtSource3)
            SplashScreenManager.CloseForm(False)
            gcSaleByWeb.DataSource = dtSource1
        Catch ex As Exception
            SplashScreenManager.CloseForm(False)
            DevExpress.XtraEditors.XtraMessageBox.Show(Me.LookAndFeel, ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub ProcessAllData(dtSource1 As DataTable, dtSource2 As DataTable, dtSource3 As DataTable)
        'PAYU (dtSource3)
        If Not dtSource1.Columns.Contains("PayU") Then
            dtSource1.Columns.Add("PayU Fecha", GetType(DateTime)).DefaultValue = Nothing
            dtSource1.Columns.Add("PayU Importe Venta", GetType(Decimal)).DefaultValue = 0
            dtSource1.Columns.Add("PayU Comision", GetType(Decimal)).DefaultValue = 0
            dtSource1.Columns.Add("PayU IGV Comision", GetType(Decimal)).DefaultValue = 0
        End If
        For r = 0 To dtSource1.Rows.Count - 1
            Dim oRow As DataRow = dtSource1.Rows(r)
            Dim dtSourceTmp3 As New DataTable
            Dim SalePayU As String = Nothing
            Try
                oRow("Vendor") = IIf(IsDBNull(oRow("Vendor")), "", oRow("Vendor"))
                If oRow("Vendor") = "" Then
                    Continue For
                End If
                If Not IsNumeric(ExtractOnlyNumbers(oRow("Vendor"))) Or Not oRow("Vendor").ToString.Contains(".") Then
                    Continue For
                End If
                SalePayU = Mid(oRow("Vendor"), 2, InStr(oRow("Vendor"), ".") - 2)
                If dtSource3.Select("DESCRIPCION LIKE '%" & SalePayU & "%'").Length = 0 Then
                    Continue For
                End If
                dtSourceTmp3 = dtSource3.Select("DESCRIPCION LIKE '%" & SalePayU & "%'").CopyToDataTable
                For p = 0 To dtSourceTmp3.Rows.Count - 1
                    Dim oRow3 As DataRow = dtSourceTmp3.Rows(p)
                    oRow("PayU Fecha") = oRow3("FECHA")
                    If oRow3("DESCRIPCION").ToString.Contains("SALES") Then
                        oRow("PayU Importe Venta") = oRow3("CREDITOS")
                    End If
                    If oRow3("DESCRIPCION").ToString.Contains("POL_COMMISSION") Then
                        oRow("PayU Comision") = oRow3("DEBITOS")
                    End If
                    If oRow3("DESCRIPCION").ToString.Contains("TAX_POL_COMMISSION") Then
                        oRow("PayU IGV Comision") = oRow3("DEBITOS")
                    End If
                Next
            Catch ex As Exception

            End Try
        Next
        'CONTABILIDAD (dtSource2)
        If Not dtSource1.Columns.Contains("FechaVenta") Then
            dtSource1.Columns.Add("FechaVenta", GetType(DateTime)).DefaultValue = Nothing
            dtSource1.Columns.Add("TipoComprobante", GetType(String)).DefaultValue = ""
            dtSource1.Columns.Add("SerieComprobante", GetType(String)).DefaultValue = ""
            dtSource1.Columns.Add("NumeroComprobante", GetType(String)).DefaultValue = ""
            dtSource1.Columns.Add("TipoDocumento", GetType(String)).DefaultValue = ""
            dtSource1.Columns.Add("NumeroDocumento", GetType(String)).DefaultValue = ""
            dtSource1.Columns.Add("NombreCliente", GetType(String)).DefaultValue = ""
            dtSource1.Columns.Add("ValorVentaTotal", GetType(Decimal)).DefaultValue = 0
        End If
        For r = 0 To dtSource1.Rows.Count - 1
            Dim oRow As DataRow = dtSource1.Rows(r)
            Dim dtSourceTmp2 As New DataTable
            Dim PedidoShopify, PedidoConta As String
            Try
                oRow("Name") = ExtractOnlyNumbers(oRow("Name"))
                PedidoShopify = oRow("Name")
                If dtSource2.Select("Descripcion LIKE '%" & PedidoShopify & "%'").Length = 0 Then
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
                    oRow("FechaVenta") = oRow2("Fecha de Venta")
                    oRow("TipoComprobante") = oRow2("Sunat Tipo Comp#")
                    oRow("SerieComprobante") = oRow2("Serie")
                    oRow("NumeroComprobante") = oRow2("Correlativo")
                    oRow("TipoDocumento") = oRow2("Sunat Tipo Doc#")
                    oRow("NumeroDocumento") = oRow2("RUC/Documento")
                    oRow("NombreCliente") = oRow2("Nombre Cliente")
                    oRow("ValorVentaTotal") = 0 'oRow2("Fecha de Venta")
                Next

            Catch ex As Exception

            End Try
        Next
    End Sub

    Private Sub bbiExportar_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiExportar.ItemClick
        ExportarExcel(gcSaleByWeb)
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

End Class