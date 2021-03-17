<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class KardexControlForm
    Inherits DevExpress.XtraEditors.XtraForm

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(KardexControlForm))
        Me.gcKardex = New DevExpress.XtraGrid.GridControl()
        Me.GridView1 = New DevExpress.XtraGrid.Views.Grid.GridView()
        Me.GridColumn1 = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.GridColumn18 = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.GridColumn2 = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.GridColumn3 = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.GridColumn4 = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.GridColumn5 = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.GridColumn6 = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.GridColumn7 = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.GridColumn8 = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.GridColumn9 = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.GridColumn10 = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.GridColumn11 = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.GridColumn12 = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.GridColumn13 = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.GridColumn15 = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.GridColumn16 = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.GridColumn14 = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.GridColumn17 = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.SplitContainerControl1 = New DevExpress.XtraEditors.SplitContainerControl()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.seFiscalPeriod = New DevExpress.XtraEditors.SpinEdit()
        Me.seFiscalYear = New DevExpress.XtraEditors.SpinEdit()
        Me.beResultado = New DevExpress.XtraEditors.ButtonEdit()
        Me.beSalidas = New DevExpress.XtraEditors.ButtonEdit()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.beImportaciones = New DevExpress.XtraEditors.ButtonEdit()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.beSaldoInicial = New DevExpress.XtraEditors.ButtonEdit()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.SplitContainerControl2 = New DevExpress.XtraEditors.SplitContainerControl()
        Me.vpInputs = New DevExpress.XtraEditors.DXErrorProvider.DXValidationProvider(Me.components)
        Me.bmActions = New DevExpress.XtraBars.BarManager(Me.components)
        Me.bar5 = New DevExpress.XtraBars.Bar()
        Me.brsDescripcion = New DevExpress.XtraBars.BarStaticItem()
        Me.brBarraAcciones = New DevExpress.XtraBars.Bar()
        Me.bbiProcesar = New DevExpress.XtraBars.BarButtonItem()
        Me.bbiBuscar = New DevExpress.XtraBars.BarButtonItem()
        Me.bbiGuardar = New DevExpress.XtraBars.BarButtonItem()
        Me.bbiExportar = New DevExpress.XtraBars.BarButtonItem()
        Me.bbiCerrar = New DevExpress.XtraBars.BarButtonItem()
        Me.BarDockControl1 = New DevExpress.XtraBars.BarDockControl()
        Me.BarDockControl2 = New DevExpress.XtraBars.BarDockControl()
        Me.BarDockControl3 = New DevExpress.XtraBars.BarDockControl()
        Me.BarDockControl4 = New DevExpress.XtraBars.BarDockControl()
        Me.barStaticItem3 = New DevExpress.XtraBars.BarStaticItem()
        Me.barStaticItem4 = New DevExpress.XtraBars.BarStaticItem()
        Me.brsEstado = New DevExpress.XtraBars.BarStaticItem()
        Me.beiProceso = New DevExpress.XtraBars.BarEditItem()
        Me.rpiProceso = New DevExpress.XtraEditors.Repository.RepositoryItemProgressBar()
        Me.BarButtonItem1 = New DevExpress.XtraBars.BarButtonItem()
        Me.bsiVistas = New DevExpress.XtraBars.BarSubItem()
        Me.bbiVistaGrilla = New DevExpress.XtraBars.BarButtonItem()
        Me.bbiTarjeta = New DevExpress.XtraBars.BarButtonItem()
        Me.bbiContrato = New DevExpress.XtraBars.BarButtonItem()
        Me.bbiCronograma = New DevExpress.XtraBars.BarButtonItem()
        Me.bbiCartaNotarial = New DevExpress.XtraBars.BarButtonItem()
        Me.bbiLetras = New DevExpress.XtraBars.BarButtonItem()
        Me.BarButtonItem3 = New DevExpress.XtraBars.BarButtonItem()
        Me.BarButtonItem4 = New DevExpress.XtraBars.BarButtonItem()
        Me.RepositoryItemLookUpEdit1 = New DevExpress.XtraEditors.Repository.RepositoryItemLookUpEdit()
        Me.RepositoryItemImageComboBox1 = New DevExpress.XtraEditors.Repository.RepositoryItemImageComboBox()
        Me.RepositoryItemRadioGroup1 = New DevExpress.XtraEditors.Repository.RepositoryItemRadioGroup()
        Me.RepositoryItemComboBox1 = New DevExpress.XtraEditors.Repository.RepositoryItemComboBox()
        Me.xofdSaldoInicial = New DevExpress.XtraEditors.XtraOpenFileDialog(Me.components)
        Me.xofdImportaciones = New DevExpress.XtraEditors.XtraOpenFileDialog(Me.components)
        Me.xofdSalidas = New DevExpress.XtraEditors.XtraOpenFileDialog(Me.components)
        Me.xfbdResultado = New DevExpress.XtraEditors.XtraFolderBrowserDialog(Me.components)
        Me.GridControl1 = New DevExpress.XtraGrid.GridControl()
        Me.GridView2 = New DevExpress.XtraGrid.Views.Grid.GridView()
        Me.GridColumn19 = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.GridColumn20 = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.GridColumn21 = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.GridColumn22 = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.GridColumn23 = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.GridColumn24 = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.GridColumn25 = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.GridColumn26 = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.GridColumn27 = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.GridColumn28 = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.GridColumn29 = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.GridColumn30 = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.GridColumn31 = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.GridColumn32 = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.GridColumn33 = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.GridColumn34 = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.GridColumn35 = New DevExpress.XtraGrid.Columns.GridColumn()
        CType(Me.gcKardex, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.GridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SplitContainerControl1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainerControl1.SuspendLayout()
        CType(Me.seFiscalPeriod.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.seFiscalYear.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.beResultado.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.beSalidas.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.beImportaciones.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.beSaldoInicial.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SplitContainerControl2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainerControl2.SuspendLayout()
        CType(Me.vpInputs, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.bmActions, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.rpiProceso, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RepositoryItemLookUpEdit1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RepositoryItemImageComboBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RepositoryItemRadioGroup1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RepositoryItemComboBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.GridControl1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.GridView2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'gcKardex
        '
        Me.gcKardex.Dock = System.Windows.Forms.DockStyle.Fill
        Me.gcKardex.Location = New System.Drawing.Point(0, 0)
        Me.gcKardex.MainView = Me.GridView1
        Me.gcKardex.Name = "gcKardex"
        Me.gcKardex.Size = New System.Drawing.Size(890, 133)
        Me.gcKardex.TabIndex = 0
        Me.gcKardex.UseEmbeddedNavigator = True
        Me.gcKardex.ViewCollection.AddRange(New DevExpress.XtraGrid.Views.Base.BaseView() {Me.GridView1})
        '
        'GridView1
        '
        Me.GridView1.Columns.AddRange(New DevExpress.XtraGrid.Columns.GridColumn() {Me.GridColumn1, Me.GridColumn18, Me.GridColumn2, Me.GridColumn3, Me.GridColumn4, Me.GridColumn5, Me.GridColumn6, Me.GridColumn7, Me.GridColumn8, Me.GridColumn9, Me.GridColumn10, Me.GridColumn11, Me.GridColumn12, Me.GridColumn13, Me.GridColumn15, Me.GridColumn16, Me.GridColumn14, Me.GridColumn17})
        Me.GridView1.GridControl = Me.gcKardex
        Me.GridView1.GroupCount = 1
        Me.GridView1.Name = "GridView1"
        Me.GridView1.OptionsView.ColumnAutoWidth = False
        Me.GridView1.OptionsView.ShowAutoFilterRow = True
        Me.GridView1.OptionsView.ShowFooter = True
        Me.GridView1.SortInfo.AddRange(New DevExpress.XtraGrid.Columns.GridColumnSortInfo() {New DevExpress.XtraGrid.Columns.GridColumnSortInfo(Me.GridColumn3, DevExpress.Data.ColumnSortOrder.Ascending)})
        '
        'GridColumn1
        '
        Me.GridColumn1.Caption = "Fecha"
        Me.GridColumn1.FieldName = "Fecha"
        Me.GridColumn1.Name = "GridColumn1"
        Me.GridColumn1.Visible = True
        Me.GridColumn1.VisibleIndex = 0
        '
        'GridColumn18
        '
        Me.GridColumn18.Caption = "Movimiento"
        Me.GridColumn18.FieldName = "Movimiento"
        Me.GridColumn18.Name = "GridColumn18"
        Me.GridColumn18.Visible = True
        Me.GridColumn18.VisibleIndex = 1
        '
        'GridColumn2
        '
        Me.GridColumn2.Caption = "Categoria"
        Me.GridColumn2.FieldName = "Categoria"
        Me.GridColumn2.Name = "GridColumn2"
        Me.GridColumn2.Visible = True
        Me.GridColumn2.VisibleIndex = 2
        '
        'GridColumn3
        '
        Me.GridColumn3.Caption = "Código de Inventario (SKU)"
        Me.GridColumn3.FieldName = "CodigoInventario"
        Me.GridColumn3.Name = "GridColumn3"
        Me.GridColumn3.Visible = True
        Me.GridColumn3.VisibleIndex = 3
        '
        'GridColumn4
        '
        Me.GridColumn4.Caption = "Descripción"
        Me.GridColumn4.FieldName = "Descripcion"
        Me.GridColumn4.Name = "GridColumn4"
        Me.GridColumn4.Visible = True
        Me.GridColumn4.VisibleIndex = 3
        '
        'GridColumn5
        '
        Me.GridColumn5.Caption = "Talla"
        Me.GridColumn5.FieldName = "Talla"
        Me.GridColumn5.Name = "GridColumn5"
        Me.GridColumn5.Visible = True
        Me.GridColumn5.VisibleIndex = 4
        '
        'GridColumn6
        '
        Me.GridColumn6.Caption = "Color"
        Me.GridColumn6.FieldName = "Color"
        Me.GridColumn6.Name = "GridColumn6"
        Me.GridColumn6.Visible = True
        Me.GridColumn6.VisibleIndex = 5
        '
        'GridColumn7
        '
        Me.GridColumn7.Caption = "Stock"
        Me.GridColumn7.FieldName = "Stock"
        Me.GridColumn7.Name = "GridColumn7"
        '
        'GridColumn8
        '
        Me.GridColumn8.Caption = "Cantidad Ingreso"
        Me.GridColumn8.FieldName = "CantidadIngreso"
        Me.GridColumn8.Name = "GridColumn8"
        Me.GridColumn8.Tag = "1"
        Me.GridColumn8.Visible = True
        Me.GridColumn8.VisibleIndex = 6
        '
        'GridColumn9
        '
        Me.GridColumn9.Caption = "Costo Unitario Ingreso"
        Me.GridColumn9.FieldName = "CostoUnitarioIngreso"
        Me.GridColumn9.Name = "GridColumn9"
        Me.GridColumn9.Tag = "1"
        Me.GridColumn9.Visible = True
        Me.GridColumn9.VisibleIndex = 7
        '
        'GridColumn10
        '
        Me.GridColumn10.Caption = "Total Ingreso"
        Me.GridColumn10.FieldName = "TotalIngreso"
        Me.GridColumn10.Name = "GridColumn10"
        Me.GridColumn10.Tag = "1"
        Me.GridColumn10.Visible = True
        Me.GridColumn10.VisibleIndex = 8
        '
        'GridColumn11
        '
        Me.GridColumn11.Caption = "Cantidad Salida"
        Me.GridColumn11.FieldName = "CantidadSalida"
        Me.GridColumn11.Name = "GridColumn11"
        Me.GridColumn11.Tag = "2"
        Me.GridColumn11.Visible = True
        Me.GridColumn11.VisibleIndex = 9
        '
        'GridColumn12
        '
        Me.GridColumn12.Caption = "Costo Unitario Salida"
        Me.GridColumn12.FieldName = "CostoUnitarioSalida"
        Me.GridColumn12.Name = "GridColumn12"
        Me.GridColumn12.Tag = "2"
        Me.GridColumn12.Visible = True
        Me.GridColumn12.VisibleIndex = 10
        '
        'GridColumn13
        '
        Me.GridColumn13.Caption = "Total Salida"
        Me.GridColumn13.FieldName = "TotalSalida"
        Me.GridColumn13.Name = "GridColumn13"
        Me.GridColumn13.Tag = "2"
        Me.GridColumn13.Visible = True
        Me.GridColumn13.VisibleIndex = 11
        '
        'GridColumn15
        '
        Me.GridColumn15.Caption = "Cantidad Total"
        Me.GridColumn15.FieldName = "CantidadTotal"
        Me.GridColumn15.Name = "GridColumn15"
        Me.GridColumn15.Tag = "3"
        Me.GridColumn15.Visible = True
        Me.GridColumn15.VisibleIndex = 12
        '
        'GridColumn16
        '
        Me.GridColumn16.Caption = "Costo Unitario Calculado"
        Me.GridColumn16.FieldName = "CostoUnitarioCalculado"
        Me.GridColumn16.Name = "GridColumn16"
        Me.GridColumn16.Tag = "3"
        Me.GridColumn16.Visible = True
        Me.GridColumn16.VisibleIndex = 13
        '
        'GridColumn14
        '
        Me.GridColumn14.Caption = "Saldo Total"
        Me.GridColumn14.FieldName = "SaldoTotal"
        Me.GridColumn14.Name = "GridColumn14"
        Me.GridColumn14.Tag = "3"
        Me.GridColumn14.Visible = True
        Me.GridColumn14.VisibleIndex = 14
        '
        'GridColumn17
        '
        Me.GridColumn17.Caption = "Importe Calculado"
        Me.GridColumn17.FieldName = "ImporteCalculado"
        Me.GridColumn17.Name = "GridColumn17"
        '
        'SplitContainerControl1
        '
        Me.SplitContainerControl1.CollapsePanel = DevExpress.XtraEditors.SplitCollapsePanel.Panel1
        Me.SplitContainerControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainerControl1.Horizontal = False
        Me.SplitContainerControl1.Location = New System.Drawing.Point(0, 47)
        Me.SplitContainerControl1.Name = "SplitContainerControl1"
        Me.SplitContainerControl1.Panel1.Controls.Add(Me.Label3)
        Me.SplitContainerControl1.Panel1.Controls.Add(Me.Label4)
        Me.SplitContainerControl1.Panel1.Controls.Add(Me.seFiscalPeriod)
        Me.SplitContainerControl1.Panel1.Controls.Add(Me.seFiscalYear)
        Me.SplitContainerControl1.Panel1.Controls.Add(Me.beResultado)
        Me.SplitContainerControl1.Panel1.Controls.Add(Me.beSalidas)
        Me.SplitContainerControl1.Panel1.Controls.Add(Me.Label6)
        Me.SplitContainerControl1.Panel1.Controls.Add(Me.beImportaciones)
        Me.SplitContainerControl1.Panel1.Controls.Add(Me.Label2)
        Me.SplitContainerControl1.Panel1.Controls.Add(Me.Label1)
        Me.SplitContainerControl1.Panel1.Controls.Add(Me.beSaldoInicial)
        Me.SplitContainerControl1.Panel1.Controls.Add(Me.Label5)
        Me.SplitContainerControl1.Panel1.Text = "Panel1"
        Me.SplitContainerControl1.Panel2.Controls.Add(Me.SplitContainerControl2)
        Me.SplitContainerControl1.Panel2.Text = "Panel2"
        Me.SplitContainerControl1.Size = New System.Drawing.Size(890, 386)
        Me.SplitContainerControl1.SplitterPosition = 163
        Me.SplitContainerControl1.TabIndex = 1
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(66, 16)
        Me.Label3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(46, 13)
        Me.Label3.TabIndex = 33
        Me.Label3.Text = "Ejercicio"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(69, 40)
        Me.Label4.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(43, 13)
        Me.Label4.TabIndex = 34
        Me.Label4.Text = "Periodo"
        '
        'seFiscalPeriod
        '
        Me.seFiscalPeriod.EditValue = New Decimal(New Integer() {1, 0, 0, 0})
        Me.seFiscalPeriod.Location = New System.Drawing.Point(118, 37)
        Me.seFiscalPeriod.Margin = New System.Windows.Forms.Padding(2)
        Me.seFiscalPeriod.Name = "seFiscalPeriod"
        Me.seFiscalPeriod.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.seFiscalPeriod.Properties.Mask.EditMask = "f0"
        Me.seFiscalPeriod.Properties.MaxValue = New Decimal(New Integer() {13, 0, 0, 0})
        Me.seFiscalPeriod.Size = New System.Drawing.Size(40, 20)
        Me.seFiscalPeriod.TabIndex = 1
        '
        'seFiscalYear
        '
        Me.seFiscalYear.EditValue = New Decimal(New Integer() {1981, 0, 0, 0})
        Me.seFiscalYear.Location = New System.Drawing.Point(118, 13)
        Me.seFiscalYear.Margin = New System.Windows.Forms.Padding(2)
        Me.seFiscalYear.Name = "seFiscalYear"
        Me.seFiscalYear.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.seFiscalYear.Properties.Mask.EditMask = "f0"
        Me.seFiscalYear.Properties.MaxValue = New Decimal(New Integer() {2099, 0, 0, 0})
        Me.seFiscalYear.Properties.MinValue = New Decimal(New Integer() {1981, 0, 0, 0})
        Me.seFiscalYear.Size = New System.Drawing.Size(54, 20)
        Me.seFiscalYear.TabIndex = 0
        '
        'beResultado
        '
        Me.beResultado.Location = New System.Drawing.Point(118, 133)
        Me.beResultado.Margin = New System.Windows.Forms.Padding(2)
        Me.beResultado.Name = "beResultado"
        Me.beResultado.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.beResultado.Size = New System.Drawing.Size(641, 20)
        Me.beResultado.TabIndex = 5
        '
        'beSalidas
        '
        Me.beSalidas.Location = New System.Drawing.Point(118, 109)
        Me.beSalidas.Margin = New System.Windows.Forms.Padding(2)
        Me.beSalidas.Name = "beSalidas"
        Me.beSalidas.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.beSalidas.Size = New System.Drawing.Size(641, 20)
        Me.beSalidas.TabIndex = 4
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(59, 136)
        Me.Label6.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(55, 13)
        Me.Label6.TabIndex = 28
        Me.Label6.Text = "Resultado"
        '
        'beImportaciones
        '
        Me.beImportaciones.Location = New System.Drawing.Point(118, 85)
        Me.beImportaciones.Margin = New System.Windows.Forms.Padding(2)
        Me.beImportaciones.Name = "beImportaciones"
        Me.beImportaciones.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.beImportaciones.Size = New System.Drawing.Size(641, 20)
        Me.beImportaciones.TabIndex = 3
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(74, 112)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(40, 13)
        Me.Label2.TabIndex = 28
        Me.Label2.Text = "Salidas"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(39, 88)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(75, 13)
        Me.Label1.TabIndex = 28
        Me.Label1.Text = "Importaciones"
        '
        'beSaldoInicial
        '
        Me.beSaldoInicial.Location = New System.Drawing.Point(118, 61)
        Me.beSaldoInicial.Margin = New System.Windows.Forms.Padding(2)
        Me.beSaldoInicial.Name = "beSaldoInicial"
        Me.beSaldoInicial.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.beSaldoInicial.Size = New System.Drawing.Size(641, 20)
        Me.beSaldoInicial.TabIndex = 2
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(51, 64)
        Me.Label5.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(63, 13)
        Me.Label5.TabIndex = 28
        Me.Label5.Text = "Saldo Inicial"
        '
        'SplitContainerControl2
        '
        Me.SplitContainerControl2.CollapsePanel = DevExpress.XtraEditors.SplitCollapsePanel.Panel2
        Me.SplitContainerControl2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainerControl2.Horizontal = False
        Me.SplitContainerControl2.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainerControl2.Name = "SplitContainerControl2"
        Me.SplitContainerControl2.Panel1.Controls.Add(Me.gcKardex)
        Me.SplitContainerControl2.Panel1.Text = "Panel1"
        Me.SplitContainerControl2.Panel2.Controls.Add(Me.GridControl1)
        Me.SplitContainerControl2.Panel2.Text = "Panel2"
        Me.SplitContainerControl2.Size = New System.Drawing.Size(890, 218)
        Me.SplitContainerControl2.SplitterPosition = 133
        Me.SplitContainerControl2.TabIndex = 1
        '
        'vpInputs
        '
        Me.vpInputs.ValidationMode = DevExpress.XtraEditors.DXErrorProvider.ValidationMode.Manual
        '
        'bmActions
        '
        Me.bmActions.Bars.AddRange(New DevExpress.XtraBars.Bar() {Me.bar5, Me.brBarraAcciones})
        Me.bmActions.DockControls.Add(Me.BarDockControl1)
        Me.bmActions.DockControls.Add(Me.BarDockControl2)
        Me.bmActions.DockControls.Add(Me.BarDockControl3)
        Me.bmActions.DockControls.Add(Me.BarDockControl4)
        Me.bmActions.Form = Me
        Me.bmActions.Items.AddRange(New DevExpress.XtraBars.BarItem() {Me.brsDescripcion, Me.barStaticItem3, Me.barStaticItem4, Me.brsEstado, Me.bbiProcesar, Me.bbiCerrar, Me.beiProceso, Me.BarButtonItem1, Me.bsiVistas, Me.bbiVistaGrilla, Me.bbiTarjeta, Me.bbiContrato, Me.bbiCronograma, Me.bbiCartaNotarial, Me.bbiLetras, Me.bbiExportar, Me.BarButtonItem3, Me.BarButtonItem4, Me.bbiGuardar, Me.bbiBuscar})
        Me.bmActions.MaxItemId = 30
        Me.bmActions.RepositoryItems.AddRange(New DevExpress.XtraEditors.Repository.RepositoryItem() {Me.rpiProceso, Me.RepositoryItemLookUpEdit1, Me.RepositoryItemImageComboBox1, Me.RepositoryItemRadioGroup1, Me.RepositoryItemComboBox1})
        '
        'bar5
        '
        Me.bar5.BarName = "Custom 3"
        Me.bar5.CanDockStyle = DevExpress.XtraBars.BarCanDockStyle.Bottom
        Me.bar5.DockCol = 0
        Me.bar5.DockRow = 0
        Me.bar5.DockStyle = DevExpress.XtraBars.BarDockStyle.Bottom
        Me.bar5.LinksPersistInfo.AddRange(New DevExpress.XtraBars.LinkPersistInfo() {New DevExpress.XtraBars.LinkPersistInfo(Me.brsDescripcion)})
        Me.bar5.OptionsBar.AllowQuickCustomization = False
        Me.bar5.OptionsBar.DrawDragBorder = False
        Me.bar5.OptionsBar.MultiLine = True
        Me.bar5.OptionsBar.UseWholeRow = True
        Me.bar5.Text = "Custom 3"
        '
        'brsDescripcion
        '
        Me.brsDescripcion.Id = 30
        Me.brsDescripcion.Name = "brsDescripcion"
        '
        'brBarraAcciones
        '
        Me.brBarraAcciones.BarName = "Custom 5"
        Me.brBarraAcciones.DockCol = 0
        Me.brBarraAcciones.DockRow = 0
        Me.brBarraAcciones.DockStyle = DevExpress.XtraBars.BarDockStyle.Top
        Me.brBarraAcciones.FloatLocation = New System.Drawing.Point(279, 188)
        Me.brBarraAcciones.LinksPersistInfo.AddRange(New DevExpress.XtraBars.LinkPersistInfo() {New DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.PaintStyle, Me.bbiProcesar, DevExpress.XtraBars.BarItemPaintStyle.CaptionGlyph), New DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.PaintStyle, Me.bbiBuscar, "", True, True, True, 0, Nothing, DevExpress.XtraBars.BarItemPaintStyle.CaptionGlyph), New DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.PaintStyle, Me.bbiGuardar, "", True, True, True, 0, Nothing, DevExpress.XtraBars.BarItemPaintStyle.CaptionGlyph), New DevExpress.XtraBars.LinkPersistInfo(Me.bbiExportar, True), New DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.PaintStyle, Me.bbiCerrar, "", True, True, True, 0, Nothing, DevExpress.XtraBars.BarItemPaintStyle.CaptionGlyph)})
        Me.brBarraAcciones.OptionsBar.AllowQuickCustomization = False
        Me.brBarraAcciones.OptionsBar.UseWholeRow = True
        Me.brBarraAcciones.Text = "Custom 5"
        '
        'bbiProcesar
        '
        Me.bbiProcesar.Caption = "&Procesar"
        Me.bbiProcesar.Id = 33
        Me.bbiProcesar.ImageOptions.Image = CType(resources.GetObject("bbiProcesar.ImageOptions.Image"), System.Drawing.Image)
        Me.bbiProcesar.ImageOptions.ImageIndex = 26
        Me.bbiProcesar.ImageOptions.LargeImageIndex = 7
        Me.bbiProcesar.Name = "bbiProcesar"
        '
        'bbiBuscar
        '
        Me.bbiBuscar.Caption = "&Buscar"
        Me.bbiBuscar.Id = 29
        Me.bbiBuscar.ImageOptions.Image = CType(resources.GetObject("bbiBuscar.ImageOptions.Image"), System.Drawing.Image)
        Me.bbiBuscar.Name = "bbiBuscar"
        Me.bbiBuscar.Visibility = DevExpress.XtraBars.BarItemVisibility.Never
        '
        'bbiGuardar
        '
        Me.bbiGuardar.Caption = "&Guardar"
        Me.bbiGuardar.Id = 27
        Me.bbiGuardar.ImageOptions.Image = CType(resources.GetObject("bbiGuardar.ImageOptions.Image"), System.Drawing.Image)
        Me.bbiGuardar.Name = "bbiGuardar"
        Me.bbiGuardar.Visibility = DevExpress.XtraBars.BarItemVisibility.Never
        '
        'bbiExportar
        '
        Me.bbiExportar.Caption = "&Exportar"
        Me.bbiExportar.Id = 21
        Me.bbiExportar.ImageOptions.Image = CType(resources.GetObject("bbiExportar.ImageOptions.Image"), System.Drawing.Image)
        Me.bbiExportar.ImageOptions.ImageIndex = 19
        Me.bbiExportar.Name = "bbiExportar"
        Me.bbiExportar.PaintStyle = DevExpress.XtraBars.BarItemPaintStyle.CaptionGlyph
        '
        'bbiCerrar
        '
        Me.bbiCerrar.Caption = "&Cerrar"
        Me.bbiCerrar.Id = 41
        Me.bbiCerrar.ImageOptions.Image = CType(resources.GetObject("bbiCerrar.ImageOptions.Image"), System.Drawing.Image)
        Me.bbiCerrar.ImageOptions.ImageIndex = 27
        Me.bbiCerrar.ImageOptions.LargeImageIndex = 0
        Me.bbiCerrar.Name = "bbiCerrar"
        Me.bbiCerrar.ShortcutKeyDisplayString = "Alt+C"
        '
        'BarDockControl1
        '
        Me.BarDockControl1.CausesValidation = False
        Me.BarDockControl1.Dock = System.Windows.Forms.DockStyle.Top
        Me.BarDockControl1.Location = New System.Drawing.Point(0, 0)
        Me.BarDockControl1.Manager = Me.bmActions
        Me.BarDockControl1.Size = New System.Drawing.Size(890, 47)
        '
        'BarDockControl2
        '
        Me.BarDockControl2.CausesValidation = False
        Me.BarDockControl2.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.BarDockControl2.Location = New System.Drawing.Point(0, 433)
        Me.BarDockControl2.Manager = Me.bmActions
        Me.BarDockControl2.Size = New System.Drawing.Size(890, 29)
        '
        'BarDockControl3
        '
        Me.BarDockControl3.CausesValidation = False
        Me.BarDockControl3.Dock = System.Windows.Forms.DockStyle.Left
        Me.BarDockControl3.Location = New System.Drawing.Point(0, 47)
        Me.BarDockControl3.Manager = Me.bmActions
        Me.BarDockControl3.Size = New System.Drawing.Size(0, 386)
        '
        'BarDockControl4
        '
        Me.BarDockControl4.CausesValidation = False
        Me.BarDockControl4.Dock = System.Windows.Forms.DockStyle.Right
        Me.BarDockControl4.Location = New System.Drawing.Point(890, 47)
        Me.BarDockControl4.Manager = Me.bmActions
        Me.BarDockControl4.Size = New System.Drawing.Size(0, 386)
        '
        'barStaticItem3
        '
        Me.barStaticItem3.Caption = "0 / 0"
        Me.barStaticItem3.Id = 31
        Me.barStaticItem3.Name = "barStaticItem3"
        '
        'barStaticItem4
        '
        Me.barStaticItem4.Caption = "Estado"
        Me.barStaticItem4.Id = 46
        Me.barStaticItem4.Name = "barStaticItem4"
        '
        'brsEstado
        '
        Me.brsEstado.Caption = "Lectura"
        Me.brsEstado.Id = 47
        Me.brsEstado.Name = "brsEstado"
        '
        'beiProceso
        '
        Me.beiProceso.Alignment = DevExpress.XtraBars.BarItemLinkAlignment.Right
        Me.beiProceso.Edit = Me.rpiProceso
        Me.beiProceso.EditWidth = 150
        Me.beiProceso.Id = 0
        Me.beiProceso.Name = "beiProceso"
        '
        'rpiProceso
        '
        Me.rpiProceso.Name = "rpiProceso"
        Me.rpiProceso.ShowTitle = True
        '
        'BarButtonItem1
        '
        Me.BarButtonItem1.Caption = "Vista"
        Me.BarButtonItem1.Id = 3
        Me.BarButtonItem1.Name = "BarButtonItem1"
        '
        'bsiVistas
        '
        Me.bsiVistas.Caption = "Vistas"
        Me.bsiVistas.Id = 6
        Me.bsiVistas.ImageOptions.ImageIndex = 20
        Me.bsiVistas.LinksPersistInfo.AddRange(New DevExpress.XtraBars.LinkPersistInfo() {New DevExpress.XtraBars.LinkPersistInfo(Me.bbiVistaGrilla), New DevExpress.XtraBars.LinkPersistInfo(Me.bbiTarjeta)})
        Me.bsiVistas.Name = "bsiVistas"
        '
        'bbiVistaGrilla
        '
        Me.bbiVistaGrilla.Caption = "Grilla"
        Me.bbiVistaGrilla.Id = 7
        Me.bbiVistaGrilla.ImageOptions.ImageIndex = 23
        Me.bbiVistaGrilla.Name = "bbiVistaGrilla"
        '
        'bbiTarjeta
        '
        Me.bbiTarjeta.Caption = "Tarjeta"
        Me.bbiTarjeta.Id = 8
        Me.bbiTarjeta.ImageOptions.ImageIndex = 21
        Me.bbiTarjeta.Name = "bbiTarjeta"
        '
        'bbiContrato
        '
        Me.bbiContrato.Caption = "Contrato Compra-Venta"
        Me.bbiContrato.Id = 13
        Me.bbiContrato.Name = "bbiContrato"
        '
        'bbiCronograma
        '
        Me.bbiCronograma.Caption = "Cronograma de Pagos"
        Me.bbiCronograma.Id = 14
        Me.bbiCronograma.Name = "bbiCronograma"
        '
        'bbiCartaNotarial
        '
        Me.bbiCartaNotarial.Caption = "Carta Notarial"
        Me.bbiCartaNotarial.Id = 15
        Me.bbiCartaNotarial.Name = "bbiCartaNotarial"
        '
        'bbiLetras
        '
        Me.bbiLetras.Caption = "Letras de Cambio"
        Me.bbiLetras.Id = 16
        Me.bbiLetras.Name = "bbiLetras"
        '
        'BarButtonItem3
        '
        Me.BarButtonItem3.Caption = "5.1 Movimiento Contable"
        Me.BarButtonItem3.Id = 25
        Me.BarButtonItem3.Name = "BarButtonItem3"
        Me.BarButtonItem3.Tag = "1"
        '
        'BarButtonItem4
        '
        Me.BarButtonItem4.Caption = "5.3 Plan Contable"
        Me.BarButtonItem4.Id = 26
        Me.BarButtonItem4.Name = "BarButtonItem4"
        Me.BarButtonItem4.Tag = "2"
        '
        'RepositoryItemLookUpEdit1
        '
        Me.RepositoryItemLookUpEdit1.AutoHeight = False
        Me.RepositoryItemLookUpEdit1.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.RepositoryItemLookUpEdit1.Name = "RepositoryItemLookUpEdit1"
        '
        'RepositoryItemImageComboBox1
        '
        Me.RepositoryItemImageComboBox1.AutoHeight = False
        Me.RepositoryItemImageComboBox1.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo), New DevExpress.XtraEditors.Controls.EditorButton(), New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.RepositoryItemImageComboBox1.Name = "RepositoryItemImageComboBox1"
        '
        'RepositoryItemRadioGroup1
        '
        Me.RepositoryItemRadioGroup1.Items.AddRange(New DevExpress.XtraEditors.Controls.RadioGroupItem() {New DevExpress.XtraEditors.Controls.RadioGroupItem(Nothing, "Nacionales"), New DevExpress.XtraEditors.Controls.RadioGroupItem(Nothing, "No Domiciliados")})
        Me.RepositoryItemRadioGroup1.Name = "RepositoryItemRadioGroup1"
        '
        'RepositoryItemComboBox1
        '
        Me.RepositoryItemComboBox1.AutoHeight = False
        Me.RepositoryItemComboBox1.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.RepositoryItemComboBox1.Items.AddRange(New Object() {"Nacionales", "No Domiciliados"})
        Me.RepositoryItemComboBox1.Name = "RepositoryItemComboBox1"
        '
        'xofdImportaciones
        '
        Me.xofdImportaciones.Multiselect = True
        '
        'GridControl1
        '
        Me.GridControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GridControl1.Location = New System.Drawing.Point(0, 0)
        Me.GridControl1.MainView = Me.GridView2
        Me.GridControl1.Name = "GridControl1"
        Me.GridControl1.Size = New System.Drawing.Size(890, 80)
        Me.GridControl1.TabIndex = 1
        Me.GridControl1.UseEmbeddedNavigator = True
        Me.GridControl1.ViewCollection.AddRange(New DevExpress.XtraGrid.Views.Base.BaseView() {Me.GridView2})
        '
        'GridView2
        '
        Me.GridView2.Columns.AddRange(New DevExpress.XtraGrid.Columns.GridColumn() {Me.GridColumn19, Me.GridColumn20, Me.GridColumn21, Me.GridColumn22, Me.GridColumn23, Me.GridColumn24, Me.GridColumn25, Me.GridColumn26, Me.GridColumn27, Me.GridColumn28, Me.GridColumn29, Me.GridColumn30, Me.GridColumn31, Me.GridColumn32, Me.GridColumn33, Me.GridColumn34, Me.GridColumn35})
        Me.GridView2.GridControl = Me.GridControl1
        Me.GridView2.GroupCount = 1
        Me.GridView2.Name = "GridView2"
        Me.GridView2.OptionsView.ColumnAutoWidth = False
        Me.GridView2.OptionsView.ShowAutoFilterRow = True
        Me.GridView2.OptionsView.ShowFooter = True
        Me.GridView2.OptionsView.ShowGroupPanel = False
        Me.GridView2.SortInfo.AddRange(New DevExpress.XtraGrid.Columns.GridColumnSortInfo() {New DevExpress.XtraGrid.Columns.GridColumnSortInfo(Me.GridColumn22, DevExpress.Data.ColumnSortOrder.Ascending)})
        '
        'GridColumn19
        '
        Me.GridColumn19.Caption = "Fecha"
        Me.GridColumn19.FieldName = "Fecha"
        Me.GridColumn19.Name = "GridColumn19"
        Me.GridColumn19.Visible = True
        Me.GridColumn19.VisibleIndex = 0
        '
        'GridColumn20
        '
        Me.GridColumn20.Caption = "Movimiento"
        Me.GridColumn20.FieldName = "Movimiento"
        Me.GridColumn20.Name = "GridColumn20"
        Me.GridColumn20.Visible = True
        Me.GridColumn20.VisibleIndex = 1
        '
        'GridColumn21
        '
        Me.GridColumn21.Caption = "Categoria"
        Me.GridColumn21.FieldName = "Categoria"
        Me.GridColumn21.Name = "GridColumn21"
        Me.GridColumn21.Visible = True
        Me.GridColumn21.VisibleIndex = 2
        '
        'GridColumn22
        '
        Me.GridColumn22.Caption = "Código de Inventario (SKU)"
        Me.GridColumn22.FieldName = "CodigoInventario"
        Me.GridColumn22.Name = "GridColumn22"
        Me.GridColumn22.Visible = True
        Me.GridColumn22.VisibleIndex = 3
        '
        'GridColumn23
        '
        Me.GridColumn23.Caption = "Descripción"
        Me.GridColumn23.FieldName = "Descripcion"
        Me.GridColumn23.Name = "GridColumn23"
        Me.GridColumn23.Visible = True
        Me.GridColumn23.VisibleIndex = 3
        '
        'GridColumn24
        '
        Me.GridColumn24.Caption = "Talla"
        Me.GridColumn24.FieldName = "Talla"
        Me.GridColumn24.Name = "GridColumn24"
        Me.GridColumn24.Visible = True
        Me.GridColumn24.VisibleIndex = 4
        '
        'GridColumn25
        '
        Me.GridColumn25.Caption = "Color"
        Me.GridColumn25.FieldName = "Color"
        Me.GridColumn25.Name = "GridColumn25"
        Me.GridColumn25.Visible = True
        Me.GridColumn25.VisibleIndex = 5
        '
        'GridColumn26
        '
        Me.GridColumn26.Caption = "Stock"
        Me.GridColumn26.FieldName = "Stock"
        Me.GridColumn26.Name = "GridColumn26"
        '
        'GridColumn27
        '
        Me.GridColumn27.Caption = "Cantidad Ingreso"
        Me.GridColumn27.FieldName = "CantidadIngreso"
        Me.GridColumn27.Name = "GridColumn27"
        Me.GridColumn27.Tag = "1"
        Me.GridColumn27.Visible = True
        Me.GridColumn27.VisibleIndex = 6
        '
        'GridColumn28
        '
        Me.GridColumn28.Caption = "Costo Unitario Ingreso"
        Me.GridColumn28.FieldName = "CostoUnitarioIngreso"
        Me.GridColumn28.Name = "GridColumn28"
        Me.GridColumn28.Tag = "1"
        Me.GridColumn28.Visible = True
        Me.GridColumn28.VisibleIndex = 7
        '
        'GridColumn29
        '
        Me.GridColumn29.Caption = "Total Ingreso"
        Me.GridColumn29.FieldName = "TotalIngreso"
        Me.GridColumn29.Name = "GridColumn29"
        Me.GridColumn29.Tag = "1"
        Me.GridColumn29.Visible = True
        Me.GridColumn29.VisibleIndex = 8
        '
        'GridColumn30
        '
        Me.GridColumn30.Caption = "Cantidad Salida"
        Me.GridColumn30.FieldName = "CantidadSalida"
        Me.GridColumn30.Name = "GridColumn30"
        Me.GridColumn30.Tag = "2"
        Me.GridColumn30.Visible = True
        Me.GridColumn30.VisibleIndex = 9
        '
        'GridColumn31
        '
        Me.GridColumn31.Caption = "Costo Unitario Salida"
        Me.GridColumn31.FieldName = "CostoUnitarioSalida"
        Me.GridColumn31.Name = "GridColumn31"
        Me.GridColumn31.Tag = "2"
        Me.GridColumn31.Visible = True
        Me.GridColumn31.VisibleIndex = 10
        '
        'GridColumn32
        '
        Me.GridColumn32.Caption = "Total Salida"
        Me.GridColumn32.FieldName = "TotalSalida"
        Me.GridColumn32.Name = "GridColumn32"
        Me.GridColumn32.Tag = "2"
        Me.GridColumn32.Visible = True
        Me.GridColumn32.VisibleIndex = 11
        '
        'GridColumn33
        '
        Me.GridColumn33.Caption = "Cantidad Total"
        Me.GridColumn33.FieldName = "CantidadTotal"
        Me.GridColumn33.Name = "GridColumn33"
        Me.GridColumn33.Tag = "3"
        Me.GridColumn33.Visible = True
        Me.GridColumn33.VisibleIndex = 12
        '
        'GridColumn34
        '
        Me.GridColumn34.Caption = "Costo Unitario Calculado"
        Me.GridColumn34.FieldName = "CostoUnitarioCalculado"
        Me.GridColumn34.Name = "GridColumn34"
        Me.GridColumn34.Tag = "3"
        Me.GridColumn34.Visible = True
        Me.GridColumn34.VisibleIndex = 13
        '
        'GridColumn35
        '
        Me.GridColumn35.Caption = "Saldo Total"
        Me.GridColumn35.FieldName = "SaldoTotal"
        Me.GridColumn35.Name = "GridColumn35"
        Me.GridColumn35.Tag = "3"
        Me.GridColumn35.Visible = True
        Me.GridColumn35.VisibleIndex = 14
        '
        'KardexControlForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(890, 462)
        Me.Controls.Add(Me.SplitContainerControl1)
        Me.Controls.Add(Me.BarDockControl3)
        Me.Controls.Add(Me.BarDockControl4)
        Me.Controls.Add(Me.BarDockControl2)
        Me.Controls.Add(Me.BarDockControl1)
        Me.Name = "KardexControlForm"
        Me.Text = "Control de Inventario"
        CType(Me.gcKardex, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.GridView1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SplitContainerControl1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainerControl1.ResumeLayout(False)
        CType(Me.seFiscalPeriod.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.seFiscalYear.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.beResultado.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.beSalidas.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.beImportaciones.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.beSaldoInicial.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SplitContainerControl2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainerControl2.ResumeLayout(False)
        CType(Me.vpInputs, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.bmActions, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.rpiProceso, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RepositoryItemLookUpEdit1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RepositoryItemImageComboBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RepositoryItemRadioGroup1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RepositoryItemComboBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.GridControl1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.GridView2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents gcKardex As DevExpress.XtraGrid.GridControl
    Friend WithEvents GridView1 As DevExpress.XtraGrid.Views.Grid.GridView
    Friend WithEvents SplitContainerControl1 As DevExpress.XtraEditors.SplitContainerControl
    Friend WithEvents vpInputs As DevExpress.XtraEditors.DXErrorProvider.DXValidationProvider
    Private WithEvents bmActions As DevExpress.XtraBars.BarManager
    Private WithEvents bar5 As DevExpress.XtraBars.Bar
    Private WithEvents brsDescripcion As DevExpress.XtraBars.BarStaticItem
    Private WithEvents brBarraAcciones As DevExpress.XtraBars.Bar
    Private WithEvents bbiProcesar As DevExpress.XtraBars.BarButtonItem
    Friend WithEvents bbiBuscar As DevExpress.XtraBars.BarButtonItem
    Friend WithEvents bbiGuardar As DevExpress.XtraBars.BarButtonItem
    Friend WithEvents bbiExportar As DevExpress.XtraBars.BarButtonItem
    Private WithEvents bbiCerrar As DevExpress.XtraBars.BarButtonItem
    Private WithEvents BarDockControl1 As DevExpress.XtraBars.BarDockControl
    Private WithEvents BarDockControl2 As DevExpress.XtraBars.BarDockControl
    Private WithEvents BarDockControl3 As DevExpress.XtraBars.BarDockControl
    Private WithEvents BarDockControl4 As DevExpress.XtraBars.BarDockControl
    Private WithEvents barStaticItem3 As DevExpress.XtraBars.BarStaticItem
    Private WithEvents barStaticItem4 As DevExpress.XtraBars.BarStaticItem
    Private WithEvents brsEstado As DevExpress.XtraBars.BarStaticItem
    Private WithEvents beiProceso As DevExpress.XtraBars.BarEditItem
    Private WithEvents rpiProceso As DevExpress.XtraEditors.Repository.RepositoryItemProgressBar
    Friend WithEvents BarButtonItem1 As DevExpress.XtraBars.BarButtonItem
    Friend WithEvents bsiVistas As DevExpress.XtraBars.BarSubItem
    Friend WithEvents bbiVistaGrilla As DevExpress.XtraBars.BarButtonItem
    Friend WithEvents bbiTarjeta As DevExpress.XtraBars.BarButtonItem
    Friend WithEvents bbiContrato As DevExpress.XtraBars.BarButtonItem
    Friend WithEvents bbiCronograma As DevExpress.XtraBars.BarButtonItem
    Friend WithEvents bbiCartaNotarial As DevExpress.XtraBars.BarButtonItem
    Friend WithEvents bbiLetras As DevExpress.XtraBars.BarButtonItem
    Friend WithEvents BarButtonItem3 As DevExpress.XtraBars.BarButtonItem
    Friend WithEvents BarButtonItem4 As DevExpress.XtraBars.BarButtonItem
    Friend WithEvents RepositoryItemLookUpEdit1 As DevExpress.XtraEditors.Repository.RepositoryItemLookUpEdit
    Friend WithEvents RepositoryItemImageComboBox1 As DevExpress.XtraEditors.Repository.RepositoryItemImageComboBox
    Friend WithEvents RepositoryItemRadioGroup1 As DevExpress.XtraEditors.Repository.RepositoryItemRadioGroup
    Friend WithEvents RepositoryItemComboBox1 As DevExpress.XtraEditors.Repository.RepositoryItemComboBox
    Friend WithEvents xofdSaldoInicial As DevExpress.XtraEditors.XtraOpenFileDialog
    Friend WithEvents beSaldoInicial As DevExpress.XtraEditors.ButtonEdit
    Friend WithEvents Label5 As Label
    Friend WithEvents SplitContainerControl2 As DevExpress.XtraEditors.SplitContainerControl
    Friend WithEvents Label3 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents seFiscalPeriod As DevExpress.XtraEditors.SpinEdit
    Friend WithEvents seFiscalYear As DevExpress.XtraEditors.SpinEdit
    Friend WithEvents beSalidas As DevExpress.XtraEditors.ButtonEdit
    Friend WithEvents beImportaciones As DevExpress.XtraEditors.ButtonEdit
    Friend WithEvents Label2 As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents beResultado As DevExpress.XtraEditors.ButtonEdit
    Friend WithEvents Label6 As Label
    Friend WithEvents xofdImportaciones As DevExpress.XtraEditors.XtraOpenFileDialog
    Friend WithEvents xofdSalidas As DevExpress.XtraEditors.XtraOpenFileDialog
    Friend WithEvents xfbdResultado As DevExpress.XtraEditors.XtraFolderBrowserDialog
    Friend WithEvents GridColumn1 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn2 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn3 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn4 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn5 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn6 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn7 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn8 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn9 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn10 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn11 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn12 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn13 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn14 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn15 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn16 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn17 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn18 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridControl1 As DevExpress.XtraGrid.GridControl
    Friend WithEvents GridView2 As DevExpress.XtraGrid.Views.Grid.GridView
    Friend WithEvents GridColumn19 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn20 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn21 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn22 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn23 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn24 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn25 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn26 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn27 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn28 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn29 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn30 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn31 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn32 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn33 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn34 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents GridColumn35 As DevExpress.XtraGrid.Columns.GridColumn
End Class
