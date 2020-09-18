<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SaleByWebForm
    Inherits DevExpress.XtraEditors.XtraForm

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(SaleByWebForm))
        Me.gcSaleByWeb = New DevExpress.XtraGrid.GridControl()
        Me.GridView1 = New DevExpress.XtraGrid.Views.Grid.GridView()
        Me.SplitContainerControl1 = New DevExpress.XtraEditors.SplitContainerControl()
        Me.beArchivoOrigen = New DevExpress.XtraEditors.ButtonEdit()
        Me.Label5 = New System.Windows.Forms.Label()
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
        Me.XtraOpenFileDialog1 = New DevExpress.XtraEditors.XtraOpenFileDialog(Me.components)
        CType(Me.gcSaleByWeb, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.GridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SplitContainerControl1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainerControl1.SuspendLayout()
        CType(Me.beArchivoOrigen.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.vpInputs, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.bmActions, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.rpiProceso, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RepositoryItemLookUpEdit1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RepositoryItemImageComboBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RepositoryItemRadioGroup1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RepositoryItemComboBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'gcSaleByWeb
        '
        Me.gcSaleByWeb.Dock = System.Windows.Forms.DockStyle.Fill
        Me.gcSaleByWeb.Location = New System.Drawing.Point(0, 0)
        Me.gcSaleByWeb.MainView = Me.GridView1
        Me.gcSaleByWeb.Name = "gcSaleByWeb"
        Me.gcSaleByWeb.Size = New System.Drawing.Size(903, 292)
        Me.gcSaleByWeb.TabIndex = 0
        Me.gcSaleByWeb.UseEmbeddedNavigator = True
        Me.gcSaleByWeb.ViewCollection.AddRange(New DevExpress.XtraGrid.Views.Base.BaseView() {Me.GridView1})
        '
        'GridView1
        '
        Me.GridView1.GridControl = Me.gcSaleByWeb
        Me.GridView1.Name = "GridView1"
        Me.GridView1.OptionsView.ColumnAutoWidth = False
        '
        'SplitContainerControl1
        '
        Me.SplitContainerControl1.CollapsePanel = DevExpress.XtraEditors.SplitCollapsePanel.Panel1
        Me.SplitContainerControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainerControl1.Horizontal = False
        Me.SplitContainerControl1.Location = New System.Drawing.Point(0, 47)
        Me.SplitContainerControl1.Name = "SplitContainerControl1"
        Me.SplitContainerControl1.Panel1.Controls.Add(Me.beArchivoOrigen)
        Me.SplitContainerControl1.Panel1.Controls.Add(Me.Label5)
        Me.SplitContainerControl1.Panel1.Text = "Panel1"
        Me.SplitContainerControl1.Panel2.Controls.Add(Me.gcSaleByWeb)
        Me.SplitContainerControl1.Panel2.Text = "Panel2"
        Me.SplitContainerControl1.Size = New System.Drawing.Size(903, 351)
        Me.SplitContainerControl1.SplitterPosition = 54
        Me.SplitContainerControl1.TabIndex = 1
        '
        'beArchivoOrigen
        '
        Me.beArchivoOrigen.Location = New System.Drawing.Point(118, 18)
        Me.beArchivoOrigen.Margin = New System.Windows.Forms.Padding(2)
        Me.beArchivoOrigen.Name = "beArchivoOrigen"
        Me.beArchivoOrigen.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
        Me.beArchivoOrigen.Size = New System.Drawing.Size(641, 20)
        Me.beArchivoOrigen.TabIndex = 27
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(29, 20)
        Me.Label5.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(85, 13)
        Me.Label5.TabIndex = 28
        Me.Label5.Text = "Origen de Datos"
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
        Me.BarDockControl1.Size = New System.Drawing.Size(903, 47)
        '
        'BarDockControl2
        '
        Me.BarDockControl2.CausesValidation = False
        Me.BarDockControl2.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.BarDockControl2.Location = New System.Drawing.Point(0, 398)
        Me.BarDockControl2.Manager = Me.bmActions
        Me.BarDockControl2.Size = New System.Drawing.Size(903, 29)
        '
        'BarDockControl3
        '
        Me.BarDockControl3.CausesValidation = False
        Me.BarDockControl3.Dock = System.Windows.Forms.DockStyle.Left
        Me.BarDockControl3.Location = New System.Drawing.Point(0, 47)
        Me.BarDockControl3.Manager = Me.bmActions
        Me.BarDockControl3.Size = New System.Drawing.Size(0, 351)
        '
        'BarDockControl4
        '
        Me.BarDockControl4.CausesValidation = False
        Me.BarDockControl4.Dock = System.Windows.Forms.DockStyle.Right
        Me.BarDockControl4.Location = New System.Drawing.Point(903, 47)
        Me.BarDockControl4.Manager = Me.bmActions
        Me.BarDockControl4.Size = New System.Drawing.Size(0, 351)
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
        'XtraOpenFileDialog1
        '
        Me.XtraOpenFileDialog1.FileName = "XtraOpenFileDialog1"
        Me.XtraOpenFileDialog1.Multiselect = True
        '
        'SaleByWebForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(903, 427)
        Me.Controls.Add(Me.SplitContainerControl1)
        Me.Controls.Add(Me.BarDockControl3)
        Me.Controls.Add(Me.BarDockControl4)
        Me.Controls.Add(Me.BarDockControl2)
        Me.Controls.Add(Me.BarDockControl1)
        Me.Name = "SaleByWebForm"
        Me.Text = "Venta por Web"
        CType(Me.gcSaleByWeb, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.GridView1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SplitContainerControl1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainerControl1.ResumeLayout(False)
        CType(Me.beArchivoOrigen.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.vpInputs, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.bmActions, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.rpiProceso, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RepositoryItemLookUpEdit1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RepositoryItemImageComboBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RepositoryItemRadioGroup1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RepositoryItemComboBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents gcSaleByWeb As DevExpress.XtraGrid.GridControl
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
    Friend WithEvents XtraOpenFileDialog1 As DevExpress.XtraEditors.XtraOpenFileDialog
    Friend WithEvents beArchivoOrigen As DevExpress.XtraEditors.ButtonEdit
    Friend WithEvents Label5 As Label
End Class
