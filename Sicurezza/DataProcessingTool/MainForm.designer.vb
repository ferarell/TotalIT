<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MainForm
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(MainForm))
        Me.nbcMainMenu = New DevExpress.XtraNavBar.NavBarControl()
        Me.NavBarGroup3 = New DevExpress.XtraNavBar.NavBarGroup()
        Me.NavBarItem1 = New DevExpress.XtraNavBar.NavBarItem()
        Me.NavBarItem2 = New DevExpress.XtraNavBar.NavBarItem()
        Me.NavBarGroup5 = New DevExpress.XtraNavBar.NavBarGroup()
        Me.NavBarItem12 = New DevExpress.XtraNavBar.NavBarItem()
        Me.NavBarGroup1 = New DevExpress.XtraNavBar.NavBarGroup()
        Me.NavBarItem3 = New DevExpress.XtraNavBar.NavBarItem()
        Me.NavBarItem5 = New DevExpress.XtraNavBar.NavBarItem()
        Me.NavBarItem11 = New DevExpress.XtraNavBar.NavBarItem()
        Me.NavBarItem6 = New DevExpress.XtraNavBar.NavBarItem()
        Me.NavBarItem7 = New DevExpress.XtraNavBar.NavBarItem()
        Me.NavBarItem8 = New DevExpress.XtraNavBar.NavBarItem()
        Me.NavBarGroup2 = New DevExpress.XtraNavBar.NavBarGroup()
        Me.NavBarItem4 = New DevExpress.XtraNavBar.NavBarItem()
        Me.NavBarItem10 = New DevExpress.XtraNavBar.NavBarItem()
        Me.NavBarGroup4 = New DevExpress.XtraNavBar.NavBarGroup()
        Me.NavBarItem9 = New DevExpress.XtraNavBar.NavBarItem()
        Me.XtraTabbedMdiManager1 = New DevExpress.XtraTabbedMdi.XtraTabbedMdiManager(Me.components)
        Me.BarManager1 = New DevExpress.XtraBars.BarManager(Me.components)
        Me.Bar3 = New DevExpress.XtraBars.Bar()
        Me.barDockControlTop = New DevExpress.XtraBars.BarDockControl()
        Me.barDockControlBottom = New DevExpress.XtraBars.BarDockControl()
        Me.barDockControlLeft = New DevExpress.XtraBars.BarDockControl()
        Me.barDockControlRight = New DevExpress.XtraBars.BarDockControl()
        Me.NavBarItem13 = New DevExpress.XtraNavBar.NavBarItem()
        CType(Me.nbcMainMenu, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.XtraTabbedMdiManager1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.BarManager1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'nbcMainMenu
        '
        Me.nbcMainMenu.ActiveGroup = Me.NavBarGroup3
        Me.nbcMainMenu.Dock = System.Windows.Forms.DockStyle.Left
        Me.nbcMainMenu.Groups.AddRange(New DevExpress.XtraNavBar.NavBarGroup() {Me.NavBarGroup5, Me.NavBarGroup1, Me.NavBarGroup2, Me.NavBarGroup3, Me.NavBarGroup4})
        Me.nbcMainMenu.Items.AddRange(New DevExpress.XtraNavBar.NavBarItem() {Me.NavBarItem1, Me.NavBarItem2, Me.NavBarItem3, Me.NavBarItem4, Me.NavBarItem5, Me.NavBarItem6, Me.NavBarItem7, Me.NavBarItem8, Me.NavBarItem9, Me.NavBarItem10, Me.NavBarItem11, Me.NavBarItem12, Me.NavBarItem13})
        Me.nbcMainMenu.Location = New System.Drawing.Point(0, 0)
        Me.nbcMainMenu.Margin = New System.Windows.Forms.Padding(2)
        Me.nbcMainMenu.Name = "nbcMainMenu"
        Me.nbcMainMenu.OptionsNavPane.ExpandedWidth = 178
        Me.nbcMainMenu.Size = New System.Drawing.Size(178, 409)
        Me.nbcMainMenu.StoreDefaultPaintStyleName = True
        Me.nbcMainMenu.TabIndex = 3
        Me.nbcMainMenu.Text = "NavBarControl1"
        '
        'NavBarGroup3
        '
        Me.NavBarGroup3.Caption = "Configuración"
        Me.NavBarGroup3.Expanded = True
        Me.NavBarGroup3.ItemLinks.AddRange(New DevExpress.XtraNavBar.NavBarItemLink() {New DevExpress.XtraNavBar.NavBarItemLink(Me.NavBarItem1), New DevExpress.XtraNavBar.NavBarItemLink(Me.NavBarItem2)})
        Me.NavBarGroup3.Name = "NavBarGroup3"
        '
        'NavBarItem1
        '
        Me.NavBarItem1.Caption = "Diseño"
        Me.NavBarItem1.Name = "NavBarItem1"
        '
        'NavBarItem2
        '
        Me.NavBarItem2.Caption = "Preferencias"
        Me.NavBarItem2.Name = "NavBarItem2"
        Me.NavBarItem2.Visible = False
        '
        'NavBarGroup5
        '
        Me.NavBarGroup5.Caption = "Conciliaciones"
        Me.NavBarGroup5.Expanded = True
        Me.NavBarGroup5.ItemLinks.AddRange(New DevExpress.XtraNavBar.NavBarItemLink() {New DevExpress.XtraNavBar.NavBarItemLink(Me.NavBarItem12), New DevExpress.XtraNavBar.NavBarItemLink(Me.NavBarItem13)})
        Me.NavBarGroup5.Name = "NavBarGroup5"
        '
        'NavBarItem12
        '
        Me.NavBarItem12.Caption = "Ventas por Web"
        Me.NavBarItem12.Name = "NavBarItem12"
        '
        'NavBarGroup1
        '
        Me.NavBarGroup1.Caption = "Informes SUNAT"
        Me.NavBarGroup1.Expanded = True
        Me.NavBarGroup1.ItemLinks.AddRange(New DevExpress.XtraNavBar.NavBarItemLink() {New DevExpress.XtraNavBar.NavBarItemLink(Me.NavBarItem3), New DevExpress.XtraNavBar.NavBarItemLink(Me.NavBarItem5), New DevExpress.XtraNavBar.NavBarItemLink(Me.NavBarItem11), New DevExpress.XtraNavBar.NavBarItemLink(Me.NavBarItem6), New DevExpress.XtraNavBar.NavBarItemLink(Me.NavBarItem7), New DevExpress.XtraNavBar.NavBarItemLink(Me.NavBarItem8)})
        Me.NavBarGroup1.Name = "NavBarGroup1"
        Me.NavBarGroup1.Visible = False
        '
        'NavBarItem3
        '
        Me.NavBarItem3.Caption = "Libro Diario"
        Me.NavBarItem3.Name = "NavBarItem3"
        '
        'NavBarItem5
        '
        Me.NavBarItem5.Caption = "Libro Mayor"
        Me.NavBarItem5.Name = "NavBarItem5"
        '
        'NavBarItem11
        '
        Me.NavBarItem11.Caption = "Libro Caja y Bancos"
        Me.NavBarItem11.Name = "NavBarItem11"
        '
        'NavBarItem6
        '
        Me.NavBarItem6.Caption = "Balance General"
        Me.NavBarItem6.Name = "NavBarItem6"
        '
        'NavBarItem7
        '
        Me.NavBarItem7.Caption = "Ganacias y Pérdidas"
        Me.NavBarItem7.Name = "NavBarItem7"
        '
        'NavBarItem8
        '
        Me.NavBarItem8.Caption = "Inventario Valorizado"
        Me.NavBarItem8.Name = "NavBarItem8"
        '
        'NavBarGroup2
        '
        Me.NavBarGroup2.Caption = "Informes Contables"
        Me.NavBarGroup2.Expanded = True
        Me.NavBarGroup2.ItemLinks.AddRange(New DevExpress.XtraNavBar.NavBarItemLink() {New DevExpress.XtraNavBar.NavBarItemLink(Me.NavBarItem4), New DevExpress.XtraNavBar.NavBarItemLink(Me.NavBarItem10)})
        Me.NavBarGroup2.Name = "NavBarGroup2"
        Me.NavBarGroup2.Visible = False
        '
        'NavBarItem4
        '
        Me.NavBarItem4.Caption = "Movimiento Contable"
        Me.NavBarItem4.Name = "NavBarItem4"
        '
        'NavBarItem10
        '
        Me.NavBarItem10.Caption = "Saldos por Cuentas"
        Me.NavBarItem10.Name = "NavBarItem10"
        '
        'NavBarGroup4
        '
        Me.NavBarGroup4.Caption = "Herramientas"
        Me.NavBarGroup4.Expanded = True
        Me.NavBarGroup4.ItemLinks.AddRange(New DevExpress.XtraNavBar.NavBarItemLink() {New DevExpress.XtraNavBar.NavBarItemLink(Me.NavBarItem9)})
        Me.NavBarGroup4.Name = "NavBarGroup4"
        Me.NavBarGroup4.Visible = False
        '
        'NavBarItem9
        '
        Me.NavBarItem9.Caption = "Consulta SQL"
        Me.NavBarItem9.Name = "NavBarItem9"
        '
        'XtraTabbedMdiManager1
        '
        Me.XtraTabbedMdiManager1.ClosePageButtonShowMode = DevExpress.XtraTab.ClosePageButtonShowMode.InAllTabPageHeaders
        Me.XtraTabbedMdiManager1.MdiParent = Me
        '
        'BarManager1
        '
        Me.BarManager1.Bars.AddRange(New DevExpress.XtraBars.Bar() {Me.Bar3})
        Me.BarManager1.DockControls.Add(Me.barDockControlTop)
        Me.BarManager1.DockControls.Add(Me.barDockControlBottom)
        Me.BarManager1.DockControls.Add(Me.barDockControlLeft)
        Me.BarManager1.DockControls.Add(Me.barDockControlRight)
        Me.BarManager1.Form = Me
        Me.BarManager1.StatusBar = Me.Bar3
        '
        'Bar3
        '
        Me.Bar3.BarName = "Barra de estado"
        Me.Bar3.CanDockStyle = DevExpress.XtraBars.BarCanDockStyle.Bottom
        Me.Bar3.DockCol = 0
        Me.Bar3.DockRow = 0
        Me.Bar3.DockStyle = DevExpress.XtraBars.BarDockStyle.Bottom
        Me.Bar3.OptionsBar.AllowQuickCustomization = False
        Me.Bar3.OptionsBar.DrawDragBorder = False
        Me.Bar3.OptionsBar.UseWholeRow = True
        Me.Bar3.Text = "Barra de estado"
        '
        'barDockControlTop
        '
        Me.barDockControlTop.CausesValidation = False
        Me.barDockControlTop.Dock = System.Windows.Forms.DockStyle.Top
        Me.barDockControlTop.Location = New System.Drawing.Point(0, 0)
        Me.barDockControlTop.Manager = Me.BarManager1
        Me.barDockControlTop.Margin = New System.Windows.Forms.Padding(2)
        Me.barDockControlTop.Size = New System.Drawing.Size(830, 0)
        '
        'barDockControlBottom
        '
        Me.barDockControlBottom.CausesValidation = False
        Me.barDockControlBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.barDockControlBottom.Location = New System.Drawing.Point(0, 409)
        Me.barDockControlBottom.Manager = Me.BarManager1
        Me.barDockControlBottom.Margin = New System.Windows.Forms.Padding(2)
        Me.barDockControlBottom.Size = New System.Drawing.Size(830, 23)
        '
        'barDockControlLeft
        '
        Me.barDockControlLeft.CausesValidation = False
        Me.barDockControlLeft.Dock = System.Windows.Forms.DockStyle.Left
        Me.barDockControlLeft.Location = New System.Drawing.Point(0, 0)
        Me.barDockControlLeft.Manager = Me.BarManager1
        Me.barDockControlLeft.Margin = New System.Windows.Forms.Padding(2)
        Me.barDockControlLeft.Size = New System.Drawing.Size(0, 409)
        '
        'barDockControlRight
        '
        Me.barDockControlRight.CausesValidation = False
        Me.barDockControlRight.Dock = System.Windows.Forms.DockStyle.Right
        Me.barDockControlRight.Location = New System.Drawing.Point(830, 0)
        Me.barDockControlRight.Manager = Me.BarManager1
        Me.barDockControlRight.Margin = New System.Windows.Forms.Padding(2)
        Me.barDockControlRight.Size = New System.Drawing.Size(0, 409)
        '
        'NavBarItem13
        '
        Me.NavBarItem13.Caption = "Control de Inventario"
        Me.NavBarItem13.Name = "NavBarItem13"
        '
        'MainForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSize = True
        Me.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.ClientSize = New System.Drawing.Size(830, 432)
        Me.Controls.Add(Me.nbcMainMenu)
        Me.Controls.Add(Me.barDockControlLeft)
        Me.Controls.Add(Me.barDockControlRight)
        Me.Controls.Add(Me.barDockControlBottom)
        Me.Controls.Add(Me.barDockControlTop)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.IsMdiContainer = True
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.Name = "MainForm"
        Me.Text = "Integración de Datos"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        CType(Me.nbcMainMenu, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.XtraTabbedMdiManager1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.BarManager1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents nbcMainMenu As DevExpress.XtraNavBar.NavBarControl
    Friend WithEvents XtraTabbedMdiManager1 As DevExpress.XtraTabbedMdi.XtraTabbedMdiManager
    Friend WithEvents NavBarGroup1 As DevExpress.XtraNavBar.NavBarGroup
    Friend WithEvents NavBarItem3 As DevExpress.XtraNavBar.NavBarItem
    Friend WithEvents NavBarGroup2 As DevExpress.XtraNavBar.NavBarGroup
    Friend WithEvents NavBarItem4 As DevExpress.XtraNavBar.NavBarItem
    Friend WithEvents NavBarGroup3 As DevExpress.XtraNavBar.NavBarGroup
    Friend WithEvents NavBarItem1 As DevExpress.XtraNavBar.NavBarItem
    Friend WithEvents NavBarItem2 As DevExpress.XtraNavBar.NavBarItem
    Friend WithEvents barDockControlLeft As DevExpress.XtraBars.BarDockControl
    Friend WithEvents barDockControlRight As DevExpress.XtraBars.BarDockControl
    Friend WithEvents barDockControlBottom As DevExpress.XtraBars.BarDockControl
    Friend WithEvents barDockControlTop As DevExpress.XtraBars.BarDockControl
    Friend WithEvents BarManager1 As DevExpress.XtraBars.BarManager
    Friend WithEvents Bar3 As DevExpress.XtraBars.Bar
    Friend WithEvents NavBarItem5 As DevExpress.XtraNavBar.NavBarItem
    Friend WithEvents NavBarItem6 As DevExpress.XtraNavBar.NavBarItem
    Friend WithEvents NavBarItem7 As DevExpress.XtraNavBar.NavBarItem
    Friend WithEvents NavBarItem8 As DevExpress.XtraNavBar.NavBarItem
    Friend WithEvents NavBarGroup4 As DevExpress.XtraNavBar.NavBarGroup
    Friend WithEvents NavBarItem9 As DevExpress.XtraNavBar.NavBarItem
    Friend WithEvents NavBarItem10 As DevExpress.XtraNavBar.NavBarItem
    Friend WithEvents NavBarItem11 As DevExpress.XtraNavBar.NavBarItem
    Friend WithEvents NavBarGroup5 As DevExpress.XtraNavBar.NavBarGroup
    Friend WithEvents NavBarItem12 As DevExpress.XtraNavBar.NavBarItem
    Friend WithEvents NavBarItem13 As DevExpress.XtraNavBar.NavBarItem
End Class
