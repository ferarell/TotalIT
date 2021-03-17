<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
    'Se puede modificar usando el Diseñador de Windows Forms.  
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim GridLevelNode1 As DevExpress.XtraGrid.GridLevelNode = New DevExpress.XtraGrid.GridLevelNode()
        Me.CardView1 = New DevExpress.XtraGrid.Views.Card.CardView()
        Me.gcExternalData = New DevExpress.XtraGrid.GridControl()
        Me.GridView1 = New DevExpress.XtraGrid.Views.Grid.GridView()
        Me.SimpleButton1 = New DevExpress.XtraEditors.SimpleButton()
        Me.NavBarControl1 = New DevExpress.XtraNavBar.NavBarControl()
        CType(Me.CardView1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.gcExternalData, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.GridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NavBarControl1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'CardView1
        '
        Me.CardView1.FocusedCardTopFieldIndex = 0
        Me.CardView1.GridControl = Me.gcExternalData
        Me.CardView1.Name = "CardView1"
        Me.CardView1.OptionsBehavior.Editable = False
        Me.CardView1.OptionsFind.AlwaysVisible = True
        Me.CardView1.OptionsView.ShowQuickCustomizeButton = False
        Me.CardView1.VertScrollVisibility = DevExpress.XtraGrid.Views.Base.ScrollVisibility.[Auto]
        '
        'gcExternalData
        '
        Me.gcExternalData.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.gcExternalData.EmbeddedNavigator.Buttons.Append.Visible = False
        Me.gcExternalData.EmbeddedNavigator.Buttons.CancelEdit.Visible = False
        Me.gcExternalData.EmbeddedNavigator.Buttons.Edit.Visible = False
        Me.gcExternalData.EmbeddedNavigator.Buttons.EndEdit.Visible = False
        Me.gcExternalData.EmbeddedNavigator.Buttons.Remove.Visible = False
        Me.gcExternalData.EmbeddedNavigator.Margin = New System.Windows.Forms.Padding(4)
        GridLevelNode1.LevelTemplate = Me.CardView1
        GridLevelNode1.RelationName = "Level1"
        Me.gcExternalData.LevelTree.Nodes.AddRange(New DevExpress.XtraGrid.GridLevelNode() {GridLevelNode1})
        Me.gcExternalData.Location = New System.Drawing.Point(0, 152)
        Me.gcExternalData.MainView = Me.GridView1
        Me.gcExternalData.Margin = New System.Windows.Forms.Padding(4)
        Me.gcExternalData.Name = "gcExternalData"
        Me.gcExternalData.Size = New System.Drawing.Size(1263, 409)
        Me.gcExternalData.TabIndex = 173
        Me.gcExternalData.UseEmbeddedNavigator = True
        Me.gcExternalData.ViewCollection.AddRange(New DevExpress.XtraGrid.Views.Base.BaseView() {Me.GridView1, Me.CardView1})
        '
        'GridView1
        '
        Me.GridView1.GridControl = Me.gcExternalData
        Me.GridView1.GroupSummary.AddRange(New DevExpress.XtraGrid.GridSummaryItem() {New DevExpress.XtraGrid.GridGroupSummaryItem(DevExpress.Data.SummaryItemType.Sum, "ImporteTotal", Nothing, "", New Decimal(New Integer() {0, 0, 0, 131072}))})
        Me.GridView1.Name = "GridView1"
        Me.GridView1.OptionsBehavior.Editable = False
        Me.GridView1.OptionsView.ShowAutoFilterRow = True
        Me.GridView1.OptionsView.ShowFilterPanelMode = DevExpress.XtraGrid.Views.Base.ShowFilterPanelMode.ShowAlways
        Me.GridView1.OptionsView.ShowFooter = True
        '
        'SimpleButton1
        '
        Me.SimpleButton1.Location = New System.Drawing.Point(130, 33)
        Me.SimpleButton1.Name = "SimpleButton1"
        Me.SimpleButton1.Size = New System.Drawing.Size(179, 54)
        Me.SimpleButton1.TabIndex = 174
        Me.SimpleButton1.Text = "SimpleButton1"
        '
        'NavBarControl1
        '
        Me.NavBarControl1.ActiveGroup = Nothing
        Me.NavBarControl1.Location = New System.Drawing.Point(404, 59)
        Me.NavBarControl1.Name = "NavBarControl1"
        Me.NavBarControl1.OptionsNavPane.ExpandedWidth = 140
        Me.NavBarControl1.Size = New System.Drawing.Size(140, 300)
        Me.NavBarControl1.TabIndex = 175
        Me.NavBarControl1.Text = "NavBarControl1"
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1263, 561)
        Me.Controls.Add(Me.NavBarControl1)
        Me.Controls.Add(Me.SimpleButton1)
        Me.Controls.Add(Me.gcExternalData)
        Me.Name = "Form1"
        Me.Text = "Form1"
        CType(Me.CardView1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.gcExternalData, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.GridView1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NavBarControl1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents gcExternalData As DevExpress.XtraGrid.GridControl
    Friend WithEvents CardView1 As DevExpress.XtraGrid.Views.Card.CardView
    Friend WithEvents GridView1 As DevExpress.XtraGrid.Views.Grid.GridView
    Friend WithEvents SimpleButton1 As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents NavBarControl1 As DevExpress.XtraNavBar.NavBarControl

End Class
