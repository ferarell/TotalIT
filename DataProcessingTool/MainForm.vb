Imports DevExpress.XtraEditors
Imports DevExpress.Skins
Imports System.Threading
Imports System.Globalization

Public Class MainForm

    Public Sub New()
        Dim currentWithOverriddenNumber As CultureInfo = New CultureInfo(CultureInfo.CurrentCulture.Name)
        currentWithOverriddenNumber.NumberFormat.CurrencyPositivePattern = 0 '; // make sure there is no space between symbol and number
        currentWithOverriddenNumber.NumberFormat.CurrencySymbol = " " '; // no currency symbol
        currentWithOverriddenNumber.NumberFormat.CurrencyDecimalSeparator = "." '; //decimal separator
        currentWithOverriddenNumber.NumberFormat.CurrencyDecimalDigits = 2
        currentWithOverriddenNumber.NumberFormat.CurrencyGroupSizes = {3} '; //no digit groupings
        currentWithOverriddenNumber.NumberFormat.CurrencyGroupSeparator = ","
        currentWithOverriddenNumber.NumberFormat.NumberGroupSizes = {3} ';
        currentWithOverriddenNumber.NumberFormat.NumberGroupSeparator = ","
        currentWithOverriddenNumber.NumberFormat.NumberDecimalSeparator = "." '; //decimal separator
        currentWithOverriddenNumber.NumberFormat.NumberDecimalDigits = 2
        currentWithOverriddenNumber.DateTimeFormat.FullDateTimePattern = "dd/MM/yyyy hh:mm"
        currentWithOverriddenNumber.DateTimeFormat.ShortDatePattern = "dd/MM/yyyy"
        Thread.CurrentThread.CurrentCulture = currentWithOverriddenNumber
        InitializeComponent()
        DevExpress.Skins.SkinManager.EnableFormSkins()
        DevExpress.UserSkins.BonusSkins.Register()
        SkinName = My.Settings.LookAndFeel
    End Sub

    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False
        'My.Settings.Upgrade()
        If My.Computer.Name <> "FERARELL" Then
            My.Settings.Upgrade()
            NavBarGroup4.Visible = False
            NavBarItem8.Visible = False
            'NavBarGroup2.Visible = False
        End If

        DevExpress.LookAndFeel.UserLookAndFeel.Default.SetSkinStyle(SkinName)
        If My.Settings.PaintStyle = "ExplorerBar" Then
            nbcMainMenu.PaintStyleKind = DevExpress.XtraNavBar.NavBarViewKind.ExplorerBar
        Else
            nbcMainMenu.PaintStyleKind = DevExpress.XtraNavBar.NavBarViewKind.NavigationPane
        End If
        nbcMainMenu.PaintStyleName = My.Settings.LookAndFeel
    End Sub

    Private Sub MainForm_TextChanged(sender As Object, e As EventArgs) Handles MyBase.TextChanged
        Me.Text = My.Application.Info.ProductName + " [" + My.Application.Info.Version.ToString + "]" ' & UserApp
    End Sub

    Private Sub SelectPage(ByVal FormName As String)
        For Each myChildForm In MdiChildren
            If myChildForm.Name = FormName Then
                myChildForm.Focus()
            End If
        Next
    End Sub

    Private Sub OpenForm(AppForm As Windows.Forms.Form)
        Try
            Dim myForm As New Windows.Forms.Form
            myForm = AppForm
            If Me.Controls.Find(myForm.Name, True).Count = 0 Then
                myForm.MdiParent = Me
                myForm.Show()
            Else
                SelectPage(myForm.Name)
            End If
        Catch ex As Exception
            DevExpress.XtraEditors.XtraMessageBox.Show(Me.LookAndFeel, ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub NavBarItem2_LinkClicked(sender As Object, e As DevExpress.XtraNavBar.NavBarLinkEventArgs) Handles NavBarItem2.LinkClicked
        OpenForm(PreferencesForm)
    End Sub

    Private Sub NavBarItem1_LinkClicked(sender As Object, e As DevExpress.XtraNavBar.NavBarLinkEventArgs) Handles NavBarItem1.LinkClicked
        OpenForm(LookAndFeelForm)
    End Sub

    Private Sub NavBarItem3_LinkClicked(sender As Object, e As DevExpress.XtraNavBar.NavBarLinkEventArgs) Handles NavBarItem3.LinkClicked
        'OpenForm(LibroDiarioForm)
    End Sub

    Private Sub NavBarItem5_LinkClicked(sender As Object, e As DevExpress.XtraNavBar.NavBarLinkEventArgs) Handles NavBarItem5.LinkClicked
        'OpenForm(LibroMayorForm)
    End Sub

    Private Sub MainForm_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If DevExpress.XtraEditors.XtraMessageBox.Show(Me.LookAndFeel, "Esta seguro de cerrar la aplicación?", "Salir", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.No Then
            e.Cancel = True
        End If
    End Sub

    Private Sub NavBarItem6_LinkClicked(sender As Object, e As DevExpress.XtraNavBar.NavBarLinkEventArgs) Handles NavBarItem6.LinkClicked
        'OpenForm(BalanceGeneralForm)
    End Sub

    Private Sub NavBarItem7_LinkClicked(sender As Object, e As DevExpress.XtraNavBar.NavBarLinkEventArgs) Handles NavBarItem7.LinkClicked
        'OpenForm(EstadoResultadoForm)
    End Sub

    Private Sub NavBarItem9_LinkClicked(sender As Object, e As DevExpress.XtraNavBar.NavBarLinkEventArgs) Handles NavBarItem9.LinkClicked
        Dim myForm As New QueryForm
        myForm.MdiParent = Me
        myForm.Show()
    End Sub

    Private Sub NavBarItem8_LinkClicked(sender As Object, e As DevExpress.XtraNavBar.NavBarLinkEventArgs) Handles NavBarItem8.LinkClicked
        'OpenForm(InventarioValorizadoForm)
    End Sub

    Private Sub NavBarItem4_LinkClicked(sender As Object, e As DevExpress.XtraNavBar.NavBarLinkEventArgs) Handles NavBarItem4.LinkClicked
        'OpenForm(MovimientoContableForm)
    End Sub

    Private Sub NavBarItem10_LinkClicked(sender As Object, e As DevExpress.XtraNavBar.NavBarLinkEventArgs) Handles NavBarItem10.LinkClicked
        'OpenForm(SaldoCuentaAnualForm)
    End Sub

    Private Sub NavBarItem11_LinkClicked(sender As Object, e As DevExpress.XtraNavBar.NavBarLinkEventArgs) Handles NavBarItem11.LinkClicked
        'OpenForm(LibroCajaBancosForm)
    End Sub

    Private Sub NavBarItem12_LinkClicked(sender As Object, e As DevExpress.XtraNavBar.NavBarLinkEventArgs) Handles NavBarItem12.LinkClicked
        OpenForm(SaleByWebForm)
    End Sub
End Class