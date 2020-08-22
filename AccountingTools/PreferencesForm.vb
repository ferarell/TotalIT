﻿Imports DevExpress.LookAndFeel

Public Class PreferencesForm

    Private Sub beArchivoOrigen_Properties_ButtonClick(sender As Object, e As DevExpress.XtraEditors.Controls.ButtonPressedEventArgs) Handles beDBDirectory.Properties.ButtonClick
        FolderBrowserDialog1.SelectedPath = sender.text
        If FolderBrowserDialog1.ShowDialog = DialogResult.OK Then
            sender.EditValue = FolderBrowserDialog1.SelectedPath
        End If
    End Sub

    Private Sub bbiCerrar_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiCerrar.ItemClick
        Close()
    End Sub

    Private Sub bbiGuardar_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiGuardar.ItemClick
        Validate()
        Try
            My.Settings.Save()
        Catch ex As Exception
            DevExpress.XtraEditors.XtraMessageBox.Show(Me.LookAndFeel, ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            DevExpress.XtraEditors.XtraMessageBox.Show(Me.LookAndFeel, "Changes applied successfully.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Try

    End Sub

    Private Sub PreferencesForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        For Each cnt In DevExpress.Skins.SkinManager.Default.Skins
            lbcEstilos.Items.Add(cnt.SkinName)
        Next
        lbcEstilos.SelectedValue = DevExpress.LookAndFeel.UserLookAndFeel.Default.ActiveLookAndFeel.ActiveSkinName
    End Sub

    Private Sub bbiReset_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles bbiReset.ItemClick
        If DevExpress.XtraEditors.XtraMessageBox.Show(Me.LookAndFeel, "Are you sure to reset?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            My.Settings.Reset()
            Me.PreferencesForm_Load(sender, e)
        End If
    End Sub

    Private Sub lbcEstilos_Click(sender As Object, e As EventArgs) Handles lbcEstilos.Click
        Dim skinName As String
        skinName = lbcEstilos.Text
        DevExpress.LookAndFeel.UserLookAndFeel.Default.SetSkinStyle("")
        DevExpress.LookAndFeel.UserLookAndFeel.Default.SetSkinStyle(skinName)
        My.Settings.LookAndFeel = DevExpress.LookAndFeel.UserLookAndFeel.Default.ActiveLookAndFeel.ActiveSkinName
        My.Settings.Save()
    End Sub

    Private Sub rgPaintStyle_SelectedIndexChanged(sender As Object, e As EventArgs) Handles rgPaintStyle.SelectedIndexChanged
        'My.Settings.PaintStyle = rgPaintStyle.EditValue
        My.Settings.Save()
        'If My.Settings.PaintStyle = "ExplorerBar" Then
        '    MainForm.nbcMainMenu.PaintStyleKind = DevExpress.XtraNavBar.NavBarViewKind.ExplorerBar
        'Else
        '    MainForm.nbcMainMenu.PaintStyleKind = DevExpress.XtraNavBar.NavBarViewKind.NavigationPane
        'End If
    End Sub

End Class