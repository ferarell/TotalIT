Imports System
Imports System.ComponentModel
Imports System.Collections
Imports System.Windows.Forms
Imports System.Data
Imports System.Drawing
Imports Npgsql

Public Class Form1
    'Dim oGeneralDAO As New GeneralDAO

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
       
    End Sub

    Private Sub SimpleButton1_Click(sender As Object, e As EventArgs) Handles SimpleButton1.Click
        Dim dtQuery As New DataTable
        dtQuery = ExecuteSQL("SELECT * FROM public.account_account_financial_report")
        gcExternalData.DataSource = dtQuery
    End Sub
End Class
