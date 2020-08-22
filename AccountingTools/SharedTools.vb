Imports System.IO
Imports Microsoft.Office.Interop
Imports DevExpress.XtraGrid.Views.Grid
Imports DevExpress.XtraGrid

Module SharedTools

    <System.Runtime.CompilerServices.Extension>
    Public Function Contains(ByVal str As String, ByVal ParamArray values As String()) As Boolean
        For Each value In values
            If str.Contains(value) Then
                Return True
            End If
        Next
        Return False
    End Function
    Friend Sub ExportToExcel(sender As System.Object)
        Dim sFileName As String = GetTmpFileName("xlsx")
        If sender.ProductName = "DevExpress.XtraVerticalGrid" Then
            Dim oVGridControl As New GridControl
            oVGridControl = sender
            oVGridControl.ExportToXlsx(sFileName)
        Else
            Dim oGridView As New GridView
            oGridView = sender.FocusedView
            oGridView.OptionsPrint.AutoWidth = False
            oGridView.BestFitColumns()
            oGridView.BestFitMaxRowCount = oGridView.RowCount
            oGridView.ExportToXlsx(sFileName)
        End If
        If IO.File.Exists(sFileName) Then
            Dim oXls As New Excel.Application 'Crea el objeto excel 
            oXls.Workbooks.Open(sFileName, , False) 'El true es para abrir el archivo en modo Solo lectura (False si lo quieres de otro modo)
            oXls.Visible = True
            oXls.WindowState = Excel.XlWindowState.xlMaximized 'Para que la ventana aparezca maximizada.
        End If
    End Sub
    Friend Function GetTmpFileName(sExtFile As String) As String
        Dim sResult As String = ""
        Dim sPath As String = Path.GetTempPath
        sResult = (FileIO.FileSystem.GetTempFileName).Replace(".tmp", ".xlsx")
        Return sResult
    End Function
    Friend Sub TextToSpeak(sText As String)
        If My.Settings.AudioEnabled Then
            Dim t As New System.Threading.Thread(AddressOf SpeechThread)
            t.Start(sText)
        End If
    End Sub
    Private Sub SpeechThread(sText As String)
        Try
            Dim sapi
            sapi = CreateObject("sapi.spvoice")
            sapi.speak(sText)
        Catch ex As Exception
            My.Settings.AudioEnabled = False
            My.Settings.Save()
        End Try
    End Sub

End Module
