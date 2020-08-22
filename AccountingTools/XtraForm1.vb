Public Class XtraForm1
    Private Sub XtraForm1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub SimpleButton1_Click(sender As Object, e As EventArgs) Handles SimpleButton1.Click
        LabelControl1.Text = GetMatchCode(TextEdit1.Text)
    End Sub

    Function GetMatchCode(StrInput As String) As String
        StrInput = Replace(StrInput, " ", "")
        Dim sResult, Str1, Str2 As String
        Dim iPos As Integer = Len(StrInput) + 1
        For p = 0 To Len(StrInput) - 1
            iPos -= 1
            If Not IsNumeric(Mid(StrInput, iPos, 1)) Then
                Exit For
            End If
        Next
        Str1 = Mid(StrInput, 1, iPos).Trim
        Str2 = CInt(Mid(StrInput, iPos + 1, Len(StrInput))).ToString
        sResult = Str1 + Space(1) + Str2
        Return sResult
    End Function
End Class