Imports MailBot.AccesoDatos
Imports MailBot.Entidades
Public Class InputSubjectBL
    Dim oInputSubjectDAO As InputSubjectDAO = New InputSubjectDAO()

    Public Sub Insert(oInputSubject As InputSubject)
        oInputSubjectDAO.Insert(oInputSubject)
    End Sub

    Public Sub Update(oInputSubject As InputSubject)
        oInputSubjectDAO.Update(oInputSubject)
    End Sub

    Public Sub Delete(oInputSubject As InputSubject)
        oInputSubjectDAO.Delete(oInputSubject)
    End Sub

    Public Function GetAll(oInputSubject As InputSubject) As List(Of InputSubject)
        Return oInputSubjectDAO.GetAll(oInputSubject)
    End Function

    Public Function GetAllDataTable(oInputSubject As InputSubject) As DataTable
        Return oInputSubjectDAO.GetAllDataTable(oInputSubject)
    End Function


End Class
