Imports MailBot.AccesoDatos
Imports MailBot.Entidades
Public Class MessageLayoutBL
    Dim oMessageLayoutDAO As MessageLayoutDAO = New MessageLayoutDAO()

    Public Sub Insert(oMessageLayout As MessageLayout)
        oMessageLayoutDAO.Insert(oMessageLayout)
    End Sub

    Public Sub Update(oMessageLayout As MessageLayout)
        oMessageLayoutDAO.Update(oMessageLayout)
    End Sub

    Public Sub Delete(oMessageLayout As MessageLayout)
        oMessageLayoutDAO.Delete(oMessageLayout)
    End Sub

    Public Function GetAll(oMessageLayout As MessageLayout) As List(Of MessageLayout)
        Return oMessageLayoutDAO.GetAll(oMessageLayout)
    End Function

    Public Function GetAllDataTable(oMessageLayout As MessageLayout) As DataTable
        Return oMessageLayoutDAO.GetAllDataTable(oMessageLayout)
    End Function
End Class
