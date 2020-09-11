Imports MailBot.AccesoDatos
Imports MailBot.Entidades
Public Class InputTypeBL
    Dim oInputTypeDAO As InputTypeDAO = New InputTypeDAO()

    Public Sub Insert(oInputType As InputType)
        oInputTypeDAO.Insert(oInputType)
    End Sub

    Public Sub Update(oInputType As InputType)
        oInputTypeDAO.Update(oInputType)
    End Sub

    Public Sub Delete(oInputType As InputType)
        oInputTypeDAO.Delete(oInputType)
    End Sub

    Public Function GetAll(oInputType As InputType) As List(Of InputType)
        Return oInputTypeDAO.GetAll(oInputType)
    End Function

    Public Function GetAllDataTable(oInputType As InputType) As DataTable
        Return oInputTypeDAO.GetAllDataTable(oInputType)
    End Function
End Class
