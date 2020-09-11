Imports MailBot.AccesoDatos
Imports MailBot.Entidades
Public Class MessageConfigurationBL
    Dim oMessageConfigurationDAO As MessageConfigurationDAO = New MessageConfigurationDAO()

    Public Sub Insert(oMessageConfiguration As MessageConfiguration)
        oMessageConfigurationDAO.Insert(oMessageConfiguration)
    End Sub

    Public Sub Update(oMessageConfiguration As MessageConfiguration)
        oMessageConfigurationDAO.Update(oMessageConfiguration)
    End Sub

    Public Sub Delete(oMessageConfiguration As MessageConfiguration)
        oMessageConfigurationDAO.Delete(oMessageConfiguration)
    End Sub

    Public Function GetAll(oMessageConfiguration As MessageConfiguration) As List(Of MessageConfiguration)
        Return oMessageConfigurationDAO.GetAll(oMessageConfiguration)
    End Function

    Public Function GetAllDataTable(oMessageConfiguration As MessageConfiguration) As DataTable
        Return oMessageConfigurationDAO.GetAllDataTable(oMessageConfiguration)
    End Function
End Class
