Imports MailBot.AccesoDatos
Imports MailBot.Entidades

Public Class InputConfigurationBL
    Dim oInputConfigurationDAO As InputConfigurationDAO = New InputConfigurationDAO()

    Public Sub Insert(oInputConfiguration As InputConfiguration)
        oInputConfigurationDAO.Insert(oInputConfiguration)
    End Sub

    Public Sub Update(oInputConfiguration As InputConfiguration)
        oInputConfigurationDAO.Update(oInputConfiguration)
    End Sub

    Public Sub Delete(oInputConfiguration As InputConfiguration)
        oInputConfigurationDAO.Delete(oInputConfiguration)
    End Sub

    Public Function GetAll(oInputConfiguration As InputConfiguration) As List(Of InputConfiguration)
        Return oInputConfigurationDAO.GetAll(oInputConfiguration)
    End Function

    Public Function GetAllDataTable(oInputConfiguration As InputConfiguration) As DataTable
        Return oInputConfigurationDAO.GetAllDataTable(oInputConfiguration)
    End Function
End Class
