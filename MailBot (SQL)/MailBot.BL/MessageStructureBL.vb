Imports MailBot.AccesoDatos
Imports MailBot.Entidades
Public Class MessageStructureBL
    Dim oMessageStructureDAO As MessageStructureDAO = New MessageStructureDAO()

    Public Sub Insert(oMessageStructure As MessageStructure)
        oMessageStructureDAO.Insert(oMessageStructure)
    End Sub

    Public Sub Update(oMessageStructure As MessageStructure)
        oMessageStructureDAO.Update(oMessageStructure)
    End Sub

    Public Sub Delete(oMessageStructure As MessageStructure)
        oMessageStructureDAO.Delete(oMessageStructure)
    End Sub

    Public Function GetAll(oMessageStructure As MessageStructure) As List(Of MessageStructure)
        Return oMessageStructureDAO.GetAll(oMessageStructure)
    End Function

    Public Function GetAllDataTable(oMessageStructure As MessageStructure) As DataTable
        Return oMessageStructureDAO.GetAllDataTable(oMessageStructure)
    End Function
End Class
