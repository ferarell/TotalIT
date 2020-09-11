' NOTA: puede usar el comando "Cambiar nombre" del menú contextual para cambiar el nombre de clase "MailBotServicesWCF" en el código, en svc y en el archivo de configuración a la vez.
' NOTA: para iniciar el Cliente de prueba WCF para probar este servicio, seleccione MailBotServicesWCF.svc o MailBotServicesWCF.svc.vb en el Explorador de soluciones e inicie la depuración.

Imports MailBot.Entidades
Imports MailBot.LogicaNegocio
Public Class MailBotServicesWCF
    Implements IMailBotServicesWCF


    Dim oInputConfigurationBL As InputConfigurationBL = New InputConfigurationBL()
    Dim oInputSubjectBL As InputSubjectBL = New InputSubjectBL()
    Dim oInputTypeBL As InputTypeBL = New InputTypeBL()
    Dim oMessageConfigurationBL As MessageConfigurationBL = New MessageConfigurationBL()
    Dim oMessageLayoutBL As MessageLayoutBL = New MessageLayoutBL()
    Dim oMessageStructureBL As MessageStructureBL = New MessageStructureBL()

    Public Function ExecuteSQL(ByVal QueryString As String) As DataSet Implements IMailBotServicesWCF.ExecuteSQL
        Try
            Dim oDataAccess As New DataAccess
            Return oDataAccess.ExecuteSQL(QueryString)
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Public Function NewExecuteSQL(ByVal QueryString As String) As DataSet Implements IMailBotServicesWCF.NewExecuteSQL
        Dim dsResult As New DataSet
        Dim oDataAccess As New DataAccess
        dsResult = oDataAccess.NewExecuteSQL(QueryString)
        Return dsResult
    End Function

    Public Function ExecuteSQLNonQuery(ByVal QueryString As String) As Boolean Implements IMailBotServicesWCF.ExecuteSQLNonQuery
        Dim bResult As Boolean = True
        Try
            Dim oDataAccess As New DataAccess
            Return oDataAccess.ExecuteSQLNonQuery(QueryString)
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Public Function NewExecuteSQLNonQuery(ByVal QueryString As String) As ArrayList Implements IMailBotServicesWCF.NewExecuteSQLNonQuery
        Dim aResult As New ArrayList
        aResult.AddRange({True, ""})
        Try
            Dim oDataAccess As New DataAccess
            aResult(0) = oDataAccess.ExecuteSQLNonQuery(QueryString)
        Catch ex As Exception
            aResult(1) = ex.Message
        End Try
        Return aResult
    End Function

#Region "InputConfiguration"


    Public Sub InsertInputConfiguration(oInputConfiguration As InputConfiguration) Implements IMailBotServicesWCF.InsertInputConfiguration
        oInputConfigurationBL.Insert(oInputConfiguration)
    End Sub


    Public Sub UpdateInputConfiguration(oInputConfiguration As InputConfiguration) Implements IMailBotServicesWCF.UpdateInputConfiguration
        oInputConfigurationBL.Update(oInputConfiguration)
    End Sub


    Public Sub DeleteInputConfiguration(oInputConfiguration As InputConfiguration) Implements IMailBotServicesWCF.DeleteInputConfiguration
        oInputConfigurationBL.Delete(oInputConfiguration)
    End Sub


    Public Function GetAllInputConfiguration(oInputConfiguration As InputConfiguration) As List(Of InputConfiguration) Implements IMailBotServicesWCF.GetAllInputConfiguration
        Return oInputConfigurationBL.GetAll(oInputConfiguration)
    End Function


    Public Function GetAllDataTableInputConfiguration(oInputConfiguration As InputConfiguration) As DataTable Implements IMailBotServicesWCF.GetAllDataTableInputConfiguration
        Return oInputConfigurationBL.GetAllDataTable(oInputConfiguration)
    End Function


#End Region

#Region "InputSubject"


    Public Sub InsertInputSubject(oInputSubject As InputSubject) Implements IMailBotServicesWCF.InsertInputSubject
        oInputSubjectBL.Insert(oInputSubject)
    End Sub


    Public Sub UpdateInputSubject(oInputSubject As InputSubject) Implements IMailBotServicesWCF.UpdateInputSubject
        oInputSubjectBL.Update(oInputSubject)
    End Sub


    Public Sub DeleteInputSubject(oInputSubject As InputSubject) Implements IMailBotServicesWCF.DeleteInputSubject
        oInputSubjectBL.Delete(oInputSubject)
    End Sub


    Public Function GetAllInputSubject(oInputSubject As InputSubject) As List(Of InputSubject) Implements IMailBotServicesWCF.GetAllInputSubject
        Return oInputSubjectBL.GetAll(oInputSubject)
    End Function


    Public Function GetAllDataTableInputSubject(oInputSubject As InputSubject) As DataTable Implements IMailBotServicesWCF.GetAllDataTableInputSubject
        Return oInputSubjectBL.GetAllDataTable(oInputSubject)
    End Function


#End Region

#Region "InputType"


    Public Sub InsertInputType(oInputType As InputType) Implements IMailBotServicesWCF.InsertInputType
        oInputTypeBL.Insert(oInputType)
    End Sub


    Public Sub UpdateInputType(oInputType As InputType) Implements IMailBotServicesWCF.UpdateInputType
        oInputTypeBL.Update(oInputType)
    End Sub


    Public Sub DeleteInputType(oInputType As InputType) Implements IMailBotServicesWCF.DeleteInputType
        oInputTypeBL.Delete(oInputType)
    End Sub


    Public Function GetAllInputType(oInputType As InputType) As List(Of InputType) Implements IMailBotServicesWCF.GetAllInputType
        Return oInputTypeBL.GetAll(oInputType)
    End Function


    Public Function GetAllDataTableInputType(oInputType As InputType) As DataTable Implements IMailBotServicesWCF.GetAllDataTableInputType
        Return oInputTypeBL.GetAllDataTable(oInputType)
    End Function


#End Region

#Region "MessageConfiguration"


    Public Sub InsertMessageConfiguration(oMessageConfiguration As MessageConfiguration) Implements IMailBotServicesWCF.InsertMessageConfiguration
        oMessageConfigurationBL.Insert(oMessageConfiguration)
    End Sub


    Public Sub UpdateMessageConfiguration(oMessageConfiguration As MessageConfiguration) Implements IMailBotServicesWCF.UpdateMessageConfiguration
        oMessageConfigurationBL.Update(oMessageConfiguration)
    End Sub


    Public Sub DeleteMessageConfiguration(oMessageConfiguration As MessageConfiguration) Implements IMailBotServicesWCF.DeleteMessageConfiguration
        oMessageConfigurationBL.Delete(oMessageConfiguration)
    End Sub


    Public Function GetAllMessageConfiguration(oMessageConfiguration As MessageConfiguration) As List(Of MessageConfiguration) Implements IMailBotServicesWCF.GetAllMessageConfiguration
        Return oMessageConfigurationBL.GetAll(oMessageConfiguration)
    End Function


    Public Function GetAllDataTableMessageConfiguration(oMessageConfiguration As MessageConfiguration) As DataTable Implements IMailBotServicesWCF.GetAllDataTableMessageConfiguration
        Return oMessageConfigurationBL.GetAllDataTable(oMessageConfiguration)
    End Function


#End Region

#Region "MessageLayout"


    Public Sub InsertMessageLayout(oMessageLayout As MessageLayout) Implements IMailBotServicesWCF.InsertMessageLayout
        oMessageLayoutBL.Insert(oMessageLayout)
    End Sub


    Public Sub UpdateMessageLayout(oMessageLayout As MessageLayout) Implements IMailBotServicesWCF.UpdateMessageLayout
        oMessageLayoutBL.Update(oMessageLayout)
    End Sub


    Public Sub DeleteMessageLayout(oMessageLayout As MessageLayout) Implements IMailBotServicesWCF.DeleteMessageLayout
        oMessageLayoutBL.Delete(oMessageLayout)
    End Sub


    Public Function GetAllMessageLayout(oMessageLayout As MessageLayout) As List(Of MessageLayout) Implements IMailBotServicesWCF.GetAllMessageLayout
        Return oMessageLayoutBL.GetAll(oMessageLayout)
    End Function


    Public Function GetAllDataTableMessageLayout(oMessageLayout As MessageLayout) As DataTable Implements IMailBotServicesWCF.GetAllDataTableMessageLayout
        Return oMessageLayoutBL.GetAllDataTable(oMessageLayout)
    End Function


#End Region

#Region "MessageStructure"


    Public Sub InsertMessageStructure(oMessageStructure As MessageStructure) Implements IMailBotServicesWCF.InsertMessageStructure
        oMessageStructureBL.Insert(oMessageStructure)
    End Sub


    Public Sub UpdateMessageStructure(oMessageStructure As MessageStructure) Implements IMailBotServicesWCF.UpdateMessageStructure
        oMessageStructureBL.Update(oMessageStructure)
    End Sub


    Public Sub DeleteMessageStructure(oMessageStructure As MessageStructure) Implements IMailBotServicesWCF.DeleteMessageStructure
        oMessageStructureBL.Delete(oMessageStructure)
    End Sub


    Public Function GetAllMessageStructure(oMessageStructure As MessageStructure) As List(Of MessageStructure) Implements IMailBotServicesWCF.GetAllMessageStructure
        Return oMessageStructureBL.GetAll(oMessageStructure)
    End Function


    Public Function GetAllDataTableMessageStructure(oMessageStructure As MessageStructure) As DataTable Implements IMailBotServicesWCF.GetAllDataTableMessageStructure
        Return oMessageStructureBL.GetAllDataTable(oMessageStructure)
    End Function


#End Region

End Class
