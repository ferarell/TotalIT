Imports System.ServiceModel
Imports MailBot.Entidades

' NOTA: puede usar el comando "Cambiar nombre" del menú contextual para cambiar el nombre de interfaz "IMailBotServicesWCF" en el código y en el archivo de configuración a la vez.
<ServiceContract()>
Public Interface IMailBotServicesWCF

    <OperationContract()>
    Function ExecuteSQL(ByVal QueryString As String) As DataSet

    <OperationContract()>
    Function NewExecuteSQL(ByVal QueryString As String) As DataSet

    <OperationContract()>
    Function ExecuteSQLNonQuery(ByVal QueryString As String) As Boolean

    <OperationContract()>
    Function NewExecuteSQLNonQuery(ByVal QueryString As String) As ArrayList

#Region "InputConfiguration"

    <OperationContract()>
    Sub InsertInputConfiguration(oInputConfiguration As InputConfiguration)

    <OperationContract()>
    Sub UpdateInputConfiguration(oInputConfiguration As InputConfiguration)

    <OperationContract()>
    Sub DeleteInputConfiguration(oInputConfiguration As InputConfiguration)

    <OperationContract()>
    Function GetAllInputConfiguration(oInputConfiguration As InputConfiguration) As List(Of InputConfiguration)

    <OperationContract()>
    Function GetAllDataTableInputConfiguration(oInputConfiguration As InputConfiguration) As DataTable



#End Region

#Region "InputSubject"

    <OperationContract()>
    Sub InsertInputSubject(oInputSubject As InputSubject)

    <OperationContract()>
    Sub UpdateInputSubject(oInputSubject As InputSubject)

    <OperationContract()>
    Sub DeleteInputSubject(oInputSubject As InputSubject)

    <OperationContract()>
    Function GetAllInputSubject(oInputSubject As InputSubject) As List(Of InputSubject)

    <OperationContract()>
    Function GetAllDataTableInputSubject(oInputSubject As InputSubject) As DataTable



#End Region

#Region "InputType"

    <OperationContract()>
    Sub InsertInputType(oInputType As InputType)

    <OperationContract()>
    Sub UpdateInputType(oInputType As InputType)

    <OperationContract()>
    Sub DeleteInputType(oInputType As InputType)

    <OperationContract()>
    Function GetAllInputType(oInputType As InputType) As List(Of InputType)

    <OperationContract()>
    Function GetAllDataTableInputType(oInputType As InputType) As DataTable


#End Region

#Region "MessageConfiguration"

    <OperationContract()>
    Sub InsertMessageConfiguration(oMessageConfiguration As MessageConfiguration)

    <OperationContract()>
    Sub UpdateMessageConfiguration(oMessageConfiguration As MessageConfiguration)

    <OperationContract()>
    Sub DeleteMessageConfiguration(oMessageConfiguration As MessageConfiguration)

    <OperationContract()>
    Function GetAllMessageConfiguration(oMessageConfiguration As MessageConfiguration) As List(Of MessageConfiguration)

    <OperationContract()>
    Function GetAllDataTableMessageConfiguration(oMessageConfiguration As MessageConfiguration) As DataTable



#End Region

#Region "MessageLayout"

    <OperationContract()>
    Sub InsertMessageLayout(oMessageLayout As MessageLayout)

    <OperationContract()>
    Sub UpdateMessageLayout(oMessageLayout As MessageLayout)

    <OperationContract()>
    Sub DeleteMessageLayout(oMessageLayout As MessageLayout)

    <OperationContract()>
    Function GetAllMessageLayout(oMessageLayout As MessageLayout) As List(Of MessageLayout)

    <OperationContract()>
    Function GetAllDataTableMessageLayout(oMessageLayout As MessageLayout) As DataTable



#End Region

#Region "MessageStructure"

    <OperationContract()>
    Sub InsertMessageStructure(oMessageStructure As MessageStructure)

    <OperationContract()>
    Sub UpdateMessageStructure(oMessageStructure As MessageStructure)

    <OperationContract()>
    Sub DeleteMessageStructure(oMessageStructure As MessageStructure)

    <OperationContract()>
    Function GetAllMessageStructure(oMessageStructure As MessageStructure) As List(Of MessageStructure)

    <OperationContract()>
    Function GetAllDataTableMessageStructure(oMessageStructure As MessageStructure) As DataTable




#End Region

End Interface
