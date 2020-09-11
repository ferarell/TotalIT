Imports System.Runtime.Serialization



<DataContract()>
Public Class MessageLayout

    Private _IdMessageLayout As Integer
    Private _IdMessageStructure As Integer
    Private _IdInputConfiguration As Integer
    Private _MessageText As String
    Private _ValidFrom As Nullable(Of DateTime)
    Private _ValidTo As Nullable(Of DateTime)
    Private _UserCreate As String
    Private _DateCreate As Nullable(Of DateTime)
    Private _UserUpdate As String
    Private _DateUpdate As Nullable(Of DateTime)
    Public Sub New()
    End Sub
    Public Sub New(IdMessageLayout As Integer, IdMessageStructure As Integer, IdInputConfiguration As Integer, MessageText As String, ValidFrom As Date, ValidTo As Date, UserCreate As String, DateCreate As Date, UserUpdate As String, DateUpdate As Date)
        _IdMessageLayout = IdMessageLayout
        _IdMessageStructure = IdMessageStructure
        _IdInputConfiguration = IdInputConfiguration
        _MessageText = MessageText
        _ValidFrom = ValidFrom
        _ValidTo = ValidTo
        _UserCreate = UserCreate
        _DateCreate = DateCreate
        _UserUpdate = UserUpdate
        _DateUpdate = DateUpdate
    End Sub
    <DataMember()>
    Public Property IdMessageLayout As Integer
        Get
            Return _IdMessageLayout
        End Get
        Set(value As Integer)
            _IdMessageLayout = value
        End Set
    End Property
    <DataMember()>
    Public Property IdMessageStructure As Integer
        Get
            Return _IdMessageStructure
        End Get
        Set(value As Integer)
            _IdMessageStructure = value
        End Set
    End Property
    <DataMember()>
    Public Property IdInputConfiguration As Integer
        Get
            Return _IdInputConfiguration
        End Get
        Set(value As Integer)
            _IdInputConfiguration = value
        End Set
    End Property
    <DataMember()>
    Public Property MessageText As String
        Get
            Return _MessageText
        End Get
        Set(value As String)
            _MessageText = value
        End Set
    End Property
    <DataMember()>
    Public Property ValidFrom As Nullable(Of DateTime)
        Get
            Return _ValidFrom
        End Get
        Set(value As Nullable(Of DateTime))
            _ValidFrom = value
        End Set
    End Property
    <DataMember()>
    Public Property ValidTo As Nullable(Of DateTime)
        Get
            Return _ValidTo
        End Get
        Set(value As Nullable(Of DateTime))
            _ValidTo = value
        End Set
    End Property
    <DataMember()>
    Public Property UserCreate As String
        Get
            Return _UserCreate
        End Get
        Set(value As String)
            _UserCreate = value
        End Set
    End Property
    <DataMember()>
    Public Property DateCreate As Nullable(Of DateTime)
        Get
            Return _DateCreate
        End Get
        Set(value As Nullable(Of DateTime))
            _DateCreate = value
        End Set
    End Property
    <DataMember()>
    Public Property UserUpdate As String
        Get
            Return _UserUpdate
        End Get
        Set(value As String)
            _UserUpdate = value
        End Set
    End Property
    <DataMember()>
    Public Property DateUpdate As Nullable(Of DateTime)
        Get
            Return _DateUpdate
        End Get
        Set(value As Nullable(Of DateTime))
            _DateUpdate = value
        End Set
    End Property
End Class
