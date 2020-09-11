Public Class InputConfiguration
    Private _IdInputConfiguration As Integer
    Private _InputConfigurationDescription As String
    Private _IdMessageConfiguration As Integer
    Private _IdInputType As Integer
    Private _ProcessFile As String
    Private _UserCreate As String
    Private _DateCreate As DateTime
    Private _UserUpdate As String
    Private _DateUpdate As DateTime

    Public Sub New()
    End Sub

    Public Sub New(IdInputConfiguration As Integer, InputConfigurationDescription As String, IdMessageConfiguration As Integer, IdInputType As Integer, ProcessFile As String, UserCreate As String, DateCreate As Date, UserUpdate As String, DateUpdate As Date)
        _IdInputConfiguration = IdInputConfiguration
        _InputConfigurationDescription = InputConfigurationDescription
        _IdMessageConfiguration = IdMessageConfiguration
        _IdInputType = IdInputType
        _ProcessFile = ProcessFile
        _UserCreate = UserCreate
        _DateCreate = DateCreate
        _UserUpdate = UserUpdate
        _DateUpdate = DateUpdate
    End Sub

    Public Property IdInputConfiguration As Integer
        Get
            Return _IdInputConfiguration
        End Get
        Set(value As Integer)
            _IdInputConfiguration = value
        End Set
    End Property

    Public Property InputConfigurationDescription As String
        Get
            Return _InputConfigurationDescription
        End Get
        Set(value As String)
            _InputConfigurationDescription = value
        End Set
    End Property

    Public Property IdMessageConfiguration As Integer
        Get
            Return _IdMessageConfiguration
        End Get
        Set(value As Integer)
            _IdMessageConfiguration = value
        End Set
    End Property

    Public Property IdInputType As Integer
        Get
            Return _IdInputType
        End Get
        Set(value As Integer)
            _IdInputType = value
        End Set
    End Property

    Public Property ProcessFile As String
        Get
            Return _ProcessFile
        End Get
        Set(value As String)
            _ProcessFile = value
        End Set
    End Property

    Public Property UserCreate As String
        Get
            Return _UserCreate
        End Get
        Set(value As String)
            _UserCreate = value
        End Set
    End Property

    Public Property DateCreate As Date
        Get
            Return _DateCreate
        End Get
        Set(value As Date)
            _DateCreate = value
        End Set
    End Property

    Public Property UserUpdate As String
        Get
            Return _UserUpdate
        End Get
        Set(value As String)
            _UserUpdate = value
        End Set
    End Property

    Public Property DateUpdate As Date
        Get
            Return _DateUpdate
        End Get
        Set(value As Date)
            _DateUpdate = value
        End Set
    End Property
End Class
