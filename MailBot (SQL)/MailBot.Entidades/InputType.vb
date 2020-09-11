Public Class InputType
    '   [] [smallint] Not NULL,
    '[] [nvarchar](150) NULL,
    '[UserCreate] [nvarchar](150) NULL,
    '[DateCreate] [datetime] NULL,
    '[UserUpdate] [nvarchar](150) NULL,
    '[DateUpdate] [datetime] NULL

    Private _IdInputType As Integer
    Private _InputTypeDescription As String
    Private _UserCreate As String
    Private _DateCreate As DateTime
    Private _UserUpdate As String
    Private _DateUpdate As DateTime
    Public Sub New()
    End Sub
    Public Sub New(IdInputType As Integer, InputTypeDescription As String, UserCreate As String, DateCreate As Date, UserUpdate As String, DateUpdate As Date)
        _IdInputType = IdInputType
        _InputTypeDescription = InputTypeDescription
        _UserCreate = UserCreate
        _DateCreate = DateCreate
        _UserUpdate = UserUpdate
        _DateUpdate = DateUpdate
    End Sub

    Public Property IdInputType As Integer
        Get
            Return _IdInputType
        End Get
        Set(value As Integer)
            _IdInputType = value
        End Set
    End Property

    Public Property InputTypeDescription As String
        Get
            Return _InputTypeDescription
        End Get
        Set(value As String)
            _InputTypeDescription = value
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
