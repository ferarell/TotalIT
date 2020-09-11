Public Class MessageStructure
    '   [] [int] Not NULL,
    '[] [nvarchar](200) NULL,
    '[] [smallint] NULL,
    '[] [bit] Not NULL,
    '[UserCreate] [nvarchar](150) NULL,
    '[DateCreate] [datetime] NULL,
    '[UserUpdate] [nvarchar](150) NULL,
    '[DateUpdate] [datetime] NULL

    Private _IdMessageStructure As Integer
    Private _MessageStructureDescription As String
    Private _IndexLocation As Integer
    Private _Signature As Boolean
    Private _UserCreate As String
    Private _DateCreate As DateTime
    Private _UserUpdate As String
    Private _DateUpdate As DateTime
    Public Sub New()
    End Sub
    Public Sub New(IdMessageStructure As Integer, MessageStructureDescription As String, IndexLocation As Integer, Signature As Boolean, UserCreate As String, DateCreate As Date, UserUpdate As String, DateUpdate As Date)
        _IdMessageStructure = IdMessageStructure
        _MessageStructureDescription = MessageStructureDescription
        _IndexLocation = IndexLocation
        _Signature = Signature
        _UserCreate = UserCreate
        _DateCreate = DateCreate
        _UserUpdate = UserUpdate
        _DateUpdate = DateUpdate
    End Sub

    Public Property IdMessageStructure As Integer
        Get
            Return _IdMessageStructure
        End Get
        Set(value As Integer)
            _IdMessageStructure = value
        End Set
    End Property

    Public Property MessageStructureDescription As String
        Get
            Return _MessageStructureDescription
        End Get
        Set(value As String)
            _MessageStructureDescription = value
        End Set
    End Property

    Public Property IndexLocation As Integer
        Get
            Return _IndexLocation
        End Get
        Set(value As Integer)
            _IndexLocation = value
        End Set
    End Property

    Public Property Signature As Boolean
        Get
            Return _Signature
        End Get
        Set(value As Boolean)
            _Signature = value
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
