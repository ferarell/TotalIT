Public Class MessageConfiguration



    Private _IdMessageConfiguration As Integer
    Private _IdInputType As Integer
    Private _Recipients As String
    Private _CopyRecipients As String
    Private _BlindCopyRecipients As String
    Private _CustomizedQuery As String
    Private _AttachedFiles As Byte()
    Private _Image As Byte()
    Private _ImageFile As Byte()
    Private _UserCreate As String
    Private _DateCreate As DateTime
    Private _UserUpdate As String
    Private _DateUpdate As DateTime
    Public Sub New()
    End Sub
    Public Sub New(IdMessageConfiguration As Integer, IdInputType As Integer, Recipients As String, CopyRecipients As String, BlindCopyRecipients As String, CustomizedQuery As String, AttachedFiles() As Byte, Image() As Byte, ImageFile() As Byte, UserCreate As String, DateCreate As Date, UserUpdate As String, DateUpdate As Date)
        _IdMessageConfiguration = IdMessageConfiguration
        _IdInputType = IdInputType
        _Recipients = Recipients
        _CopyRecipients = CopyRecipients
        _BlindCopyRecipients = BlindCopyRecipients
        _CustomizedQuery = CustomizedQuery
        _AttachedFiles = AttachedFiles
        _Image = Image
        _ImageFile = ImageFile
        _UserCreate = UserCreate
        _DateCreate = DateCreate
        _UserUpdate = UserUpdate
        _DateUpdate = DateUpdate
    End Sub

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

    Public Property Recipients As String
        Get
            Return _Recipients
        End Get
        Set(value As String)
            _Recipients = value
        End Set
    End Property

    Public Property CopyRecipients As String
        Get
            Return _CopyRecipients
        End Get
        Set(value As String)
            _CopyRecipients = value
        End Set
    End Property

    Public Property BlindCopyRecipients As String
        Get
            Return _BlindCopyRecipients
        End Get
        Set(value As String)
            _BlindCopyRecipients = value
        End Set
    End Property

    Public Property CustomizedQuery As String
        Get
            Return _CustomizedQuery
        End Get
        Set(value As String)
            _CustomizedQuery = value
        End Set
    End Property

    Public Property AttachedFiles As Byte()
        Get
            Return _AttachedFiles
        End Get
        Set(value As Byte())
            _AttachedFiles = value
        End Set
    End Property

    Public Property Image As Byte()
        Get
            Return _Image
        End Get
        Set(value As Byte())
            _Image = value
        End Set
    End Property

    Public Property ImageFile As Byte()
        Get
            Return _ImageFile
        End Get
        Set(value As Byte())
            _ImageFile = value
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
