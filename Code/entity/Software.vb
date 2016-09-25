<Serializable()> _
Public Class Software
    Implements IComparable

    Private _id As Guid
    Private _name As String
    Private _versions As New AMP.VersionCollection
    Private _url As String
    Private _free As Boolean
    Private _publisher As AMP.Publisher
    Private _type As Software.Types      ' bitmask
    Private _extensions As New ArrayList

#Region " Enumerations "

    <Flags()> _
    Public Enum Types
        Plugin = VideoPlugin Or VstiPlugin Or DirectXPlugin
        VideoPlugin = &H1
        VSTiPlugin = &H2
        DirectXPlugin = &H4
        Application = &H10
    End Enum

#End Region

#Region " Properties "

    Public ReadOnly Property FullName() As String
        Get
            Return String.Format("{0} {1}", Me.Publisher.Name, Me.Name)
        End Get
    End Property

    Public Property Extensions() As ArrayList
        Get
            Return _extensions
        End Get
        Set(ByVal Value As ArrayList)
            _extensions = Value
        End Set
    End Property

    Public Property Publisher() As AMP.Publisher
        Get
            Return _publisher
        End Get
        Set(ByVal Value As AMP.Publisher)
            _publisher = Value
        End Set
    End Property

    Public Property Free() As Boolean
        Get
            Return _free
        End Get
        Set(ByVal Value As Boolean)
            _free = Value
        End Set
    End Property

    Public Property Type() As Software.Types
        Get
            Return _type
        End Get
        Set(ByVal Value As Software.Types)
            _type = Value
        End Set
    End Property

    Public ReadOnly Property ID() As Guid
        Get
            Return _id
        End Get
    End Property

    Public Property Name() As String
        Get
            Return _name
        End Get
        Set(ByVal Value As String)
            _name = Security.SafeString(Value, 100)
        End Set
    End Property

    Public Property Versions() As AMP.VersionCollection
        Get
            Return _versions
        End Get
        Set(ByVal Value As AMP.VersionCollection)
            _versions = Value
        End Set
    End Property

    Public Property Url() As String
        Get
            Return _url
        End Get
        Set(ByVal Value As String)
            _url = Security.SafeString(Value, 100)
        End Set
    End Property

#End Region

    Public Sub New()
        _id = Guid.NewGuid
    End Sub

    '---COMMENT---------------------------------------------------------------
    '	get publisher plus software name
    '
    '	Date:		Name:	Description:
    '	1/13/05    JEA		Creation
    '-------------------------------------------------------------------------
    Public Function FullNameLink() As String
        Return String.Format("<a href=""http://{0}"">{1}</a>", Me.Url, Me.FullName)
    End Function

    Public Function NameLink() As String
        Return String.Format("<a href=""http://{0}"">{1}</a>", Me.Url, Me.Name)
    End Function

    '---COMMENT---------------------------------------------------------------
    '	can this software open file of given extension
    '
    '	Date:		Name:	Description:
    '	12/28/04    JEA		Creation
    '-------------------------------------------------------------------------
    Public Function CanOpen(ByVal extension As String) As Boolean
        For Each e As String In _extensions
            If e = extension Then Return True
        Next
        Return False
    End Function

    Public Function CompareTo(ByVal entity As Object) As Integer Implements System.IComparable.CompareTo
        Dim s As AMP.Software = DirectCast(entity, AMP.Software)
        Dim compare As Integer
        compare = String.Compare(Me.Publisher.Name, s.Publisher.Name)
        If compare = 0 Then compare = String.Compare(Me.Name, s.Name)
        Return compare
    End Function
End Class