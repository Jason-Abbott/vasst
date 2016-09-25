Imports System.Text

<Serializable()> _
Public Class Category
    Implements IComparable

    Private _id As Guid
    Private _name As String
    Private _section As Integer = 0         ' bitmask
    Private _forSoftwareType As Integer = 0 ' bitmask
    Private _forAssetType As Integer = 0    ' bitmask
    Private _forEntity As Integer = 0       ' bitmask

#Region " Properties "

    Public Property ID() As Guid
        Get
            Return _id
        End Get
        Set(ByVal Value As Guid)
            _id = Value
        End Set
    End Property

    Public Property Name() As String
        Get
            Return _name
        End Get
        Set(ByVal Value As String)
            _name = Security.SafeString(Value, 50)
        End Set
    End Property

    Public Property Section() As Integer
        Get
            Return _section
        End Get
        Set(ByVal Value As Integer)
            _section = Value
        End Set
    End Property

    '---COMMENT---------------------------------------------------------------
    '	bitmask of asset.types values to filter which publisher items
    '   appear associated with different asset types
    '
    '	Date:		Name:	Description:
    '	12/26/04	JEA		Creation
    '-------------------------------------------------------------------------
    Public Property ForAssetType() As Integer
        Get
            Return _forAssetType
        End Get
        Set(ByVal Value As Integer)
            _forAssetType = Value
        End Set
    End Property

    '---COMMENT---------------------------------------------------------------
    '	bitmask of software.types values
    '
    '	Date:		Name:	Description:
    '	12/26/04	JEA		Creation
    '-------------------------------------------------------------------------
    Public Property ForSoftwareType() As Integer
        Get
            Return _forSoftwareType
        End Get
        Set(ByVal Value As Integer)
            _forSoftwareType = Value
        End Set
    End Property

    '---COMMENT---------------------------------------------------------------
    '	bitmask of site.entity values
    '
    '	Date:		Name:	Description:
    '	1/12/05 	JEA		Creation
    '-------------------------------------------------------------------------
    Public Property ForEntity() As Integer
        Get
            Return _forEntity
        End Get
        Set(ByVal Value As Integer)
            _forEntity = Value
        End Set
    End Property

#End Region

    Public Sub New()
        _id = Guid.NewGuid
    End Sub

    Public Function SearchLink() As String
        Return String.Format("<a href=""{0}/search.aspx?category={1}"">{2}</a>", _
            Global.BasePath, HttpUtility.UrlEncode(Me.Name), Me.Name)
    End Function

    Public Function CompareTo(ByVal entity As Object) As Integer Implements System.IComparable.CompareTo
        Return String.Compare(Me.Name, DirectCast(entity, AMP.Category).Name)
    End Function

End Class
