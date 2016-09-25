<Serializable()> _
Public Class Publisher
    Inherits AMP.Company
    Implements IComparable

    Private _id As Guid
    Private _description As String
    Private _logoFile As String
    Private _software As AMP.SoftwareCollection
    Private _forAssetType As Integer = 0    ' bitmask
    Private _section As Integer = 0         ' bitmask

#Region " Properties "

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

    Public Property Software() As AMP.SoftwareCollection
        Get
            Return _software
        End Get
        Set(ByVal Value As AMP.SoftwareCollection)
            _software = Value
        End Set
    End Property

    '---COMMENT---------------------------------------------------------------
    '	bitmask of site.section values
    '
    '	Date:		Name:	Description:
    '	12/26/04	JEA		Creation
    '-------------------------------------------------------------------------
    Public Property Section() As Integer
        Get
            Return _section
        End Get
        Set(ByVal Value As Integer)
            _section = Value
        End Set
    End Property

    Public Property LogoFile() As String
        Get
            Return _logoFile
        End Get
        Set(ByVal Value As String)
            _logoFile = Security.SafeString(Value, 50)
        End Set
    End Property

    Public Property Description() As String
        Get
            Return _description
        End Get
        Set(ByVal Value As String)
            _description = Security.SafeString(Value, 1500)
        End Set
    End Property

    Public ReadOnly Property ID() As Guid
        Get
            Return _id
        End Get
    End Property

#End Region

    Public Sub New()
        _id = Guid.NewGuid
    End Sub

    '---COMMENT---------------------------------------------------------------
    '	default compare sorts by name
    '
    '	Date:		Name:	Description:
    '	12/26/04	JEA		Creation
    '-------------------------------------------------------------------------
    Public Function CompareTo(ByVal entity As Object) As Integer Implements System.IComparable.CompareTo
        String.Compare(Me.Name, DirectCast(entity, AMP.Publisher).Name)
    End Function
End Class
