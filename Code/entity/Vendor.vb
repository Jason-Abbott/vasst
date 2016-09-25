<Serializable()> _
Public Class Vendor
    Inherits AMP.Company
    Implements IComparable

    Private _id As Guid
    Private _description As String
    Private _fulfillmentEmail As String
    Private _contact As AMP.Person
    Private _products As AMP.ProductCollection

#Region " Properties "

    Public Property FulFillmentEmail() As String
        Get
            Return _fulfillmentEmail
        End Get
        Set(ByVal Value As String)
            _fulfillmentEmail = Value
        End Set
    End Property

    Public Property Contact() As AMP.Person
        Get
            Return _contact
        End Get
        Set(ByVal Value As AMP.Person)
            _contact = Value
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

    Public Property Products() As AMP.ProductCollection
        Get
            Return _products
        End Get
        Set(ByVal Value As AMP.ProductCollection)
            _products = Value
        End Set
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
        String.Compare(Me.Name, DirectCast(entity, AMP.Vendor).Name)
    End Function
End Class
