<Serializable()> _
Public Class ProductFormat

    Private _id As Guid
    Private _isbn As Integer
    Private _upc As Integer
    Private _product As AMP.Product
    Private _weight As Single
    Private _description As String
    Private _preOrderUntil As DateTime
    Private _inventory As Integer
    Private _file As AMP.File
    Private _tangible As Boolean
    Private _enabled As Boolean = True
    Private _maxPurchasable As Integer
    Private _maxInCart As Integer
    Private _price As AMP.PriceCollection

#Region " Properties "

    Public Property PreOrderUntil() As DateTime
        Get
            Return _preOrderUntil
        End Get
        Set(ByVal Value As DateTime)
            _preOrderUntil = Value
        End Set
    End Property

    Public ReadOnly Property ID() As Guid
        Get
            Return _id
        End Get
    End Property

    Public Property UPC() As Integer
        Get
            Return _upc
        End Get
        Set(ByVal Value As Integer)
            _upc = Value
        End Set
    End Property

    Public Property ISBN() As Integer
        Get
            Return _isbn
        End Get
        Set(ByVal Value As Integer)
            _isbn = Value
        End Set
    End Property

    '---COMMENT---------------------------------------------------------------
    '   weight can be used to calculate shipping
    '
    '	Date:		Name:	Description:
    '	3/7/05	    JEA		Creation
    '-------------------------------------------------------------------------
    Public Property Weight() As Single
        Get
            Return _weight
        End Get
        Set(ByVal Value As Single)
            _weight = Value
        End Set
    End Property

    '---COMMENT---------------------------------------------------------------
    '   reference to the product to which this format belongs
    '
    '	Date:		Name:	Description:
    '	3/7/05	    JEA		Creation
    '-------------------------------------------------------------------------
    Public Property Product() As AMP.Product
        Get
            Return _product
        End Get
        Set(ByVal Value As AMP.Product)
            _product = Value
        End Set
    End Property

    '---COMMENT---------------------------------------------------------------
    '   tangible usually indicates some shipping required
    '
    '	Date:		Name:	Description:
    '	3/7/05	    JEA		Creation
    '-------------------------------------------------------------------------
    Public Property Tangible() As Boolean
        Get
            Return _tangible
        End Get
        Set(ByVal Value As Boolean)
            _tangible = Value
        End Set
    End Property

    '---COMMENT---------------------------------------------------------------
    '   limit how many customer may place in cart
    '
    '	Date:		Name:	Description:
    '	3/7/05	    JEA		Creation
    '-------------------------------------------------------------------------
    Public Property MaximumInCart() As Integer
        Get
            Return _maxInCart
        End Get
        Set(ByVal Value As Integer)
            _maxInCart = Value
        End Set
    End Property

    '---COMMENT---------------------------------------------------------------
    '   limit how many customer may purchase ever
    '
    '	Date:		Name:	Description:
    '	3/7/05	    JEA		Creation
    '-------------------------------------------------------------------------
    Public Property MaximumPurchasable() As Integer
        Get
            Return _maxPurchasable
        End Get
        Set(ByVal Value As Integer)
            _maxPurchasable = Value
        End Set
    End Property

    Public Property Enabled() As Boolean
        Get
            Return _enabled
        End Get
        Set(ByVal Value As Boolean)
            _enabled = Value
        End Set
    End Property

    Public Property Inventory() As Integer
        Get
            Return _inventory
        End Get
        Set(ByVal Value As Integer)
            _inventory = Value
        End Set
    End Property

    Public Property Price() As AMP.PriceCollection
        Get
            Return _price
        End Get
        Set(ByVal Value As AMP.PriceCollection)
            _price = Value
        End Set
    End Property

    Public Property Description() As String
        Get
            Return _description
        End Get
        Set(ByVal Value As String)
            _description = Security.SafeString(Value, 100)
        End Set
    End Property

#End Region

    Public Sub New()
        _id = Guid.NewGuid
    End Sub

    '---COMMENT---------------------------------------------------------------
    '   figure out how many of this format are available for purchase
    '
    '	Date:		Name:	Description:
    '	3/7/05	    JEA		Creation
    '-------------------------------------------------------------------------
    Public Function Available(ByVal user As AMP.Person) As Integer
        If Not _enabled Then Return 0
        If _price.Current = Nothing Then Return 0

        Dim count As Integer = _inventory

        ' TODO: get user counts for real comparisons
        If _maxPurchasable <> Nothing AndAlso _
            count > _maxPurchasable Then count = _maxPurchasable

        If _maxInCart <> Nothing AndAlso _
            count > _maxInCart Then count = _maxInCart

        Return count
    End Function

    Public Function Available() As Integer
        Return Me.Available(Profile.User)
    End Function
End Class
