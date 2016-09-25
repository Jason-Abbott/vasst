<Serializable()> _
Public Class CartItem
    Dim _product As AMP.ProductFormat
    Dim _addedOn As DateTime
    Dim _quantity As Integer

#Region " Properties "

    Public Property AddedOn() As DateTime
        Get
            Return _addedOn
        End Get
        Set(ByVal Value As DateTime)
            _addedOn = Value
        End Set
    End Property

    '---COMMENT---------------------------------------------------------------
    '	use product formats since that's what the customer actually buys
    '
    '	Date:		Name:	Description:
    '   3/7/05	    JEA		Creation
    '-------------------------------------------------------------------------
    Public Property Product() As AMP.ProductFormat
        Get
            Return _product
        End Get
        Set(ByVal Value As AMP.ProductFormat)
            _product = Value
        End Set
    End Property

    Public Property Quantity() As Integer
        Get
            Return _quantity
        End Get
        Set(ByVal Value As Integer)
            _quantity = Value
        End Set
    End Property

#End Region

    Public Function ExtendedPrice() As Single
        Return _product.Price.Current * Quantity
    End Function

    Public Function Weight() As Single
        Return _product.Weight * Quantity
    End Function
End Class
