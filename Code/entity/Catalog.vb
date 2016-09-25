<Serializable()> _
Public Class Catalog
    Inherits ProductCollection

    Private _promotions As AMP.PromotionCollection
    Private _products As AMP.ProductCollection

#Region " Properties "

    Public Property Products() As AMP.ProductCollection
        Get
            Return _products
        End Get
        Set(ByVal Value As AMP.ProductCollection)
            _products = Value
        End Set
    End Property

    Public Property Promotions() As AMP.PromotionCollection
        Get
            Return _promotions
        End Get
        Set(ByVal Value As AMP.PromotionCollection)
            _promotions = Value
        End Set
    End Property

#End Region

End Class