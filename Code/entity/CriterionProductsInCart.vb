Namespace Criteria
    <Serializable()> _
    Public Class CriterionProductsInCart
        Implements ICriterion

        Private _products As AMP.ProductFormatCollection

        Public Property Products() As AMP.ProductFormatCollection
            Get
                Return _products
            End Get
            Set(ByVal Value As AMP.ProductFormatCollection)
                _products = Value
            End Set
        End Property

        Public Function Qualified(ByVal user As Person) As Boolean Implements ICriterion.Qualified
            If _products Is Nothing Then Return True
            If Not user.Cart Is Nothing Then
                For Each p As AMP.ProductFormat In _products
                    If Not user.Cart.HasItem(p.ID) Then Return False
                Next
            End If
            Return False
        End Function

        Public Function Qualified() As Boolean Implements ICriterion.Qualified
            Return Me.Qualified(Profile.User)
        End Function
    End Class
End Namespace