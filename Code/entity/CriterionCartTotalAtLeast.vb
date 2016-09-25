Namespace Criteria
    <Serializable()> _
    Public Class CriterionCartTotalAtLeast
        Implements ICriterion

        Private _threshold As Single

        Public Property Threshold() As Single
            Get
                Return _threshold
            End Get
            Set(ByVal Value As Single)
                _threshold = Value
            End Set
        End Property

        Public Function Qualified(ByVal user As Person) As Boolean Implements ICriterion.Qualified
            If _threshold = Nothing Then Return True
            If Not user.Cart Is Nothing AndAlso user.Cart.Total >= _threshold Then Return True
            Return False
        End Function

        Public Function Qualified() As Boolean Implements ICriterion.Qualified
            Return Me.Qualified(Profile.User)
        End Function
    End Class
End Namespace