Namespace Criteria
    <Serializable()> _
    Public Class CriterionSignupBefore
        Implements ICriterion

        Private _threshold As DateTime

        Public Property Threshold() As DateTime
            Get
                Return _threshold
            End Get
            Set(ByVal Value As DateTime)
                _threshold = Value
            End Set
        End Property

        Public Function Qualified(ByVal user As Person) As Boolean Implements ICriterion.Qualified
            If _threshold = Nothing Then Return True
            Return (user.RegisteredOn < _threshold)
        End Function

        Public Function Qualified() As Boolean Implements ICriterion.Qualified
            Return Me.Qualified(Profile.User)
        End Function
    End Class
End Namespace