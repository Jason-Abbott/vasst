Namespace Criteria
    <Serializable()> _
    Public Class CriterionCodeEntered
        Implements ICriterion

        Private _code As String

        Public Property Code() As String
            Get
                Return _code
            End Get
            Set(ByVal Value As String)
                _code = Value
            End Set
        End Property

        Public Function Qualified(ByVal user As Person) As Boolean Implements ICriterion.Qualified
            Throw New NotImplementedException
        End Function

        Public Function Qualified() As Boolean Implements ICriterion.Qualified
            If _code = Nothing Then Return True
            Return (Profile.PromotionCode.ToLower = _code.ToLower)
        End Function
    End Class
End Namespace