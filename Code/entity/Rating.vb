<Serializable()> _
Public Class Rating

    Private _rankedBy As AMP.Person
    Private _rank As Single
    Private _comment As String
    Private _rankedOn As DateTime

#Region " Properties "

    Public Property [Date]() As DateTime
        Get
            Return _rankedOn
        End Get
        Set(ByVal Value As DateTime)
            _rankedOn = Value
        End Set
    End Property

    Public Property Comment() As String
        Get
            Return _comment
        End Get
        Set(ByVal Value As String)
            _comment = Security.SafeString(Value, 1500)
        End Set
    End Property

    Public Property Value() As Single
        Get
            Return _rank
        End Get
        Set(ByVal Value As Single)
            _rank = Value
        End Set
    End Property

    Public Property Person() As AMP.Person
        Get
            Return _rankedBy
        End Get
        Set(ByVal Value As AMP.Person)
            _rankedBy = Value
        End Set
    End Property

#End Region

End Class
