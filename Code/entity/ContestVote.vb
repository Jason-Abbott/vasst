<Serializable()> _
Public Class ContestVote
    Private _person As AMP.Person
    Private _rank As Integer
    Private _date As DateTime

#Region " Properties "

    Public Property [Date]() As DateTime
        Get
            Return _date
        End Get
        Set(ByVal Value As DateTime)
            _date = Value
        End Set
    End Property

    Public Property Rank() As Integer
        Get
            Return _rank
        End Get
        Set(ByVal Value As Integer)
            _rank = Value
        End Set
    End Property

    Public Property Person() As AMP.Person
        Get
            Return _person
        End Get
        Set(ByVal Value As AMP.Person)
            _person = Value
        End Set
    End Property

#End Region

    Public Sub New()
        _date = DateTime.Now
    End Sub
End Class
