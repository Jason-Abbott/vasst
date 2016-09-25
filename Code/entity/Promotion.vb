<Serializable()> _
Public Class Promotion
    Private _name As String
    Private _description As String
    Private _criteria As AMP.CriteriaCollection
    Private _startOn As DateTime
    Private _endAfter As DateTime

#Region " Properties "

    Public Property StartOn() As DateTime
        Get
            Return _startOn
        End Get
        Set(ByVal Value As DateTime)
            _startOn = Value
        End Set
    End Property

    Public Property EndAfter() As DateTime
        Get
            Return _endAfter
        End Get
        Set(ByVal Value As DateTime)
            _endAfter = Value
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

    Public Property Name() As String
        Get
            Return _name
        End Get
        Set(ByVal Value As String)
            _name = Security.SafeString(Value, 100)
        End Set
    End Property

#End Region

End Class
