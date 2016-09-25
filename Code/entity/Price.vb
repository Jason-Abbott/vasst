<Serializable()> _
Public Class Price
    Private _description As String
    Private _startOn As DateTime = DateTime.Now
    Private _endAfter As DateTime
    Private _enabled As Boolean = True
    Private _value As Single

#Region " Properties "

    '---COMMENT---------------------------------------------------------------
    '   price description, like "pre-order"
    '
    '	Date:		Name:	Description:
    '	3/7/05	    JEA		Creation
    '-------------------------------------------------------------------------
    Public Property Description() As String
        Get
            Return _description
        End Get
        Set(ByVal Value As String)
            _description = Security.SafeString(Value, 100)
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

    '---COMMENT---------------------------------------------------------------
    '   the monetary price
    '
    '	Date:		Name:	Description:
    '	3/7/05	    JEA		Creation
    '-------------------------------------------------------------------------
    Public Property Value() As Single
        Get
            Return _value
        End Get
        Set(ByVal Value As Single)
            _value = Value
        End Set
    End Property

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

#End Region

End Class
