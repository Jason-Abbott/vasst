<Serializable()> _
Public Class Seminar
    Private _id As Guid
    Private _name As String
    Private _date As DateTime
    Private _address As AMP.Address
    Private _description As String
    Private _notes As String

    ' collection of time ranges for start/stop ?

#Region " Properties "

    Public Property Notes() As String
        Get
            Return _notes
        End Get
        Set(ByVal Value As String)
            _notes = Security.SafeString(Value, 1000)
        End Set
    End Property

    Public Property ID() As Guid
        Get
            Return _id
        End Get
        Set(ByVal Value As Guid)
            _id = Value
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

    Public Property Address() As AMP.Address
        Get
            Return _address
        End Get
        Set(ByVal Value As AMP.Address)
            _address = Value
        End Set
    End Property

    Public Property [Date]() As DateTime
        Get
            Return _date
        End Get
        Set(ByVal Value As DateTime)
            _date = Value
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
