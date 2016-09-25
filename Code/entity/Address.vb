<Serializable()> _
Public Class Address

    Private _street As String
    Private _city As String
    Private _state As Address.States
    Private _country As Address.Countries
    Private _name As String
    Private _zipCode As Integer
    Private _validated As Boolean = False
    Private _type As Integer                ' bitmask

#Region " Enumerations "

    <Flags()> _
    Public Enum Types
        Any = Billing Or Delivery
        Billing = &H2
        Delivery = &H4
    End Enum

    Public Enum States
        Alaska
        Alabama
        Arkansas
        Arizona
        California
        Colorado
        Connecticut
        Deleware
        Florida
        Georgia
        Hawaii
        Idaho
        Iowa
        Illinois
        Indiana
        Kansas
        Kentucky
        Louisiana
        Maine
        Maryland
        Massachusetts
        Michigan
        Minnesota
        Mississippi
        Missouri
        Montana
        North_Carolina
        North_Dakota
        Nebraska
        Nevada
        New_Hampshire
        New_Jersey
        New_Mexico
        New_York
        Ohio
        Oklahoma
        Oregon
        Pennsylvania
        Rhode_Island
        South_Carolina
        South_Dakota
        Tennessee
        Texas
        Utah
        Virginia
        Vermont
        Washington
        Wisconsin
        WestVirginia
        Wyoming
        Washington_DC
    End Enum

    Public Enum Countries
        UnitedStates
    End Enum

#End Region

#Region " Properties "

    Public Property Type() As Integer
        Get
            Return _type
        End Get
        Set(ByVal Value As Integer)
            _type = Value
        End Set
    End Property

    Public Property Validated() As Boolean
        Get
            Return _validated
        End Get
        Set(ByVal Value As Boolean)
            _validated = Value
        End Set
    End Property

    Public Property State() As Address.States
        Get
            Return _state
        End Get
        Set(ByVal Value As Address.States)
            _state = Value
        End Set
    End Property

    Public Property ZipCode() As Integer
        Get
            Return _zipCode
        End Get
        Set(ByVal Value As Integer)
            _zipCode = Value
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

    Public Property City() As String
        Get
            Return _city
        End Get
        Set(ByVal Value As String)
            _city = Security.SafeString(Value, 100)
        End Set
    End Property

    Public Property Street() As String
        Get
            Return _street
        End Get
        Set(ByVal Value As String)
            _street = Security.SafeString(Value, 150)
        End Set
    End Property

#End Region

End Class
