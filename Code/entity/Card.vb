<Serializable()> _
Public Class Card
    Private _number As Integer
    Private _cvv As Integer
    Private _expireDate As DateTime
    Private _type As Card.Types

    Public Enum Types
        VISA
        Discover
        MasterCard
        AmericanExpress
    End Enum

#Region " Properties "

    Public Property Type() As Card.Types
        Get
            Return _type
        End Get
        Set(ByVal Value As Card.Types)
            _type = Value
        End Set
    End Property

    Public Property Expire() As DateTime
        Get
            Return _expireDate
        End Get
        Set(ByVal Value As DateTime)
            _expireDate = Value
        End Set
    End Property

    Public Property CVV() As Integer
        Get
            Return _cvv
        End Get
        Set(ByVal Value As Integer)
            _cvv = Value
        End Set
    End Property


    Public Property Number() As Integer
        Get
            Return _number
        End Get
        Set(ByVal Value As Integer)
            _number = Value
        End Set
    End Property

#End Region

End Class
