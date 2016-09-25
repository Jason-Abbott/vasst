<Serializable()> _
Public Class Phone
    Private _countryCode As Integer = 1     ' USA country code
    Private _areaCode As Integer
    Private _number As Integer
    Private _description As String
    Private _type As Phone.Types            ' bitmask

    Public Enum Types
        Home = &H1
        Mobile = &H2
        Office = &H4
        FAX = &H8
    End Enum

#Region " Properties "

    Public Property Type() As Phone.Types
        Get
            Return _type
        End Get
        Set(ByVal Value As Phone.Types)
            _type = Value
        End Set
    End Property

    Public Property Description() As String
        Get
            Return _description
        End Get
        Set(ByVal Value As String)
            _description = Security.SafeString(Value, 100)
        End Set
    End Property

    Public Property CountryCode() As Integer
        Get
            Return _countryCode
        End Get
        Set(ByVal Value As Integer)
            _countryCode = Value
        End Set
    End Property

    Public Property AreaCode() As Integer
        Get
            Return _areaCode
        End Get
        Set(ByVal Value As Integer)
            _areaCode = Value
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

    '---COMMENT---------------------------------------------------------------
    '	get sum of current prices for product formats in cart
    '
    '	Date:		Name:	Description:
    '   3/7/05	    JEA		Creation
    '-------------------------------------------------------------------------
    Public Function DisplayNumber(ByVal includeCountry As Boolean) As String
        ' TODO: format this
        Return _number.ToString
    End Function

    Public Function DisplayNumber() As String
        Return Me.DisplayNumber(False)
    End Function

End Class
