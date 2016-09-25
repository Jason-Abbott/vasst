<Serializable()> _
Public Class Company
    Private _name As String
    Private _webSite As String
    Private _address As AMP.Address

#Region " Properties "

    Public Property Address() As AMP.Address
        Get
            Return _address
        End Get
        Set(ByVal Value As AMP.Address)
            _address = Value
        End Set
    End Property

    Public Property WebSite() As String
        Get
            Return _webSite
        End Get
        Set(ByVal Value As String)
            _webSite = Security.SafeString(Value, 100)
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
