Imports System.Web.HttpContext
Imports System.Collections.Specialized
Imports AMP.Global

Public Class Profile

    Private Shared _useCookies As Boolean = True

#Region " Session and Cookie Keys "

    Private Shared _userKey As String = "User"
    Private Shared _personKey As String = "PersonID"
    Private Shared _messageKey As String = "Message"
    Private Shared _offsetKey As String = "TimeOffset"
    Private Shared _lastLoginKey As String = "LastLogin"
    Private Shared _authenticatedKey As String = "Authenticated"
    Private Shared _destinationKey As String = "DestinationPage"
    Private Shared _resultsKey As String = "SearchResults"
    Private Shared _contributeKey As String = "Contribution"
    Private Shared _formValuesKey As String = "FormValues"
    Private Shared _resumeDownload As String = "ResumeDownloadID"
    Private Shared _permissionsKey As String = "Permissions"
    Private Shared _testCookie As String = "test"
    Private Shared _promoCodeKey As String = "PromoCode"

#End Region

#Region " Properties "

    '---COMMENT---------------------------------------------------------------
    '	hold any promotion code entered
    '
    '	Date:		Name:	Description:
    '	3/8/05      JEA		Creation
    '-------------------------------------------------------------------------
    Public Shared Property PromotionCode() As String
        Get
            Return Current.Session(_promoCodeKey).ToString
        End Get
        Set(ByVal Value As String)
            If Value Is Nothing Then
                Current.Session.Remove(_promoCodeKey)
            Else
                Current.Session(_promoCodeKey) = Security.SafeString(Value, 25)
            End If
        End Set
    End Property

    '---COMMENT---------------------------------------------------------------
    '	used to cache permissions array for duration of session
    '
    '	Date:		Name:	Description:
    '	1/29/05     JEA		Creation
    '-------------------------------------------------------------------------
    Public Shared Property Permissions() As AMP.Site.Permission()
        Get
            If Current.Session(_permissionsKey) Is Nothing Then
                Return Nothing
            Else
                Return DirectCast(Current.Session(_permissionsKey), AMP.Site.Permission())
            End If
        End Get
        Set(ByVal Value As AMP.Site.Permission())
            If Value Is Nothing Then
                Current.Session.Remove(_permissionsKey)
            Else
                Current.Session(_permissionsKey) = Value
            End If
        End Set
    End Property

    '---COMMENT---------------------------------------------------------------
    '	hold viewed asset ID when redirected to login
    '
    '	Date:		Name:	Description:
    '	1/23/05     JEA		Creation
    '-------------------------------------------------------------------------
    Public Shared Property ResumeDownload() As Guid
        Get
            If Current.Session(_resumeDownload) Is Nothing Then
                Return Nothing
            Else
                Return DirectCast(Current.Session(_resumeDownload), Guid)
            End If
        End Get
        Set(ByVal Value As Guid)
            If Value.Equals(Guid.Empty) Then
                Current.Session.Remove(_resumeDownload)
            Else
                Current.Session(_resumeDownload) = Value
            End If
        End Set
    End Property

    '---COMMENT---------------------------------------------------------------
    '	hold form post values when redirected to login
    '
    '	Date:		Name:	Description:
    '	1/19/05     JEA		Creation
    '-------------------------------------------------------------------------
    Public Shared Property FormValues() As NameValueCollection
        Get
            If Current.Session(_formValuesKey) Is Nothing Then
                Return Nothing
            Else
                Return DirectCast(Current.Session(_formValuesKey), NameValueCollection)
            End If
        End Get
        Set(ByVal Value As NameValueCollection)
            If Value Is Nothing Then
                Current.Session.Remove(_formValuesKey)
            Else
                Current.Session(_formValuesKey) = Value
            End If
        End Set
    End Property

    '---COMMENT---------------------------------------------------------------
    '	holds asset contribution during wizard steps
    '
    '	Date:		Name:	Description:
    '	12/20/04    JEA		Creation
    '-------------------------------------------------------------------------
    Public Shared Property Contribution() As AMP.AssetContribution
        Get
            If Current.Session(_contributeKey) Is Nothing Then
                Dim ac As New AMP.AssetContribution
                Current.Session(_contributeKey) = ac
                Return ac
            Else
                Return DirectCast(Current.Session(_contributeKey), AMP.AssetContribution)
            End If
        End Get
        Set(ByVal Value As AMP.AssetContribution)
            If Value Is Nothing Then
                Current.Session(_contributeKey) = New AMP.AssetContribution
            Else
                Current.Session(_contributeKey) = Value
            End If
        End Set
    End Property

    Public Shared Property SearchResults() As ArrayList
        Get
            If Current.Session(_resultsKey) Is Nothing Then
                Return Nothing
            Else
                Return DirectCast(Current.Session(_resultsKey), ArrayList)
            End If
        End Get
        Set(ByVal Value As ArrayList)
            If Value Is Nothing Then
                Current.Session.Remove(_resultsKey)
            Else
                Current.Session(_resultsKey) = Value
            End If
        End Set
    End Property

    Public Shared Property UseCookies() As Boolean
        Get
            Return _useCookies
        End Get
        Set(ByVal Value As Boolean)
            _useCookies = Value
        End Set
    End Property

    Public Shared Property Message() As String
        Get
            If Current.Session(_messageKey) Is Nothing Then
                Return Nothing
            Else
                Dim Value As String
                Value = CStr(Current.Session(_messageKey))
                Current.Session.Remove(_messageKey)
                Current.Response.Cookies.Remove("message")
                Return Value
            End If
        End Get
        Set(ByVal Value As String)
            If Value = Nothing Then
                Current.Session.Remove(_messageKey)
            Else
                Current.Session(_messageKey) = Value
            End If
            ' store in cookie to be used for output cache variance
            Current.Response.Cookies.Add(New HttpCookie("message", Value))
        End Set
    End Property

    '---COMMENT---------------------------------------------------------------
    '	Tracks time difference between client and server so times can be
    '   displayed relative to client locale
    '
    '	Date:		Name:	Description:
    '	12/3/04     JEA		Creation
    '-------------------------------------------------------------------------
    Public Shared Property TimeOffset() As TimeSpan
        Get
            If Current.Session(_offsetKey) Is Nothing Then
                If Current.Request.Cookies(_offsetKey) Is Nothing Then
                    Return New TimeSpan(0)
                Else
                    ' get offset from cookie and store in session
                    Dim offset As New TimeSpan
                    offset = DirectCast(Current.Request.Cookies(_offsetKey).Value, TimeSpan)
                    Current.Session(_offsetKey) = offset
                    Return offset
                End If
            Else
                Return DirectCast(Current.Session(_offsetKey), TimeSpan)
            End If
        End Get
        Set(ByVal Value As TimeSpan)
            Current.Session(_offsetKey) = Value
            SetCookie(_offsetKey, Value)
        End Set
    End Property

    '---COMMENT---------------------------------------------------------------
    '   Use last login to determine which assets are new to this user
    '
    '	Date:		Name:	Description:
    '	12/3/04     JEA		Creation
    '-------------------------------------------------------------------------
    Public Shared Property LastLogin() As DateTime
        Get
            If Current.Session(_offsetKey) Is Nothing Then
                Return Nothing
            Else
                Return DirectCast(Current.Session(_offsetKey), DateTime)
            End If
        End Get
        Set(ByVal Value As DateTime)
            Current.Session(_offsetKey) = Value
        End Set
    End Property

    Public Shared Property Authenticated() As Boolean
        Get
            If Current.Session(_authenticatedKey) Is Nothing Then
                Return False
            Else
                Return CBool(Current.Session(_authenticatedKey))
            End If
        End Get
        Set(ByVal Value As Boolean)
            Current.Session(_authenticatedKey) = Value
        End Set
    End Property

    Public Shared Property User() As AMP.Person
        Get
            If Current.Session(_userKey) Is Nothing Then
                Return Nothing
            Else
                Return DirectCast(Current.Session(_userKey), AMP.Person)
            End If
        End Get
        Set(ByVal Value As AMP.Person)
            If Value Is Nothing Then
                Current.Session.Remove(_userKey)
            Else
                Current.Session(_userKey) = Value
                Profile.PersonID = Value.ID
            End If
        End Set
    End Property

    '---COMMENT---------------------------------------------------------------
    '   Get person ID from session or cookie
    '
    '	Date:		Name:	Description:
    '	12/3/04     JEA		Creation
    '   12/7/04     JEA     Also persist to cookie
    '-------------------------------------------------------------------------
    Public Shared Property PersonID() As Guid
        Get
            If Current.Session(_personKey) Is Nothing Then
                If Current.Request.Cookies(_personKey) Is Nothing OrElse _
                    Current.Request.Cookies(_personKey).Value.Length < 32 Then

                    Return Nothing
                Else
                    ' looks like a guid
                    Return New Guid(Current.Request.Cookies(_personKey).Value)
                End If
            Else
                Return New Guid(Current.Session(_personKey).ToString)
            End If
        End Get
        Set(ByVal Value As Guid)
            If Value.Equals(Guid.Empty) Then
                Current.Session.Remove(_personKey)
            Else
                Current.Session(_personKey) = Value
                SetCookie(_personKey, Value)
            End If
        End Set
    End Property

    '---COMMENT---------------------------------------------------------------
    '	Stores intended page when redirected to login for authentication
    '
    '	Date:		Name:	Description:
    '	9/22/04     JEA		Creation
    '-------------------------------------------------------------------------
    Public Shared Property DestinationPage() As String
        Get
            If Current.Session(_destinationKey) Is Nothing Then
                Return Nothing
            Else
                Return CStr(Current.Session(_destinationKey))
            End If
        End Get
        Set(ByVal Value As String)
            If Value = Nothing Then
                Current.Session.Remove(_destinationKey)
            Else
                Current.Session(_destinationKey) = Value
            End If
        End Set
    End Property

#End Region

    Public Sub Clear()
        Me.PersonID = Nothing
        Me.Contribution.Cancel()
        Me.Permissions = Nothing
        Me.ResumeDownload = Nothing
        Me.DestinationPage = Nothing
        Me.Authenticated = False
        Me.CreateGuest()
    End Sub

    '---COMMENT---------------------------------------------------------------
    '	Set a persistent cookie for given name-value
    '
    '	Date:		Name:	Description:
    '	12/10/04    JEA		Creation
    '-------------------------------------------------------------------------
    Private Shared Sub SetCookie(ByVal name As String, ByVal value As Object)
        If _useCookies Then
            Dim cookie As New HttpCookie(name, value.ToString)
            cookie.Expires = New Date(2010, 1, 1)
            Current.Response.Cookies.Add(cookie)
        End If
    End Sub

    '---COMMENT---------------------------------------------------------------
    '	write a cookie to be tested later
    '
    '	Date:		Name:	Description:
    '	1/24/05     JEA		Creation
    '-------------------------------------------------------------------------
    Public Shared Sub WriteTestCookie()
        Dim cookie As HttpCookie = New HttpCookie(_testCookie, HttpContext.Current.Session.SessionID)
        'cookie.Path = Global.BasePath
        Current.Response.Cookies.Add(cookie)
    End Sub

    Public Shared Function SupportsCookies() As Boolean
        Dim cookie As HttpCookie = Current.Request.Cookies(_testCookie)
        Dim supported As Boolean = ((Not cookie Is Nothing) AndAlso _
            cookie.Value = HttpContext.Current.Session.SessionID)
        If supported Then Current.Request.Cookies.Remove(_testCookie)
        Return supported
    End Function

    '---COMMENT---------------------------------------------------------------
    '	create a guest account
    '
    '	Date:		Name:	Description:
    '	12/10/04    JEA		Creation
    '   1/12/05     JEA     Handle empty web site
    '   2/18/05     JEA     Update cache control cookies
    '-------------------------------------------------------------------------
    Public Shared Sub CreateGuest()
        Dim user As New AMP.Person

        If Not WebSite Is Nothing Then
            user.Roles.Add(WebSite.Roles(AMP.Site.Role.Anonymous))
        End If
        user.Section = AMP.Site.Section.All

        Profile.User = user
        Profile.UseCookies = _useCookies

        ' store in cookie to be used for output cache variance
        Current.Response.Cookies.Add(New HttpCookie("section", user.Section.ToString))
        Current.Response.Cookies.Add(New HttpCookie("role", user.Roles.ToString))
    End Sub

End Class
