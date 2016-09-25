Imports AMP.Site
Imports AMP.Common
Imports System.Web
Imports System.Data
Imports System.Text
Imports System.Collections
Imports System.Web.Security
Imports System.Web.HttpContext
Imports System.Text.RegularExpressions
Imports System.Configuration.ConfigurationSettings

Public Class Security

    Private _denyPermissionType As Integer
    Private _saveCookie As Boolean = True
    Private Const _hashMethod As String = "MD5"
    Private Const _salt As String = "93A05358-A5A1-4cbb-ADA2-D8B63E4708AB"

#Region "Properties"

    Public Property SaveCookie() As Boolean
        Get
            Return _saveCookie
        End Get
        Set(ByVal Value As Boolean)
            _saveCookie = Value
        End Set
    End Property

#End Region

    '---COMMENT---------------------------------------------------------------
    '	Find the user ID matching the given credentials
    '
    '	Date:		Name:	Description:
    '	12/2/04     JEA		Creation
    '   2/11/05     JEA     Check status of person
    '   2/28/05     JEA     Add temporary check for clear-text passwords
    '-------------------------------------------------------------------------
    Public Function Authenticate(ByVal email As String, ByVal password As String) As Boolean
        Dim person As AMP.Person = WebSite.Persons.WithEmail(email)
        If person Is Nothing Then
            AMP.Log.Activity(Activity.FailedLogin, email)
            Return False
        ElseIf person.Status = Status.Disabled Then
            ' person is not approved
            AMP.Log.Activity(Activity.TriedDisabledAccountLogin, person.ID)
            Profile.Message = String.Format("That account is not active")
            Return False
        ElseIf person.Password = Me.Encrypt(password) Then
            AMP.Log.Activity(Activity.Login, person.ID)
            Return CreateSession(person)
        ElseIf person.Password = password Then
            ' TODO: delete this option after a month or so
            AMP.Log.Activity(Activity.Login, person.ID)
            ' update clear-text password
            person.Password = Me.Encrypt(person.Password)
            Return CreateSession(person)
        Else
            AMP.Log.Activity(Activity.FailedLogin, email)
            Return False
        End If
    End Function

    '---COMMENT---------------------------------------------------------------
    '	Create session with user ID alone, such as from cookie
    '
    '	Date:		Name:	Description:
    '	12/3/04     JEA		Creation
    '   1/4/05      JEA     Added overload and changed to GUID
    '   2/11/05     JEA     Check status of person
    '-------------------------------------------------------------------------
    Public Function Authenticate(ByVal id As Guid) As Boolean
        If WebSite Is Nothing Then Return False
        Dim person As AMP.Person = WebSite.Persons(id)
        If person Is Nothing Then
            AMP.Log.Activity(Activity.FailedLogin, id)
            Return False
        ElseIf person.Status <> Status.Disabled Then
            AMP.Log.Activity(Activity.AutoLoginFromCookie, person.ID)
            Profile.Message = String.Format("Welcome Back {0}", person.DisplayName)
            Return CreateSession(person)
        Else
            ' person is not approved
            AMP.Log.Activity(Activity.TriedDisabledAccountLogin, person.ID)
            Profile.Message = String.Format("That account is not active")
            Return False
        End If
    End Function

    Public Function Authenticate(ByVal id As String) As Boolean
        Return Me.Authenticate(New Guid(id))
    End Function

    '---COMMENT---------------------------------------------------------------
    '	Create new person entity
    '
    '	Date:		Name:	Description:
    '	1/7/05      JEA		Creation
    '   2/26/05     JEA     Set privacy
    '-------------------------------------------------------------------------
    Public Function Register(ByVal firstName As String, ByVal lastName As String, _
        ByVal screenName As String, ByVal email As String, ByVal password As String, _
        ByVal url As String, ByVal privacy As Boolean) As Boolean

        If Not WebSite.Persons.WithEmail(email) Is Nothing Then
            Return False
        End If

        Dim person As New AMP.Person
        Dim mail As New AMP.Email

        With person
            .FirstName = firstName
            .LastName = lastName
            .NickName = screenName
            .Email = email
            .Password = Me.Encrypt(password)
            .PrivateEmail = privacy
            .WebSite = url
            .Status = Status.Approved
            .Roles.Add(WebSite.Roles(AMP.Site.Role.VerifiedGuest))
            .LastLogin = DateTime.Now
        End With

        WebSite.Persons.Add(person)
        Me.CreateSession(person)

        AMP.Log.Activity(Activity.Register, person.ID)

        Return True
    End Function

    '---COMMENT---------------------------------------------------------------
    '	reset password to random value
    '
    '	Date:		Name:	Description:
    '	1/10/05     JEA		Creation
    '-------------------------------------------------------------------------
    Public Function ResetPassword(ByVal person As AMP.Person) As Boolean
        Dim password As String = Me.RandomText(7)
        Dim mail As New AMP.Email

        person.Password = Me.Encrypt(password)
        mail.Password(person, password)
        Return True
    End Function

    '---COMMENT---------------------------------------------------------------
    '	Create session based on given person
    '
    '	Date:		Name:	Description:
    '	12/3/04     JEA		Creation
    '   2/18/04     JEA     Write cache control cookies
    '   2/27/05     JEA     For any session creation, switch status to approved
    '-------------------------------------------------------------------------
    Private Function CreateSession(ByVal person As AMP.Person) As Boolean
        ' setup profile with successful login
        Dim profile As New AMP.Profile

        BugOut("Creating session for {0}", person.FullName)

        With profile
            .User = person
            .UseCookies = Me.SaveCookie
            .Authenticated = True
            .LastLogin = person.LastLogin
            .Permissions = Nothing  ' clear to regenerate
        End With

        ' store in cookie to be used for output cache variance
        Current.Response.Cookies.Add(New HttpCookie("section", person.Section.ToString))
        Current.Response.Cookies.Add(New HttpCookie("role", person.Roles.ToString))

        person.Status = Status.Approved
        person.LastLogin = DateTime.Now
        WebSite.Save()

        Return True
    End Function

    '---COMMENT---------------------------------------------------------------
    '	Encrypt a string with salt
    '
    '	Date:		Name:	Description:
    '   10/28/04    JEA     Creation
    '-------------------------------------------------------------------------
    Shared Function Encrypt(ByVal text As String) As String
        Return FormsAuthentication.HashPasswordForStoringInConfigFile(text & _salt, _hashMethod)
    End Function

    '---COMMENT---------------------------------------------------------------
    '	Generate a simple, random set of letters and numbers
    '
    '	Date:		Name:	Description:
    '   10/28/04    JEA     Creation
    '-------------------------------------------------------------------------
    Shared Function RandomText(ByVal length As Integer) As String
        Dim random As New StringBuilder
        Dim randomNumber As Integer

        Randomize()
        For x As Integer = 1 To length
            ' alternate random numbers and letters
            randomNumber = CInt(IIf((x Mod 2 = 0), CInt((25 * Rnd()) + 1) + 65, CInt((10 * Rnd()) + 1) + 47))
            random.Append(LCase(Chr(randomNumber)))
        Next
        Return random.ToString
    End Function

    '---COMMENT---------------------------------------------------------------
    '	strip potentially harmful characters from string and ensure length limit
    '
    '	Date:		Name:	Description:
    '   1/13/05     JEA     Creation
    '-------------------------------------------------------------------------
    Shared Function SafeString(ByVal text As String, ByVal length As Integer) As String
        If text <> Nothing Then
            ' strip HTML and escape sequences
            ' e.g. precludes <div> or &#340; or %13
            Dim re As New Regex("<\/?\w+>|\&\#?\w{1,10}\;|\%\d{2,3}")
            text = re.Replace(text, "")
            If text.Length > length Then text = text.Substring(0, length)
        End If
        Return text
    End Function
End Class
