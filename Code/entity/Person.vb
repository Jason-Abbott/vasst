Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Runtime.Serialization
Imports System.Security.Permissions
Imports System.Runtime.Serialization.Formatters.Binary
Imports AMP.Data
Imports AMP.Site
Imports AMP.Global

<Serializable()> _
Public Class Person
    Implements IComparable, ICloneable

    Private _id As Guid
    Private _firstName As String
    Private _lastName As String
    Private _nickName As String
    Private _imageFile As String
    Private _email As String
    Private _password As String
    Private _jobTitle As String
    Private _webSite As String
    Private _registeredOn As DateTime
    Private _description As String
    Private _permission As AMP.Site.Permission()
    Private _lastLogin As DateTime
    Private _status As AMP.Site.Status
    Private _address As New AMP.AddressCollection
    Private _phone As New AMP.PhoneCollection
    Private _privateEmail As Boolean
    Private _employer As New AMP.Company
    Private _cart As New AMP.Cart
    Private _cards As New AMP.CardCollection
    Private _roles As New AMP.RoleCollection
    Private _section As Integer = AMP.Site.Section.All

#Region " Properties "

    Public Property Phone() As AMP.PhoneCollection
        Get
            Return _phone
        End Get
        Set(ByVal Value As AMP.PhoneCollection)
            _phone = Value
        End Set
    End Property

    Public Property ImageFile() As String
        Get
            Return _imageFile
        End Get
        Set(ByVal Value As String)
            _imageFile = Value
        End Set
    End Property

    Public Property Cart() As AMP.Cart
        Get
            Return _cart
        End Get
        Set(ByVal Value As AMP.Cart)
            _cart = Value
        End Set
    End Property

    Public Property Permissions() As AMP.Site.Permission()
        Get
            Return _permission
        End Get
        Set(ByVal Value As AMP.Site.Permission())
            _permission = Value
        End Set
    End Property

    Public Property Section() As Integer
        Get
            Return _section
        End Get
        Set(ByVal Value As Integer)
            _section = Value
        End Set
    End Property

    Public Property Roles() As AMP.RoleCollection
        Get
            Return _roles
        End Get
        Set(ByVal Value As AMP.RoleCollection)
            _roles = Value
        End Set
    End Property

    Public Property Cards() As AMP.CardCollection
        Get
            Return _cards
        End Get
        Set(ByVal Value As AMP.CardCollection)
            _cards = Value
        End Set
    End Property

    Public Property JobTitle() As String
        Get
            Return _jobTitle
        End Get
        Set(ByVal Value As String)
            _jobTitle = Security.SafeString(Value, 75)
        End Set
    End Property

    Public Property Employer() As AMP.Company
        Get
            Return _employer
        End Get
        Set(ByVal Value As AMP.Company)
            _employer = Value
        End Set
    End Property

    Public Property PrivateEmail() As Boolean
        Get
            Return _privateEmail
        End Get
        Set(ByVal Value As Boolean)
            _privateEmail = Value
        End Set
    End Property

    Public Property LastLogin() As DateTime
        Get
            Return _lastLogin
        End Get
        Set(ByVal Value As DateTime)
            _lastLogin = Value
        End Set
    End Property

    'Public Property ConfirmationCode() As String
    '    Get
    '        Return _confirmationCode
    '    End Get
    '    Set(ByVal Value As String)
    '        _confirmationCode = Security.SafeString(Value, 20)
    '    End Set
    'End Property

    Public Property Status() As AMP.Site.Status
        Get
            Return _status
        End Get
        Set(ByVal Value As AMP.Site.Status)
            _status = Value
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

    Public Property RegisteredOn() As DateTime
        Get
            Return _registeredOn
        End Get
        Set(ByVal Value As DateTime)
            _registeredOn = Value
        End Set
    End Property

    Public Property WebSite() As String
        Get
            Return _webSite
        End Get
        Set(ByVal Value As String)
            _webSite = Security.SafeString(Value, 150).Replace("http://", Nothing)
        End Set
    End Property

    Public Property Password() As String
        Get
            Return _password
        End Get
        Set(ByVal Value As String)
            _password = Security.SafeString(Value, 32)
        End Set
    End Property

    Public Property Email() As String
        Get
            Return _email
        End Get
        Set(ByVal Value As String)
            _email = Security.SafeString(Value, 50)
        End Set
    End Property

    Public Property Address() As AddressCollection
        Get
            Return _address
        End Get
        Set(ByVal Value As AddressCollection)
            _address = Value
        End Set
    End Property

    Public ReadOnly Property ID() As Guid
        Get
            Return _id
        End Get
    End Property

    Public Property FirstName() As String
        Get
            Return _firstName
        End Get
        Set(ByVal Value As String)
            _firstName = Security.SafeString(Value, 100)
        End Set
    End Property

    Public Property LastName() As String
        Get
            Return _lastName
        End Get
        Set(ByVal Value As String)
            _lastName = Security.SafeString(Value, 100)
        End Set
    End Property

    Public Property NickName() As String
        Get
            Return _nickName
        End Get
        Set(ByVal Value As String)
            _nickName = Security.SafeString(Value, 100)
        End Set
    End Property

    Public ReadOnly Property FullName() As String
        Get
            Return String.Format("{0} {1}", _firstName, _lastName).Trim
        End Get
    End Property

    Public ReadOnly Property DisplayName() As String
        Get
            Return IIf(_nickName = Nothing, Me.FullName, _nickName).ToString
        End Get
    End Property

    Public ReadOnly Property DetailLink() As String
        Get
            Return String.Format("<a href=""{0}/person.aspx?id={1}"">{2}</a>", _
                Global.BasePath, Me.ID, HttpUtility.HtmlEncode(Me.DisplayName))
        End Get
    End Property

    Public ReadOnly Property EmailLink() As String
        Get
            Return String.Format("<a href=""mailto:{0}"">{0}</a>", Me.Email)
        End Get
    End Property

    Public ReadOnly Property WebSiteLink() As String
        Get
            If Me.WebSite <> Nothing Then
                Return String.Format("<a href=""http://{0}"">{0}</a>", Me.WebSite)
            End If
        End Get
    End Property

#End Region

    Public Sub New()
        _id = Guid.NewGuid
    End Sub

    '---COMMENT---------------------------------------------------------------
    '   do other entities depend on this user
    '
    '	Date:		Name:	Description:
    '	3/8/05	    JEA		Creation
    '-------------------------------------------------------------------------
    Public Function Dependencies() As Boolean
        If Global.WebSite.Assets.FromUser(Me).Count > 0 Then Return True
        If Global.WebSite.Contests.EntriesFromUser(Me).Count > 0 Then Return True
        ' TODO: check for outstanding orders when commerce up, and possibly votes
        Return False
    End Function

    '---COMMENT---------------------------------------------------------------
    '   can this person be edited by current user
    '
    '	Date:		Name:	Description:
    '	2/9/05	    JEA		Creation
    '-------------------------------------------------------------------------
    Public Function CanEdit() As Boolean
        Return ((Me.ID.Equals(Profile.User.ID) AndAlso _
            Profile.User.HasPermission(Permission.EditMyself)) OrElse _
            Profile.User.HasPermission(Permission.EditAnyUser))
    End Function

    '---COMMENT---------------------------------------------------------------
    '	get all permissions for this person
    '
    '	Date:		Name:	Description:
    '	12/7/04	    JEA		Creation
    '   1/29/05     JEA     Cache permissions list per session
    '   3/1/05      JEA     Only cache for active user
    '-------------------------------------------------------------------------
    Public Function EffectivePermissions() As AMP.Site.Permission()
        Dim cache As Boolean = Profile.PersonID.Equals(_id)

        If cache AndAlso Not Profile.Permissions Is Nothing Then
            Return Profile.Permissions
        End If

        Dim pc As New PermissionCollection
        Dim permissions As AMP.Site.Permission()
        Dim mine As AMP.Site.Permission() = Me.Permissions
        Dim fromRole As AMP.Site.Permission() = Me.Roles.Permissions

        ' get personal permissions
        If Not mine Is Nothing Then
            For x As Integer = 0 To mine.Length - 1
                pc.Add(mine(x))
            Next
        End If

        ' get role permissions
        If Not fromRole Is Nothing Then
            For x As Integer = 0 To fromRole.Length - 1
                pc.Add(fromRole(x))
            Next
        End If

        permissions = pc.All
        If cache Then Profile.Permissions = permissions
        Return permissions
    End Function

    '---COMMENT---------------------------------------------------------------
    '	does person have given permission either directly or from role
    '
    '	Date:		Name:	Description:
    '	12/7/04	    JEA		Creation
    '   1/29/05     JEA     Cache permission answer per request
    '   3/1/05      JEA     Only cache for active user
    '-------------------------------------------------------------------------
    Public Function HasPermission(ByVal permission As AMP.Site.Permission) As Boolean
        Dim cache As Boolean = Profile.PersonID.Equals(_id)
        Dim key As String = permission.ToString

        If cache AndAlso HttpContext.Current.Items.Contains(key) Then
            Return CBool(HttpContext.Current.Items(key))
        End If

        Dim permissions As AMP.Site.Permission() = Me.EffectivePermissions
        Dim permitted As Boolean

        If Not permissions Is Nothing Then
            permitted = (Array.IndexOf(permissions, permission) >= 0)
        Else
            permitted = False
        End If

        If cache Then HttpContext.Current.Items.Add(key, permitted)
        Return permitted
    End Function


    '---COMMENT---------------------------------------------------------------
    '	does person have text in name or other string
    '
    '	Date:		Name:	Description:
    '	12/7/04	    JEA		Creation
    '   2/16/05     JEA     Use regex
    '-------------------------------------------------------------------------
    Public Function HasText(ByVal text As String) As Boolean
        Dim re As New Regex(text, RegexOptions.IgnoreCase)
        If re.IsMatch(_firstName) OrElse re.IsMatch(_lastName) _
            OrElse (_nickName <> Nothing AndAlso re.IsMatch(_nickName)) Then

            Return True
        Else
            Return False
        End If
    End Function

    Public Function CompareTo(ByVal entity As Object) As Integer Implements System.IComparable.CompareTo
        Dim p As AMP.Person = DirectCast(entity, AMP.Person)
        Dim compare As Integer
        compare = String.Compare(Me.LastName, p.LastName)
        If compare = 0 Then compare = String.Compare(Me.FirstName, p.FirstName)
        Return compare
    End Function

#Region " Serialization "

    Public Function Clone() As Object Implements System.ICloneable.Clone
        Return Serialization.Clone(Me)
    End Function

#End Region

End Class
