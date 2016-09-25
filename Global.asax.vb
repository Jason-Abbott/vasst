Imports System.IO
Imports System.Web
Imports System.Text
Imports System.Web.SessionState
Imports System.Configuration.ConfigurationSettings
Imports AMP.Site

Public Class Global
    Inherits System.Web.HttpApplication

    Private Shared _basePath As String
    Private Shared _rootEntity As AMP.Site
    Private Shared _startTime As DateTime
    Private Shared _saveTime As DateTime
    Private Shared _errorEmail As New Hashtable
    Private Shared _activeDataFile As FileInfo
    Private Shared _dailyLog As Boolean = False

#Region " Global Properties "

    '---COMMENT---------------------------------------------------------------
    '	indicates whether daily visitor log has been started
    '
    '	Date:		Name:	Description:
    '	2/18/05 	JEA		Creation
    '-------------------------------------------------------------------------
    Public Shared Property DailyLog() As Boolean
        Get
            Return _dailyLog
        End Get
        Set(ByVal Value As Boolean)
            _dailyLog = Value
        End Set
    End Property

    '---COMMENT---------------------------------------------------------------
    '	tracks the current data file
    '
    '	Date:		Name:	Description:
    '	2/13/05 	JEA		Creation
    '-------------------------------------------------------------------------
    Public Shared Property ActiveDataFile() As FileInfo
        Get
            Return _activeDataFile
        End Get
        Set(ByVal Value As FileInfo)
            _activeDataFile = Value
        End Set
    End Property

    '---COMMENT---------------------------------------------------------------
    '	prevent flood of e-mail by tracking which errors have been e-mailed
    '   and only resending periodically
    '
    '	Date:		Name:	Description:
    '	2/11/05 	JEA		Creation
    '-------------------------------------------------------------------------
    Public Shared Property ErrorEmail() As Hashtable
        Get
            Return _errorEmail
        End Get
        Set(ByVal Value As Hashtable)
            _errorEmail = Value
        End Set
    End Property

    Public Shared Property WebSite() As AMP.Site
        Get
            Return _rootEntity
        End Get
        Set(ByVal Value As AMP.Site)
            _rootEntity = Value
        End Set
    End Property

    Public Shared ReadOnly Property BasePath() As String
        Get
            Return _basePath
        End Get
    End Property

    Public Shared ReadOnly Property ApplicationStart() As DateTime
        Get
            Return _startTime
        End Get
    End Property

    Public Shared Property LastSave() As DateTime
        Get
            Return _saveTime
        End Get
        Set(ByVal Value As DateTime)
            _saveTime = Value
        End Set
    End Property

#End Region

    '---COMMENT---------------------------------------------------------------
    '	retrieve or create site base entity
    '
    '	Date:		Name:	Description:
    '	12/3/04     JEA		Creation
    '-------------------------------------------------------------------------
    Sub Application_Start(ByVal sender As Object, ByVal e As EventArgs)
        _basePath = HttpRuntime.AppDomainAppVirtualPath
        _startTime = DateTime.Now

        If _basePath = "/" Then _basePath = ""

        ' load data
        Dim schemaChange As Boolean = False
        Dim file As New AMP.Data.File
        Dim type As System.Type = System.Type.GetType("AMP.Site")
        Dim path As FileInfo = file.Newest(AppSettings("DataFolder"), type)

        BugOut("Starting application")
        Try
            _rootEntity = DirectCast(file.Load(path, type, schemaChange), AMP.Site)
            _activeDataFile = path
        Catch
            _rootEntity = Nothing
        End Try

        If _rootEntity Is Nothing OrElse _rootEntity.Persons Is Nothing Then
            Log.Error("Unable to load root entity", Log.ErrorType.Custom)
            BugOut("Unable to retrieve valid root entity")
            _rootEntity = New AMP.Site
        ElseIf schemaChange Then
            BugOut("Saving root entity to persist new schema")
            _rootEntity.Save()
        End If

        Log.Activity(Activity.StartedApplication)
    End Sub

    '---COMMENT---------------------------------------------------------------
    '	Auto-login from cookie, if cookie found
    '
    '	Date:		Name:	Description:
    '	12/3/04     JEA		Creation
    '   1/24/05     JEA     Write test cookie
    '   2/18/05     JEA     Update visit count
    '-------------------------------------------------------------------------
    Sub Session_Start(ByVal sender As Object, ByVal e As EventArgs)
        Log.AddVisit()
        Dim personID As Guid = Profile.PersonID
        If Not personID.Equals(Guid.Empty) Then
            ' populated person ID would have to come from cookie, attempt auto-login
            BugOut("Starting session for Person ID {0}", personID)
            Dim security As New AMP.Security
            If security.Authenticate(personID) Then Return
            BugOut("Failed to auto-login")
        End If
        ' create empty person for profile
        BugOut("Starting session with new user")
        Profile.CreateGuest()
    End Sub

    '---COMMENT---------------------------------------------------------------
    '	cleanup persisted session data
    '
    '	Date:		Name:	Description:
    '	1/15/05     JEA		Creation
    '-------------------------------------------------------------------------
    Sub Session_End(ByVal sender As Object, ByVal e As EventArgs)
        Profile.Contribution.Cancel()
    End Sub

    '---COMMENT---------------------------------------------------------------
    '	not sure when this fires but if it does then let's save the root entity
    '
    '	Date:		Name:	Description:
    '	12/3/04     JEA		Creation
    '-------------------------------------------------------------------------
    Sub Application_End(ByVal sender As Object, ByVal e As EventArgs)
        _rootEntity.Save()
    End Sub

#Region " Unhandled events "

    Sub Application_BeginRequest(ByVal sender As Object, ByVal e As EventArgs)
        ' Fires at the beginning of each request
    End Sub

    Sub Application_AuthenticateRequest(ByVal sender As Object, ByVal e As EventArgs)
        ' Fires upon attempting to authenticate the use
    End Sub

    Sub Application_Error(ByVal sender As Object, ByVal e As EventArgs)
        ' Fires when an error occurs
    End Sub

#End Region

    '---COMMENT---------------------------------------------------------------
    '	strings used to differentiate output caching
    '
    '	Date:		Name:	Description:
    '	12/3/04     JEA		Creation
    '   1/18/05     JEA     Add "message"
    '   2/2/05      JEA     Get values from cookies since session may not exist
    '-------------------------------------------------------------------------
    Public Overrides Function GetVaryByCustomString(ByVal context As System.Web.HttpContext, _
        ByVal custom As String) As String

        Dim keys As String() = custom.Split(","c)
        Dim pattern As New StringBuilder

        For Each key As String In keys
            key = key.Trim.ToLower
            If key = "browser" Then
                pattern.Append("_")
                pattern.Append(context.Request.Browser.Type.ToLower)
            Else
                Dim cookie As HttpCookie = context.Request.Cookies(key)
                If Not cookie Is Nothing Then
                    pattern.Append("_")
                    pattern.Append(cookie.Value)
                End If
            End If
        Next

        Return pattern.ToString
    End Function
End Class
