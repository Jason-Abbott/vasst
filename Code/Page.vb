Imports System.Text
Imports System.Web.Caching
Imports System.IO
Imports System.Configuration.ConfigurationSettings

Public Class Page
    Inherits System.Web.UI.Page

    Private _profile As New AMP.Profile
    Private _say As New AMP.Locale
    Private _security As AMP.Security
    Private _requireAuthentication As Boolean = False
    Private _title As String
    Private _templateFile As String = "~/template/ThreeColumn.ascx"
    Private _template As AMP.Template
    Private _render As Boolean = True

#Region " Controls "

    Protected pnlBody As Panel
    Protected pnlWindows As Panel

#End Region

#Region " Properties "

    Public Property Form() As AMP.Controls.Form
        Get
            Return _template.frmMain
        End Get
        Set(ByVal Value As AMP.Controls.Form)
            _template.frmMain = Value
        End Set
    End Property

    Public WriteOnly Property ScriptBlock() As String
        Set(ByVal Value As String)
            With _template.ScriptBlock
                .Append(Environment.NewLine)
                .Append(Value)
                .Append(Environment.NewLine)
            End With
        End Set
    End Property

    Public Property Security() As AMP.Security
        Get
            If _security Is Nothing Then _security = New AMP.Security
            Return _security
        End Get
        Set(ByVal Value As AMP.Security)
            _security = Value
        End Set
    End Property

    Public Property Profile() As AMP.Profile
        Get
            Return _profile
        End Get
        Set(ByVal Value As AMP.Profile)
            _profile = Value
        End Set
    End Property

    Public Property Say() As AMP.Locale
        Get
            Return _say
        End Get
        Set(ByVal Value As AMP.Locale)
            _say = Value
        End Set
    End Property

    Public Property StyleSheet() As ArrayList
        Get
            Return _template.StyleSheet
        End Get
        Set(ByVal Value As ArrayList)
            _template.StyleSheet = Value
        End Set
    End Property

    Public Property ScriptFile() As ArrayList
        Get
            Return _template.ScriptFile
        End Get
        Set(ByVal Value As ArrayList)
            _template.ScriptFile = Value
        End Set
    End Property

    Public Property Template() As AMP.Template
        Get
            Return _template
        End Get
        Set(ByVal Value As AMP.Template)
            _template = Value
        End Set
    End Property

    Public Property TemplateFile() As String
        Get
            Return _templateFile
        End Get
        Set(ByVal Value As String)
            _templateFile = Value
        End Set
    End Property

    Public WriteOnly Property Title() As String
        Set(ByVal Value As String)
            _title = Value
        End Set
    End Property

    '---COMMENT---------------------------------------------------------------
    '	return cleaned copy of referring page
    '
    '	Date:		Name:	Description:
    '	10/4/04		JEA		Creation
    '-------------------------------------------------------------------------
    Public ReadOnly Property ReferringPage() As String
        Get
            If Not Request.UrlReferrer Is Nothing Then
                Dim name As String = Request.UrlReferrer.ToString
                Return name.Substring(name.LastIndexOf("/") + 1)
            Else
                Return Nothing
            End If
        End Get
    End Property

    '---COMMENT---------------------------------------------------------------
    '	Causes derived page to redirect to login if user isn't authenticated
    '   Set this property in derived Page_Init() in "Web Form Designer Generated Code" region
    '
    '	Date:		Name:	Description:
    '	9/21/04		JEA		Creation
    '-------------------------------------------------------------------------
    Public WriteOnly Property RequireAuthentication() As Boolean
        Set(ByVal Value As Boolean)
            _requireAuthentication = Value
        End Set
    End Property

    '---COMMENT---------------------------------------------------------------
    '	return name of derived .aspx page
    '
    '	Date:		Name:	Description:
    '	9/22/04		JEA		Creation
    '-------------------------------------------------------------------------
    Public ReadOnly Property PageName() As String
        Get
            'Dim name As String = HttpContext.Current.Request.ServerVariables("SCRIPT_NAME")
            Dim name As String = Request.Path
            Return name.Substring(name.LastIndexOf("/") + 1)
        End Get
    End Property

#End Region

    Public Sub SendBack(ByVal defaultPage As String)
        Dim sendTo As String = Me.ReferringPage
        If sendTo = Nothing Or sendTo = Me.PageName Then sendTo = defaultPage
        Response.Redirect(sendTo, True)
    End Sub

    Public Sub SendBack()
        SendBack("default.aspx")
    End Sub

    Public Sub SendToLogin()
        Profile.DestinationPage = Request.Url.PathAndQuery
        Profile.WriteTestCookie()
        _render = False
        Response.Redirect(AppSettings("LoginPage"), True)
    End Sub

    '---COMMENT---------------------------------------------------------------
    '	add content to template
    '
    '	Date:		Name:	Description:
    '	2/22/05		JEA		Creation
    '-------------------------------------------------------------------------
    Protected Overrides Sub AddParsedSubObject(ByVal control As Object)
        If _template Is Nothing Then
            _template = DirectCast(Me.Page.LoadControl(_templateFile), AMP.Template)
        End If

        If control.GetType Is GetType(Panel) Then
            Dim panel As Panel = DirectCast(control, Panel)
            Select Case panel.ID
                Case "pnlWindows"
                    _template.CopyToWebPart(panel)
                Case "pnlBody"
                    _template.CopyToBody(panel)
            End Select
        End If
    End Sub

    '---COMMENT---------------------------------------------------------------
    '	replace page with template
    '
    '	Date:		Name:	Description:
    '	1/11/05		JEA		Creation
    '-------------------------------------------------------------------------
    Private Sub Page_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Init
        Me.Controls.Clear()
        Me.Controls.Add(_template)

        Me.ScriptFile.Add("global")
        Me.ScriptFile.Add("dom")
        Me.ScriptFile.Add("menu")
        Me.ScriptFile.Add("cookies")
        Me.ScriptFile.Add("webpart")
    End Sub

    '---COMMENT---------------------------------------------------------------
    '	allow time for control loads to update message
    '
    '	Date:		Name:	Description:
    '	1/11/05		JEA		Creation
    '-------------------------------------------------------------------------
    Private Sub Page_PreRender(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.PreRender
        If _requireAuthentication And Not Profile.Authenticated Then Me.SendToLogin()

        If _render Then
            Me.StyleSheet.Add("menu")
            Me.StyleSheet.Add("webpart")

            _template.Message = Profile.Message
            If _title <> Nothing Then _template.Title = _title
        End If
    End Sub

#Region " Custom Form Support "

    '---COMMENT---------------------------------------------------------------
    '	override state methods to support XHTML strict elements
    '
    '	Date:		Name:	Description:
    '   11/15/04    JEA     Creation
    '-------------------------------------------------------------------------
    Protected Overrides Function LoadPageStateFromPersistenceMedium() As Object
        Dim format As LosFormatter = New LosFormatter
        Dim viewState As String = Request.Form("__VIEWSTATE").ToString()
        Return format.Deserialize(viewState)
    End Function

    Protected Overrides Sub SavePageStateToPersistenceMedium(ByVal viewState As Object)
        Dim format As LosFormatter = New LosFormatter
        Dim writer As New StringWriter
        Dim stateField As New Literal
        format.Serialize(writer, viewState)
        stateField.Text = "<input type=""hidden"" name=""__VIEWSTATE"" value=""" & writer.ToString & """ />"
        _template.StateValue = stateField
    End Sub

    Public Overrides Sub VerifyRenderingInServerForm(ByVal control As System.Web.UI.Control)
        ' do nothing
    End Sub

#End Region

End Class
