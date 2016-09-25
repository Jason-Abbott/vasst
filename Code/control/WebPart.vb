Imports System.Web.UI
Imports System.Configuration.ConfigurationSettings

Namespace Controls
    Public Class WebPart
        Inherits System.Web.UI.UserControl

        Private _cssClass As String = "webPart"
        Private _minimized As Boolean
        Private _minimizable As Boolean = False
        Private _maximizable As Boolean = False
        Private _maximizeURL As String
        Private _stateSet As Boolean = False
        Private _height As String   ' optionally define css height and width
        Private _width As String
        Private _title As String
        Private _hostPage As AMP.Page

#Region " Properties "

        Public Shadows Property Page() As AMP.Page
            Get
                Return _hostPage
            End Get
            Set(ByVal Value As AMP.Page)
                _hostPage = Value
            End Set
        End Property

        Public WriteOnly Property Minimized() As Boolean
            Set(ByVal Value As Boolean)
                _minimized = Value
                _stateSet = True
                If _minimized Then _minimizable = True
            End Set
        End Property

        Public WriteOnly Property MaximizeURL() As String
            Set(ByVal Value As String)
                _maximizeURL = Value
            End Set
        End Property


        Public WriteOnly Property Height() As String
            Set(ByVal Value As String)
                _height = Value
            End Set
        End Property

        Public WriteOnly Property Width() As String
            Set(ByVal Value As String)
                _width = Value
            End Set
        End Property

        Public WriteOnly Property CssClass() As String
            Set(ByVal Value As String)
                _cssClass = Value
            End Set
        End Property

        Public WriteOnly Property Title() As String
            Set(ByVal Value As String)
                _title = Value
            End Set
        End Property

        Public WriteOnly Property Minimizable() As Boolean
            Set(ByVal Value As Boolean)
                _minimizable = Value
                If Not _minimizable Then
                    _minimized = False
                    _stateSet = True
                End If
            End Set
        End Property

        Public WriteOnly Property Maximizable() As Boolean
            Set(ByVal Value As Boolean)
                _maximizable = Value
            End Set
        End Property

#End Region

        Private Sub Page_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Init
            _hostPage = DirectCast(MyBase.Page, AMP.Page)
        End Sub

        '---COMMENT---------------------------------------------------------------
        '	create control tree
        '
        '	Date:		Name:	Description:
        '	11/4/04	    JEA		Creation
        '-------------------------------------------------------------------------
        Private Sub Page_PreRender(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.PreRender
            Dim container As New WebControls.Panel
            Dim titleBar As New WebControls.Panel
            Dim title As New WebControls.Label
            Dim body As New WebControls.Panel
            Dim bodyID As String

            If Not _stateSet Then _minimized = Boolean.Parse(AppSettings("WebPartsMinimized"))

            If Me.HasControls Then
                bodyID = Me.ClientID & "_body"
            Else
                bodyID = Me.ClientID
                _minimizable = False
            End If

            container.CssClass = _cssClass
            If _width <> Nothing Then container.Style.Item("width") = _width

            ' build title bar
            titleBar.CssClass = "titleBar"
            If _minimizable OrElse _maximizable Then
                Dim toggle As New AMP.Controls.Toggle
                toggle.NodeID = bodyID
                If _minimized Then toggle.InitialState = Draw.ToggleState.Closed
                titleBar.Controls.Add(toggle)
            End If
            title.CssClass = "title"
            title.Text = _title
            titleBar.Controls.Add(title)

            ' populate body with literal content
            body.CssClass = "body"
            body.ID = "body"
            If _minimized Then body.Style.Add("display", "none")
            If _height <> Nothing Then body.Style.Item("height") = _height

            If Me.HasControls Then
                Dim count As Integer = Me.Controls.Count
                For x As Integer = 0 To count - 1
                    body.Controls.Add(Me.Controls(0))
                Next
            End If

            ' build container control
            container.Controls.Add(titleBar)
            container.Controls.Add(body)

            Me.Controls.Clear()
            Me.Controls.Add(container)
        End Sub
    End Class
End Namespace
