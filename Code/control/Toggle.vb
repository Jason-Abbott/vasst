Option Strict On

Imports System.Web.UI
Imports System.Configuration.ConfigurationSettings

Namespace Controls
    Public Class Toggle
        Inherits AMP.Controls.Image

        Private _color As String
        Private _backGroundColor As String
        Private _initialState As Draw.ToggleState = Draw.ToggleState.Open
        Private _nodeID As String
        Private _draw As AMP.Draw
        Private _click As String

#Region " Properties "

        Public WriteOnly Property NodeID() As String
            Set(ByVal Value As String)
                _nodeID = Value
            End Set
        End Property

        Public WriteOnly Property InitialState() As Draw.ToggleState
            Set(ByVal Value As Draw.ToggleState)
                _initialState = Value
            End Set
        End Property

        Public WriteOnly Property ForceNew() As Boolean
            Set(ByVal Value As Boolean)
                _draw.ForceRegeneration = Value
            End Set
        End Property

        Public WriteOnly Property Color() As String
            Set(ByVal Value As String)
                _color = Value
            End Set
        End Property

        Public WriteOnly Property BackGround() As String
            Set(ByVal Value As String)
                _backGroundColor = Value
            End Set
        End Property

#End Region

        Public Sub New()
            _draw = New AMP.Draw
        End Sub

        Private Sub Toggle_PreRender(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.PreRender
            Me.OnClick = String.Format("WebPart.Toggle(this, '{0}');", _nodeID)
            Me.Prepare()
        End Sub

        '---COMMENT---------------------------------------------------------------
        '	method to force control to render
        '
        '	Date:		Name:	Description:
        '	11/4/04	    JEA		Creation
        '-------------------------------------------------------------------------
        Public Sub Write(ByVal writer As HtmlTextWriter, ByVal nodeID As String)
            If nodeID <> Nothing Then _nodeID = nodeID
            Me.Prepare()
            Me.RenderControl(writer)
        End Sub

        Public Sub write(ByVal writer As HtmlTextWriter)
            Me.Write(writer, Nothing)
        End Sub

        '---COMMENT---------------------------------------------------------------
        '	setup image object
        '
        '	Date:		Name:	Description:
        '	11/4/04	    JEA		Creation
        '-------------------------------------------------------------------------
        Private Sub Prepare()
            If Me.Height = Nothing Then Me.Height = CInt(AppSettings("ToggleHeight"))
            If _color = Nothing Then _color = AppSettings("ToggleColor")
            Me.Width = Me.Height    ' toggler is square

            With _draw
                .Width = Me.Width
                .Height = Me.Height
                If _backGroundColor <> Nothing Then
                    .Color.BackGround = ColorTranslator.FromHtml(_backGroundColor)
                End If
                .Color.ForeGround = ColorTranslator.FromHtml(_color)
            End With

            With Me
                .Src = _draw.Toggler(_initialState)
                .Transparency = True
            End With
        End Sub
    End Class
End Namespace
