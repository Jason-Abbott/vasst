Option Strict On

Imports System.Web.UI
Imports System.Configuration.ConfigurationSettings

Namespace Controls
    Public Class Button
        Inherits AMP.Controls.Image

        Private _text As String
        Private _resx As String
        Private _url As String
        Private _foreGroundColor As String
        Private _backGroundColor As String
        Private _highLightColor As String
        Private _textColor As String
        Private _draw As AMP.Draw

#Region " Properties "

        Public WriteOnly Property Url() As String
            Set(ByVal Value As String)
                _url = Value
            End Set
        End Property

        Public WriteOnly Property HighLightColor() As String
            Set(ByVal Value As String)
                _highLightColor = Value
            End Set
        End Property

        Public WriteOnly Property Text() As String
            Set(ByVal Value As String)
                _text = Value
            End Set
        End Property

        Public WriteOnly Property ForceNew() As Boolean
            Set(ByVal Value As Boolean)
                _draw.ForceRegeneration = Value
            End Set
        End Property

        Public WriteOnly Property Color() As String
            Set(ByVal Value As String)
                _foreGroundColor = Value
            End Set
        End Property

        Public WriteOnly Property BackGround() As String
            Set(ByVal Value As String)
                _backGroundColor = Value
            End Set
        End Property

        Public WriteOnly Property TextColor() As String
            Set(ByVal Value As String)
                _textColor = Value
            End Set
        End Property

        Public WriteOnly Property Resx() As String
            Set(ByVal Value As String)
                _resx = Value
            End Set
        End Property

#End Region

        Public Sub New()
            _draw = New AMP.Draw
        End Sub

        '---COMMENT---------------------------------------------------------------
        '	get values needed to render button
        '
        '	Date:		Name:	Description:
        '	10/22/04  	JEA		Creation
        '   1/13/05     JEA     Get text from resource
        '-------------------------------------------------------------------------
        Private Sub Button_PreRender(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.PreRender
            With AppSettings
                If _highLightColor = Nothing Then _highLightColor = .Item("ButtonHighlightColor")
                If _foreGroundColor = Nothing Then _foreGroundColor = .Item("ButtonColor")
                If _textColor = Nothing Then _textColor = .Item("ButtonTextColor")
            End With

            With _draw
                .Quantize = False
                .Width = Nothing
                .Color.ForeGround = ColorTranslator.FromHtml(_foreGroundColor)
                .Color.Text = ColorTranslator.FromHtml(_textColor)
                .Color.Highlight = ColorTranslator.FromHtml(_highLightColor)
            End With

            If _text = Nothing AndAlso _resx <> Nothing Then
                _text = DirectCast(Me.Page, AMP.Page).Say(String.Format("Action.{0}", _resx))
            End If

            With Me
                .Transparency = True
                If _url <> Nothing Then .OnClick = String.Format("location.href='{0}';", _url)
                .CssClass = "button"
                .RollOver = "DOM.Button(this);"
                If .AlternateText = Nothing Then .AlternateText = _text
                .Src = _draw.Button(_text, True)
                .Height = _draw.Height
            End With
        End Sub
    End Class
End Namespace