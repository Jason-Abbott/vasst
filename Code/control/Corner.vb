Option Strict On

Imports System.Web.UI
Imports System.Configuration.ConfigurationSettings

Namespace Controls
    Public Class Corner
        Inherits HtmlControl

        Private _radius As Integer
        Private _foreGroundColor As String
        Private _backGroundColor As String
        Private _borderColor As String
        Private _borderWidth As Integer
        Private _forTitle As Boolean
        Private _draw As AMP.Draw

#Region " Properties "

        Public WriteOnly Property ForceNew() As Boolean
            Set(ByVal Value As Boolean)
                _draw.ForceRegeneration = Value
            End Set
        End Property

        Public WriteOnly Property Radius() As Integer
            Set(ByVal Value As Integer)
                _radius = Value
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

        Public WriteOnly Property BorderColor() As String
            Set(ByVal Value As String)
                _borderColor = Value
            End Set
        End Property

        Public WriteOnly Property BorderWidth() As Integer
            Set(ByVal Value As Integer)
                _borderWidth = Value
            End Set
        End Property

        Public WriteOnly Property Orientation() As String
            Set(ByVal Value As String)
                Select Case Value
                    Case "tl"
                        _draw.Orientation = Draw.Position.TopLeft
                    Case "tr"
                        _draw.Orientation = Draw.Position.TopRight
                    Case "br"
                        _draw.Orientation = Draw.Position.BottomRight
                    Case "bl"
                        _draw.Orientation = Draw.Position.BottomLeft
                    Case Else
                        _draw.Orientation = Draw.Position.TopLeft
                End Select
            End Set
        End Property

#End Region

        Public Sub New()
            _draw = New AMP.Draw
        End Sub

        '---COMMENT---------------------------------------------------------------
        '	get default values when none specified
        '
        '	Date:		Name:	Description:
        '	10/22/04	JEA		Creation
        '-------------------------------------------------------------------------
        Private Sub Corner_PreRender(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.PreRender
            With AppSettings
                If _radius = Nothing Then _radius = CInt(.Item("CornerRadius"))
                If _borderWidth = Nothing Then _borderWidth = CInt(.Item("CornerBorderWidth"))
                If _foreGroundColor = Nothing Then _foreGroundColor = .Item("CornerColor")
                'If _backGroundColor = Nothing Then _backGroundColor = .Item("BackGroundColor")
                If _borderColor = Nothing Then _borderColor = .Item("CornerBorderColor")
            End With

            With _draw
                .Width = _radius
                .Height = _radius
                .Color.BackGround = ColorTranslator.FromHtml(_backGroundColor)
                .Color.ForeGround = ColorTranslator.FromHtml(_foreGroundColor)
                .Color.Border = ColorTranslator.FromHtml(_borderColor)
                .BorderWidth = _borderWidth
            End With
        End Sub

        '---COMMENT---------------------------------------------------------------
        '	generate image tag for corner graphic
        '
        '	Date:		Name:	Description:
        '	10/22/04	JEA		Creation
        '-------------------------------------------------------------------------
        Protected Overrides Sub Render(ByVal writer As System.Web.UI.HtmlTextWriter)
            With writer
                .Write("<img src=""")
                .Write(_draw.Corner)
                .Write(""" width=""")
                .Write(_radius)
                .Write(""" height=""")
                .Write(_radius)
                .Write(""" alt="""" />")
            End With
        End Sub
    End Class
End Namespace
