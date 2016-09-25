Option Strict On

Imports System.Web.UI
Imports System.Configuration.ConfigurationSettings

Namespace Controls
    Public Class Rating
        Inherits AMP.Controls.Image

        Private _rating As Single = 0
        Private _button As Boolean = False
        Private _radius As Integer
        Private _foreGroundColor As String
        Private _burstColor As String
        Private _borderColor As String
        Private _borderWidth As Integer
        Private _highLightColor As String
        Private _points As Integer = 5
        Private _sharpness As Single = 0.58
        Protected Draw As AMP.Draw

#Region " Properties "

        Public WriteOnly Property Button() As Boolean
            Set(ByVal Value As Boolean)
                _button = Value
            End Set
        End Property

        Public WriteOnly Property HighLightColor() As String
            Set(ByVal Value As String)
                _highLightColor = Value
            End Set
        End Property

        Public WriteOnly Property Rating() As Single
            Set(ByVal Value As Single)
                _rating = Value
            End Set
        End Property

        Public WriteOnly Property Sharpness() As Unit
            Set(ByVal sharp As Unit)
                Select Case sharp.Type
                    Case UnitType.Percentage
                        _sharpness = CSng(sharp.Value / 100)
                    Case Else
                        _sharpness = CSng(sharp.Value)
                End Select
            End Set
        End Property

        Public WriteOnly Property Points() As Integer
            Set(ByVal Value As Integer)
                _points = Value
            End Set
        End Property

        Public WriteOnly Property ForceNew() As Boolean
            Set(ByVal Value As Boolean)
                Me.Draw.ForceRegeneration = Value
            End Set
        End Property

        Public Property Radius() As Integer
            Set(ByVal Value As Integer)
                _radius = Value
            End Set
            Get
                Return _radius
            End Get
        End Property

        Public WriteOnly Property Color() As String
            Set(ByVal Value As String)
                _foreGroundColor = Value
            End Set
        End Property

        Public WriteOnly Property BurstColor() As String
            Set(ByVal Value As String)
                _burstColor = Value
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

#End Region

        Public Sub New()
            Me.Draw = New AMP.Draw
        End Sub

        '---COMMENT---------------------------------------------------------------
        '	generate image tag for rating stars
        '
        '	Date:		Name:	Description:
        '	10/24/04	JEA		Creation
        '   12/23/04    JEA     Add burst color for gradient fill
        '-------------------------------------------------------------------------
        Private Sub Rating_PreRender(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.PreRender
            If _rating > 0 Then
                With AppSettings
                    If _radius = Nothing Then _radius = CInt(.Item("StarRadius"))
                    If _borderWidth = Nothing Then _borderWidth = CInt(.Item("StarBorderWidth"))
                    If _foreGroundColor = Nothing Then _foreGroundColor = .Item("StarColor")
                    If _borderColor = Nothing Then _borderColor = .Item("StarBorderColor")
                    If _burstColor = Nothing Then _burstColor = .Item("StarBurstColor")
                End With

                With Me
                    .Draw.Width = _radius * 2
                    .Draw.Height = _radius * 2

                    .Draw.Color.ForeGround = ColorTranslator.FromHtml(_foreGroundColor)
                    .Draw.Color.Border = ColorTranslator.FromHtml(_borderColor)
                    .Draw.Color.Highlight = ColorTranslator.FromHtml(_burstColor)
                    .Draw.BorderWidth = _borderWidth

                    .Style.Add("width", String.Format("{0}px", _radius * 10))
                    .Style.Add("height", String.Format("{0}px", _radius * 2))
                    .Src = Me.Draw.RatingStars(_rating, _points, _sharpness)

                    If _button Then
                        If _highLightColor = Nothing Then _highLightColor = AppSettings.Item("ButtonHighlightColor")
                        .Draw.Color.ForeGround = ColorTranslator.FromHtml(_highLightColor)
                        .Draw.Color.Border = .Draw.Color.ForeGround
                        Dim overSrc As String = .Draw.RatingStars(_rating, _points, _sharpness)
                        .OnMouseOver = String.Format("this.src='{0}';", overSrc)
                        .OnMouseOut = String.Format("this.src='{0}';", .Src)
                    End If
                    .Transparency = True
                End With
            End If
        End Sub

        Protected Overrides Sub Render(ByVal writer As System.Web.UI.HtmlTextWriter)
            If _rating > 0 Then MyBase.Render(writer)
        End Sub
    End Class
End Namespace