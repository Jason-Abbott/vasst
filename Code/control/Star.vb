Imports System.Web.UI
Imports System.Configuration.ConfigurationSettings

Namespace Controls
    Public Class Star
        Inherits HtmlControl

        Private _radius As Integer
        Private _foreGroundColor As String
        Private _burstColor As String
        Private _borderColor As String
        Private _borderWidth As Integer
        Private _sharpness As Single
        Private _points As Integer
        Protected Draw As AMP.Draw

#Region " Properties "

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
        '	get default values when none specified
        '
        '	Date:		Name:	Description:
        '	10/23/04	JEA		Creation
        '   12/23/04    JEA     Add burst color for gradient fill
        '-------------------------------------------------------------------------
        Private Sub Star_PreRender(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.PreRender
            With AppSettings
                If _radius = Nothing Then _radius = CInt(.Item("StarRadius"))
                If _borderWidth = Nothing Then _borderWidth = CInt(.Item("StarBorderWidth"))
                If _foreGroundColor = Nothing Then _foreGroundColor = .Item("StarColor")
                If _burstColor = Nothing Then _burstColor = .Item("StarBurstColor")
                If _borderColor = Nothing Then _borderColor = .Item("StarBorderColor")
            End With

            With Me.Draw
                .Width = _radius * 2
                .Height = _radius * 2
                .Color.Highlight = ColorTranslator.FromHtml(_burstColor)
                .Color.ForeGround = ColorTranslator.FromHtml(_foreGroundColor)
                .Color.Border = ColorTranslator.FromHtml(_borderColor)
                .BorderWidth = _borderWidth
            End With
        End Sub

        '---COMMENT---------------------------------------------------------------
        '	generate image tag for rating starss
        '
        '	Date:		Name:	Description:
        '	10/23/04	JEA		Creation
        '-------------------------------------------------------------------------
        Protected Overrides Sub Render(ByVal writer As System.Web.UI.HtmlTextWriter)
            With writer
                .Write("<img src=""")
                .Write(Me.Draw.Star(_points, _sharpness))
                .Write(""" width=""")
                .Write(_radius * 2)
                .Write(""" height=""")
                .Write(_radius * 2)
                .Write(""" alt="""" />")
            End With
        End Sub
    End Class
End Namespace