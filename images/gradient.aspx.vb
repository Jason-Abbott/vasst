Imports System.IO
Imports System.Web.HttpContext
Imports System.Web.Caching
Imports System.Text
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Drawing.Imaging
Imports System.Configuration.ConfigurationSettings

Public Class Gradient
    Inherits System.Web.UI.Page

    Private _startColor As Color
    Private _endColor As Color
    Private _startAlpha As Integer = 100
    Private _endAlpha As Integer = 100
    Private _width As Integer = 1
    Private _height As Integer = 1
    Private _angle As Single
    Private _orthogonal As Boolean = True

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub

    'NOTE: The following placeholder declaration is required by the Web Form Designer.
    'Do not delete or move it.
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: This method call is required by the Web Form Designer
        'Do not modify it using the code editor.
        InitializeComponent()
    End Sub

#End Region

    '---COMMENT---------------------------------------------------------------
    '	build gradient for output
    '
    '	Date:		Name:	Description:
    '	11/7/04	    JEA		Creation
    '   1/4/05      JEA     Added caching
    '-------------------------------------------------------------------------
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim bitmap As Bitmap
        Dim cacheKey As String = Request.Url.Query
        Dim transparency As Boolean = False

        With Request
            If .QueryString("a1") <> Nothing Then _startAlpha = CInt(.QueryString("a1"))
            If .QueryString("c2") = Nothing Then
                _endAlpha = 0
            ElseIf .QueryString("a2") <> Nothing Then
                _endAlpha = CInt(.QueryString("a2"))
            End If
            transparency = (_startAlpha < 100 OrElse _endAlpha < 100)

            If Page.Cache(cacheKey) Is Nothing Then

                If .QueryString("c1") <> Nothing Then
                    _startColor = ColorTranslator.FromHtml("#" & .QueryString("c1"))
                Else
                    Exit Sub      ' nothing to draw without color
                End If
                If .QueryString("c2") <> Nothing Then
                    _endColor = ColorTranslator.FromHtml("#" & .QueryString("c2"))
                Else
                    ' if no end color specified then fade to transparent
                    _endColor = _startColor
                End If

                If .QueryString("w") <> Nothing Then _width = CInt(.QueryString("w"))
                If .QueryString("h") <> Nothing Then _height = CInt(.QueryString("h"))

                If _startAlpha < 100 Then _startColor = Draw.AdjustAlpha(_startColor, _startAlpha)
                If _endAlpha < 100 Then _endColor = Draw.AdjustAlpha(_endColor, _endAlpha)

                Select Case .QueryString("d")
                    Case "ttb"
                        _angle = 90
                    Case "rtl"
                        _angle = 180
                    Case "btt"
                        _angle = 270
                    Case "ltr"
                        _angle = 0
                    Case "sunrise"
                        _angle = CSng(Math.Atan(_width / _height) * (180 / Math.PI))
                        _orthogonal = False
                    Case "sunset"
                        _angle = CSng(Math.Atan(_height / _width) * (180 / Math.PI)) + 90
                        _orthogonal = False
                    Case Else
                        '_angle = Integer.Parse(.QueryString("d"))
                        'If _angle = Nothing Then _angle = 0
                        _angle = 0
                End Select

                bitmap = MakeGradient()
                Me.CacheImage(cacheKey, bitmap)
            Else
                bitmap = DirectCast(Page.Cache(cacheKey), Bitmap)
            End If
        End With

        Response.Cache.AppendCacheExtension("post-check=3600,pre-check=43200")

        If transparency Then
            ' use PNG to support alpha
            Dim memory As New MemoryStream
            Response.ContentType = "image/png"
            bitmap.Save(memory, ImageFormat.Png)
            memory.WriteTo(Response.OutputStream)
        Else
            ' PNG gamma support varies by browser so otherwise use JPEG
            Response.ContentType = "image/jpeg"
            bitmap.Save(Response.OutputStream, ImageFormat.Jpeg)
        End If
        'memory.Close()
        'bitmap.Dispose()
    End Sub

    Private Function MakeGradient() As Bitmap
        Dim bitmap As New Bitmap(_width, _height)
        Dim graphic As Graphics = Graphics.FromImage(bitmap)
        Dim background As New Rectangle(0, 0, _width, _height)
        Dim gb As New LinearGradientBrush(background, _startColor, _endColor, _angle)

        If Not _orthogonal Then
            Dim cb As New ColorBlend
            cb.Positions = New Single() {0, 0.5, 1}
            cb.Colors = New Color() {_startColor, _endColor, _endColor}
            gb.InterpolationColors = cb
            'BugOut("Ending at {0} for {1}x{2} image of angle {3}", _endColor.ToString, _width, _height, _angle)
        End If

        graphic.FillRectangle(gb, background)
        gb.Dispose()

        Return bitmap

        'bitmap.Dispose()
        'graphic.Dispose()
    End Function

    Private Sub CacheImage(ByVal key As String, ByVal image As Bitmap)
        Page.Cache.Add(key, image, Nothing, _
            System.Web.Caching.Cache.NoAbsoluteExpiration, Nothing, _
            CacheItemPriority.Normal, Nothing)
    End Sub
End Class
