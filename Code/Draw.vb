Option Explicit On 
Option Strict On

Imports System
Imports System.Web.HttpContext
Imports System.Text
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Drawing.Imaging
Imports System.IO
Imports System.Configuration.ConfigurationSettings

Public Class Draw

#Region " Collections "

    Public Enum Position
        TopLeft
        TopRight
        BottomRight
        BottomLeft
    End Enum

    Public Enum ToggleState
        Open
        Closed
    End Enum

    Public Class Colors
        Public BackGround As Color = System.Drawing.Color.Transparent()
        Public ForeGround As Color
        Public Border As Color
        Public Text As Color
        Public Highlight As Color
    End Class

#End Region

    Private _imageName As String
    Private _saveName As String
    Private _format As ImageFormat = ImageFormat.Png
    Private _longName As Boolean = False
    Private _quantize As Boolean = False
    Private _paletteSize As Integer = 255
    Private _width As Integer
    Private _height As Integer
    Private _borderWidth As Integer = 2
    Private _orientation As Position
    ' list colors in adjacent order for anti-aliasing, starting with background
    Private _adjacentColors As New ArrayList
    Private _forceRegeneration As Boolean = False
    Private Const _minPixelWidth As Single = 0.3
    Public Color As New Draw.Colors

#Region " Properties "

    Public WriteOnly Property LongName() As Boolean
        Set(ByVal Value As Boolean)
            _longName = Value
        End Set
    End Property


    Public ReadOnly Property Transparent() As Boolean
        Get
            Return Color.BackGround.Equals(System.Drawing.Color.Transparent)
        End Get
    End Property

    Public ReadOnly Property Extension() As String
        Get
            Return _format.ToString.ToLower
        End Get
    End Property

    Public WriteOnly Property Format() As ImageFormat
        Set(ByVal Value As ImageFormat)
            _format = Value
            If _format.Equals(ImageFormat.Gif) Then Me.Quantize = True
        End Set
    End Property

    Public WriteOnly Property ForceRegeneration() As Boolean
        Set(ByVal Value As Boolean)
            _forceRegeneration = Value
        End Set
    End Property

    Public WriteOnly Property BorderWidth() As Integer
        Set(ByVal Value As Integer)
            _borderWidth = Value
        End Set
    End Property

    Public WriteOnly Property Orientation() As Position
        Set(ByVal Value As Position)
            _orientation = Value
        End Set
    End Property

    Public WriteOnly Property Width() As Integer
        Set(ByVal Value As Integer)
            _width = Value
        End Set
    End Property

    Public Property Height() As Integer
        Get
            Return _height
        End Get
        Set(ByVal Value As Integer)
            _height = Value
        End Set
    End Property

    Public WriteOnly Property PaletteSize() As Integer
        Set(ByVal Value As Integer)
            _paletteSize = Value
        End Set
    End Property

    Public WriteOnly Property Quantize() As Boolean
        Set(ByVal Value As Boolean)
            _quantize = Value
        End Set
    End Property

#End Region

#Region " Toggler "

    '---COMMENT---------------------------------------------------------------
    '   retrieve +/- toggle image path
    ' 
    '	Date:		Name    Description:
    '	11/4/04     JEA     Created
    '-------------------------------------------------------------------------
    Public Function Toggler(ByVal type As ToggleState) As String
        Dim fileName As New StringBuilder

        If _longName Then
            With _adjacentColors
                If Not Me.Transparent Then .Add(Me.Color.BackGround)
                .Add(Me.Color.ForeGround)
            End With
        End If

        With fileName
            .Append("toggle")
            If _longName Then
                .Append("_")
                .Append(AdjacentColorValues())
                .Append("-")
                .Append(_height)
            End If
            .Append(IIf(type = ToggleState.Open, "-.", "+."))
            .Append(Me.Extension)
        End With

        If Me.NeedNew(fileName) Then Me.CreateToggler()
        Return _imageName
    End Function

    '---COMMENT---------------------------------------------------------------
    '   create both graphics for +/- toggle
    ' 
    '	Date:		Name    Description:
    '	11/4/04     JEA     Created
    '-------------------------------------------------------------------------
    Private Sub CreateToggler()
        Dim toggle As New Bitmap(_height, _height)
        Dim graphic As Graphics = Graphics.FromImage(toggle)
        Dim pen As New Pen(Me.Color.ForeGround, 1)
        Dim point1 As PointF
        Dim point2 As PointF

        With graphic
            point1 = New PointF(3, CSng(_height / 2))
            point2 = New PointF(_height - 4, CSng(_height / 2))

            If Not Me.Transparent Then .Clear(Me.Color.BackGround)
            .DrawRectangle(pen, 0, 0, _height - 1, _height - 1)
            .DrawLine(pen, point1, point2)

            _saveName = _saveName.Replace("+.", "-.")
            Me.Save(DirectCast(toggle.Clone, Bitmap))

            ' draw vertical line to form "+"
            point1 = New PointF(CSng(_height / 2), 3)
            point2 = New PointF(CSng(_height / 2), _height - 4)
            .DrawLine(pen, point1, point2)

            _saveName = _saveName.Replace("-.", "+.")
            Me.Save(toggle)
        End With

        pen.Dispose()
        graphic.Dispose()
        toggle.Dispose()
    End Sub

#End Region

#Region " Rating Stars "

    '---COMMENT---------------------------------------------------------------
    '   build a star image name and return or generate the image
    ' 
    '	Date:		Name    Description:
    '	10/24/04    JEA     Created
    '-------------------------------------------------------------------------
    Public Function RatingStars(ByVal rating As Single, ByVal points As Integer, ByVal sharpness As Double) As String
        Dim fileName As New StringBuilder

        If _longName Then
            With _adjacentColors
                If Not Me.Transparent Then .Add(Me.Color.BackGround)
                .Add(Me.Color.Border)
                .Add(Me.Color.ForeGround)
            End With
        End If

        With fileName
            .Append("rating_")
            .Append(rating)
            .Append("-")
            If _longName Then
                .Append(AdjacentColorValues())
                .Append("-")
            End If
            .Append(_height)
            .Append(".")
            .Append(Me.Extension)
        End With

        If Me.NeedNew(fileName) Then Me.CreateRatingStars(rating, points, sharpness)
        Return _imageName
    End Function

    '---COMMENT---------------------------------------------------------------
    '   create graphic of five stars representing rating
    ' 
    '	Date:		Name    Description:
    '	10/24/04    JEA     Created
    '   12/23/04    JEA     Use gradient brush
    '-------------------------------------------------------------------------
    Private Sub CreateRatingStars(ByVal rating As Single, ByVal points As Integer, ByVal sharpness As Double)
        Dim remainingRating As Single = rating
        Dim star As New Bitmap(_height * 5, _height)
        Dim graphic As Graphics = Graphics.FromImage(star)
        'Dim brush As New SolidBrush(Me.Color.ForeGround)
        Dim pen As New Pen(Me.Color.Border, _borderWidth)
        Dim path As GraphicsPath = Me.BuildStarPath(points, sharpness)
        Dim brush As PathGradientBrush = Me.GetGradientBrush(path)
        Dim move As New Matrix(1, 0, 0, 1, _height, 0)

        With graphic
            .SmoothingMode = SmoothingMode.HighQuality

            ' position and draw remaining stars
            For x As Integer = 0 To 5
                If x > 0 Then
                    path.Transform(move)
                    brush.TranslateTransform(_height, 0)
                End If

                If remainingRating > 0 Then
                    If remainingRating < 1 Then
                        ' make clipping region to show partial star
                        Dim partWidth As Single = CSng((x + remainingRating) * _height) + 1
                        .Clip = New Region(New RectangleF(New PointF(0, 0), New SizeF(partWidth, _height)))
                    End If
                    ' draw the outline
                    .FillPolygon(brush, path.PathPoints)
                End If
                .ResetClip()
                .DrawPolygon(pen, path.PathPoints)
                '.FillPath(brush, path)
                remainingRating -= 1
            Next
        End With

        Me.Save(star)

        brush.Dispose()
        pen.Dispose()
        graphic.Dispose()
        star.Dispose()
    End Sub

#End Region

#Region " Star "

    '---COMMENT---------------------------------------------------------------
    '   build a star image name and return or generate the image
    ' 
    '	Date:		Name    Description:
    '	10/23/04    JEA     Created
    '-------------------------------------------------------------------------
    Public Function Star(ByVal points As Integer, ByVal sharpness As Double) As String
        Dim fileName As New StringBuilder

        If _longName Then
            With _adjacentColors
                If Not Me.Transparent Then .Add(Me.Color.BackGround)
                .Add(Me.Color.Border)
                .Add(Me.Color.ForeGround)
            End With
        End If

        With fileName
            .Append("star_")
            .Append(points)
            .Append("-")
            .Append(sharpness)
            If _longName Then
                .Append(AdjacentColorValues())
                .Append("-")
            End If
            .Append(_height)
            .Append(".")
            .Append(Me.Extension)
        End With

        If Me.NeedNew(fileName) Then Me.CreateStar(points, sharpness)
        Return _imageName
    End Function

    '---COMMENT---------------------------------------------------------------
    '   create single star and save
    ' 
    '	Date:		Name    Description:
    '	10/23/04    JEA     Created
    '   12/23/04    JEA     Use gradient brush
    '-------------------------------------------------------------------------
    Private Sub CreateStar(ByVal points As Integer, ByVal sharpness As Double)
        Dim star As New Bitmap(_height, _height)
        Dim graphic As Graphics = Graphics.FromImage(star)
        Dim pen As New Pen(Me.Color.Border, _borderWidth)
        Dim path As GraphicsPath = Me.BuildStarPath(points, sharpness)
        Dim brush As PathGradientBrush = Me.GetGradientBrush(path)

        With graphic
            .SmoothingMode = SmoothingMode.HighQuality
            If Not Me.Transparent Then .Clear(Me.Color.BackGround)
            '.FillPath(brush, path)
            .FillPolygon(brush, path.PathPoints)
            .DrawPolygon(pen, path.PathPoints)
        End With

        Me.Save(star)

        brush.Dispose()
        pen.Dispose()
        graphic.Dispose()
        star.Dispose()
    End Sub

    '---COMMENT---------------------------------------------------------------
    '   create path representing star edge
    ' 
    '	Date:		Name    Description:
    '	12/23/04    JEA     Created
    '-------------------------------------------------------------------------
    Private Function BuildStarPath(ByVal points As Integer, ByVal sharpness As Double) As GraphicsPath
        Dim path As New GraphicsPath
        path.AddLines(Me.BuildStarPoints(points, sharpness))
        Return path
    End Function

    '---COMMENT---------------------------------------------------------------
    '   create star graphic; angles measured in radians
    ' 
    '	Date:		Name    Description:
    '	10/23/04    JEA     Created
    '   10/25/04    JEA     Subtract any border width from radius
    '-------------------------------------------------------------------------
    Private Function BuildStarPoints(ByVal points As Integer, ByVal sharpness As Double) As PointF()
        Dim bleed As Double = 0

        If _borderWidth > 1 Then
            ' need to reduce star radius to accomodate border width
            ' compute based on triangle formed in unit circle (radius = 1)
            Dim legEdge As Double
            Dim crossMultiple As Double
            Dim sineOfPointAngle As Double
            Dim redux As Double

            ' law of cosines
            legEdge = Math.Sqrt((sharpness ^ 2 + 1) - (2 * sharpness * Math.Cos(Math.PI / points)))
            ' law of sines
            crossMultiple = sharpness * Math.Sin(Math.PI / points)
            sineOfPointAngle = crossMultiple / legEdge
            ' height of point formed at vertex by pen width
            bleed = ((_borderWidth - 1) * legEdge) / (2 * crossMultiple)

            ' bleed is mathematically correct but pixel fractions don't display for
            ' excessively sharp points, so reduce bleed point to minimum pixel width
            redux = (_minPixelWidth / 2) / Math.Tan(Math.Asin(sineOfPointAngle))

            bleed -= redux
            If bleed < 0 Then bleed = 0
        End If

        Dim radius As Single = CSng((_height / 2) - bleed)
        Dim innerRadius As Single = radius - CSng(radius * sharpness)
        Dim vertices As PointF()
        Dim factor As Single
        ReDim vertices(points * 2 - 1)
        Dim cosine As Double
        Dim sine As Double

        For x As Integer = 1 To vertices.Length
            ' alternating inner and outer points
            factor = CInt(IIf(x Mod 2 = 0, innerRadius, radius))
            With vertices(x - 1)
                ' compute coordinates on unit circle; pi/2 radians = 90 degree start position
                .X = CSng(Math.Cos((x * Math.PI / points) + (Math.PI / 2)))
                .Y = CSng(Math.Sin((x * Math.PI / points) + (Math.PI / 2)))
                ' multiply size to match specified radius
                .X *= factor
                .Y *= factor
                ' reposition so star center is in center of image
                .X += CSng(_height / 2)
                .Y += CSng(_height / 2)
            End With
        Next
        Return vertices
    End Function

    Function BuildStarPoints2(ByVal points As Integer, ByVal sharpness As Double) As PointF()
        ' from http://www.bobpowell.net/pgb.htm
        Dim vertices(points - 1) As PointF
        Dim inner As Boolean = False
        Dim point As Integer = 0
        Dim angle As Single

        Do Until angle <= Math.PI * 2
            Dim length As Single = 50 + (80 * CInt(IIf(inner, 0, 1)))
            vertices(point) = New PointF(CSng(Math.Cos(angle - Math.PI / 2) * length), _
                CSng(Math.Sin(angle - Math.PI / 2) * length))
            angle += CSng(Math.PI / 5)
            inner = Not inner
            point += point
        Loop
        vertices(point) = vertices(0)

        Return vertices
    End Function

#End Region

#Region " Button "

    '---COMMENT---------------------------------------------------------------
    '   build a button image name and return or generate the image
    ' 
    '	Date:		Name    Description:
    '	10/21/04    JEA     Created
    '-------------------------------------------------------------------------
    Public Function Button(ByVal name As String, ByVal drawHighlight As Boolean) As String
        Dim fileName As New StringBuilder

        If _longName Then
            With _adjacentColors
                .Clear()
                If Not Me.Transparent Then .Add(Me.Color.BackGround)
                .Add(Me.Color.ForeGround)
                .Add(Me.Color.Text)
            End With
        End If

        With fileName
            .Append("btn_")
            .Append(name.Replace(" ", ""))
            If _longName Then
                .Append("_")
                .Append(AdjacentColorValues())
                .Append("-")
                If _width = Nothing Then
                    .Append(_height)
                Else
                    .AppendFormat("{0}by{1}", _width, _height)
                End If
            End If
            'If _highLight Then .Append("_on")
            .Append(".")
            .Append(Me.Extension)
        End With

        'If Me.NeedNew(fileName) Then Me.CreateRoundedButton(name, drawHighlight)
        If Me.NeedNew(fileName) Then Me.CreateButtonFromTemplate(name, Nothing)
        Return _imageName
    End Function

    '---COMMENT---------------------------------------------------------------
    '   create button graphic
    ' 
    '	Date:		Name    Description:
    '	10/21/04    JEA     Created
    '-------------------------------------------------------------------------
    Private Sub CreateRoundedButton(ByVal name As String, ByVal drawHighlight As Boolean)

        Dim pixelsFromLeft As Integer
        Dim button As New Bitmap(1, 1)
        Dim graphic As Graphics = Graphics.FromImage(button)
        Dim brush As New SolidBrush(Me.Color.ForeGround)
        Dim pen As New SolidBrush(Me.Color.Text)
        Dim penFont As New Font("Tahoma", _height - 2, FontStyle.Bold, GraphicsUnit.Pixel)
        Dim stringSize As New SizeF

        graphic.SmoothingMode = SmoothingMode.HighQuality
        graphic.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias
        stringSize = graphic.MeasureString(name, penFont)

        If _width = Nothing Then
            ' compute width based on rendered string length
            _width = CInt(stringSize.Width) + _height - 4
            pixelsFromLeft = CInt(_height / 2) - 2
        Else
            ' center string within fixed width
            pixelsFromLeft = CInt((_width - CInt(stringSize.Width)) / 2)
        End If

        button = New Bitmap(_width, _height)
        graphic = Graphics.FromImage(button)

        graphic.SmoothingMode = SmoothingMode.HighQuality
        graphic.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias

        ' draw left arc
        graphic.FillEllipse(brush, 0, 0, _height - 1, _height - 1)
        ' draw right arc
        graphic.FillEllipse(brush, (_width - _height), 0, _height - 1, _height - 1)
        ' draw background
        graphic.FillRectangle(brush, CInt(_height / 2) - 1, -1, (_width - _height) + 1, _height + 1)
        ' write text
        graphic.DrawString(name, penFont, pen, pixelsFromLeft, -1)

        Me.Save(button)

        If drawHighlight Then
            graphic.Clear(System.Drawing.Color.FromArgb(0, 0, 0, 0))

            brush = New SolidBrush(Me.Color.Highlight)
            ' draw left arc
            graphic.FillEllipse(brush, 0, 0, _height - 1, _height - 1)
            ' draw right arc
            graphic.FillEllipse(brush, (_width - _height), 0, _height - 1, _height - 1)
            ' draw background
            graphic.FillRectangle(brush, CInt(_height / 2) - 1, -1, (_width - _height) + 1, _height + 1)
            ' write text
            graphic.DrawString(name, penFont, pen, pixelsFromLeft, -1)

            _saveName = _saveName.Replace("." & Me.Extension, "_on." & Me.Extension)
            Me.Save(button)
        End If

        brush.Dispose()
        graphic.Dispose()
        button.Dispose()
    End Sub

    '---COMMENT---------------------------------------------------------------
    '   create button graphic from template graphic
    ' 
    '	Date:		Name    Description:
    '	10/21/04    JEA     Created
    '-------------------------------------------------------------------------
    Private Sub CreateButtonFromTemplate(ByVal text As String, ByVal file As String)
        Dim pixelsFromLeft As Integer
        Dim button As New Bitmap(1, 1)
        Dim buttonOn As Bitmap
        Dim graphic As Graphics = Graphics.FromImage(button)
        Dim graphicOn As Graphics
        Dim brush As New SolidBrush(Me.Color.ForeGround)
        Dim pen As New SolidBrush(Me.Color.Text)
        Dim image As System.Drawing.Image
        Dim imageOn As System.Drawing.Image
        Dim stringSize As New SizeF
        Dim penFont As Font
        Dim endWidth As Integer = 9

        file = Current.Server.MapPath(AppSettings("ButtonTemplate"))

        ' get image to find height and determine font size
        image = System.Drawing.Image.FromFile(file)
        imageOn = System.Drawing.Image.FromFile(file.Replace(".png", "_on.png"))

        _height = image.Height      ' get height from template

        'Dim bmp As New Bitmap(image) !

        penFont = New Font("Tahoma", _height - 6, FontStyle.Bold, GraphicsUnit.Pixel)

        graphic.SmoothingMode = SmoothingMode.HighQuality
        graphic.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias
        stringSize = graphic.MeasureString(text, penFont)

        If _width = Nothing Then
            ' compute width based on rendered string length
            _width = CInt(stringSize.Width) + _height + 8 ' - 4
            'pixelsFromLeft = CInt(_height / 2) - 2
            'pixelsFromLeft = cint(
        Else
            ' center string within fixed width
            pixelsFromLeft = CInt((_width - CInt(stringSize.Width)) / 2)
        End If
        pixelsFromLeft = CInt((_width - CInt(stringSize.Width)) / 2)

        'Dim leftEnd As New Bitmap(5, _height)
        'Dim rightEnd As New Bitmap(5, _height)

        ' specify rectangles for button ends
        Dim left As New Rectangle(0, 0, endWidth, _height)
        Dim right As New Rectangle(_width - endWidth, 0, endWidth, _height)
        Dim leftSrc As New Rectangle(0, 0, endWidth, _height)
        Dim rightSrc As New Rectangle(image.Width - endWidth, 0, endWidth, _height)
        Dim midSrc As New Rectangle(CInt(image.Width / 2), 0, 1, _height)

        button = New Bitmap(_width, _height)
        graphic = Graphics.FromImage(button)

        buttonOn = New Bitmap(_width, _height)
        graphicOn = Graphics.FromImage(buttonOn)

        With graphicOn
            .SmoothingMode = SmoothingMode.HighQuality
            .TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias

            ' draw the image template ends onto the button bitmap
            .DrawImage(imageOn, left, leftSrc, GraphicsUnit.Pixel)
            .DrawImage(imageOn, right, rightSrc, GraphicsUnit.Pixel)

            For x As Integer = endWidth To (_width - endWidth) - 1
                .DrawImage(imageOn, New Rectangle(x, 0, 1, _height), midSrc, GraphicsUnit.Pixel)
            Next

            ' write text
            .DrawString(text, penFont, pen, pixelsFromLeft, 1)
        End With

        With graphic
            .SmoothingMode = SmoothingMode.HighQuality
            .TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias

            ' draw the image template ends onto the button bitmap
            .DrawImage(image, left, leftSrc, GraphicsUnit.Pixel)
            .DrawImage(image, right, rightSrc, GraphicsUnit.Pixel)

            For x As Integer = endWidth To (_width - endWidth) - 1
                .DrawImage(image, New Rectangle(x, 0, 1, _height), midSrc, GraphicsUnit.Pixel)
            Next

            ' write text
            .DrawString(text, penFont, pen, pixelsFromLeft, 1)
        End With

        'image.
        Me.Save(button)
        _saveName = _saveName.Replace("." & Me.Extension, "_on." & Me.Extension)
        Me.Save(buttonOn)

        pen.Dispose()
        brush.Dispose()
        graphic.Dispose()
        image.Dispose()
        button.Dispose()
    End Sub

#End Region

#Region " Corner "

    '---COMMENT---------------------------------------------------------------
    '   build an image name and return or generate the image
    ' 
    '	Date:		Name    Description:
    '	10/21/04    JEA     Created
    '-------------------------------------------------------------------------
    Public Function Corner() As String
        Dim fileName As New StringBuilder

        With _adjacentColors
            If Not Me.Transparent Then .Add(Me.Color.BackGround)
            .Add(Me.Color.Border)
            .Add(Me.Color.ForeGround)
        End With

        With fileName
            .Append("corner_")
            Select Case _orientation
                Case Position.TopLeft
                    .Append("tl-")
                Case Position.TopRight
                    .Append("tr-")
                Case Position.BottomRight
                    .Append("br-")
                Case Position.BottomLeft
                    .Append("bl-")
            End Select
            .Append(AdjacentColorValues())
            .AppendFormat("-{0}by{1}", _width, _height)
            .Append(".")
            .Append(Me.Extension)
        End With

        If Me.NeedNew(fileName) Then Me.CreateRoundedCorner()
        Return _imageName
    End Function

    '---COMMENT---------------------------------------------------------------
    '   get the next color evenly spaced between two colors
    ' 
    '	Date:		Name    Description:
    '   ?                   CodeProject
    '	10/21/04    JEA     Cleaned up logic
    '-------------------------------------------------------------------------
    Private Sub CreateRoundedCorner()
        ' generate initial bitmap
        Dim corner As New Bitmap(_width, _height)
        Dim graphic As Graphics = Graphics.FromImage(corner)

        graphic.SmoothingMode = SmoothingMode.HighQuality

        ' give the new image our background color
        If Not Me.Transparent Then graphic.Clear(Me.Color.BackGround)

        ' define pen to draw border edge
        Dim pen As New Pen(Me.Color.Border, _borderWidth)
        pen.DashStyle = DashStyle.Solid

        ' define a brush for the foreground
        Dim brush As New SolidBrush(Me.Color.ForeGround)

        ' compute arc start coordinates
        Dim x As Integer = 0
        Dim y As Integer = 0
        Select Case _orientation
            Case Position.TopLeft
                x = CInt(_borderWidth / 3)
                y = CInt(_borderWidth / 3)
            Case Position.TopRight
                x = -_width - 1
                x -= CInt(_borderWidth / 3)
                y += CInt(_borderWidth / 3)
            Case Position.BottomRight
                x = -_width - 1
                y = -_height - 1
                x -= CInt(_borderWidth / 3)
                y -= CInt(_borderWidth / 3)
            Case Position.BottomLeft
                y = -_height - 1
                x += CInt(_borderWidth / 3)
                y -= CInt(_borderWidth / 3)
        End Select

        ' draw arc
        graphic.FillEllipse(brush, x, y, 2 * _width, 2 * _height)
        graphic.DrawArc(pen, x, y, 2 * _width, 2 * _height, 0, 360)

        Me.Save(corner)

        brush.Dispose()
        pen.Dispose()
        graphic.Dispose()
        corner.Dispose()
    End Sub

#End Region

    '---COMMENT---------------------------------------------------------------
    '   get gradient brush along given path with color properties
    ' 
    '	Date:		Name    Description:
    '	12/23/04    JEA     Created
    '-------------------------------------------------------------------------
    Private Function GetGradientBrush(ByVal path As GraphicsPath) As PathGradientBrush
        Dim brush As New PathGradientBrush(path)
        Dim b As New Blend

        brush.CenterColor = Me.Color.ForeGround
        brush.SurroundColors = New Color() {Me.Color.Highlight}

        b.Factors = New Single() {1, 0}
        b.Positions = New Single() {0, 1}

        brush.Blend = b
        '.Transform = New Matrix(1, 0, 0, 1, CSng(_height / 2), CSng(_height / 2))
        Return brush
    End Function

    '---COMMENT---------------------------------------------------------------
    '   return path object for given points
    ' 
    '	Date:		Name    Description:
    '	12/23/04    JEA     Created
    '-------------------------------------------------------------------------
    Private Function GetPointsPath(ByVal points As PointF()) As GraphicsPath
        Dim path As New GraphicsPath
        path.AddLines(points)
        Return path
    End Function

    Private Function NeedNew(ByVal name As String) As Boolean
        Me.GetNames(name)
        Return (_forceRegeneration OrElse Not IO.File.Exists(_saveName))
    End Function

    Private Function NeedNew(ByVal name As StringBuilder) As Boolean
        Return NeedNew(name.ToString)
    End Function

    Private Sub GetNames(ByVal name As String)
        _imageName = String.Format("{0}/{1}/{2}", Global.BasePath, _
            AppSettings("GeneratedImageFolder"), name)
        _saveName = Current.Server.MapPath(_imageName)
    End Sub

    Private Sub GetNames(ByVal name As StringBuilder)
        GetNames(name.ToString)
    End Sub

    '---COMMENT---------------------------------------------------------------
    '   set alpha value on given color
    ' 
    '	Date:		Name    Description:
    '	11/18/04    JEA     Created
    '-------------------------------------------------------------------------
    Public Shared Function AdjustAlpha(ByVal c As Color, ByVal alpha As Integer) As System.Drawing.Color
        Dim r As Integer = c.R
        Dim g As Integer = c.G
        Dim b As Integer = c.B
        Return System.Drawing.Color.FromArgb(CInt(2.55 * alpha), r, g, b)
    End Function

    '---COMMENT---------------------------------------------------------------
    '   build string of adjacent color values to supply unique file name
    ' 
    '	Date:		Name    Description:
    '	10/22/04    JEA     Created
    '-------------------------------------------------------------------------
    Private Function AdjacentColorValues() As String
        Dim colors As New StringBuilder
        Dim color As System.Drawing.Color
        With colors
            For x As Integer = 0 To _adjacentColors.Count - 1
                color = DirectCast(_adjacentColors(x), System.Drawing.Color)
                .Append(ColorTranslator.ToHtml(color).Replace("#", "").TrimStart("0".ToCharArray))
                .Append(",")
            Next
        End With
        Return colors.ToString.TrimEnd(",".ToCharArray)
    End Function

    '---COMMENT---------------------------------------------------------------
    '   get the next color evenly spaced between two colors; use to create gradient
    ' 
    '	Date:		Name    Description:
    '	10/21/04    JEA     Created
    '-------------------------------------------------------------------------
    Private Function PickColorBetween(ByVal c1 As Color, ByVal c2 As Color, _
        ByVal thisStep As Integer, ByVal steps As Integer) As Color

        Dim r, g, b As Integer

        r = c1.R + CInt(thisStep * ((-c1.R + c2.R) / steps))
        g = c1.G + CInt(thisStep * ((-c1.G + c2.G) / steps))
        b = c1.B + CInt(thisStep * ((-c1.B + c2.B) / steps))

        Return System.Drawing.Color.FromArgb(r, g, b)
    End Function

    '---COMMENT---------------------------------------------------------------
    '   create custom three color palette to be used by the custom quantizer
    '   for anti-aliasing, colors displayed adjacently should be passed adjacently
    ' 
    '	Date:	    Name:	Description:
    '	10/21/04    JEA	    Created
    '   11/19/04    JEA     Handle variable length color list
    '-------------------------------------------------------------------------
    Private Function CreatePalette() As ArrayList
        Dim palette As New ArrayList(_paletteSize)
        Dim color1 As System.Drawing.Color
        Dim color2 As System.Drawing.Color

        For x As Integer = 0 To _adjacentColors.Count - 1
            palette.Add(_adjacentColors(x))
        Next

        With palette
            ' compute gradient steps based on palette size
            Dim blends As Integer = palette.Count - 1

            If blends > 0 Then
                ' build the palette up with color blends between those adjacent to each other
                Dim steps As Integer = CInt((_paletteSize - .Count) / blends)

                For thisStep As Integer = 1 To steps
                    For x As Integer = 0 To blends - 1
                        color1 = DirectCast(palette(x), System.Drawing.Color)
                        color2 = DirectCast(palette(x + 1), System.Drawing.Color)
                        .Add(PickColorBetween(color1, color2, thisStep, steps))
                    Next
                Next
            End If

            ' arbitrarily fill remainder of palette with first color
            For x As Integer = 0 To _paletteSize - .Count
                .Add(_adjacentColors(0))
            Next
        End With

        Return palette
    End Function

    '---COMMENT---------------------------------------------------------------
    '   save generated images
    ' 
    '	Date:	    Name:	Description:
    '	10/21/04    JEA	    Created
    '   2/11/05     JEA     Added error check
    '-------------------------------------------------------------------------
    Public Sub Save(ByRef image As Bitmap)
        Try
            image.Save(_saveName, _format)
        Catch e As Exception
            Log.Error(e, Log.ErrorType.FileSystem, Profile.User)
        End Try
    End Sub

#Region " Resize "

    '--------------------------------------------------------------------
    '   resize image to fit given dimensions, maintaining aspect ratio
    '
    '	Date:		Name:	Description:
    '	4/22/04 	JEA	    Created
    '   4/30/04     JEA     If missing dimension parameter, set to ratio instead of equal
    '--------------------------------------------------------------------
    Public Function Resize(ByVal image As Drawing.Image, _
        ByVal newWidth As Double, ByVal newHeight As Double) As Drawing.Image

        Dim width As Double = image.Width
        Dim height As Double = image.Height
        Dim widthHeightRatio As Double = width / height

        ' set missing parameter
        If newWidth = 0 Then
            newWidth = newHeight * widthHeightRatio
        ElseIf newHeight = 0 Then
            newHeight = newWidth / widthHeightRatio
        End If

        If width > newWidth Or height > newHeight Then
            ' image needs to be downsized
            If width / newWidth > height / newHeight Then
                ' to maintain aspect ratio, height must be smaller than passed
                newHeight = CType(newWidth / widthHeightRatio, Integer)
            Else
                ' to maintain aspect ratio, width must be smaller than passed
                newWidth = CType(newHeight * widthHeightRatio, Integer)
            End If
            image.RotateFlip(RotateFlipType.Rotate180FlipNone)
            image.RotateFlip(RotateFlipType.Rotate180FlipNone)
            Resize = image.GetThumbnailImage(CType(newWidth, Integer), CType(newHeight, Integer), Nothing, IntPtr.Zero)

            'Dim g As Graphics = Graphics.FromImage(image)
            'With g
            '    .CompositingMode = CompositingMode.SourceOver
            '    .CompositingQuality = CompositingQuality.HighQuality
            '    .SmoothingMode = SmoothingMode.HighQuality
            '    .InterpolationMode = InterpolationMode.HighQualityBicubic
            '    .PixelOffsetMode = Drawing2D.PixelOffsetMode.HighQuality
            '    .DrawImage(image, 0, 0, newWidth, newHeight)
            '    .Dispose()
            'End With
            'Resize = image
        Else
            Resize = image
        End If
    End Function

    ' overload accepting stream
    Public Function Resize(ByVal imageStream As System.IO.Stream, _
        ByVal newWidth As Integer, ByVal newHeight As Integer) As Drawing.Image

        Dim image As Drawing.Image = Drawing.Image.FromStream(imageStream)
        Resize = Resize(image, newWidth, newHeight)
    End Function

    ' overload accepting byte array
    Public Function Resize(ByVal imageBytes As Byte(), _
        ByVal newWidth As Integer, ByVal newHeight As Integer) As Drawing.Image

        Resize = Resize(New MemoryStream(imageBytes), newWidth, newHeight)
    End Function

    ' overload accepting file name
    Public Function Resize(ByVal imagePath As String, _
        ByVal newWidth As Integer, ByVal newHeight As Integer) As Drawing.Image

        Resize = Resize(Image.FromFile(imagePath), newWidth, newHeight)
    End Function

#End Region

    '---COMMENT---------------------------------------------------------------
    '   get encoder information for given MIME type
    ' 
    '	Date:	    Name:	Description:
    '	4/26/04     JEA	    Copied from Q324788
    '-------------------------------------------------------------------------
    Private Function GetEncoderInfo(ByVal mimeType As String) As ImageCodecInfo
        Dim j As Integer
        Dim encoders As ImageCodecInfo()
        encoders = ImageCodecInfo.GetImageEncoders()
        For j = 0 To encoders.Length
            If encoders(j).MimeType = mimeType Then
                Return encoders(j)
            End If
        Next j
        Return Nothing
    End Function

    '---COMMENT---------------------------------------------------------------
    '   save JPEG with 1-100 quality setting
    ' 
    '	Name:	Date:		Description:
    '	JEA	    4/26/04 	Copied from Q324788
    '-------------------------------------------------------------------------
    Private Sub SaveJPGWithCompressionSetting(ByVal image As Drawing.Image, _
        ByVal fileName As String, ByVal compression As Long)

        Dim eps As EncoderParameters = New EncoderParameters(1)
        eps.Param(0) = New EncoderParameter(Imaging.Encoder.Quality, compression)
        Dim ici As ImageCodecInfo = GetEncoderInfo("image/jpeg")
        image.Save(fileName, ici, eps)
    End Sub
End Class
