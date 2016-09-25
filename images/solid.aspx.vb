Imports System.IO
Imports System.Web.HttpContext
Imports System.Web.Caching
Imports System.Text
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Drawing.Imaging
Imports System.Configuration.ConfigurationSettings

Public Class Solid
    Inherits System.Web.UI.Page

    Private _color As Color = ColorTranslator.FromHtml(AppSettings.Item("BackGroundColor"))
    Private _width As Integer = 5
    Private _height As Integer = 5
    Private _alpha As Integer = 100

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
    '	build solid image for output
    '
    '	Date:		Name:	Description:
    '	11/10/04	JEA		Creation
    '   1/4/05      JEA     Added caching
    '-------------------------------------------------------------------------
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim bitmap As Bitmap
        Dim cacheKey As String = Request.Url.Query

        If Page.Cache(cacheKey) Is Nothing Then
            With Request
                If .QueryString("c") <> Nothing Then
                    _color = ColorTranslator.FromHtml("#" & .QueryString("c"))
                End If
                If .QueryString("w") <> Nothing Then _width = CInt(.QueryString("w"))
                If .QueryString("h") <> Nothing Then _height = CInt(.QueryString("h"))
                If .QueryString("a") <> Nothing Then _alpha = CInt(.QueryString("a"))
            End With

            bitmap = New Bitmap(_width, _height)
            Dim graphic As Graphics = Graphics.FromImage(bitmap)

            If _alpha < 100 Then _color = Draw.AdjustAlpha(_color, _alpha)

            Dim background As New Rectangle(0, 0, _width, _height)
            Dim brush As New SolidBrush(_color)

            graphic.FillRectangle(brush, background)
            Me.CacheImage(cacheKey, bitmap)
            graphic.Dispose()
        Else
            bitmap = DirectCast(Page.Cache(cacheKey), Bitmap)
        End If

        Dim memory As New MemoryStream
        ' http://www.aspnetresources.com/blog/cache_control_extensions.aspx
        Response.Cache.AppendCacheExtension("post-check=3600,pre-check=43200")
        Response.ContentType = "image/png"
        bitmap.Save(memory, ImageFormat.Png)
        memory.WriteTo(Response.OutputStream)
    End Sub

    Private Sub CacheImage(ByVal key As String, ByVal image As Bitmap)
        Page.Cache.Add(key, image, Nothing, _
            System.Web.Caching.Cache.NoAbsoluteExpiration, Nothing, _
            CacheItemPriority.Normal, Nothing)
    End Sub

End Class
