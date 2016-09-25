Imports System.IO
Imports System.Web.UI
Imports System.Web.Caching
Imports System.Web.HttpContext
Imports System.Drawing.Imaging
Imports System.Configuration.ConfigurationSettings

Namespace Controls
    Public Class Image
        Inherits AMP.Controls.HtmlControl

        Private _mouseOver As String
        Private _mouseOut As String
        Private _submitForm As Boolean = False
        Private _resize As Boolean = False
        Private _transparency As Boolean = False
        Private _src As String                      ' image source as file path
        Private _imageBytes As Byte()               ' image source as byte array
        Private _image As System.Drawing.Image      ' image source as object
        Private _height As Integer
        Private _width As Integer
        Private Const _blankImage As String = "/images/blank.gif"
        Private Const _cacheUrl As String = "cachedimage.axd?key={0}"
        Private Const _alphaFilter As String = "DXImageTransform.Microsoft.AlphaImageLoader"

#Region " Properties "

        Public Property Image() As System.Drawing.Image
            Get
                Return _image
            End Get
            Set(ByVal Value As System.Drawing.Image)
                _image = Value
            End Set
        End Property

        Public Property ImageBytes() As Byte()
            Get
                Return _imageBytes
            End Get
            Set(ByVal Value As Byte())
                _imageBytes = Value
            End Set
        End Property

        Public Property CacheKey() As String
            Get
                If Me.ViewState("CacheKey") Is Nothing Then
                    Return Nothing
                Else
                    Return Me.ViewState("CacheKey").ToString
                End If
            End Get
            Set(ByVal Value As String)
                Me.ViewState("CacheKey") = Value
            End Set
        End Property

        Public WriteOnly Property Resize() As Boolean
            Set(ByVal Value As Boolean)
                _resize = Value
            End Set
        End Property

        Public WriteOnly Property RollOver() As String
            Set(ByVal Value As String)
                MyBase.OnMouseOut = Value
                MyBase.OnMouseOver = Value
            End Set
        End Property

        Public WriteOnly Property Transparency() As Boolean
            Set(ByVal Value As Boolean)
                _transparency = Value
            End Set
        End Property

        Public WriteOnly Property Alt() As String
            Set(ByVal Value As String)
                MyBase.Attributes.Add("alt", Value)
            End Set
        End Property

        Public Property AlternateText() As String
            Get
                Return MyBase.Attributes.Item("alt")
            End Get
            Set(ByVal Value As String)
                MyBase.Attributes.Add("alt", Value)
            End Set
        End Property

        Public WriteOnly Property SubmitForm() As Boolean
            Set(ByVal Value As Boolean)
                _submitForm = Value
            End Set
        End Property

        Public Property Src() As String
            Get
                If _src <> Nothing Then
                    Return _src
                ElseIf Not (Me.Image Is Nothing AndAlso Me.ImageBytes Is Nothing) Then
                    ' generate src from cached image
                    If Me.CacheKey = Nothing Then Me.CacheKey = Guid.NewGuid.ToString
                    Return String.Format(_cacheUrl, Me.CacheKey)
                End If
            End Get
            Set(ByVal Value As String)
                _src = Value.Replace("~", Global.BasePath)
            End Set
        End Property

        Public Property Width() As Integer
            Get
                Return _width
            End Get
            Set(ByVal Value As Integer)
                If Value <> 0 Then
                    MyBase.Style.Add("width", String.Format("{0}px", Value))
                    _width = Value
                End If
            End Set
        End Property

        Public Property Height() As Integer
            Get
                Return _height
            End Get
            Set(ByVal Value As Integer)
                If Value <> 0 Then
                    MyBase.Style.Add("height", String.Format("{0}px", Value))
                    _height = Value
                End If
            End Set
        End Property

#End Region

        '---COMMENT---------------------------------------------------------------
        '	render image tag
        '
        '	Date:		Name:	Description:
        '	11/13/04	JEA		Creation
        '   11/17/04    JEA     Change sizing method if size specified
        '   1/4/05      JEA     Handle cached images
        '   2/28/05     JEA     Move some members to base class
        '-------------------------------------------------------------------------
        Protected Overrides Sub Render(ByVal writer As System.Web.UI.HtmlTextWriter)
            If Me.Src <> Nothing Then
                If Not Me.Image Is Nothing Then
                    Me.CacheImage(Me.Image)
                ElseIf Not Me.ImageBytes Is Nothing Then
                    Me.CacheImage(Me.ImageBytes)
                End If

                With writer
                    .Write(IIf(_submitForm, "<input type=""image"" ", "<img "))
                    .Write("src=""")
                    If _transparency AndAlso _
                        HttpContext.Current.Request.Browser.Browser = "IE" Then
                        ' WinIE needs DX filter to display png
                        .Write(Global.BasePath)
                        .Write(_blankImage)
                        MyBase.Style.Add("filter", _
                            String.Format("progid:{0}(src='{1}', sizingMethod='image')", _
                            _alphaFilter, Me.Src))
                    Else
                        .Write(Me.Src)
                    End If
                    .Write("""")
                    MyBase.RenderAttributes(writer)
                    .Write(" />")
                End With
            End If
        End Sub

        Private Sub CacheImage(ByVal image As Object)
            If Page.Cache(Me.CacheKey) Is Nothing Then
                Page.Cache.Add(Me.CacheKey, image, Nothing, _
                    System.Web.Caching.Cache.NoAbsoluteExpiration, TimeSpan.FromMinutes(5), _
                    CacheItemPriority.High, Nothing)
            End If
        End Sub
    End Class
End Namespace

