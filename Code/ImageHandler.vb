Imports System.Web
Imports System.Drawing.Imaging

Public Class ImageHandler
    Implements IHttpHandler

    ' http://msdn.microsoft.com/msdnmag/issues/04/04/CuttingEdge/default.aspx
    Public ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return True
        End Get
    End Property

    Public Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest
        Dim image As Object = HttpContext.Current.Cache(context.Request("key"))
        If image Is Nothing Then Return

        ' TODO: fix this
        If image.GetType.Name = "Bitmap" Then
            Me.WriteImage(DirectCast(image, System.Drawing.Image))
        Else
            Me.WriteImage(DirectCast(image, Byte()))
        End If
    End Sub

    Private Sub WriteImage(ByVal image As Byte())
        With HttpContext.Current.Response
            .ContentType = "image/jpeg"
            .OutputStream.Write(image, 0, image.Length)
        End With
    End Sub

    Private Sub WriteImage(ByVal image As System.Drawing.Image)
        With HttpContext.Current.Response
            .ContentType = "image/jpeg"
            image.Save(.OutputStream, ImageFormat.Jpeg)
        End With
    End Sub

End Class
