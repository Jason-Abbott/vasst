Namespace Pages
    Public Class Privacy
        Inherits AMP.Page

        Protected phNoCookies As PlaceHolder

        Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
            phNoCookies.Visible = (Request.QueryString("cookies") = "no")
        End Sub

    End Class
End Namespace
