Imports AMP.Site

Namespace Pages.Administration
    Public Class Home
        Inherits AMP.AdminPage

        Private Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
            If Page.IsPostBack Then
                WebSite.UseSaveBuffer = False
                WebSite.Save()
                Profile.Message = "A new data file has been written"
            End If

            If Profile.User.HasPermission(Permission.EditProduct) Then
                ' write legacy authentication cookie
                Response.Cookies.Add(New HttpCookie("loggedin", "vr8735slkdj8e#!421%2598326^#s"))
            End If
        End Sub
    End Class
End Namespace