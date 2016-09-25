Namespace Pages
    Public Class Import
        Inherits System.Web.UI.Page

        Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
            WebSite.ImportData()
        End Sub

    End Class
End Namespace
