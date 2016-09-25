Namespace Pages
    Public Class Home
        Inherits AMP.Page

        Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
            Me.Title = "Learn it FASST!"
            Profile.SearchResults = Nothing
            Profile.WriteTestCookie()
        End Sub
    End Class
End Namespace
