Namespace Pages
    Public Class Signin
        Inherits AMP.Page

        Protected ampSignInUp As AMP.Controls.SignInUp

        Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
            Me.StyleSheet.Add("form")
            Me.StyleSheet.Add("share")
        End Sub

    End Class
End Namespace