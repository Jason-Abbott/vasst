Public Class AdminPage
    Inherits AMP.Page

    Public Sub New()
        MyBase.TemplateFile = "~/template/ThreeColumnAdmin.ascx"
        MyBase.RequireAuthentication = True
    End Sub

    Private Sub Page_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Init
        MyBase.StyleSheet.Add("admin/common")
        MyBase.StyleSheet.Add("form")
    End Sub

End Class
