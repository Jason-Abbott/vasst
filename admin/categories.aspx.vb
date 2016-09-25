Namespace Pages.Administration
    Public Class Categories
        Inherits AMP.AdminPage

#Region " Controls "

        Protected ampSectionList As AMP.Controls.EnumCheckbox
        Protected rptCategories As Repeater

#End Region

        Private Sub Page_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Init
            If Not Profile.User.HasPermission(AMP.Site.Permission.EditCategories) Then
                Profile.Message = Me.Say("Error.Permissions")
                Me.SendBack()
            End If
            ampSectionList.Type = GetType(AMP.Site.Section)
        End Sub

        Private Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
            Me.StyleSheet.Add("admin/categories")
            Me.ScriptFile.Add("broker/common")
            Me.ScriptFile.Add("categories")

            rptCategories.DataSource = WebSite.Categories
            rptCategories.DataBind()
        End Sub
    End Class
End Namespace
