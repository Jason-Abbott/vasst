Namespace Pages.Administration
    Public Class Products
        Inherits AMP.AdminPage

#Region " Controls "

        Protected lbCategories As AMP.Controls.ListBox
        Protected ampSectionList As AMP.Controls.EnumList
        Protected fldTitle As AMP.Controls.Field
        Protected tbDescription As TextBox
        ' dates
        Protected fldShowOn As AMP.Controls.Field
        Protected fldHideAfter As AMP.Controls.Field

#End Region

        Private Sub Page_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Init
            Me.StyleSheet.Add("admin/product")

            ampSectionList.Type = GetType(AMP.Site.Section)

            lbCategories.DataSource = WebSite.Categories.ForEntity(AMP.Site.Entity.Product)
            lbCategories.DataBind()
        End Sub

    End Class
End Namespace
