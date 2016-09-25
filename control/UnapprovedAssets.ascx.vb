Namespace Controls
    Public Class UnapprovedAssets
        Inherits AMP.Controls.WebPart

        Protected rptAssets As WebControls.Repeater

        Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
            Dim assets As ArrayList = WebSite.Assets.Unapproved

            Me.Minimizable = True

            If assets.Count > 0 Then
                Me.Title = Me.Page.Say("Title.UnapprovedResources")
            Else
                assets = WebSite.Assets.Newest(15, AMP.Site.Section.All)
                If assets.Count > 0 Then Me.Title = Me.Page.Say("Title.NewResources")
            End If

            If assets.Count > 0 Then
                rptAssets.DataSource = assets
                rptAssets.DataBind()
            Else
                Me.Visible = False
            End If
        End Sub

    End Class
End Namespace