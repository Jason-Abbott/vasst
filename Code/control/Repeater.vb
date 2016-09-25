Namespace Controls
    Public Class Repeater
        Inherits System.Web.UI.webcontrols.Repeater

        Private Sub Repeater_PreRender(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.PreRender
            Me.Visible = (Me.Items.Count <> 0)
        End Sub
    End Class
End Namespace