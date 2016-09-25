Namespace Controls
    Public Class SearchResults
        Inherits AMP.Controls.WebPart

        Protected rptResults As WebControls.Repeater

        Public Sub New()
            Me.Minimizable = True
        End Sub

        Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
            Me.Title = Me.Page.Say("Title.SearchResults")
            Dim matches As ArrayList

            matches = Profile.SearchResults
            If matches Is Nothing Then
                rptResults.Visible = False
            Else
                rptResults.DataSource = matches
                rptResults.DataBind()
            End If
        End Sub

    End Class
End Namespace
