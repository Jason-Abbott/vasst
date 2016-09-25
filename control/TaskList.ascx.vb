Imports System.Text

Namespace Controls
    Public Class TaskList
        Inherits AMP.Controls.WebPart

        Protected list As WebControls.DataList

        Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
            Dim entries As Boolean = Me.ContestEntries
            Dim assets As Boolean = Me.AssetContributions
            Dim list As New StringBuilder

            With list
                .Append("<ul id=""tasks"">")
                If entries OrElse assets Then
                    If entries Then .Append("<li><a href=""contest-entries.aspx"">There are new contest entries</a></li>")
                    If assets Then .Append("<li><a href=""resources.aspx"">There are new resource uploads</a></li>")
                Else
                    .Append("<li>No outstanding tasks</li>")
                End If
                .Append("</ul>")
            End With

            Me.Controls.Add(New LiteralControl(list.ToString))
            Me.Title = "Task List"
            Me.Minimizable = True
        End Sub

        Private Function ContestEntries() As Boolean
            For Each c As AMP.Contest In WebSite.Contests.Active
                If c.Entries.Unapproved.Count > 0 Then Return True
            Next
            Return False
        End Function

        Private Function AssetContributions() As Boolean
            Return (WebSite.Assets.Unapproved.Count > 0)
        End Function
    End Class
End Namespace