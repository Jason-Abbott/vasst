Namespace Pages.Administration
    Public Class Users
        Inherits AMP.AdminPage

#Region " Controls "

        'Protected rptRoles As Repeater

#End Region

        Private Sub Page_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Init
            If Not Profile.User.HasPermission(AMP.Site.Permission.EditAnyUser) Then
                Profile.Message = Me.Say("Error.Permissions")
                Me.SendBack()
            End If
        End Sub

        Private Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
            'Me.ScriptFile.Add("broker/common")
            'Me.ScriptFile.Add("broker/roles")
            'Me.ScriptFile.Add("roles")
            'Me.ScriptFile.Add("drag")
        End Sub
    End Class
End Namespace