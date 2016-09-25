Namespace Pages
    Public Class Unsubscribe
        Inherits AMP.Page

        Protected tbEmail As TextBox

        '---COMMENT---------------------------------------------------------------
        '   unsubscribe user from mailing list and cancel membership if file sharer
        ' 
        '   Date:       Name:   Description:
        '	2/17/05     JEA     Created
        '-------------------------------------------------------------------------
        Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
            Me.Title = "Unsubscribe"
            Me.StyleSheet.Add("form")
            Me.StyleSheet.Add("unsubscribe")

            If Page.IsPostBack Then
                Dim legacy As New AMP.Data.Legacy
                If legacy.DisableEmail(tbEmail.Text) Then
                    ' also disable account in this db
                    Dim user As AMP.Person = WebSite.Persons.WithEmail(tbEmail.Text)
                    If Not user Is Nothing Then
                        user.Status = AMP.Site.Status.Disabled
                        Log.Activity(AMP.Site.Activity.Unsubscribed, user.ID)
                        WebSite.Save()
                        Profile.Clear()
                        Session.Abandon()
                    Else
                        Log.Activity(AMP.Site.Activity.Unsubscribed, tbEmail.Text)
                    End If
                    Profile.Message = String.Format("{0} has been unsubscribed", tbEmail.Text)
                    Response.Redirect("~/", True)
                Else
                    Profile.Message = String.Format("{0} could not be found", tbEmail.Text)
                End If
            Else
                If Profile.Authenticated Then tbEmail.Text = Profile.User.Email
            End If
        End Sub

    End Class
End Namespace