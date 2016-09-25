Imports AMP.Site
Imports System.Configuration.ConfigurationSettings

Namespace Pages.Administration
    Public Class ContestEntries
        Inherits AMP.AdminPage

#Region " Controls "

        Protected rptContests As Repeater
        Protected fldAction As HtmlInputHidden
        Protected fldEntryID As HtmlInputHidden

#End Region

        Private Sub Page_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Init
            If Not Profile.User.HasPermission(Permission.ApproveEntry) Then
                Profile.Message = Me.Say("Error.Permissions")
                Me.SendBack()
            End If
        End Sub

        '---COMMENT---------------------------------------------------------------
        '	approve or deny contest entries
        '
        '	Date:		Name:	Description:
        '	2/27/05 	JEA		Creation
        '-------------------------------------------------------------------------
        Private Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
            Me.StyleSheet.Add("admin/contest")
            Me.ScriptFile.Add("entry")
            Me.ScriptFile.Add("validation/common")

            If Page.IsPostBack Then
                Dim entry As AMP.ContestEntry = WebSite.Contests.EntryWithID(fldEntryID.Value)
                If Not entry Is Nothing Then
                    Dim email As New AMP.Email

                    Select Case fldAction.Value
                        Case "approve"
                            If entry.Approve() Then
                                Log.Activity(Activity.ApproveContestEntry, Profile.PersonID)
                                email.EntryApproval(entry)
                                Profile.Message = String.Format("""{0}"" has been approved", entry.Title)
                                WebSite.Save()
                            End If
                        Case "deny"
                            Profile.Message = String.Format("""{0}"" has been deleted", entry.Title)
                            entry.Delete()
                            Log.Activity(Activity.DenyContestEntry, Profile.PersonID)
                            email.EntryDenial(entry)
                            WebSite.Save()
                    End Select
                End If
            End If

            rptContests.DataSource = WebSite.Contests.Active
            rptContests.DataBind()
        End Sub

#Region " Build Links "

        Protected Function ApproveLink(ByVal data As Object) As String
            Dim entry As AMP.ContestEntry = DirectCast(data, AMP.ContestEntry)
            Return String.Format("Entry.Approve('{0}')", entry.ID)
        End Function

        Protected Function DenyLink(ByVal data As Object) As String
            Dim entry As AMP.ContestEntry = DirectCast(data, AMP.ContestEntry)
            Return String.Format("Entry.Deny('{0}')", entry.ID)
        End Function

#End Region

    End Class
End Namespace
