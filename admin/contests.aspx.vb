Imports System.Text.RegularExpressions
Imports System.Configuration.ConfigurationSettings

Namespace Pages.Administration
    Public Class Contests
        Inherits AMP.AdminPage

#Region " Controls "

        Protected ddlContests As DropDownList
        Protected ampSectionList As AMP.Controls.EnumList
        Protected ampFileList As AMP.Controls.EnumList
        Protected topLabel As HtmlControls.HtmlContainerControl
        Protected fldTitle As AMP.Controls.Field
        Protected fldStartDate As AMP.Controls.Field
        Protected fldEndVoteDate As AMP.Controls.Field
        Protected fldStopDate As AMP.Controls.Field
        Protected fldWinners As AMP.Controls.Field
        Protected fldVotes As AMP.Controls.Field
        Protected fldVoteWeight As AMP.Controls.Field
        Protected tbDescription As TextBox
        Protected tbPrizes As TextBox
        ' restrictions
        Protected fldEntriesAllowed As AMP.Controls.Field
        Protected fldMaxFileSize As AMP.Controls.Field
        Protected fldFreePlugins As AMP.Controls.Field
        Protected fldGeneratedMediaOnly As AMP.Controls.Field
        Protected fldAnonymous As AMP.Controls.Field
        Protected tbRules As TextBox

#End Region

        Private Sub Page_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Init
            If Not Profile.User.HasPermission(AMP.Site.Permission.EditContest) Then
                Profile.Message = Me.Say("Error.Permissions")
                Me.SendBack()
            End If
        End Sub

        '---COMMENT---------------------------------------------------------------
        '	display and save contest
        '
        '	Date:		Name:	Description:
        '	2/15/05 	JEA		Creation
        '   2/24/05     JEA     Insert default values for new contest
        '   3/1/05      JEA     Handle rules
        '-------------------------------------------------------------------------
        Private Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
            Me.StyleSheet.Add("admin/contest")
            Me.ScriptFile.Add("contest")
            Me.ScriptFile.Add("validation/common")

            Dim contest As AMP.Contest

            If Request.QueryString("id") <> Nothing Then
                contest = WebSite.Contests(Request.QueryString("id"))
            End If

            ampSectionList.Type = GetType(AMP.Site.Section)
            ampFileList.Type = GetType(AMP.File.Types)
            ampFileList.Attributes.Add("onchange", "Contest.FileClick(this);")

            If Page.IsPostBack Then
                ' save contest
                Dim newContest As Boolean = False
                Const delimiter As String = "\n\s*-"

                If contest Is Nothing Then
                    contest = New AMP.Contest
                    newContest = True
                End If

                With contest
                    .Title = fldTitle.Value
                    .Description = tbDescription.Text
                    .Section = ampSectionList.Selected
                    .Prizes = Regex.Split(tbPrizes.Text, delimiter)
                    ' dates
                    .Start = Date.Parse(fldStartDate.Value)
                    .FinishVote = Date.Parse(fldEndVoteDate.Value)
                    .Finish = Date.Parse(fldStopDate.Value)
                    ' setup
                    .WinnerCount = CInt(fldWinners.Value)
                    .VotesAllowed = CInt(fldVotes.Value)
                    .WeightFactor = CInt(fldVoteWeight.Value)
                    .EntriesAllowed = CInt(fldEntriesAllowed.Value)
                    ' restrictions
                    .FreePluginsOnly = fldFreePlugins.Checked
                    .GeneratedMediaOnly = fldGeneratedMediaOnly.Checked
                    .FileType = ampFileList.Selected
                    .MaximumFileSize = CInt(fldMaxFileSize.Value)
                    .Anonymous = fldAnonymous.Checked
                    .Rules = Regex.Split(tbRules.Text, delimiter)
                End With

                If newContest Then WebSite.Contests.Add(contest)
                WebSite.Save()
                Response.Redirect(String.Format("../contest.aspx?id={0}", contest.ID), False)
                Return
            End If

            ' main contest selector
            WebSite.Contests.Sort(New Compare.ContestTitle)
            With ddlContests
                .Attributes.Add("onchange", "Contest.Change(this)")
                .DataSource = WebSite.Contests
                .DataBind()
                .Items.Insert(0, New ListItem("New Contest ...", "0"))
            End With

            If Not contest Is Nothing Then
                ' load contest
                With contest
                    ddlContests.SelectedValue = .ID.ToString
                    ampFileList.Selected = .FileType
                    ampSectionList.Selected = .Section
                    topLabel.InnerText = "Edit Contest"
                    fldTitle.Value = .Title
                    ' dates
                    fldStartDate.Value = String.Format("{0:d}", .Start)
                    fldEndVoteDate.Value = String.Format("{0:d}", .FinishVote)
                    fldStopDate.Value = String.Format("{0:d}", .Finish)
                    ' voting and winners
                    fldWinners.Value = .WinnerCount.ToString
                    fldVotes.Value = .VotesAllowed.ToString
                    fldVoteWeight.Value = .WeightFactor.ToString
                    tbDescription.Text = .Description
                    ' restrictions
                    If .MaximumFileSize = Nothing Then
                        fldMaxFileSize.Value = AppSettings("MaxFileUploadKB")
                    Else
                        fldMaxFileSize.Value = .MaximumFileSize.ToString
                    End If
                    fldEntriesAllowed.Value = .EntriesAllowed.ToString
                    fldFreePlugins.Checked = .FreePluginsOnly
                    fldGeneratedMediaOnly.Checked = .GeneratedMediaOnly
                    fldAnonymous.Checked = .Anonymous
                    tbRules.Text = "-" & .RuleList(Environment.NewLine & "-")
                    tbPrizes.Text = "-" & .PrizeList(Environment.NewLine & "-")
                End With
            Else
                ' set default values for contest
                fldMaxFileSize.Value = AppSettings("MaxFileUploadKB")
                ' dates
                fldStartDate.Value = String.Format("{0:d}", DateTime.Now)
                fldEndVoteDate.Value = String.Format("{0:d}", DateTime.Now.AddDays(21))
                fldStopDate.Value = String.Format("{0:d}", DateTime.Now.AddDays(35))
                ' voting and winners
                fldWinners.Value = "1"
                fldVotes.Value = "3"
                fldVoteWeight.Value = "1"
                ' entry restrictions
                fldEntriesAllowed.Value = "1"
                fldFreePlugins.Checked = True
                fldGeneratedMediaOnly.Checked = True
            End If
        End Sub
    End Class
End Namespace
