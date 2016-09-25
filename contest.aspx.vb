Imports AMP.Global
Imports AMP.Site
Imports System.Text
Imports System.Collections.Specialized

Namespace Pages
    Public Class Contest
        Inherits AMP.Page

        Private _contest As AMP.Contest

#Region " Controls "

        Protected lblStatus As Label
        Protected lblTitle As Label
        Protected lblAbout As Label
        Protected lblInstructions As Label
        Protected lblDescription As Label
        Protected btnEdit As AMP.Controls.Button
        Protected rptVotes As Repeater
        Protected rptRules As Repeater
        Protected rptPrizes As Repeater
        ' ballot
        Protected fsBallot As HtmlControls.HtmlContainerControl
        Protected pnlCanVote As Panel
        Protected lblNoVote As Label
        Protected ampBallot As AMP.Controls.Ballot
        ' buttons
        Protected btnEnter As AMP.Controls.Button
        Protected btnVote As AMP.Controls.Button

#End Region

#Region " Properties "

        Protected Property Entity() As AMP.Contest
            Get
                Return _contest
            End Get
            Set(ByVal Value As AMP.Contest)
                _contest = Value
            End Set
        End Property

#End Region

        Private Sub Page_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Init
            If Request.QueryString("id") <> Nothing Then
                _contest = WebSite.Contests(Request.QueryString("id"))
                If _contest Is Nothing OrElse _
                    Not (_contest.Active OrElse _
                    Profile.User.HasPermission(Permission.EditContest)) Then

                    Me.SendBack()
                End If

                ' check for download to resume
                Dim id As Guid = Profile.ResumeDownload
                If Not id.Equals(Guid.Empty) Then
                    Dim entry As AMP.ContestEntry = _contest.Entries(id)
                    Dim asset As AMP.Asset = WebSite.Assets(id)
                    Dim url As String

                    If Not entry Is Nothing Then
                        url = entry.ViewURL
                    ElseIf Not asset Is Nothing Then
                        url = asset.ViewURL
                    End If

                    MyBase.ScriptBlock = Common.JSRedirect(url)
                    Profile.ResumeDownload = Nothing
                End If
            End If
            ampBallot.Contest = _contest
        End Sub

        Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
            Me.StyleSheet.Add("contest")
            Me.StyleSheet.Add("form")
            Me.ScriptFile.Add("contest")
            Me.ScriptFile.Add("validation/common")

            Dim user As AMP.Person = Profile.User
            Dim approved As Integer

            If Page.IsPostBack Then
                ' store form in session to handle potential login redirection
                Profile.FormValues = Request.Form
                If Not Profile.Authenticated Then
                    Me.SendToLogin()
                Else
                    Me.SaveVote(_contest, user)
                End If
            Else
                ' check for session form values
                Dim form As NameValueCollection = Profile.FormValues
                If (Not form Is Nothing) AndAlso _
                    Array.IndexOf(form.AllKeys, ampBallot.UniqueID) >= 0 Then
                    ' manually update ballot with form data
                    ampBallot.LoadPostData(ampBallot.UniqueID, form)
                    Me.SaveVote(_contest, user)
                End If
            End If

            ' load contest details
            Dim _votesPossible As Integer

            With _contest
                lblTitle.Text = .Title
                lblDescription.Text = Format.ToHtml(.Description)
                Me.Title = .Title
                _votesPossible = .VotesPossible
                approved = .Entries.Approved.Count
            End With

            If user.HasPermission(Permission.EditContest) Then
                btnEdit.Visible = True
                btnEdit.Url = String.Format("./admin/contests.aspx?id={0}", _contest.ID)
            End If

            If _contest.CanEnter Then
                btnEnter.Visible = True
                btnEnter.Url = String.Format("contest-enter.aspx?id={0}", _contest.ID)
            End If

            If _contest.Active Then
                ' display rules and such
                Dim about As New StringBuilder
                With about
                    .Append(Format.WordForNumber(_contest.WinnerCount, True))
                    .Append(" winner")
                    If _contest.WinnerCount > 1 Then .Append("s")
                    .Append(" will be selected by your votes.  Vote")
                    If _contest.VotesAllowed > 1 Then
                        .Append(" for up to ")
                        .Append(Format.WordForNumber(_contest.VotesAllowed))
                    End If
                    .Append(" or ")
                    If _contest.EntriesAllowed > 1 Then
                        .Append("submit up to ")
                        .Append(Format.WordForNumber(_contest.EntriesAllowed))
                        .Append(" entries")
                    Else
                        .Append("enter")
                    End If
                    .Append(" by ")
                    .Append(String.Format("{0:MMMM d.}", _contest.FinishVote))
                    lblAbout.Text = .ToString
                End With
                lblAbout.Visible = True

                If approved > 0 Then
                    lblStatus.Text = "Contestants"
                Else
                    lblStatus.Visible = False
                End If

                If _contest.Prizes.Length > 0 Then
                    rptPrizes.DataSource = _contest.Prizes
                    rptPrizes.DataBind()
                    rptPrizes.Visible = True
                End If

                If _votesPossible > 0 Then
                    ' show ballot
                    Dim text As New StringBuilder
                    fsBallot.Visible = True

                    If approved < 2 Then
                        ' prevent vote for just one entry
                        lblNoVote.Visible = True
                        pnlCanVote.Visible = False
                    Else
                        ' show vote options
                        With text
                            ' instructions
                            .Append("You may change your ")
                            If _votesPossible > 1 Then
                                .Append(Format.WordForNumber(_votesPossible))
                                .Append(" votes anytime.<br/>The order")
                                If _contest.WeightFactor > 1 Then
                                    .Append(" matters.")
                                Else
                                    .Append(" doesn't matter.")
                                End If
                            Else
                                .Append("vote anytime")
                            End If
                            lblInstructions.Text = .ToString
                            .Length = 0
                        End With
                    End If
                Else
                    ' show rules in place of ballot
                    rptRules.DataSource = _contest.Rules
                    rptRules.DataBind()
                    rptRules.Visible = True
                End If
                rptVotes.DataSource = _contest.RankedEntries
            Else
                ' display conclusion (winners)
                lblStatus.Text = "Final Results"
                rptVotes.DataSource = _contest.TopEntries
            End If
            rptVotes.DataBind()
        End Sub

        '---COMMENT---------------------------------------------------------------
        '	save contest vote
        '
        '	Date:		Name:	Description:
        '	2/23/05     JEA		Creation
        '-------------------------------------------------------------------------
        Private Sub SaveVote(ByVal contest As AMP.Contest, ByVal user As AMP.Person)
            Dim entry As AMP.ContestEntry
            Dim selfVote As Boolean = False
            ' remove old votes
            contest.RemoveUserVotes(user)
            ' add new votes
            For Each id As Guid In ampBallot.Votes.Keys
                entry = contest.Entries(id)
                If Not entry Is Nothing Then
                    If entry.Contestant.ID.Equals(Profile.PersonID) Then
                        selfVote = True
                    Else
                        entry.Votes.Add(DirectCast(ampBallot.Votes.Item(id), AMP.ContestVote))
                    End If
                End If
            Next
            Profile.FormValues = Nothing
            Profile.Message = Me.Say(IIf(selfVote, "Error.SelfVote", "Msg.AfterVoting").ToString)
            Log.Activity(Activity.Vote)
            WebSite.Save()
        End Sub
    End Class
End Namespace