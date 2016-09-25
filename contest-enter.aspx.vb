Imports AMP.Global
Imports AMP.Site
Imports System.Text
Imports System.Configuration.ConfigurationSettings

Namespace Pages
    Public Class EnterContest
        Inherits AMP.Page

        Private _contest As AMP.Contest

#Region " Controls "

        Protected lblTitle As Label
        Protected lblDescription As Label
        Protected lblLimit As Label
        Protected fldRules As AMP.Controls.Field
        Protected rptRules As Repeater
        ' file upload
        Protected pnlFileUpload As Panel
        Protected ampUpload As AMP.Controls.Upload
        Protected fldTitle As AMP.Controls.Field

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

        Public Sub New()
            Me.RequireAuthentication = True
        End Sub

        Private Sub Page_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Init
            If Request.QueryString("id") <> Nothing Then
                _contest = WebSite.Contests(Request.QueryString("id"))
                If _contest Is Nothing Then Me.SendBack()
                ampUpload.AllowedTypes = _contest.FileType
                ampUpload.MaxFileSize = _contest.MaximumFileSize
            End If
        End Sub

        Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
            Me.StyleSheet.Add("contestEnter")
            Me.StyleSheet.Add("form")
            Me.ScriptFile.Add("contest")
            Me.ScriptFile.Add("entry")
            Me.ScriptFile.Add("validation/common")

            If Page.IsPostBack AndAlso Not ampUpload.File Is Nothing Then
                If _contest.CanEnter(Profile.User) Then
                    Dim email As New AMP.Email
                    Dim entry As New AMP.ContestEntry
                    With entry
                        .Contest = _contest
                        .File = ampUpload.File
                        .Contestant = Profile.User
                        .Date = DateTime.Now
                        .Title = fldTitle.Value
                        .Status = Status.Pending
                    End With
                    _contest.Entries.Add(entry)
                    WebSite.Save()
                    email.ContestEntry(entry)
                    Profile.Message = Me.Say("Msg.EnterContest")
                Else
                    Profile.Message = Me.Say("Error.EntryCount")
                End If
                Response.Redirect(String.Format("contest.aspx?id={0}", _contest.ID), True)
            End If

            fldRules.Label = "This entry complies with <a style=""cursor:pointer;"" href=""#"" onclick=""javascript:DOM.Show('rules');"">the contest rules</a>"
            rptRules.DataSource = _contest.Rules
            rptRules.DataBind()

            With _contest
                lblTitle.Text = .Title
                lblDescription.Text = Format.ToHtml(.Description)
                lblLimit.Text = String.Format("You have submitted {0} of {1} possible entries", _
                    Format.WordForNumber(.Entries.ByUser.Count), _
                    Format.WordForNumber(.EntriesAllowed))
                Me.Title = Me.Say("Action.EnterContest")

                If (.FileType And AMP.File.Types.Media) > 0 Then
                    ' show upload thingy
                    Dim count As Integer = 0
                    Dim note As New StringBuilder
                    Dim type As System.Type = GetType(AMP.File.Types)
                    Dim values As Integer() = CType([Enum].GetValues(type), Integer())
                    Dim names As String() = [Enum].GetNames(type)

                    Array.Sort(names, values)

                    With note
                        .Append("contest allows ")
                        For x As Integer = 0 To values.Length - 1
                            If (values(x) And _contest.FileType) > 0 AndAlso _
                                Common.IsFlag(values(x)) Then

                                If count > 0 Then .Append(", ")
                                .Append(names(x).ToLower)
                                count += 1
                            End If
                        Next
                        .Append(" files")
                        ampUpload.Note = .ToString
                    End With
                    pnlFileUpload.Visible = True
                End If
            End With
        End Sub
    End Class
End Namespace
