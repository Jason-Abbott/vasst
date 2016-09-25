Namespace Controls
    Public Class Ballot
        Inherits AMP.Controls.HtmlControl
        Implements IPostBackDataHandler

        Private _contestID As Guid
        Private _contest As AMP.Contest
        Private _entries As ArrayList
        Private _votes As Hashtable

#Region " Properties "

        '---COMMENT---------------------------------------------------------------
        '	store votes in entryID keyed hashtable on postback
        '
        '	Date:		Name:	Description:
        '	2/22/05 	JEA		Creation
        '-------------------------------------------------------------------------
        Public ReadOnly Property Votes() As Hashtable
            Get
                Return _votes
            End Get
        End Property

        Public WriteOnly Property ContestID() As Guid
            Set(ByVal Value As Guid)
                _contestID = Value
                _contest = WebSite.Contests(_contestID)
            End Set
        End Property

        Public WriteOnly Property Contest() As AMP.Contest
            Set(ByVal Value As AMP.Contest)
                _contest = Value
            End Set
        End Property

#End Region

        '---COMMENT---------------------------------------------------------------
        '	get ballot values
        '
        '	Date:		Name:	Description:
        '	2/22/05 	JEA		Creation
        '-------------------------------------------------------------------------
        Public Function LoadPostData(ByVal key As String, ByVal posted As System.Collections.Specialized.NameValueCollection) As Boolean Implements IPostBackDataHandler.LoadPostData
            Dim entryID As String
            Dim rank As Integer = 1
            Dim vote As AMP.ContestVote
            _votes = New Hashtable

            For x As Integer = 1 To CType(posted(key), Integer)
                entryID = posted(String.Format("{0}_{1}", key, x))

                If entryID.Length = 36 Then
                    vote = New AMP.ContestVote
                    vote.Person = Profile.User
                    vote.Rank = rank
                    _votes.Add(New Guid(entryID), vote)
                    rank += 1
                End If
            Next
            Return False
        End Function

        Public Sub RaisePostDataChangedEvent() Implements IPostBackDataHandler.RaisePostDataChangedEvent
            ' this is called if LoadPostData returns true
        End Sub

        '---COMMENT---------------------------------------------------------------
        '	render a contest ballot
        '
        '	Date:		Name:	Description:
        '	2/22/05 	JEA		Creation
        '-------------------------------------------------------------------------
        Protected Overrides Sub Render(ByVal writer As HtmlTextWriter)
            Dim votesPossible As Integer = _contest.VotesPossible
            _entries = _contest.Entries.NotByUser
            _entries.Sort(New AMP.Compare.ContestEntryTitle)

            With writer
                If votesPossible > 1 Then
                    .Write("<ol id=""")
                    .Write(Me.ClientID)
                    .Write(""" name=""")
                    .Write(Me.UniqueID)
                    .Write("""")
                    MyBase.RenderCss(writer)
                    .Write(">")
                    For x As Integer = 1 To votesPossible
                        .Write("<li>")
                        Me.RenderEntryList(writer, x)
                        .Write("</li>")
                    Next
                    .Write("</ol>")
                Else
                    Me.RenderEntryList(writer, 1)
                End If
                .Write("<input type=""hidden"" id=""")
                .Write(Me.ClientID)
                .Write(""" name=""")
                .Write(Me.UniqueID)
                .Write(""" value=""")
                .Write(votesPossible)
                .Write(""">")
            End With
        End Sub

        '---COMMENT---------------------------------------------------------------
        '	render individual select lists
        '
        '	Date:		Name:	Description:
        '	2/22/05 	JEA		Creation
        '-------------------------------------------------------------------------
        Private Sub RenderEntryList(ByVal writer As HtmlTextWriter, ByVal rank As Integer)
            Dim vote As AMP.ContestVote
            Dim user As AMP.Person = Profile.User

            With writer
                .Write("<select id=""")
                .Write(Me.ClientID)
                .Write("_")
                .Write(rank)
                .Write(""" name=""")
                .Write(Me.UniqueID)
                .Write("_")
                .Write(rank)
                .Write(""" onchange=""Contest.CheckBallot(this)"">")
                .Write("<option value=""0"" class=""choose"">no vote</option>")
                For Each e As AMP.ContestEntry In _entries
                    vote = e.Votes.ByUser(user)
                    .Write("<option value=""")
                    .Write(e.ID)
                    .Write("""")
                    If Not vote Is Nothing AndAlso vote.Rank = rank Then
                        .Write(" selected=""selected""")
                    End If
                    .Write(">")
                    .Write(e.Title)
                    .Write("</option>")
                Next
                .Write("</select>")
            End With
        End Sub
    End Class
End Namespace