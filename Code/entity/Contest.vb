Imports AMP.Data
Imports AMP.Site
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Runtime.Serialization
Imports System.Configuration.ConfigurationSettings

<Serializable()> _
Public Class Contest
    Implements IComparable
    'Implements IDeserializationCallback

    Private _id As Guid
    Private _title As String
    Private _start As DateTime
    Private _finish As DateTime
    Private _finishVote As DateTime
    Private _rules As String() = {}
    Private _prizes As String() = {}
    Private _description As String
    Private _votesAllowed As Integer
    Private _entriesAllowed As Integer = 1
    Private _winnerCount As Integer
    Private _weightFactor As Single
    Private _section As Integer = 0                     ' bitmask
    Private _fileType As Integer = AMP.File.Types.Any   ' bitmask
    Private _anonymousEntryAuthors As Boolean = False
    Private _freePluginsOnly As Boolean = True
    Private _maxFileSize As Integer
    Private _generatedMediaOnly As Boolean = True
    Private _entries As New AMP.ContestEntryCollection

#Region " Properties "

    Public Property Prizes() As String()
        Get
            Return _prizes
        End Get
        Set(ByVal Value As String())
            _prizes = New String() {}
            For x As Integer = 0 To Value.Length - 1
                Me.AddPrize(Value(x))
            Next
        End Set
    End Property

    '---COMMENT---------------------------------------------------------------
    '	array of rules strings; use .AddRule to clean
    '
    '	Date:		Name:	Description:
    '	3/1/05 	    JEA		Creation
    '-------------------------------------------------------------------------
    Public Property Rules() As String()
        Get
            Return _rules
        End Get
        Set(ByVal Value As String())
            _rules = New String() {}
            For x As Integer = 0 To Value.Length - 1
                Me.AddRule(Value(x))
            Next
        End Set
    End Property

    '---COMMENT---------------------------------------------------------------
    '	option to override default file size limit
    '
    '	Date:		Name:	Description:
    '	2/20/05 	JEA		Creation
    '-------------------------------------------------------------------------
    Public Property MaximumFileSize() As Integer
        Get
            Return _maxFileSize
        End Get
        Set(ByVal Value As Integer)
            _maxFileSize = Value
        End Set
    End Property

    '---COMMENT---------------------------------------------------------------
    '	if true, only allows voters to see file link, not asset/author detail
    '
    '	Date:		Name:	Description:
    '	2/18/05 	JEA		Creation
    '-------------------------------------------------------------------------
    Public Property Anonymous() As Boolean
        Get
            Return _anonymousEntryAuthors
        End Get
        Set(ByVal Value As Boolean)
            _anonymousEntryAuthors = Value
        End Set
    End Property

    Public Property FileType() As Integer
        Get
            Return _fileType
        End Get
        Set(ByVal Value As Integer)
            _fileType = Value
        End Set
    End Property

    Public Property GeneratedMediaOnly() As Boolean
        Get
            Return _generatedMediaOnly
        End Get
        Set(ByVal Value As Boolean)
            _generatedMediaOnly = Value
        End Set
    End Property

    Public Property EntriesAllowed() As Integer
        Get
            Return _entriesAllowed
        End Get
        Set(ByVal Value As Integer)
            _entriesAllowed = Value
        End Set
    End Property

    Public ReadOnly Property Active() As Boolean
        Get
            Return Me.Start <= Date.Now AndAlso Me.Finish >= Date.Now
        End Get
    End Property

    ' collection of assets entered in the contest
    Public Property Entries() As AMP.ContestEntryCollection
        Get
            Return _entries
        End Get
        Set(ByVal Value As AMP.ContestEntryCollection)
            _entries = Value
        End Set
    End Property

    ' allow only assets with free plugins
    Public Property FreePluginsOnly() As Boolean
        Get
            Return _freePluginsOnly
        End Get
        Set(ByVal Value As Boolean)
            _freePluginsOnly = Value
        End Set
    End Property

    Public Property Section() As Integer
        Get
            Return _section
        End Get
        Set(ByVal Value As Integer)
            _section = Value
        End Set
    End Property

    ' when multiple votes are allowed, how are they weighted
    ' The last choice always receives one point with successively higher
    ' choices differentiated by the weight factor.
    Public Property WeightFactor() As Single
        Get
            Return _weightFactor
        End Get
        Set(ByVal Value As Single)
            _weightFactor = Value
        End Set
    End Property

    ' how many entries can a single person vote for
    Public Property VotesAllowed() As Integer
        Get
            Return _votesAllowed
        End Get
        Set(ByVal Value As Integer)
            _votesAllowed = Value
        End Set
    End Property

    ' number of winners that will be picked
    Public Property WinnerCount() As Integer
        Get
            Return _winnerCount
        End Get
        Set(ByVal Value As Integer)
            _winnerCount = Value
        End Set
    End Property

    Public Property Description() As String
        Get
            Return _description
        End Get
        Set(ByVal Value As String)
            _description = Security.SafeString(Value, 1500)
        End Set
    End Property

    ' date on which voting finishes
    Public Property FinishVote() As DateTime
        Get
            Return _finishVote
        End Get
        Set(ByVal Value As DateTime)
            _finishVote = Value
        End Set
    End Property

    ' date on which contest disappears from site
    Public Property Finish() As DateTime
        Get
            Return _finish
        End Get
        Set(ByVal Value As DateTime)
            _finish = Value
        End Set
    End Property

    Public Property Start() As DateTime
        Get
            Return _start
        End Get
        Set(ByVal Value As DateTime)
            _start = Value
        End Set
    End Property

    Public Property Title() As String
        Get
            Return _title
        End Get
        Set(ByVal Value As String)
            _title = Security.SafeString(Value, 75)
        End Set
    End Property

    Public ReadOnly Property ID() As Guid
        Get
            Return _id
        End Get
    End Property

#End Region

    Public Sub New()
        _id = Guid.NewGuid
    End Sub


#Region " Prizes "

    '---COMMENT---------------------------------------------------------------
    '	add prize string
    '
    '	Date:		Name:	Description:
    '	3/5/05 	    JEA		Creation
    '-------------------------------------------------------------------------
    Public Sub AddPrize(ByVal prize As String)
        prize = prize.Trim.Replace(vbCrLf, "")
        prize = Regex.Replace(prize, "^\s*-\s*", "")
        If prize <> "" Then
            Dim index As Integer = _prizes.Length
            ReDim Preserve _prizes(index)
            _prizes(index) = Security.SafeString(prize, 150)
        End If
    End Sub

    '---COMMENT---------------------------------------------------------------
    '	return prizes as string list
    '
    '	Date:		Name:	Description:
    '	3/5/05   	JEA		Creation
    '-------------------------------------------------------------------------
    Public Function PrizeList(ByVal delimiter As String) As String
        Dim prizes As New StringBuilder
        For x As Integer = 0 To _prizes.Length - 1
            If x > 0 Then prizes.Append(delimiter)
            prizes.Append(_prizes(x))
        Next
        Return prizes.ToString
    End Function

#End Region

#Region " Rules "

    '---COMMENT---------------------------------------------------------------
    '	add rule string
    '
    '	Date:		Name:	Description:
    '	3/1/05 	    JEA		Creation
    '-------------------------------------------------------------------------
    Public Sub AddRule(ByVal rule As String)
        rule = rule.Trim.Replace(vbCrLf, "")
        rule = Regex.Replace(rule, "^\s*-\s*", "")
        If rule <> "" Then
            Dim index As Integer = _rules.Length
            ReDim Preserve _rules(index)
            _rules(index) = Security.SafeString(rule, 100)
        End If
    End Sub

    '---COMMENT---------------------------------------------------------------
    '	return rules as string list
    '
    '	Date:		Name:	Description:
    '	3/1/05   	JEA		Creation
    '-------------------------------------------------------------------------
    Public Function RuleList(ByVal delimiter As String) As String
        Dim rules As New StringBuilder
        For x As Integer = 0 To _rules.Length - 1
            If x > 0 Then rules.Append(delimiter)
            rules.Append(_rules(x))
        Next
        Return rules.ToString
    End Function

#End Region

    '---COMMENT---------------------------------------------------------------
    '	remove all votes by given user
    '
    '	Date:		Name:	Description:
    '	2/23/05 	JEA		Creation
    '-------------------------------------------------------------------------
    Public Function RemoveUserVotes(ByVal user As AMP.Person) As Boolean
        For Each e As AMP.ContestEntry In _entries
            e.Votes.RemoveByUser(user)
        Next
    End Function

    '---COMMENT---------------------------------------------------------------
    '	determine if user can enter
    '   if not authenticated then assume true so they can at least click link
    '
    '	Date:		Name:	Description:
    '	2/22/05 	JEA		Creation
    '-------------------------------------------------------------------------
    Public Function CanEnter(ByVal user As AMP.Person) As Boolean
        Return (DateTime.Now < _finishVote AndAlso _entries.ByUser(user).Count < _entriesAllowed)
    End Function

    Public Function CanEnter() As Boolean
        If Not Profile.Authenticated Then Return True
        Return Me.CanEnter(Profile.User)
    End Function

    '---COMMENT---------------------------------------------------------------
    '	limit allowed votes to number of entries
    '
    '	Date:		Name:	Description:
    '	2/19/05 	JEA		Creation
    '-------------------------------------------------------------------------
    Public Function VotesPossible() As Integer
        Dim possible As Integer
        Dim approved As Integer = _entries.Approved.Count
        If approved < _votesAllowed Then
            possible = approved
        Else
            possible = _votesAllowed
        End If
        ' user can't vote for own entries so subtract those
        Return possible - _entries.ByUser(Status.Approved).Count
    End Function

    '---COMMENT---------------------------------------------------------------
    '	return entries with highest votes
    '
    '	Date:		Name:	Description:
    '	12/15/04	JEA		Creation
    '-------------------------------------------------------------------------
    Public Function TopEntries() As ArrayList
        Dim entries As New ArrayList
        Dim comparer As New AMP.Compare.ContestEntryVotes(Me.VotesAllowed, Me.WeightFactor)
        Dim count As Integer = Me.WinnerCount
        Dim x As Integer = 0

        _entries.Sort(comparer)
        If count < _entries.Count Then count = _entries.Count

        For Each entry As AMP.ContestEntry In _entries
            If entry.Status = Site.Status.Approved AndAlso entry.Votes.Count > 0 Then
                entries.Add(entry)
                x += 1
            End If
            If x = count Then Exit For
        Next

        Return entries
    End Function

    '---COMMENT---------------------------------------------------------------
    '	return all entries sorted by rank
    '
    '	Date:		Name:	Description:
    '	2/23/05 	JEA		Creation
    '-------------------------------------------------------------------------
    Public Function RankedEntries() As ArrayList
        Dim entries As New ArrayList
        Dim comparer As New AMP.Compare.ContestEntryVotes(Me.VotesAllowed, Me.WeightFactor)

        For Each e As AMP.ContestEntry In _entries
            If e.Status = Site.Status.Approved Then entries.Add(e)
        Next

        entries.Sort(comparer)
        Return entries
    End Function

#Region " Links "

    Public Function DetailURL() As String
        Return String.Format("{0}/contest.aspx?id={1}", _
            Global.BasePath, Me.ID, HttpUtility.HtmlEncode(Me.Title))
    End Function

    Public Function DetailLink() As String
        Return String.Format("<a href=""{0}"">{1}</a>", _
            Me.DetailURL, HttpUtility.HtmlEncode(Me.Title))
    End Function

    Public Function FullDetailUrl() As String
        Return String.Format("http://{0}{1}", _
            HttpContext.Current.Request.ServerVariables("SERVER_NAME"), Me.DetailURL)
    End Function

#End Region

#Region " Percentages "

    '---COMMENT---------------------------------------------------------------
    '	get entry points as percentage of total contest entry points
    '
    '	Date:		Name:	Description:
    '	12/17/04	JEA		Creation
    '   2/27/05     JEA     Handle 0 points
    '-------------------------------------------------------------------------
    Public Function VotePercent(ByVal votes As AMP.ContestVoteCollection, ByVal places As Integer) As Single
        Dim points As Integer = Me.Points(votes)
        If points = 0 Then Return 0
        Dim total As Integer = Me.PointTotal
        If total = 0 Then Return 0
        Return CSng(Math.Round((points / total) * 100, places))
    End Function

    Public Function VotePercent(ByVal votes As AMP.ContestVoteCollection) As Single
        Return Me.VotePercent(votes, 0)
    End Function

    '---COMMENT---------------------------------------------------------------
    '	get value of vote points as percentage of entry with highest points
    '
    '	Date:		Name:	Description:
    '	12/17/04	JEA		Creation
    '   2/27/05     JEA     Handle 0 max points
    '-------------------------------------------------------------------------
    Public Function PercentOfHighest(ByVal votes As AMP.ContestVoteCollection) As Integer
        Dim maxPoints As Integer = 0
        Dim entryPoints As Integer = 0
        Dim points As Integer = Me.Points(votes)

        If points = 0 Then Return 0

        For Each entry As AMP.ContestEntry In _entries
            entryPoints = Me.Points(entry.Votes)
            If entryPoints > maxPoints Then maxPoints = entryPoints
        Next

        If maxPoints = 0 Then Return 0
        Return CInt(Math.Round((points / maxPoints) * 100, 0))
    End Function

#End Region

#Region " Points "

    Public Function PointTotal() As Integer
        Dim totalPoints As Integer = 0
        For Each entry As AMP.ContestEntry In _entries
            For Each vote As AMP.ContestVote In entry.Votes
                totalPoints += Me.Points(vote.Rank)
            Next
        Next
        Return totalPoints
    End Function

    '---COMMENT---------------------------------------------------------------
    '	compute points for given votes
    '   if a contest allows multiple votes and a weight factor other than 1
    '   is specified, then the voter's last choice receives one point with
    '   successively higher choices differentiated by the factor
    '
    '	Date:		Name:	Description:
    '	12/17/04	JEA		Creation
    '-------------------------------------------------------------------------
    Public Function Points(ByVal rank As Integer) As Integer
        Return Points(rank, Me.WeightFactor, Me.VotesAllowed)
    End Function

    Public Shared Function Points(ByVal rank As Integer, ByVal weightFactor As Single, _
        ByVal votesAllowed As Integer) As Integer

        Return CInt((weightFactor * (votesAllowed - rank)) + 1)
    End Function

    Public Function Points(ByVal votes As AMP.ContestVoteCollection) As Integer
        Dim pointTotal As Integer = 0
        For Each vote As AMP.ContestVote In votes
            pointTotal += Me.Points(vote.Rank)
        Next
        Return pointTotal
    End Function

#End Region

    '---COMMENT---------------------------------------------------------------
    '	default compare sorts from newest to oldest
    '
    '	Date:		Name:	Description:
    '	11/30/04	JEA		Creation
    '-------------------------------------------------------------------------
    Public Function CompareTo(ByVal entity As Object) As Integer Implements System.IComparable.CompareTo
        Return Date.Compare(Me.Start, DirectCast(entity, AMP.Contest).Start)
    End Function

    'Public Sub OnDeserialization(ByVal sender As Object) Implements System.Runtime.Serialization.IDeserializationCallback.OnDeserialization
    '    ' TODO: remove this
    '    If _prizes Is Nothing Then _prizes = New String() {}
    'End Sub
End Class
