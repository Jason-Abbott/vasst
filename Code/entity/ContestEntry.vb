Imports System.Runtime.Serialization
Imports System.Configuration.ConfigurationSettings

<Serializable()> _
Public Class ContestEntry
    'Implements IDeserializationCallback

    Private _id As Guid
    Private _contest As AMP.Contest
    Private _asset As AMP.Asset
    Private _file As AMP.File
    Private _contestant As AMP.Person
    Private _title As String
    Private _date As DateTime = DateTime.Now
    Private _status As AMP.Site.Status
    Private _votes As New AMP.ContestVoteCollection

#Region " Properties "

    '---COMMENT---------------------------------------------------------------
    '	the contest this entry belongs to
    '
    '	Date:		Name:	Description:
    '	2/23/05 	JEA		Creation
    '-------------------------------------------------------------------------
    Public Property Contest() As AMP.Contest
        Get
            Return _contest
        End Get
        Set(ByVal Value As AMP.Contest)
            _contest = Value
        End Set
    End Property

    Public ReadOnly Property ID() As Guid
        Get
            Return _id
        End Get
    End Property

    Public Property Status() As AMP.Site.Status
        Get
            Return _status
        End Get
        Set(ByVal Value As AMP.Site.Status)
            _status = Value
        End Set
    End Property

    '---COMMENT---------------------------------------------------------------
    '	store entry details if no asset reference
    '
    '	Date:		Name:	Description:
    '	2/19/05 	JEA		Creation
    '-------------------------------------------------------------------------
    Public Property Contestant() As AMP.Person
        Get
            If _contestant Is Nothing AndAlso Not _asset Is Nothing Then
                _contestant = _asset.SubmittedBy
            End If
            Return _contestant
        End Get
        Set(ByVal Value As AMP.Person)
            _contestant = Value
        End Set
    End Property

    Public Property File() As AMP.File
        Get
            If _file Is Nothing AndAlso Not _asset Is Nothing Then
                _file = _asset.File
            End If
            Return _file
        End Get
        Set(ByVal Value As AMP.File)
            _file = Value
        End Set
    End Property

    Public Property Title() As String
        Get
            If _title = Nothing AndAlso Not _asset Is Nothing Then
                _title = _asset.Title
            End If
            Return _title
        End Get
        Set(ByVal Value As String)
            _title = Security.SafeString(Value, 75)
        End Set
    End Property

    Public Property Votes() As AMP.ContestVoteCollection
        Get
            Return _votes
        End Get
        Set(ByVal Value As AMP.ContestVoteCollection)
            _votes = Value
        End Set
    End Property

    Public Property [Date]() As DateTime
        Get
            Return _date
        End Get
        Set(ByVal Value As DateTime)
            _date = Value
        End Set
    End Property

    Public Property Asset() As AMP.Asset
        Get
            Return _asset
        End Get
        Set(ByVal Value As AMP.Asset)
            _asset = Value
        End Set
    End Property
#End Region

    Public Sub New()
        _id = Guid.NewGuid
    End Sub

    '---COMMENT---------------------------------------------------------------
    '	approve contest entry
    '
    '	Date:		Name:	Description:
    '	2/27/05 	JEA		Creation
    '-------------------------------------------------------------------------
    Public Function Approve() As Boolean
        Dim success As Boolean = False
        If Not _asset Is Nothing Then
            success = _asset.Approve
        ElseIf Not _file Is Nothing Then
            success = _file.Approve(AppSettings("ContestFolder"))
        End If
        If success Then
            _status = Status.Approved
            Return True
        Else
            Profile.Message = "Failed to approve"
            Return False
        End If
    End Function

    '---COMMENT---------------------------------------------------------------
    '	delete contest entry
    '
    '	Date:		Name:	Description:
    '	2/27/05 	JEA		Creation
    '-------------------------------------------------------------------------
    Public Sub Delete()
        If Not _asset Is Nothing Then
            _asset.Delete()
        ElseIf Not _file Is Nothing Then
            _file.Delete()
        End If
        _contest.Entries.Remove(Me)
    End Sub

    '---COMMENT---------------------------------------------------------------
    '	build url for viewing this entry
    '
    '	Date:		Name:	Description:
    '	2/23/05 	JEA		Creation
    '-------------------------------------------------------------------------
    Public Function ViewURL() As String
        If _asset Is Nothing Then
            Return String.Format("{0}/file.aspx?id={1}", Global.BasePath, Me.ID)
        Else
            Return IIf(_contest.Anonymous, _asset.ViewURL, _asset.DetailURL).ToString
        End If
    End Function

    '---COMMENT---------------------------------------------------------------
    '	build HTML link for viewing this entry
    '
    '	Date:		Name:	Description:
    '	2/23/05 	JEA		Creation
    '-------------------------------------------------------------------------
    Public Function ViewLink() As String
        Dim title As String = String.Empty
        If Not _contest.Anonymous Then title = String.Format("By {0}", _contestant.DisplayName)
        Return String.Format("<a href=""{0}"" title=""{1}"">{2}</a>", _
            Me.ViewURL, title, Me.Title)
    End Function

    'Public Sub OnDeserialization(ByVal sender As Object) Implements System.Runtime.Serialization.IDeserializationCallback.OnDeserialization
    '    _status = Site.Status.Approved
    'End Sub
End Class
