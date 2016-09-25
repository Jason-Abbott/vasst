<Serializable()> _
Public Class ContestEntryCollection
    Inherits CollectionBase

#Region " Properties "

    Public Property Item(ByVal index As Integer) As AMP.ContestEntry
        Get
            Return DirectCast(MyBase.InnerList(index), AMP.ContestEntry)
        End Get
        Set(ByVal Value As AMP.ContestEntry)
            MyBase.InnerList(index) = Value
        End Set
    End Property

    Default Public ReadOnly Property WithID(ByVal id As Guid) As AMP.ContestEntry
        Get
            SyncLock Me.InnerList
                For Each item As AMP.ContestEntry In Me.InnerList
                    If item.ID.Equals(id) Then
                        Return item
                        Exit For
                    End If
                Next
            End SyncLock
        End Get
    End Property

    Default Public ReadOnly Property WithID(ByVal id As String) As AMP.ContestEntry
        Get
            Return Me.WithID(New Guid(id))
        End Get
    End Property

#End Region

    '---COMMENT---------------------------------------------------------------
    '	get all unapproved entries
    '
    '	Date:		Name:	Description:
    '	2/27/05 	JEA		Creation
    '-------------------------------------------------------------------------
    Public Function Unapproved() As ArrayList
        Dim entries As New ArrayList
        For Each e As AMP.ContestEntry In MyBase.InnerList
            If e.Status = Site.Status.Pending Then entries.Add(e)
        Next
        Return entries
    End Function

    '---COMMENT---------------------------------------------------------------
    '	get all approved entries
    '
    '	Date:		Name:	Description:
    '	2/27/05 	JEA		Creation
    '-------------------------------------------------------------------------
    Public Function Approved() As ArrayList
        Dim entries As New ArrayList
        For Each e As AMP.ContestEntry In MyBase.InnerList
            If e.Status = Site.Status.Approved Then entries.Add(e)
        Next
        Return entries
    End Function

    '---COMMENT---------------------------------------------------------------
    '	get entries by given user
    '
    '	Date:		Name:	Description:
    '	2/22/05 	JEA		Creation
    '-------------------------------------------------------------------------
    Public Function ByUser(ByVal user As AMP.Person, ByVal status As Site.Status) As ArrayList
        Dim entries As New ArrayList
        For Each e As AMP.ContestEntry In MyBase.InnerList
            If user.ID.Equals(e.Contestant.ID) AndAlso _
                (status = Nothing OrElse e.Status = status) Then entries.Add(e)
        Next
        Return entries
    End Function

    Public Function ByUser() As ArrayList
        Return Me.ByUser(Profile.User, Nothing)
    End Function

    Public Function ByUser(ByVal user As AMP.Person) As ArrayList
        Return Me.ByUser(user, Nothing)
    End Function

    Public Function ByUser(ByVal status As Site.Status) As ArrayList
        Return Me.ByUser(Profile.User, status)
    End Function

    '---COMMENT---------------------------------------------------------------
    '	get approved entries not by given user
    '
    '	Date:		Name:	Description:
    '	2/22/05 	JEA		Creation
    '   3/1/05      JEA     Also must be approved
    '-------------------------------------------------------------------------
    Public Function NotByUser(ByVal user As AMP.Person) As ArrayList
        Dim entries As New ArrayList
        For Each e As AMP.ContestEntry In MyBase.InnerList
            If e.Status = Site.Status.Approved AndAlso _
                Not user.ID.Equals(e.Contestant.ID) Then entries.Add(e)
        Next
        Return entries
    End Function

    Public Function NotByUser() As ArrayList
        Return Me.NotByUser(Profile.User)
    End Function

    '---COMMENT---------------------------------------------------------------
    '	basic list methods
    '
    '	Date:		Name:	Description:
    '	11/30/04	JEA		Creation
    '-------------------------------------------------------------------------
    Public Function Add(ByVal entity As AMP.ContestEntry) As Integer
        Return MyBase.InnerList.Add(entity)
    End Function

    Public Sub Remove(ByVal entity As AMP.ContestEntry)
        MyBase.InnerList.Remove(entity)
    End Sub

    Public Sub Sort()
        MyBase.InnerList.Sort()
    End Sub

    Public Sub Sort(ByVal comparer As IComparer)
        MyBase.InnerList.Sort(comparer)
    End Sub
End Class