Imports System.Collections

<Serializable()> _
Public Class ContestCollection
    Inherits CollectionBase

#Region " Properties "

    Public Property Item(ByVal index As Integer) As AMP.Contest
        Get
            Return DirectCast(MyBase.InnerList(index), AMP.Contest)
        End Get
        Set(ByVal Value As AMP.Contest)
            MyBase.InnerList(index) = Value
        End Set
    End Property

    Default Public ReadOnly Property WithID(ByVal id As Guid) As AMP.Contest
        Get
            For Each contest As AMP.Contest In Me.InnerList
                If contest.ID.Equals(id) Then
                    Return contest
                    Exit For
                End If
            Next
        End Get
    End Property

    Default Public ReadOnly Property WithID(ByVal id As String) As AMP.Contest
        Get
            Return Me.WithID(New Guid(id))
        End Get
    End Property

#End Region

    '---COMMENT---------------------------------------------------------------
    '	get contest entry by ID
    '
    '	Date:		Name:	Description:
    '	2/23/05 	JEA		Creation
    '-------------------------------------------------------------------------
    Public Function EntryWithID(ByVal id As Guid) As AMP.ContestEntry
        For Each c As AMP.Contest In MyBase.InnerList
            For Each e As AMP.ContestEntry In c.Entries
                If e.ID.Equals(id) Then Return e
            Next
        Next
        Return Nothing
    End Function

    Public Function EntryWithID(ByVal id As String) As AMP.ContestEntry
        Return Me.EntryWithID(New Guid(id))
    End Function

    '---COMMENT---------------------------------------------------------------
    '	get all entries from given user
    '
    '	Date:		Name:	Description:
    '	3/8/05 	    JEA		Creation
    '-------------------------------------------------------------------------
    Public Function EntriesFromUser(ByVal user As AMP.Person) As ArrayList
        Dim matches As New ArrayList
        For Each c As AMP.Contest In MyBase.InnerList
            For Each e As AMP.ContestEntry In c.Entries
                If e.Contestant.ID.Equals(user.ID) Then matches.Add(e)
            Next
        Next
        Return matches
    End Function

    '---COMMENT---------------------------------------------------------------
    '	get contest with given name
    '
    '	Date:		Name:	Description:
    '	1/27/05 	JEA		Creation
    '-------------------------------------------------------------------------
    Public Function WithTitle(ByVal title As String) As AMP.Contest
        title = title.ToLower
        For Each c As AMP.Contest In MyBase.InnerList
            If c.Title.ToLower = title Then Return c
        Next
        Return Nothing
    End Function

    '---COMMENT---------------------------------------------------------------
    '	get active contests
    '
    '	Date:		Name:	Description:
    '	12/12/04	JEA		Creation
    '   12/29/04    JEA     Create overload to pass section
    '-------------------------------------------------------------------------
    Public Function Active(ByVal section As Integer) As ArrayList
        Dim result As New ArrayList

        If section = Nothing Then section = Profile.User.Section

        ' default sort is date descending
        MyBase.InnerList.Sort()

        For Each c As AMP.Contest In MyBase.InnerList
            If c.Active AndAlso (c.Section And section) > 0 Then
                result.Add(c)
            End If
        Next

        Return result
    End Function

    Public Function Active() As ArrayList
        Return Me.Active(Profile.User.Section)
    End Function

    '---COMMENT---------------------------------------------------------------
    '	basic list methods
    '
    '	Date:		Name:	Description:
    '	11/23/04	JEA		Creation
    '-------------------------------------------------------------------------
    Public Function Add(ByVal entity As AMP.Contest) As Integer
        Return MyBase.InnerList.Add(entity)
    End Function

    Public Sub Remove(ByVal entity As AMP.Contest)
        MyBase.InnerList.Remove(entity)
    End Sub

    Public Sub Sort()
        MyBase.InnerList.Sort()
    End Sub

    Public Sub Sort(ByVal comparer As IComparer)
        MyBase.InnerList.Sort(comparer)
    End Sub
End Class
