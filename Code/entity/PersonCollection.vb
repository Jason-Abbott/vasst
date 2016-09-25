Imports System.Collections

<Serializable()> _
Public Class PersonCollection
    Inherits CollectionBase

    <NonSerialized()> Private _matches As New Hashtable

#Region " Properties "

    Public Property Item(ByVal index As Integer) As AMP.Person
        Get
            Return DirectCast(MyBase.InnerList(index), AMP.Person)
        End Get
        Set(ByVal Value As AMP.Person)
            MyBase.InnerList(index) = Value
        End Set
    End Property

    Default Public ReadOnly Property WithID(ByVal id As Guid) As AMP.Person
        Get
            For Each p As AMP.Person In Me.InnerList
                If p.ID.Equals(id) Then
                    Return p
                    Exit For
                End If
            Next
            BugOut("No person found for {0}", id.ToString)
        End Get
    End Property

    Default Public ReadOnly Property WithID(ByVal id As String) As AMP.Person
        Get
            Return Me.WithID(New Guid(id))
        End Get
    End Property

#End Region

    '---COMMENT---------------------------------------------------------------
    '	get users who haven't logged in since a given date
    '
    '	Date:		Name:	Description:
    '	3/8/05 	    JEA		Created
    '-------------------------------------------------------------------------
    Public Sub Archive(ByVal olderThan As DateTime)
        Dim candidates As ArrayList = Me.NoLoginSince(olderThan)
        For Each p As AMP.Person In candidates
            ' TODO: copy user data to external database before removing
            'if Not p.Dependencies then
            ' remove contest entries, contest votes, comments and ratings, assets
        Next
    End Sub

    '---COMMENT---------------------------------------------------------------
    '	get users who haven't logged in since a given date
    '
    '	Date:		Name:	Description:
    '	3/8/05 	    JEA		Created
    '-------------------------------------------------------------------------
    Public Function NoLoginSince(ByVal threshold As DateTime) As ArrayList
        Dim old As New ArrayList
        For Each p As AMP.Person In Me.InnerList
            If p.LastLogin < threshold Then old.Add(p)
        Next
        Return old
    End Function

    '---COMMENT---------------------------------------------------------------
    '	check if address is already used
    '
    '	Date:		Name:	Description:
    '	2/28/05 	JEA		Creation
    '-------------------------------------------------------------------------
    Public Function EmailUsed(ByVal email As String) As Boolean
        For Each p As AMP.Person In Me.InnerList
            If p.Email.ToLower = email.ToLower Then Return True
        Next
        Return False
    End Function

    '---COMMENT---------------------------------------------------------------
    '	get all persons with given permission
    '
    '	Date:		Name:	Description:
    '	3/1/05  	JEA		Creation
    '-------------------------------------------------------------------------
    Public Function WithPermission(ByVal permission As AMP.Site.Permission) As ArrayList
        Dim persons As New ArrayList
        For Each p As AMP.Person In Me.InnerList
            If p.HasPermission(permission) Then persons.Add(p)
        Next
        Return persons
    End Function

    '---COMMENT---------------------------------------------------------------
    '	get person with e-mail address
    '
    '	Date:		Name:	Description:
    '	11/23/04	JEA		Creation
    '-------------------------------------------------------------------------
    Public Function WithEmail(ByVal email As String) As AMP.Person
        For Each item As AMP.Person In Me.InnerList
            If item.Email.ToLower = email.ToLower Then
                Return item
                Exit For
            End If
        Next
        Return Nothing
    End Function

    '---COMMENT---------------------------------------------------------------
    '	basic list methods
    '
    '	Date:		Name:	Description:
    '	11/23/04	JEA		Creation
    '-------------------------------------------------------------------------
    Public Function Add(ByVal entity As AMP.Person) As Integer
        Return MyBase.InnerList.Add(entity)
    End Function

    Public Sub Remove(ByVal entity As AMP.Person)
        MyBase.InnerList.Remove(entity)
    End Sub

    Public Sub Remove(ByVal id As Guid)
        Me.Remove(Me.WithID(id))
    End Sub

End Class
