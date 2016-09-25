<Serializable()> _
Public Class ContestVoteCollection
    Inherits CollectionBase

#Region " Properties "

    Default Public Property Item(ByVal index As Integer) As AMP.ContestVote
        Get
            Return DirectCast(MyBase.InnerList(index), AMP.ContestVote)
        End Get
        Set(ByVal Value As AMP.ContestVote)
            MyBase.InnerList(index) = Value
        End Set
    End Property

#End Region

    '---COMMENT---------------------------------------------------------------
    '	get any vote by given user
    '
    '	Date:		Name:	Description:
    '	2/22/05 	JEA		Creation
    '-------------------------------------------------------------------------
    Public Function ByUser(ByVal user As AMP.Person) As AMP.ContestVote
        For Each v As AMP.ContestVote In MyBase.InnerList
            If user.ID.Equals(v.Person.ID) Then Return v
        Next
        Return Nothing
    End Function

    '---COMMENT---------------------------------------------------------------
    '	remove all votes by given user
    '
    '	Date:		Name:	Description:
    '	2/23/05 	JEA		Creation
    '-------------------------------------------------------------------------
    Public Function RemoveByUser(ByVal user As AMP.Person) As Boolean
        ' synclock issues?
        For Each v As AMP.ContestVote In MyBase.InnerList
            If user.ID.Equals(v.Person.ID) Then
                Me.Remove(v)
                Return True
            End If
        Next
        Return False
    End Function

    '---COMMENT---------------------------------------------------------------
    '	basic list methods
    '
    '	Date:		Name:	Description:
    '	11/30/04	JEA		Creation
    '-------------------------------------------------------------------------
    Public Function Add(ByVal entity As AMP.ContestVote) As Integer
        Return MyBase.InnerList.Add(entity)
    End Function

    Public Sub Remove(ByVal entity As AMP.ContestVote)
        MyBase.InnerList.Remove(entity)
    End Sub
End Class
