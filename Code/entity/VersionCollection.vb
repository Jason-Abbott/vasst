<Serializable()> _
Public Class VersionCollection
    Inherits CollectionBase

#Region " Properties "

    Public Property Item(ByVal index As Integer) As AMP.Version
        Get
            Return DirectCast(MyBase.InnerList(index), AMP.Version)
        End Get
        Set(ByVal Value As AMP.Version)
            MyBase.InnerList(index) = Value
        End Set
    End Property

    Default Public ReadOnly Property WithNumber(ByVal number As String) As AMP.Version
        Get
            number = number.ToLower
            For Each item As AMP.Version In Me.InnerList
                If item.Number = number Then
                    Return item
                    Exit For
                End If
            Next
            'BugOut("No match for version text ""{0}""", number)
            Return Nothing
        End Get
    End Property

#End Region

    '---COMMENT---------------------------------------------------------------
    '	basic list methods
    '
    '	Date:		Name:	Description:
    '	12/28/04	JEA		Creation
    '-------------------------------------------------------------------------
    Public Function Add(ByVal entity As AMP.Version) As Integer
        Return MyBase.InnerList.Add(entity)
    End Function

    Public Sub Remove(ByVal entity As AMP.Version)
        MyBase.InnerList.Remove(entity)
    End Sub

    Public Sub Remove(ByVal number As String)
        Me.Remove(Me.WithNumber(number))
    End Sub

    Public Sub Sort()
        Me.InnerList.Sort()
    End Sub

    '---COMMENT---------------------------------------------------------------
    '	get newest version; after sorting, should be last item in collection
    '
    '	Date:		Name:	Description:
    '	12/28/04	JEA		Creation
    '-------------------------------------------------------------------------
    Public Function Latest() As AMP.Version
        If Me.InnerList.Count > 0 Then
            Me.InnerList.Sort()
            Return DirectCast(Me.InnerList.Item(Me.InnerList.Count - 1), AMP.Version)
        Else
            Return Nothing
        End If
    End Function
End Class
