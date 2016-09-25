<Serializable()> _
Public Class PublisherCollection
    Inherits CollectionBase

#Region " Properties "

    Public Property Item(ByVal index As Integer) As AMP.Publisher
        Get
            Return DirectCast(MyBase.InnerList(index), AMP.Publisher)
        End Get
        Set(ByVal Value As AMP.Publisher)
            MyBase.InnerList(index) = Value
        End Set
    End Property

    Default Public ReadOnly Property WithID(ByVal id As Guid) As AMP.Publisher
        Get
            For Each p As AMP.Publisher In Me.InnerList
                If p.ID.Equals(id) Then Return p
            Next
            Return Nothing
        End Get
    End Property

    Default Public ReadOnly Property WithID(ByVal id As String) As AMP.Publisher
        Get
            Return Me.WithID(New Guid(id))
        End Get
    End Property

#End Region

    Public Function SoftwareWithID(ByVal id As Guid) As AMP.Software
        For Each p As AMP.Publisher In Me.InnerList
            For Each s As AMP.Software In p.Software
                If s.ID.Equals(id) Then Return s
            Next
        Next
        Return Nothing
    End Function

    Public Function SoftwareWithID(ByVal id As String) As AMP.Software
        Return Me.SoftwareWithID(New Guid(id))
    End Function

    '---COMMENT---------------------------------------------------------------
    '	find publisher with given name
    '
    '	Date:		Name:	Description:
    '	1/27/05	    JEA		Creation
    '-------------------------------------------------------------------------
    Public Function WithName(ByVal name As String) As AMP.Publisher
        name = name.ToLower
        For Each p As AMP.Publisher In Me.InnerList
            If p.Name.ToLower = name Then Return p
        Next
        Return Nothing
    End Function

    '---COMMENT---------------------------------------------------------------
    '	find software with given name
    '
    '	Date:		Name:	Description:
    '	1/27/05	    JEA		Creation
    '-------------------------------------------------------------------------
    Public Function SoftwareWithName(ByVal name As String) As AMP.Software
        name = name.ToLower
        For Each p As AMP.Publisher In Me.InnerList
            For Each s As AMP.Software In p.Software
                If s.Name.ToLower = name Then Return s
            Next
        Next
        Return Nothing
    End Function

    '---COMMENT---------------------------------------------------------------
    '	find the best software match for given extension
    '   first search software of publishers in user preferred sections
    '
    '	Date:		Name:	Description:
    '	12/30/04	JEA		Creation
    '-------------------------------------------------------------------------
    Public Function SoftwareForExtension(ByVal extension As String) As AMP.Software
        Dim software As AMP.Software
        Dim preferred As Boolean

        For Each p As AMP.Publisher In Me.InnerList
            preferred = (p.Section And Profile.User.Section) > 0
            For Each s As AMP.Software In p.Software
                If s.CanOpen(extension) Then
                    If preferred Then
                        Return s
                    Else
                        ' keep track of non-preferred matches
                        software = s
                    End If
                End If
            Next
        Next
        ' if we get here then return any non-preferred match
        Return software
    End Function

    '---COMMENT---------------------------------------------------------------
    '	get all software of type plugin
    '
    '	Date:		Name:	Description:
    '	12/26/04	JEA		Creation
    '-------------------------------------------------------------------------
    Public Function Plugins() As ArrayList
        Dim matches As New ArrayList
        For Each p As AMP.Publisher In Me.InnerList
            For Each plugin As AMP.Software In p.Software.Plugins
                matches.Add(plugin)
            Next
        Next
        matches.Sort()
        Return matches
    End Function

    '---COMMENT---------------------------------------------------------------
    '	get all software of type application
    '
    '	Date:		Name:	Description:
    '	12/26/04	JEA		Creation
    '-------------------------------------------------------------------------
    Public Function Applications() As ArrayList
        Dim matches As New ArrayList
        For Each p As AMP.Publisher In Me.InnerList
            For Each title As AMP.Software In p.Software.Applications
                matches.Add(title)
            Next
        Next
        matches.Sort()
        Return matches
    End Function

    '---COMMENT---------------------------------------------------------------
    '	basic list methods
    '
    '	Date:		Name:	Description:
    '	12/26/04	JEA		Creation
    '-------------------------------------------------------------------------
    Public Function Add(ByVal entity As AMP.Publisher) As Integer
        Return MyBase.InnerList.Add(entity)
    End Function

    Public Sub Remove(ByVal entity As AMP.Publisher)
        MyBase.InnerList.Remove(entity)
    End Sub

    Public Sub Remove(ByVal id As Guid)
        Me.Remove(Me.WithID(id))
    End Sub
End Class
