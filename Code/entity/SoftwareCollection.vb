<Serializable()> _
Public Class SoftwareCollection
    Inherits CollectionBase

#Region " Properties "

    Public Property Item(ByVal index As Integer) As AMP.Software
        Get
            Return DirectCast(MyBase.InnerList(index), AMP.Software)
        End Get
        Set(ByVal Value As AMP.Software)
            MyBase.InnerList(index) = Value
        End Set
    End Property

    Default Public ReadOnly Property WithID(ByVal id As Guid) As AMP.Software
        Get
            For Each item As AMP.Software In Me.InnerList
                If item.ID.Equals(id) Then
                    Return item
                    Exit For
                End If
            Next
        End Get
    End Property

    Default Public ReadOnly Property WithID(ByVal id As String) As AMP.Software
        Get
            Return Me.WithID(New Guid(id))
        End Get
    End Property

#End Region

    '---COMMENT---------------------------------------------------------------
    '	get all software of type plugin
    '
    '	Date:		Name:	Description:
    '	12/26/04	JEA		Creation
    '-------------------------------------------------------------------------
    Public Function Plugins() As ArrayList
        Return Me.OfType(Software.Types.Plugin)
    End Function

    '---COMMENT---------------------------------------------------------------
    '	get all software of type application
    '
    '	Date:		Name:	Description:
    '	12/26/04	JEA		Creation
    '-------------------------------------------------------------------------
    Public Function Applications() As ArrayList
        Return Me.OfType(Software.Types.Application)
    End Function

    '---COMMENT---------------------------------------------------------------
    '	return string array of IDs typically for using with Controls.SelectList
    '
    '	Date:		Name:	Description:
    '	2/11/05 	JEA		Creation
    '-------------------------------------------------------------------------
    Public Function IdArray() As String()
        Dim list(Me.InnerList.Count - 1) As String
        Dim x As Integer = 0

        For Each s As AMP.Software In Me.InnerList
            list(x) = s.ID.ToString
            x = x + 1
        Next
        Return list
    End Function

    '---COMMENT---------------------------------------------------------------
    '	get software of given bitmask type
    '
    '	Date:		Name:	Description:
    '	12/30/04	JEA		Creation
    '-------------------------------------------------------------------------
    Private Function OfType(ByVal type As Integer) As ArrayList
        Dim matches As New ArrayList
        For Each s As AMP.Software In Me.InnerList
            If (s.Type And type) > 0 Then
                s.Versions.Sort()
                matches.Add(s)
            End If
        Next
        Return matches
    End Function

    '---COMMENT---------------------------------------------------------------
    '	return software that opens given extension
    '
    '	Date:		Name:	Description:
    '	12/29/04	JEA		Creation
    '-------------------------------------------------------------------------
    Public Function OpensExtension(ByVal extension As String) As ArrayList
        Dim matches As New ArrayList
        For Each s As AMP.Software In Me.InnerList
            If s.CanOpen(extension) Then matches.Add(s)
        Next
        Return matches
    End Function

    '---COMMENT---------------------------------------------------------------
    '	return software with given name
    '
    '	Date:		Name:	Description:
    '	1/27/05 	JEA		Creation
    '-------------------------------------------------------------------------
    Public Function WithName(ByVal name As String) As AMP.Software
        For Each s As AMP.Software In Me.InnerList
            If s.Name = name Then Return s
        Next
        Return Nothing
    End Function

    '---COMMENT---------------------------------------------------------------
    '	basic list methods
    '
    '	Date:		Name:	Description:
    '	12/26/04	JEA		Creation
    '-------------------------------------------------------------------------
    Public Function Add(ByVal entity As AMP.Software) As Integer
        Return MyBase.InnerList.Add(entity)
    End Function

    Public Sub Remove(ByVal entity As AMP.Software)
        MyBase.InnerList.Remove(entity)
    End Sub

    Public Sub Remove(ByVal id As Guid)
        Me.Remove(Me.WithID(id))
    End Sub

End Class
