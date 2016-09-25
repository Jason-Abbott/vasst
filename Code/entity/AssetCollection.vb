Imports System.Text.RegularExpressions
Imports System.Collections
Imports System.Runtime.Serialization
Imports System.Configuration.ConfigurationSettings

<Serializable()> _
Public Class AssetCollection
    Inherits CollectionBase
    Implements IDeserializationCallback

    <NonSerialized()> Private _matches As New Hashtable

#Region " Enumerations "

    Public Enum SortBy
        Name
        Description
        SubmitDate
        Rank
        Popularity
        Author
        VersionDate
    End Enum

    Public Enum SortDirection
        Ascending
        Descending
    End Enum

#End Region

#Region " Properties "

    Public Property Item(ByVal index As Integer) As AMP.Asset
        Get
            Return DirectCast(MyBase.InnerList(index), AMP.Asset)
        End Get
        Set(ByVal Value As AMP.Asset)
            MyBase.InnerList(index) = Value
        End Set
    End Property

    Default Public ReadOnly Property WithID(ByVal id As Guid) As AMP.Asset
        Get
            'Me.InnerList.Synchronized()
            SyncLock Me.InnerList
                For Each item As AMP.Asset In Me.InnerList
                    If item.ID.Equals(id) Then
                        Return item
                        Exit For
                    End If
                Next
            End SyncLock
        End Get
    End Property

    Default Public ReadOnly Property WithID(ByVal id As String) As AMP.Asset
        Get
            Return Me.WithID(New Guid(id))
        End Get
    End Property

#End Region

    '---COMMENT---------------------------------------------------------------
    '	return the newest assets
    '
    '	Date:		Name:	Description:
    '	11/28/04	JEA		Creation
    '   12/3/04     JEA     Limit count to length of list
    '   12/11/04    JEA     Match asset section to user section
    '   12/29/04    JEA     Pass section as integer
    '-------------------------------------------------------------------------
    Public Function Newest(ByVal count As Integer, ByVal section As Integer, _
        ByVal type As Asset.Types) As ArrayList

        Return Me.Search(1, count, Nothing, section, Nothing, type, _
            SortBy.VersionDate, SortDirection.Descending, Nothing)
    End Function

    Public Function Newest(ByVal count As Integer, ByVal section As Integer) As ArrayList
        Return Me.Newest(count, section, Nothing)
    End Function

    '---COMMENT---------------------------------------------------------------
    '	return all unapproved assets
    '
    '	Date:		Name:	Description:
    '	1/27/05 	JEA		Creation
    '-------------------------------------------------------------------------
    Public Function Unapproved() As ArrayList
        Dim result As New ArrayList
        For Each a As AMP.Asset In MyBase.InnerList
            If a.Status = Site.Status.Pending Then result.Add(a)
        Next
        result.Sort()
        Return result
    End Function

    '---COMMENT---------------------------------------------------------------
    '	clear search cache, usually when assets added or removed
    '
    '	Date:		Name:	Description:
    '   3/1/05      JEA     Creation
    '-------------------------------------------------------------------------
    Public Sub ClearSearchCache()
        _matches.Clear()
    End Sub

    '---COMMENT---------------------------------------------------------------
    '	return all assets matching criteria
    '
    '	Date:		Name:	Description:
    '	1/14/05	    JEA		Combination of previous methods
    '   3/1/05      JEA     Support paged results
    '-------------------------------------------------------------------------
    Function Search(ByVal page As Integer, ByVal pageSize As Integer, _
        ByVal text As String, ByVal section As Integer, _
        ByVal category As AMP.Category, ByVal type As Asset.Types, _
        ByVal sort As SortBy, ByVal direction As SortDirection, ByRef total As Integer) As ArrayList

        Dim flexibility As Integer = CType(AppSettings("PageSizeFlexible"), Integer)
        Dim result As ArrayList
        Dim key As String

        If page = Nothing Then page = 1
        If pageSize = Nothing Then pageSize = CType(AppSettings("SearchPageSize"), Integer)
        If section = Nothing Then section = Profile.User.Section

        key = String.Format("{0}_{1}_{2}_{3}_{4}_{5}", _
            text, section, category, type, sort, direction)
        result = DirectCast(_matches(key), ArrayList)

        If result Is Nothing Then
            ' search results have not been cached
            Dim x As Integer = 0
            Dim re As Regex

            If text <> Nothing Then re = New Regex(text, RegexOptions.IgnoreCase)

            result = New ArrayList
            For Each a As AMP.Asset In MyBase.InnerList
                If a.Status = AMP.Site.Status.Approved AndAlso _
                    (a.Section And section) > 0 AndAlso _
                    (type = Nothing OrElse a.Type = type) AndAlso _
                    (category Is Nothing OrElse a.Categories.Contains(category)) AndAlso _
                    (text = Nothing OrElse ( _
                        re.IsMatch(a.Description) _
                        OrElse re.IsMatch(a.Title) _
                        OrElse (category Is Nothing And a.Categories.HasText(text)) _
                        OrElse a.AuthoredBy.HasText(text))) Then

                    result.Add(a)
                End If
            Next
            result.Sort(Me.GetComparer(sort, direction))
            _matches.Add(key, result)
        End If

        ' return page of results
        total = result.Count
        If total <= pageSize + flexibility Then
            Return result
        Else
            Dim index As Integer = (page - 1) * pageSize
            If index + pageSize > total Then pageSize = total - index
            Return result.GetRange(index, pageSize)
        End If
    End Function

    '---COMMENT---------------------------------------------------------------
    '	get all assets from person
    '
    '	Date:		Name:	Description:
    '	1/24/05 	JEA		Creation
    '-------------------------------------------------------------------------
    Public Function FromUser(ByVal person As AMP.Person) As ArrayList
        Dim result As New ArrayList

        For Each a As AMP.Asset In MyBase.InnerList
            If (a.Status And Site.Status.Approved) > 0 AndAlso _
                a.AuthoredBy Is person Then
                result.Add(a)
            End If
        Next
        result.Sort()
        Return result
    End Function

    Public Function FromUser(ByVal id As Guid) As ArrayList
        Return Me.FromUser(WebSite.Persons.WithID(id))
    End Function

    Public Function FromUser(ByVal id As String) As ArrayList
        Return Me.FromUser(WebSite.Persons.WithID(id))
    End Function

    '---COMMENT---------------------------------------------------------------
    '   return asset with name
    '
    '	Date:		Name:	Description:
    '	1/27/05	    JEA		Creation
    '-------------------------------------------------------------------------
    Public Function WithTitle(ByVal title As String) As AMP.Asset
        title = title.ToLower
        For Each a As AMP.Asset In MyBase.InnerList
            If a.Title <> Nothing AndAlso a.Title.ToLower = title Then Return a
        Next
        Return Nothing
    End Function

    '---COMMENT---------------------------------------------------------------
    '	basic list methods
    '
    '	Date:		Name:	Description:
    '	11/23/04	JEA		Creation
    '   3/1/05      JEA     Clear search cache for approved assets
    '-------------------------------------------------------------------------
    Public Function Add(ByVal entity As AMP.Asset) As Integer
        If entity.Status = Site.Status.Approved Then Me.ClearSearchCache()
        Return MyBase.InnerList.Add(entity)
    End Function

    Public Sub Remove(ByVal entity As AMP.Asset)
        If entity.Status = Site.Status.Approved Then Me.ClearSearchCache()
        MyBase.InnerList.Remove(entity)
    End Sub

    Public Sub Remove(ByVal id As Guid)
        Me.ClearSearchCache()
        Me.Remove(Me.WithID(id))
    End Sub

    '---COMMENT---------------------------------------------------------------
    '	get comparer object for given sort property and direction
    '
    '	Date:		Name:	Description:
    '	1/14/05 	JEA		Creation
    '   2/16/05     JEA     Add version date comparison
    '-------------------------------------------------------------------------
    Private Function GetComparer(ByVal sort As SortBy, ByVal direction As SortDirection) As IComparer
        ' sort ascending if no direction specified
        If direction = Nothing Then direction = SortDirection.Ascending

        Select Case sort
            Case SortBy.VersionDate
                Return New AMP.Compare.AssetVersionDate(direction)
            Case SortBy.Description
                Return New AMP.Compare.AssetDescription(direction)
            Case SortBy.Name
                Return New AMP.Compare.AssetName(direction)
            Case SortBy.Popularity
                Return New AMP.Compare.AssetPopularity(direction)
            Case SortBy.Rank
                Return New AMP.Compare.AssetRank(direction)
            Case SortBy.SubmitDate
                Return New AMP.Compare.AssetSubmitDate(direction)
            Case Else
                ' if no sort member specified then sort descending date
                Return New AMP.Compare.AssetSubmitDate(SortDirection.Descending)
        End Select
    End Function

    '---COMMENT---------------------------------------------------------------
    '	sort collection by specified member
    '
    '	Date:		Name:	Description:
    '	11/23/04	JEA		Creation
    '-------------------------------------------------------------------------
    Public Sub Sort(ByVal field As SortBy, ByVal direction As SortDirection)
        MyBase.InnerList.Sort(Me.GetComparer(field, direction))
    End Sub
    ' overload to sort ascending by default
    Public Sub Sort(ByVal field As SortBy)
        Me.Sort(field, SortDirection.Ascending)
    End Sub

    Public Sub OnDeserialization(ByVal sender As Object) Implements System.Runtime.Serialization.IDeserializationCallback.OnDeserialization
        _matches = New Hashtable
    End Sub
End Class