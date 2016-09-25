Imports AMP.AssetCollection

Namespace Compare
    '---COMMENT---------------------------------------------------------------
    '	base class for asset sorting
    '
    '	Date:		Name:	Description:
    '	11/2/04	    JEA		Creation
    '-------------------------------------------------------------------------
    Public MustInherit Class AssetCompare
        Implements IComparer
        Protected _sortDirection As SortDirection = AssetCollection.SortDirection.Ascending

        Public Sub New(ByVal sortDirection As AssetCollection.SortDirection)
            _sortDirection = sortDirection
        End Sub

        Public WriteOnly Property SortDirection() As AssetCollection.SortDirection
            Set(ByVal Value As AssetCollection.SortDirection)
                _sortDirection = Value
            End Set
        End Property

        Protected Function TitleSort(ByVal a1 As AMP.Asset, ByVal a2 As AMP.Asset) As Integer
            Return Me.StringSort(a1.Title, a2.Title)
        End Function

        Protected Function StringSort(ByVal string1 As String, ByVal string2 As String) As Integer
            If _sortDirection = AssetCollection.SortDirection.Ascending Then
                Return String.Compare(string1, string2)
            Else
                Return String.Compare(string2, string1)
            End If
        End Function

        Public MustOverride Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements System.Collections.IComparer.Compare
    End Class

    '---COMMENT---------------------------------------------------------------
    '	sort by name
    '
    '	Date:		Name:	Description:
    '	11/2/04	    JEA		Creation
    '-------------------------------------------------------------------------
    Public Class AssetName
        Inherits AssetCompare

        Public Sub New(ByVal sortDirection As AssetCollection.SortDirection)
            MyBase.New(sortDirection)
        End Sub

        Public Overrides Function Compare(ByVal x As Object, ByVal y As Object) As Integer
            Return Me.TitleSort(DirectCast(x, AMP.Asset), DirectCast(y, AMP.Asset))
        End Function
    End Class

    '---COMMENT---------------------------------------------------------------
    '	sort by submit date
    '
    '	Date:		Name:	Description:
    '	11/2/04	    JEA		Creation
    '-------------------------------------------------------------------------
    Public Class AssetSubmitDate
        Inherits AssetCompare

        Public Sub New(ByVal sortDirection As AssetCollection.SortDirection)
            MyBase.New(sortDirection)
        End Sub

        Public Overrides Function Compare(ByVal x As Object, ByVal y As Object) As Integer
            Dim asset1 As AMP.Asset = DirectCast(x, AMP.Asset)
            Dim asset2 As AMP.Asset = DirectCast(y, AMP.Asset)

            If _sortDirection = AssetCollection.SortDirection.Ascending Then
                Return Date.Compare(asset1.SubmitDate, asset2.SubmitDate)
            Else
                Return Date.Compare(asset2.SubmitDate, asset1.SubmitDate)
            End If
        End Function
    End Class

    '---COMMENT---------------------------------------------------------------
    '	sort by version date
    '
    '	Date:		Name:	Description:
    '	2/16/05	    JEA		Creation
    '-------------------------------------------------------------------------
    Public Class AssetVersionDate
        Inherits AssetCompare

        Public Sub New(ByVal sortDirection As AssetCollection.SortDirection)
            MyBase.New(sortDirection)
        End Sub

        Public Overrides Function Compare(ByVal x As Object, ByVal y As Object) As Integer
            Dim asset1 As AMP.Asset = DirectCast(x, AMP.Asset)
            Dim asset2 As AMP.Asset = DirectCast(y, AMP.Asset)

            If _sortDirection = AssetCollection.SortDirection.Ascending Then
                Return Date.Compare(asset1.VersionDate, asset2.VersionDate)
            Else
                Return Date.Compare(asset2.VersionDate, asset1.VersionDate)
            End If
        End Function
    End Class

    '---COMMENT---------------------------------------------------------------
    '	sort by description
    '
    '	Date:		Name:	Description:
    '	11/2/04	    JEA		Creation
    '-------------------------------------------------------------------------
    Public Class AssetDescription
        Inherits AssetCompare

        Public Sub New(ByVal sortDirection As AssetCollection.SortDirection)
            MyBase.New(sortDirection)
        End Sub

        Public Overrides Function Compare(ByVal x As Object, ByVal y As Object) As Integer
            Return StringSort(DirectCast(x, AMP.Asset).Description, _
                DirectCast(y, AMP.Asset).Description)
        End Function
    End Class

    '---COMMENT---------------------------------------------------------------
    '	sort by owner
    '
    '	Date:		Name:	Description:
    '	11/2/04	    JEA		Creation
    '-------------------------------------------------------------------------
    Public Class AssetOwner
        Inherits AssetCompare

        Public Sub New(ByVal sortDirection As AssetCollection.SortDirection)
            MyBase.New(sortDirection)
        End Sub

        Public Overrides Function Compare(ByVal x As Object, ByVal y As Object) As Integer
            Dim result As Integer = StringSort(DirectCast(x, AMP.Asset).AuthoredBy.DisplayName, _
                DirectCast(y, AMP.Asset).AuthoredBy.DisplayName)

            If result = 0 Then
                Me.SortDirection = AMP.AssetCollection.SortDirection.Ascending
                Return Me.TitleSort(DirectCast(x, AMP.Asset), DirectCast(y, AMP.Asset))
            Else
                Return result
            End If
        End Function
    End Class

    '---COMMENT---------------------------------------------------------------
    '	sort by ranking
    '
    '	Date:		Name:	Description:
    '	11/2/04	    JEA		Creation
    '-------------------------------------------------------------------------
    Public Class AssetRank
        Inherits AssetCompare

        Public Sub New(ByVal sortDirection As AssetCollection.SortDirection)
            MyBase.New(sortDirection)
        End Sub

        Private Function Rank(ByVal asset As AMP.Asset) As Single
            If asset.Ratings Is Nothing Then
                Return 0
            Else
                Return asset.Ratings.WeightedAverage
            End If
        End Function

        Public Overrides Function Compare(ByVal x As Object, ByVal y As Object) As Integer
            Dim asset1 As AMP.Asset = DirectCast(x, AMP.Asset)
            Dim asset2 As AMP.Asset = DirectCast(y, AMP.Asset)
            Dim result As Integer

            If Me.Rank(asset1) > Me.Rank(asset2) Then
                result = 1
            ElseIf Me.Rank(asset1) = Me.Rank(asset2) Then
                result = Me.TitleSort(DirectCast(x, AMP.Asset), DirectCast(y, AMP.Asset))
            Else
                result = -1
            End If

            If _sortDirection = AssetCollection.SortDirection.Descending Then
                Return -(result)
            Else
                Return result
            End If
        End Function
    End Class

    '---COMMENT---------------------------------------------------------------
    '	sort by popularity
    '
    '	Date:		Name:	Description:
    '	11/2/04	    JEA		Creation
    '-------------------------------------------------------------------------
    Public Class AssetPopularity
        Inherits AssetCompare

        Public Sub New(ByVal sortDirection As AssetCollection.SortDirection)
            MyBase.New(sortDirection)
        End Sub

        Private Function Popularity(ByVal asset As AMP.Asset) As Integer
            If (asset.Type And AMP.Asset.Types.File) > 0 Then
                Return asset.File.Downloads
            ElseIf (asset.Type And AMP.Asset.Types.Link) > 0 Then
                Return asset.Link.Views
            Else
                Return 0
            End If
        End Function

        Public Overrides Function Compare(ByVal x As Object, ByVal y As Object) As Integer
            Dim p1 As Integer = Popularity(DirectCast(x, AMP.Asset))
            Dim p2 As Integer = Popularity(DirectCast(y, AMP.Asset))
            Dim direction As AssetCollection.SortDirection = _sortDirection
            Dim result As Integer

            If p1 > p2 Then
                result = 1
            ElseIf p1 = p2 Then
                result = Me.TitleSort(DirectCast(x, AMP.Asset), DirectCast(y, AMP.Asset))
            Else
                result = -1
            End If

            If _sortDirection = AssetCollection.SortDirection.Descending Then
                Return -result
            Else
                Return result
            End If

        End Function
    End Class
End Namespace