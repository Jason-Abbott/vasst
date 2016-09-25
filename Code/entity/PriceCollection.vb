Imports System.Collections
Imports System.Runtime.Serialization

<Serializable()> _
Public Class PriceCollection
    Inherits CollectionBase

    '---COMMENT---------------------------------------------------------------
    '	get current price
    '
    '	Date:		Name:	Description:
    '   3/7/05	    JEA		Creation
    '-------------------------------------------------------------------------
    Public Function Current() As Single
        ' TODO: find best match if multiple
        For Each p As AMP.Price In Me.InnerList
            If p.Enabled AndAlso p.StartOn <= DateTime.Now AndAlso p.EndAfter > DateTime.Now Then
                Return p.Value
            End If
        Next
        Return Nothing
    End Function

    '---COMMENT---------------------------------------------------------------
    '	basic list methods
    '
    '	Date:		Name:	Description:
    '   3/7/05	    JEA		Creation
    '-------------------------------------------------------------------------
    Public Function Add(ByVal entity As AMP.Price) As Integer
        Return MyBase.InnerList.Add(entity)
    End Function

    Public Sub Remove(ByVal entity As AMP.Price)
        MyBase.InnerList.Remove(entity)
    End Sub
End Class
