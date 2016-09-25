Imports System.Collections

<Serializable()> _
Public Class Cart
    Inherits CollectionBase

    '---COMMENT---------------------------------------------------------------
    '	determine when cart was last updated
    '
    '	Date:		Name:	Description:
    '   3/7/05	    JEA		Creation
    '-------------------------------------------------------------------------
    Public Function LastUpdate() As DateTime
        ' TODO: use this to dump old carts
        Dim updated As DateTime
        For Each i As AMP.CartItem In Me.InnerList
            If updated = Nothing OrElse updated < i.AddedOn Then updated = i.AddedOn
        Next
        Return updated
    End Function

    '---COMMENT---------------------------------------------------------------
    '	get sum of current prices for product formats in cart
    '
    '	Date:		Name:	Description:
    '   3/7/05	    JEA		Creation
    '-------------------------------------------------------------------------
    Public Function Total() As Single
        Dim sum As Single
        For Each c As AMP.CartItem In Me.InnerList
            sum += c.ExtendedPrice
        Next
        Return sum
    End Function

    Public Function Shipping() As Single
        Dim weight As Single = Me.Weight
        ' TODO: compute shipping for given weight
    End Function

    Public Function Weight() As Single
        Dim sum As Single
        For Each c As AMP.CartItem In Me.InnerList
            sum += c.Weight
        Next
        Return sum
    End Function

    '---COMMENT---------------------------------------------------------------
    '	does cart contain item of given product format
    '
    '	Date:		Name:	Description:
    '   3/7/05	    JEA		Creation
    '-------------------------------------------------------------------------
    Public Function HasItem(ByVal id As Guid) As Boolean
        For Each c As AMP.CartItem In Me.InnerList
            If c.Product.ID.Equals(id) Then Return True
        Next
        Return False
    End Function

    '---COMMENT---------------------------------------------------------------
    '	basic list methods
    '
    '	Date:		Name:	Description:
    '	3/7/05	    JEA		Creation
    '-------------------------------------------------------------------------
    Public Function Add(ByVal entity As AMP.CartItem) As Integer
        Return MyBase.InnerList.Add(entity)
    End Function

    Public Sub Remove(ByVal entity As AMP.CartItem)
        MyBase.InnerList.Remove(entity)
    End Sub
End Class
