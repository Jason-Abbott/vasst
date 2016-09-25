<Serializable()> _
Public Class PromotionCriterion

    Private _products As AMP.ProductFormatCollection
    Private _cartTotal As Single
    Private _signupBefore As DateTime
    Private _signupAfter As DateTime

    '---COMMENT---------------------------------------------------------------
    '	does given user meet this criterion
    '
    '	Date:		Name:	Description:
    '   3/7/05	    JEA		Creation
    '-------------------------------------------------------------------------
    Public Function Met(ByVal user As AMP.Person) As Boolean
        If Not Me.HasProducts(user) Then Return False
        If Not Me.HasCartTotal(user) Then Return False
        If Not Me.HasSignupAfter(user) Then Return False
        If Not Me.HasSignupBefore(user) Then Return False
        ' all conditions met if here
        Return True
    End Function

    Public Function Met() As Boolean
        Return Me.Met(Profile.User)
    End Function

    '---COMMENT---------------------------------------------------------------
    '	if criteria products specified, does cart have them
    '
    '	Date:		Name:	Description:
    '   3/7/05	    JEA		Creation
    '-------------------------------------------------------------------------
    Private Function HasProducts(ByVal user As AMP.Person) As Boolean
        If _products Is Nothing Then Return True
        If Not user.Cart Is Nothing Then
            For Each p As AMP.ProductFormat In _products
                If Not user.Cart.HasItem(p.ID) Then Return False
            Next
        End If
        Return False
    End Function

    '---COMMENT---------------------------------------------------------------
    '	if cart total specified, does user meet it
    '
    '	Date:		Name:	Description:
    '   3/7/05	    JEA		Creation
    '-------------------------------------------------------------------------
    Private Function HasCartTotal(ByVal user As AMP.Person) As Boolean
        If _cartTotal = Nothing Then Return True
        If Not user.Cart Is Nothing AndAlso user.Cart.Total >= _cartTotal Then Return True
        Return False
    End Function

    '---COMMENT---------------------------------------------------------------
    '	if signup date specified, does user meet it
    '
    '	Date:		Name:	Description:
    '   3/7/05	    JEA		Creation
    '-------------------------------------------------------------------------
    Private Function HasSignupAfter(ByVal user As AMP.Person) As Boolean
        If _signupAfter = Nothing Then Return True
        Return (user.RegisteredOn >= _signupAfter)
    End Function

    '---COMMENT---------------------------------------------------------------
    '	if signup date specified, does user meet it
    '
    '	Date:		Name:	Description:
    '   3/7/05	    JEA		Creation
    '-------------------------------------------------------------------------
    Private Function HasSignupBefore(ByVal user As AMP.Person) As Boolean
        If _signupBefore = Nothing Then Return True
        Return (user.RegisteredOn >= _signupBefore)
    End Function
End Class
