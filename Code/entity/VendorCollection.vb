<Serializable()> _
Public Class VendorCollection
    Inherits CollectionBase

#Region " Properties "

    Public Property Item(ByVal index As Integer) As AMP.Vendor
        Get
            Return DirectCast(MyBase.InnerList(index), AMP.Vendor)
        End Get
        Set(ByVal Value As AMP.Vendor)
            MyBase.InnerList(index) = Value
        End Set
    End Property

    Default Public ReadOnly Property WithID(ByVal id As Guid) As AMP.Vendor
        Get
            For Each v As AMP.Vendor In Me.InnerList
                If v.ID.Equals(id) Then Return v
            Next
            Return Nothing
        End Get
    End Property

    Default Public ReadOnly Property WithID(ByVal id As String) As AMP.Vendor
        Get
            Return Me.WithID(New Guid(id))
        End Get
    End Property

#End Region

    Public Function ProductWithID(ByVal id As Guid) As AMP.Product
        For Each v As AMP.Vendor In Me.InnerList
            For Each p As AMP.Product In v.Products
                If p.ID.Equals(id) Then Return p
            Next
        Next
        Return Nothing
    End Function

    Public Function ProductWithID(ByVal id As String) As AMP.Product
        Return Me.ProductWithID(New Guid(id))
    End Function

    '---COMMENT---------------------------------------------------------------
    '	find vendor with given name
    '
    '	Date:		Name:	Description:
    '	3/5/05	    JEA		Creation
    '-------------------------------------------------------------------------
    Public Function WithName(ByVal name As String) As AMP.Vendor
        name = name.ToLower
        For Each v As AMP.Vendor In Me.InnerList
            If v.Name.ToLower = name Then Return v
        Next
        Return Nothing
    End Function

    '---COMMENT---------------------------------------------------------------
    '	basic list methods
    '
    '	Date:		Name:	Description:
    '	3/5/05  	JEA		Creation
    '-------------------------------------------------------------------------
    Public Function Add(ByVal entity As AMP.Vendor) As Integer
        Return MyBase.InnerList.Add(entity)
    End Function

    Public Sub Remove(ByVal entity As AMP.Vendor)
        MyBase.InnerList.Remove(entity)
    End Sub

    Public Sub Remove(ByVal id As Guid)
        Me.Remove(Me.WithID(id))
    End Sub
End Class
