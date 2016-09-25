<Serializable()> _
Public Class PromotionCollection
    Inherits CollectionBase

    '---COMMENT---------------------------------------------------------------
    '	basic list methods
    '
    '	Date:		Name:	Description:
    '	12/20/04	JEA		Creation
    '-------------------------------------------------------------------------
    Public Function Add(ByVal entity As AMP.Promotion) As Integer
        Return MyBase.InnerList.Add(entity)
    End Function

    Public Sub Remove(ByVal entity As AMP.Promotion)
        MyBase.InnerList.Remove(entity)
    End Sub

End Class
