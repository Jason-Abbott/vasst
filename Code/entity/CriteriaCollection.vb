<Serializable()> _
Public Class CriteriaCollection
    Inherits CollectionBase

    '---COMMENT---------------------------------------------------------------
    '	basic list methods
    '
    '	Date:		Name:	Description:
    '	3/8/05  	JEA		Creation
    '-------------------------------------------------------------------------
    Public Function Add(ByVal entity As AMP.ICriterion) As Integer
        Return MyBase.InnerList.Add(entity)
    End Function

    Public Sub Remove(ByVal entity As AMP.ICriterion)
        MyBase.InnerList.Remove(entity)
    End Sub
End Class
