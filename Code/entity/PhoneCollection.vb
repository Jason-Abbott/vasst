Imports System.Collections

<Serializable()> _
Public Class PhoneCollection
    Inherits CollectionBase

    '---COMMENT---------------------------------------------------------------
    '	basic list methods
    '
    '	Date:		Name:	Description:
    '	3/7/05	    JEA		Creation
    '-------------------------------------------------------------------------
    Public Function Add(ByVal entity As AMP.Phone) As Integer
        Return MyBase.InnerList.Add(entity)
    End Function

    Public Sub Remove(ByVal entity As AMP.Phone)
        MyBase.InnerList.Remove(entity)
    End Sub

End Class
