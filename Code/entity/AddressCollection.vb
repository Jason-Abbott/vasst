Imports System.Collections

<Serializable()> _
Public Class AddressCollection
    Inherits CollectionBase

    '---COMMENT---------------------------------------------------------------
    '	basic list methods
    '
    '	Date:		Name:	Description:
    '	11/30/04	JEA		Creation
    '-------------------------------------------------------------------------
    Public Function Add(ByVal entity As AMP.Address) As Integer
        Return MyBase.InnerList.Add(entity)
    End Function

    Public Sub Remove(ByVal entity As AMP.Address)
        MyBase.InnerList.Remove(entity)
    End Sub
End Class
