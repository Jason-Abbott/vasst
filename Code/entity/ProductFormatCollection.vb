Imports System.Collections
Imports System.Runtime.Serialization

<Serializable()> _
Public Class ProductFormatCollection
    Inherits CollectionBase

    '---COMMENT---------------------------------------------------------------
    '	basic list methods
    '
    '	Date:		Name:	Description:
    '   3/7/05	    JEA		Creation
    '-------------------------------------------------------------------------
    Public Function Add(ByVal entity As AMP.ProductFormat) As Integer
        Return MyBase.InnerList.Add(entity)
    End Function

    Public Sub Remove(ByVal entity As AMP.ProductFormat)
        MyBase.InnerList.Remove(entity)
    End Sub
End Class
