Imports System.Collections

<Serializable()> _
Public Class PermissionCollection
    Inherits DictionaryBase

#Region " Properties "

    Public ReadOnly Property All() As AMP.Site.Permission()
        Get
            Dim x As Integer = 0
            Dim permission As AMP.Site.Permission()
            ReDim permission(MyBase.Dictionary.Keys.Count)

            For Each key As AMP.Site.Permission In MyBase.Dictionary.Keys
                permission(x) = key
                x += 1
            Next

            Array.Sort(permission)
            Return permission
        End Get
    End Property

    Public ReadOnly Property Keys() As ICollection
        Get
            Return MyBase.Dictionary.Keys
        End Get
    End Property

#End Region

    '---COMMENT---------------------------------------------------------------
    '	basic list methods
    '
    '	Date:		Name:	Description:
    '	12/13/04	JEA		Creation
    '-------------------------------------------------------------------------
    Public Function Add(ByVal permission As AMP.Site.Permission) As Integer
        If Not MyBase.Dictionary.Contains(permission) Then
            MyBase.Dictionary.Add(permission, permission)
        End If
    End Function

    Public Sub Remove(ByVal permission As AMP.Site.Permission)
        MyBase.Dictionary.Remove(permission)
    End Sub

End Class
