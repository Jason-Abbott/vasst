Imports System.Collections
Imports System.Runtime.Serialization

<Serializable()> _
Public Class ProductCollection
    Inherits CollectionBase

#Region " Properties "

    Default Public Property Item(ByVal index As Integer) As AMP.Product
        Get
            Return DirectCast(MyBase.InnerList(index), AMP.Product)
        End Get
        Set(ByVal Value As AMP.Product)
            MyBase.InnerList(index) = Value
        End Set
    End Property

    Public ReadOnly Property ByID(ByVal id As Guid) As AMP.Product
        Get
            'Me.InnerList.Synchronized()
            'SyncLock Me.InnerList
            For Each item As AMP.Product In Me.InnerList
                If item.ID.Equals(id) Then
                    Return item
                    Exit For
                End If
            Next
            'End SyncLock
        End Get
    End Property

    Public ReadOnly Property ByID(ByVal id As String) As AMP.Product
        Get
            Return Me.ByID(New Guid(id))
        End Get
    End Property

#End Region

    '---COMMENT---------------------------------------------------------------
    '	basic list methods
    '
    '	Date:		Name:	Description:
    '	12/19/04	JEA		Creation
    '-------------------------------------------------------------------------
    Public Function Add(ByVal entity As AMP.Product) As Integer
        Return MyBase.InnerList.Add(entity)
    End Function

    Public Sub Remove(ByVal entity As AMP.Product)
        MyBase.InnerList.Remove(entity)
    End Sub

    Public Sub Remove(ByVal id As Guid)
        Me.Remove(Me.ByID(id))
    End Sub

End Class
