<Serializable()> _
Public Class Tour
    Inherits CollectionBase

    Private _description As String
    Private _name As String
    'Private _forPublisher as AMP.Publisher

#Region " Properties "

    Public Property Name() As String
        Get
            Return _name
        End Get
        Set(ByVal Value As String)
            _name = Security.SafeString(Value, 100)
        End Set
    End Property

    Public Property Description() As String
        Get
            Return _description
        End Get
        Set(ByVal Value As String)
            _description = Security.SafeString(Value, 1500)
        End Set
    End Property

#End Region

    Public Function Seminar(ByVal id As Guid) As AMP.Seminar
        For Each item As AMP.Seminar In Me.InnerList
            If item.ID.Equals(id) Then
                Return item
                Exit For
            End If
        Next
    End Function

    Public Function Seminar(ByVal id As String) As AMP.Seminar
        Return Me.Seminar(New Guid(id))
    End Function

    '---COMMENT---------------------------------------------------------------
    '	basic list methods
    '
    '	Date:		Name:	Description:
    '	12/20/04	JEA		Creation
    '-------------------------------------------------------------------------
    Public Function Add(ByVal entity As AMP.Seminar) As Integer
        Return MyBase.InnerList.Add(entity)
    End Function

    Public Sub Remove(ByVal entity As AMP.Seminar)
        MyBase.InnerList.Remove(entity)
    End Sub

    Public Sub Remove(ByVal id As Guid)
        Me.Remove(Me.Seminar(id))
    End Sub
End Class
