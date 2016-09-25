<Serializable()> _
Public Class Version
    Implements IComparable

    Private _number As String
    Private _iconFile As String

#Region " Properties "

    Public Property Number() As String
        Get
            Return _number
        End Get
        Set(ByVal Value As String)
            _number = Security.SafeString(Value, 10).ToLower
        End Set
    End Property

    Public Property IconFile() As String
        Get
            Return _iconFile
        End Get
        Set(ByVal Value As String)
            _iconFile = Security.SafeString(Value, 50)
        End Set
    End Property

#End Region

    Public Sub New(ByVal number As String, ByVal iconFile As String)
        _number = number
        _iconFile = iconFile
    End Sub

    Public Sub New()

    End Sub

    Public Function CompareTo(ByVal entity As Object) As Integer Implements System.IComparable.CompareTo
        Return String.Compare(Me.Number, DirectCast(entity, AMP.Version).Number)
    End Function
End Class
