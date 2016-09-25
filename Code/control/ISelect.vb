Namespace Controls
    Public Interface ISelect

        Property Selected() As String()
        Property ShowLink() As Boolean
        'ReadOnly Property BitMask() As Integer
        ReadOnly Property TopSelection() As String
        ReadOnly Property Posted() As Boolean

    End Interface
End Namespace
