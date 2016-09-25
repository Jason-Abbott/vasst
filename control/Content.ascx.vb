Namespace Controls
    Public Class ContentWebPart
        Inherits AMP.Controls.WebPart

        Protected ampContent As AMP.Controls.Content

        Public WriteOnly Property File() As String
            Set(ByVal Value As String)
                ampContent.File = Value
            End Set
        End Property

        Public Sub New()
            Me.Minimizable = False
        End Sub

    End Class
End Namespace