Namespace Controls
    Public Class Tours
        Inherits AMP.Controls.WebPart

        Private _newerThan As DateTime

#Region " Properties "

        Public WriteOnly Property NewerThan() As DateTime
            Set(ByVal Value As DateTime)
                _newerThan = Value
            End Set
        End Property

#End Region

        Public Sub New()
            Me.Minimizable = True
        End Sub

        Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
            Me.Title = Me.Page.Say("Title.Tours")
            Me.Visible = False
        End Sub

    End Class
End Namespace