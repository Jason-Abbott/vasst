Namespace Templates
    Public MustInherit Class ThreeColumn
        Inherits AMP.Template

#Region " Controls "

        Protected phWebParts As PlaceHolder
        Protected phBody As PlaceHolder
        Protected tbSearch As TextBox

#End Region

#Region " Properties "

        Public Overrides Property Body() As WebControls.PlaceHolder
            Get
                Return phBody
            End Get
            Set(ByVal Value As WebControls.PlaceHolder)
                phBody = Value
            End Set
        End Property

        Public Overrides Property WebParts() As WebControls.PlaceHolder
            Get
                Return phWebParts
            End Get
            Set(ByVal Value As WebControls.PlaceHolder)
                phWebParts = Value
            End Set
        End Property

#End Region

        Private Sub Page_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Init
            Me.ID = "3Col"
        End Sub

        Private Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
            Me.StyleSheet.Add("ThreeColumn")
        End Sub
    End Class
End Namespace