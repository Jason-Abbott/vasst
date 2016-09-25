Namespace Templates
    Public MustInherit Class ThreeColumnAdmin
        Inherits AMP.Template

#Region " Controls "

        Protected phWebParts As PlaceHolder
        Protected phBody As PlaceHolder

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
            Me.ID = "3ColAdmin"
        End Sub

        Private Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
            Me.StyleSheet.Add("ThreeColumn")
        End Sub
    End Class
End Namespace