Namespace Templates
    Public MustInherit Class NewsPaper
        Inherits AMP.Template

#Region " Controls "

        Protected phHeadline As PlaceHolder
        Protected phRightColumn As PlaceHolder
        Protected phLeftColumn As PlaceHolder

#End Region

#Region " Properties "

        Public Overrides WriteOnly Property Headline() As WebControls.PlaceHolder
            Set(ByVal Value As WebControls.PlaceHolder)
                Dim count As Integer = Value.Controls.Count
                For x As Integer = 0 To count - 1
                    phHeadline.Controls.Add(Value.Controls(0))
                Next
            End Set
        End Property

        Public Overrides WriteOnly Property LeftColumn() As WebControls.PlaceHolder
            Set(ByVal Value As WebControls.PlaceHolder)
                Dim count As Integer = Value.Controls.Count
                For x As Integer = 0 To count - 1
                    phLeftColumn.Controls.Add(Value.Controls(0))
                Next
            End Set
        End Property

        'Public Overrides WriteOnly Property RightColumn() As Control
        '    Set(ByVal Value As Control)
        '        Dim count As Integer = Value.Controls.Count
        '        For x As Integer = 0 To count - 1
        '            phRightColumn.Controls.Add(Value.Controls(0))
        '        Next
        '    End Set
        'End Property

#End Region

        Private Sub Page_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Init
            Me.ID = "NewsPaper"
        End Sub

        Private Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
            Me.StyleSheet.Add("NewsPaper")
        End Sub
    End Class
End Namespace