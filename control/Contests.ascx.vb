Namespace Controls
    Public Class Contests
        Inherits AMP.Controls.WebPart

        Protected rptContests As WebControls.Repeater
        Private _section As Integer

#Region " Properties "

        '---COMMENT---------------------------------------------------------------
        '	only set this property to look at a specific section
        '   otherwise the profile section(s) is used
        '
        '	Date:		Name:	Description:
        '	12/29/04	JEA		Creation
        '-------------------------------------------------------------------------
        Public WriteOnly Property Section() As String
            Set(ByVal Value As String)
                Dim type As Type = GetType(AMP.Site.Section)
                Dim values As AMP.Site.Section() = CType([Enum].GetValues(type), AMP.Site.Section())

                For x As Integer = 0 To values.Length - 1
                    If Value.ToLower = values(x).ToString.ToLower Then
                        _section = values(x)
                        Exit For
                    End If
                Next
            End Set
        End Property

#End Region

        Public Sub New()
            Me.Minimizable = True
        End Sub

        Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
            Me.Title = DirectCast(Me.Page, AMP.Page).Say("Title.Contests")
            Dim contests As ArrayList = WebSite.Contests.Active

            If contests.Count > 0 Then
                rptContests.DataSource = contests
                rptContests.DataBind()
            Else
                Me.Visible = False
            End If
        End Sub

        Protected Function Subtext(ByVal data As Object) As String
            Dim contest As AMP.Contest = DirectCast(data, AMP.Contest)
            If contest.Active Then
                Return String.Format("vote or enter by <span class=""date"">{0:MMMM d}</span>", _
                    contest.FinishVote)
            Else
                Return "see who won"
            End If
        End Function
    End Class
End Namespace