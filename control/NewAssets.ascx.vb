Imports System.Configuration.ConfigurationSettings

Namespace Controls
    Public Class NewAssets
        Inherits AMP.Controls.WebPart

        Private _count As Integer = CInt(AppSettings("NewAssetCount"))
        Private _newerThan As DateTime
        Private _section As AMP.Site.Section
        Protected rptAssets As WebControls.Repeater

#Region " Properties "

        Public WriteOnly Property Count() As Integer
            Set(ByVal Value As Integer)
                _count = Value
            End Set
        End Property

        Public WriteOnly Property NewerThan() As DateTime
            Set(ByVal Value As DateTime)
                _newerThan = Value
            End Set
        End Property

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
            Me.Title = Me.Page.Say("Title.NewResources")

            Dim assets As ArrayList = WebSite.Assets.Newest(_count, _section)

            If assets.Count > 0 Then
                rptAssets.DataSource = assets
                rptAssets.DataBind()
            Else
                Me.Visible = False
            End If
        End Sub
    End Class
End Namespace