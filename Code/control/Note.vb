Imports System.IO
Imports System.Web.UI
Imports System.Web.Caching
Imports System.Web.HttpContext
Imports System.Configuration.ConfigurationSettings

Namespace Controls
    Public Class Note
        Inherits HtmlContainerControl

        Private _text As String
        Private _resx As String
        Private _valid As Boolean = False

#Region " Properties "

        Public Property Text() As String
            Get
                Return _text
            End Get
            Set(ByVal Value As String)
                _text = Value
            End Set
        End Property

        Public WriteOnly Property Resx() As String
            Set(ByVal Value As String)
                _resx = Value
            End Set
        End Property

#End Region

        '---COMMENT---------------------------------------------------------------
        '	get text from resource string or property
        '
        '	Date:		Name:	Description:
        '	1/13/05  	JEA		Creation
        '-------------------------------------------------------------------------
        Private Sub Note_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
            If _text = Nothing AndAlso _resx <> Nothing Then
                _text = DirectCast(Me.Page, AMP.Page).Say("Note." & _resx)
            End If
            Me.Visible = (_text <> Nothing OrElse Me.HasControls)
        End Sub

        Private Sub Note_PreRender(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.PreRender
            If _text <> Nothing Then
                Me.Controls.Clear()
                Me.Controls.Add(New LiteralControl(_text))
            End If
        End Sub

        Protected Overrides Sub RenderBeginTag(ByVal writer As System.Web.UI.HtmlTextWriter)
            With writer
                .Write("<div class=""note"" id=""")
                .Write(Me.ID)
                .Write("""><div class=""body"" id=""")
                .Write(Me.ID)
                .Write("_body"">")
            End With
        End Sub

        Protected Overrides Sub RenderEndTag(ByVal writer As System.Web.UI.HtmlTextWriter)
            writer.Write("</div></div>")
        End Sub

    End Class
End Namespace
