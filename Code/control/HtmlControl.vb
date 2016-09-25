Option Strict On

Imports System.Web.UI

Namespace Controls
    Public MustInherit Class HtmlControl
        Inherits System.Web.UI.HtmlControls.HtmlControl

        Private _showNote As Boolean = True
        Private _hostPage As AMP.Page
        Private _required As Boolean = False

#Region " Properties "

        Public Property Required() As Boolean
            Get
                Return _required
            End Get
            Set(ByVal Value As Boolean)
                _required = Value
            End Set
        End Property

        Protected Shadows Property Page() As AMP.Page
            Get
                Return _hostPage
            End Get
            Set(ByVal Value As AMP.Page)
                _hostPage = Value
            End Set
        End Property

        Public Property ShowNote() As Boolean
            Get
                Return _showNote
            End Get
            Set(ByVal Value As Boolean)
                _showNote = Value
            End Set
        End Property

        Public WriteOnly Property CssClass() As String
            Set(ByVal Value As String)
                Me.Attributes.Add("class", Value)
            End Set
        End Property

#End Region

#Region " Event Properties "

        Public WriteOnly Property OnClick() As String
            Set(ByVal Value As String)
                MyBase.Attributes.Add("onclick", Value)
            End Set
        End Property

        Public WriteOnly Property OnBlur() As String
            Set(ByVal Value As String)
                MyBase.Attributes.Add("onblur", Value)
            End Set
        End Property

        Public WriteOnly Property OnChange() As String
            Set(ByVal Value As String)
                MyBase.Attributes.Add("onchange", Value)
            End Set
        End Property

        Public WriteOnly Property OnFocus() As String
            Set(ByVal Value As String)
                MyBase.Attributes.Add("onfocus", Value)
            End Set
        End Property

        Public WriteOnly Property OnMouseOver() As String
            Set(ByVal Value As String)
                MyBase.Attributes.Add("onmouseover", Value)
            End Set
        End Property

        Public WriteOnly Property OnMouseOut() As String
            Set(ByVal Value As String)
                MyBase.Attributes.Add("onmouseout", Value)
            End Set
        End Property

#End Region

        Private Sub HtmlControl_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Init
            _hostPage = DirectCast(MyBase.Page, AMP.Page)
        End Sub

        '---COMMENT---------------------------------------------------------------
        '	render css information for tag
        '
        '	Date:		Name:	Description:
        '	11/1/04	    JEA		Creation
        '-------------------------------------------------------------------------
        Protected Sub RenderCss(ByVal writer As System.Web.UI.HtmlTextWriter)
            With writer
                'If _cssClass <> Nothing Then
                '    .Write(" class=""")
                '    .Write(_cssClass)
                '    .Write("""")
                'End If
                If Me.Style.Count > 0 Then
                    .Write(" style=""")
                    For Each key As String In Me.Style.Keys
                        .Write(key)
                        .Write(": ")
                        .Write(Me.Style(key))
                        .Write(";")
                    Next
                    .Write("""")
                End If
            End With
        End Sub

        '---COMMENT---------------------------------------------------------------
        '   render standard form labels and notes
        ' 
        '   Date:       Name:   Description:
        '	2/24/05     JEA     Created
        '-------------------------------------------------------------------------
        Public Sub RenderLabel(ByVal label As String, ByVal writer As HtmlTextWriter)
            With writer
                .Write("<label")
                If Me.Required Then .Write(" class=""required""")
                .Write(" for=""")
                .Write(Me.ClientID)
                .Write(""">")
                .Write(label)
                .Write("</label>")
            End With
        End Sub

        Public Sub RenderNote(ByVal note As String, ByVal inline As Boolean, _
            ByVal writer As HtmlTextWriter)

            If Me.ShowNote AndAlso note <> Nothing Then
                With writer
                    .Write(IIf(inline, "<span", "<div"))
                    .Write(" class=""note"">")
                    .Write(note)
                    .Write(IIf(inline, "</span>", "</div>"))
                End With
            End If
        End Sub

        '---COMMENT---------------------------------------------------------------
        '	manually add validation control to form
        '
        '	Date:		Name:	Description:
        '	2/24/05	    JEA		Creation
        '-------------------------------------------------------------------------
        Protected Sub RegisterValidation(ByVal type As String, ByVal message As String)
            Dim validation As New AMP.Controls.Validation
            With validation
                .Type = type
                .Message = message
                .Required = Me.Required
                .Control = Me
                .Register(Me.Page)
            End With
        End Sub

    End Class
End Namespace