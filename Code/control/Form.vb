Namespace Controls
    Public Class Form
        Inherits System.Web.UI.HtmlControls.HtmlForm

        Private _action As String = HttpContext.Current.Request.Url.PathAndQuery
        Private _validation As ArrayList

#Region " Properties "

        Public Property Validation() As ArrayList
            Get
                If _validation Is Nothing Then _validation = New ArrayList
                Return _validation
            End Get
            Set(ByVal Value As ArrayList)
                _validation = Value
            End Set
        End Property

        Public Property Action() As String
            Get
                Return _action
            End Get
            Set(ByVal Value As String)
                _action = Value
            End Set
        End Property

#End Region

        Protected Overrides Sub RenderAttributes(ByVal writer As System.Web.UI.HtmlTextWriter)
            With writer
                .WriteAttribute("id", Me.ID)
                .WriteAttribute("method", Me.Method)
                .WriteAttribute("action", Me.Action)
                If Me.Enctype <> Nothing Then .WriteAttribute("enctype", Me.Enctype)
            End With
        End Sub

        Protected Overrides Sub RenderChildren(ByVal writer As System.Web.UI.HtmlTextWriter)
            For Each c As Control In MyBase.Controls
                c.RenderControl(writer)
            Next
        End Sub

        '---COMMENT---------------------------------------------------------------
        '   build XHTML compliant script block to validate form fields
        '   type values must equal functions in validation.js
        ' 
        '   Date:       Name:   Description:
        '	1/2/05      JEA     Created
        '   1/5/05      JEA     Prevent trailing comma
        '-------------------------------------------------------------------------
        Public Function ValidationScriptBlock() As String
            Dim scriptBlock As New System.Text.StringBuilder
            Dim first As Boolean = True
            With scriptBlock
                .Append("function Validators() { return [")
                For Each v As AMP.Controls.Validation In _validation
                    If Not v.Control Is Nothing Then
                        If Not first Then
                            .Append(",")
                        Else
                            first = False
                        End If
                        .Append(Environment.NewLine)
                        .Append(vbTab)
                        .Append("new Validation.Field(""")
                        .Append(v.Control.ClientID)
                        .Append(""",""")
                        .Append(v.Type)
                        .Append(""",""")
                        .Append(v.Message)
                        .Append(""",")
                        .Append(IIf(v.Required, "true", "false"))
                        .Append(")")
                    End If
                Next
                .Append("]; }")
                Return .ToString
            End With
        End Function
    End Class
End Namespace
