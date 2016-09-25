Namespace Controls
    Public Class Field
        Inherits AMP.Controls.HtmlControl
        Implements IPostBackDataHandler

        Private _note As String
        Private _inline As Boolean = False
        Private _maxLength As Integer
        Private _validate As String
        Private _resx As String
        Private _type As String
        Private _value As String
        Private _checked As Boolean
        Private _label As String

#Region " Properties "

        Public WriteOnly Property Label() As String
            Set(ByVal Value As String)
                _label = Value
            End Set
        End Property

        Public WriteOnly Property Note() As String
            Set(ByVal Value As String)
                _note = Value
            End Set
        End Property

        Public WriteOnly Property MaxLength() As Integer
            Set(ByVal Value As Integer)
                _maxLength = Value
            End Set
        End Property

        Public Property Checked() As Boolean
            Get
                Return _checked
            End Get
            Set(ByVal Value As Boolean)
                _checked = Value
            End Set
        End Property

        Public WriteOnly Property Inline() As Boolean
            Set(ByVal Value As Boolean)
                _inline = Value
            End Set
        End Property

        Public WriteOnly Property Resx() As String
            Set(ByVal Value As String)
                _resx = Value
            End Set
        End Property

        Public Property Value() As String
            Get
                Return _value
            End Get
            Set(ByVal Value As String)
                _value = Value
            End Set
        End Property

        Public WriteOnly Property Validate() As String
            Set(ByVal Value As String)
                _validate = Value
            End Set
        End Property

        Public WriteOnly Property Type() As String
            Set(ByVal Value As String)
                _type = Value
            End Set
        End Property

#End Region

        '---COMMENT---------------------------------------------------------------
        '	get field value
        '
        '	Date:		Name:	Description:
        '	2/22/05 	JEA		Creation
        '-------------------------------------------------------------------------
        Public Function LoadPostData(ByVal key As String, ByVal posted As System.Collections.Specialized.NameValueCollection) As Boolean Implements IPostBackDataHandler.LoadPostData
            Select Case _type.ToLower
                Case "text"
                    _value = posted(key)
                Case "checkbox"
                    _checked = (posted(key) = "on")
                Case "password"
                    _value = posted(key)
            End Select
            Return False
        End Function

        Public Sub RaisePostDataChangedEvent() Implements IPostBackDataHandler.RaisePostDataChangedEvent
            ' this is called if LoadPostData returns true
        End Sub

        '---COMMENT---------------------------------------------------------------
        '	sanity checks and setup validation
        '
        '	Date:		Name:	Description:
        '	2/15/05 	JEA		Creation
        '-------------------------------------------------------------------------
        Private Sub Field_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Init
            If Me.Required AndAlso _validate = Nothing Then
                Throw New Exception("No validation specified for required field")
            End If

            If _label = Nothing Then _label = Me.Page.Say(String.Format("Label.{0}", _resx))

            If _label = Nothing Then
                Throw New Exception(String.Format("No label resource found for ""{0}""", _resx))
            End If

        End Sub

        '---COMMENT---------------------------------------------------------------
        '	register validation; script created during page pre-render
        '
        '	Date:		Name:	Description:
        '	2/26/05 	JEA		Creation
        '-------------------------------------------------------------------------
        Private Sub Field_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
            ' validation
            If _validate <> Nothing AndAlso Me.Visible Then
                Dim alert As String = Me.Page.Say(String.Format("Validate.{0}", _resx))
                If alert = Nothing Then alert = _label
                Me.RegisterValidation(_validate, alert)
            End If
        End Sub

        '---COMMENT---------------------------------------------------------------
        '	write labeled control with validation and notes
        '
        '	Date:		Name:	Description:
        '	2/15/05 	JEA		Creation
        '-------------------------------------------------------------------------
        Protected Overrides Sub Render(ByVal writer As System.Web.UI.HtmlTextWriter)
            If _note = Nothing Then _note = Me.Page.Say(String.Format("Note.{0}", _resx))
            With writer
                Select Case _type.ToLower
                    Case "text", "password"
                        Me.RenderLabel(_label, writer)
                        .Write(Environment.NewLine)
                        .Write("<input type=""")
                        .Write(_type)
                        .Write("""")
                        Me.RenderAttributes(writer)
                        If _maxLength <> Nothing Then
                            .Write(" maxlength=""")
                            .Write(_maxLength)
                            .Write("""")
                        End If
                        If _value <> Nothing Then
                            .Write(" value=""")
                            .Write(HttpUtility.HtmlEncode(_value))
                            .Write("""")
                        End If
                        .Write(" />")
                        Me.RenderNote(_note, _inline, writer)

                    Case "checkbox"
                        .Write("<div class=""checkbox""><input type=""checkbox""")
                        Me.RenderAttributes(writer)
                        If _checked Then .Write(" checked=""checked""")
                        .Write(" />")
                        Me.RenderLabel(_label, writer)
                        .Write("</div>")
                        Me.RenderNote(_note, _inline, writer)

                    Case Else
                        Throw New Exception("Unsupported field type")
                End Select
            End With
        End Sub

        Private Shadows Sub RenderAttributes(ByVal writer As HtmlTextWriter)
            With writer
                .Write(""" name=""")
                .Write(Me.UniqueID)
                .Write("""")
                'MyBase.RenderCss(writer)
                MyBase.RenderAttributes(writer)
            End With
        End Sub
    End Class
End Namespace