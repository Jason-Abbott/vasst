Imports System.Collections.Specialized

Namespace Controls
    Public MustInherit Class SelectList
        Inherits AMP.Controls.HtmlControl
        Implements ISelect, IPostBackDataHandler

        Private _selected As String()
        Private _multiple As Boolean = True
        Private _chooseText As String = "Choose one"
        Private _rows As Integer = 1
        Private _showLink As Boolean = False

#Region " Properties "

        Public Property Multiple() As Boolean
            Get
                Return _multiple
            End Get
            Set(ByVal Value As Boolean)
                _multiple = Value
            End Set
        End Property

        Public ReadOnly Property Posted() As Boolean Implements ISelect.Posted
            Get
                Return ((Not Me.Selected Is Nothing) AndAlso Me.Selected.Length > 0)
            End Get
        End Property

        Public Property ShowLink() As Boolean Implements ISelect.ShowLink
            Get
                Return _showLink
            End Get
            Set(ByVal Value As Boolean)
                _showLink = Value
            End Set
        End Property

        Public Property Rows() As Integer
            Get
                Return _rows
            End Get
            Set(ByVal Value As Integer)
                _rows = Value
            End Set
        End Property

        Public Overridable Property Selected() As String() Implements ISelect.Selected
            Get
                Return _selected
            End Get
            Set(ByVal Value As String())
                _selected = Value
            End Set
        End Property

        Public ReadOnly Property TopSelection() As String Implements ISelect.TopSelection
            Get
                If Me.Posted Then
                    Return Me.Selected(0)
                Else
                    Return Nothing
                End If
            End Get
        End Property

        Public Property ChooseText() As String
            Get
                Return _chooseText
            End Get
            Set(ByVal Value As String)
                _chooseText = Value
            End Set
        End Property

#End Region

        '---COMMENT---------------------------------------------------------------
        '	get selected values as array
        '
        '	Date:		Name:	Description:
        '	2/22/05 	JEA		Creation
        '-------------------------------------------------------------------------
        Public Overridable Function LoadPostData(ByVal key As String, ByVal posted As NameValueCollection) As Boolean Implements IPostBackDataHandler.LoadPostData
            If posted(key) <> Nothing Then
                _selected = posted(key).Split(","c)
            End If
            Return False
        End Function

        Public Sub RaisePostDataChangedEvent() Implements IPostBackDataHandler.RaisePostDataChangedEvent
            ' this is called if LoadPostData returns true
        End Sub

        '---COMMENT---------------------------------------------------------------
        '	indicate if the given value has been selected
        '
        '	Date:		Name:	Description:
        '	1/5/05	    JEA		Creation
        '-------------------------------------------------------------------------
        Protected Overridable Function IsSelected(ByVal value As String) As Boolean
            If _selected Is Nothing Then
                Return False
            Else
                Return (Array.IndexOf(_selected, value) >= 0)
            End If
        End Function

        '---COMMENT---------------------------------------------------------------
        '	write opening select tag
        '
        '	Date:		Name:	Description:
        '	1/5/05	    JEA		Creation
        '   2/24/05     JEA     Make "multiple" optional
        '-------------------------------------------------------------------------
        Protected Sub StartTag(ByVal writer As System.Web.UI.HtmlTextWriter)
            Dim keys As IEnumerator = Me.Attributes.Keys.GetEnumerator()
            With writer
                .Write("<select name=""")
                .Write(Me.UniqueID)
                .Write(""" id=""")
                .Write(Me.ClientID)
                .Write("""")

                If _rows > 1 Then
                    .Write(" size=""")
                    .Write(_rows)
                    .Write("""")
                    If _multiple Then .Write(" multiple=""multiple""")
                End If

                Me.RenderCss(writer)

                While keys.MoveNext()
                    Dim key As String = DirectCast(keys.Current, String)
                    .Write(" ")
                    .Write(key)
                    .Write("=""")
                    .Write(Me.Attributes(key))
                    .Write("""")
                End While

                .Write(">")
                If _rows = 1 Then
                    .Write("<option value=""0"" class=""choose"">")
                    .Write(_chooseText)
                    .Write("</option>")
                End If
            End With
        End Sub

        '---COMMENT---------------------------------------------------------------
        '	write closing select tag
        '
        '	Date:		Name:	Description:
        '	1/13/05	    JEA		Creation
        '-------------------------------------------------------------------------
        Protected Sub EndTag(ByVal writer As HtmlTextWriter)
            With writer
                .Write("</select>")
                If _rows > 1 Then
                    Dim hostPage As AMP.Page = DirectCast(Me.Page, AMP.Page)
                    If _showLink Then
                        .Write("<a class=""listAction"" href=""javascript:DOM.ClearSelection('")
                        .Write(Me.ClientID.Trim)
                        .Write("')"">")
                        .Write(hostPage.Say("Action.Clear"))
                        .Write("</a>")
                    End If
                    If Me.ShowNote Then
                        .Write("<div class=""note"">")
                        .Write(hostPage.Say("Note.MultiSelectNote"))
                        .Write("</div>")
                    End If
                End If
            End With
        End Sub
    End Class
End Namespace