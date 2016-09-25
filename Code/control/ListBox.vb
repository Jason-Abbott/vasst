Namespace Controls
    Public Class ListBox
        Inherits WebControls.ListBox
        Implements ISelect

        Private _showLink As Boolean = False
        Private _showNote As Boolean = False

#Region " Properties "

        Public ReadOnly Property Posted() As Boolean Implements ISelect.Posted
            Get
                Return ((Not Me.Selected Is Nothing) AndAlso Me.Selected.Length > 0)
            End Get
        End Property

        Public Property ShowNote() As Boolean
            Get
                Return _showNote
            End Get
            Set(ByVal Value As Boolean)
                _showNote = Value
            End Set
        End Property

        Public Property ShowLink() As Boolean Implements ISelect.ShowLink
            Get
                Return _showLink
            End Get
            Set(ByVal Value As Boolean)
                _showLink = Value
            End Set
        End Property

        Public Property Selected() As String() Implements ISelect.Selected
            Get
                If Page.Request.Form(Me.UniqueID) <> Nothing Then
                    Return Page.Request.Form(Me.UniqueID).Split(","c)
                Else
                    Return Nothing
                End If
            End Get
            Set(ByVal Value As String())
                If Value.Length > 0 Then
                    For Each item As ListItem In Me.Items
                        item.Selected = (Array.IndexOf(Value, item.Value) >= 0)
                        BugOut("item {0} selection: {1}", item.Value, item.Selected)
                    Next
                End If
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

#End Region

        Protected Overrides Sub Render(ByVal writer As System.Web.UI.HtmlTextWriter)
            With writer
                .Write("<select id=""")
                .Write(Me.ClientID)
                .Write(""" name=""")
                .Write(Me.UniqueID)
                .Write("""")
                If Me.Rows > 1 Then
                    .Write(" size=""")
                    .Write(Me.Rows)
                    .Write("""")
                End If
                If Me.SelectionMode = ListSelectionMode.Multiple Then
                    .Write(" multiple=""multiple""")
                End If
                .Write(">")
                For Each li As ListItem In Me.Items
                    .Write("<option value=""")
                    .Write(li.Value)
                    If li.Selected Then .Write(""" selected=""selected""")
                    .Write(""">")
                    .Write(li.Text)
                    .Write("</option>")
                Next
                .Write("</select>")
                If Me.Rows > 1 Then
                    Dim hostPage As AMP.Page = DirectCast(Me.Page, AMP.Page)
                    If _showLink Then
                        .Write("<a class=""listAction"" href=""javascript:ClearSelection('")
                        .Write(Me.ClientID)
                        .Write("')"">")
                        .Write(hostPage.Say("Action.Clear"))
                        .Write("</a>")
                    End If
                    If _showNote Then
                        .Write("<div class=""note"">")
                        .Write(hostPage.Say("Note.MultiSelectNote"))
                        .Write("</div>")
                    End If
                End If
            End With
        End Sub
    End Class
End Namespace
