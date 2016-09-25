Namespace Controls
    Public Class EnumCheckbox
        Inherits AMP.Controls.HtmlControl
        Implements IPostBackDataHandler

        Private _type As Type
        Private _group As Boolean = True
        Private _mask As Integer = -1
        Private _selected As Integer   ' bitmask

#Region " Properties "

        Public Property Group() As Boolean
            Get
                Return _group
            End Get
            Set(ByVal Value As Boolean)
                _group = Value
            End Set
        End Property

        Public Property Selected() As Integer
            Get
                Return _selected
            End Get
            Set(ByVal Value As Integer)
                _selected = Value
            End Set
        End Property

        Public Property Type() As Type
            Get
                Return _type
            End Get
            Set(ByVal Value As Type)
                _type = Value
            End Set
        End Property

        '---COMMENT---------------------------------------------------------------
        '	bitmaks to filter returned enums
        '
        '	Date:		Name:	Description:
        '	1/12/05	    JEA		Creation
        '-------------------------------------------------------------------------
        Public Property Mask() As Integer
            Get
                Return _mask
            End Get
            Set(ByVal Value As Integer)
                _mask = Value
            End Set
        End Property

#End Region

        '---COMMENT---------------------------------------------------------------
        '	get enum values
        '
        '	Date:		Name:	Description:
        '	2/22/05 	JEA		Creation
        '-------------------------------------------------------------------------
        Public Function LoadPostData(ByVal key As String, ByVal posted As System.Collections.Specialized.NameValueCollection) As Boolean Implements IPostBackDataHandler.LoadPostData
            If posted(key) <> Nothing Then
                Dim selected As String() = posted(key).Split(","c)
                _selected = 0
                For x As Integer = 0 To selected.Length - 1
                    _selected = _selected Or CInt(selected(x))
                Next
            End If
            Return False
        End Function

        Public Sub RaisePostDataChangedEvent() Implements IPostBackDataHandler.RaisePostDataChangedEvent
            ' this is called if LoadPostData returns true
        End Sub

        '---COMMENT---------------------------------------------------------------
        '	check if value is present in bitmask of selected enumerations
        '
        '	Date:		Name:	Description:
        '	1/5/05	    JEA		Creation
        '-------------------------------------------------------------------------
        Protected Function IsSelected(ByVal value As Integer) As Boolean
            If _selected = Nothing Then
                Return False
            Else
                Return (_selected And value) > 0
            End If
        End Function

        '---COMMENT---------------------------------------------------------------
        '	render a list of enumerations
        '
        '	Date:		Name:	Description:
        '	12/30/04	JEA		Creation
        '   1/12/05     JEA     Add bitmask filtering
        '   2/9/05      JEA     Render as checkboxes
        '   2/20/05     JEA     Use label tag
        '-------------------------------------------------------------------------
        Protected Overrides Sub Render(ByVal writer As System.Web.UI.HtmlTextWriter)
            Dim values As Integer() = CType([Enum].GetValues(_type), Integer())
            Dim names As String() = [Enum].GetNames(_type)
            Dim id As String

            Array.Sort(names, values)

            With writer
                For x As Integer = 0 To values.Length - 1
                    If _mask = -1 OrElse (values(x) And _mask) > 0 Then
                        If Common.IsFlag(values(x)) Then
                            id = String.Format("{0}_{1}", Me.ClientID, names(x))
                            .Write("<div class=""enum""><input id=""")
                            .Write(id)
                            If _group Then  ' treat as group
                                .Write(""" name=""")
                                .Write(Me.UniqueID)
                            End If
                            .Write(""" type=""checkbox"" value=""")
                            .Write(values(x))
                            .Write("""")
                            If Me.IsSelected(values(x)) Then .Write(" checked=""checked""")
                            .Write("><label for=""")
                            .Write(id)
                            .Write(""">")
                            .Write(Format.NormalSpacing(names(x)))
                            .Write("</label></div>")
                        End If
                    End If
                Next
            End With
        End Sub
    End Class
End Namespace