Namespace Controls
    Public Class EnumList
        Inherits AMP.Controls.SelectList

        Private _type As Type
        Private _mask As Integer = -1
        Private _selected As Integer    ' bitmask

#Region " Properties "

        Public Shadows Property Selected() As Integer
            Get
                Return _selected
            End Get
            Set(ByVal Value As Integer)
                _selected = Value
            End Set
        End Property

        Public Shadows ReadOnly Property Posted() As Boolean
            Get
                Return (Not _selected = Nothing)
            End Get
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
        Public Overrides Function LoadPostData(ByVal key As String, ByVal posted As System.Collections.Specialized.NameValueCollection) As Boolean
            If posted(key) <> Nothing Then
                If Me.Rows = 1 Then
                    _selected = CInt(posted(key))
                Else
                    Dim selected As String() = posted(key).Split(","c)
                    _selected = 0
                    For x As Integer = 0 To selected.Length - 1
                        _selected = _selected Or CInt(selected(x))
                    Next
                End If
            End If
            Return False
        End Function

        '---COMMENT---------------------------------------------------------------
        '	check if value is present in bitmask of selected enumerations
        '
        '	Date:		Name:	Description:
        '	1/5/05	    JEA		Creation
        '-------------------------------------------------------------------------
        Protected Overloads Function IsSelected(ByVal value As Integer) As Boolean
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
        '-------------------------------------------------------------------------
        Protected Overrides Sub Render(ByVal writer As System.Web.UI.HtmlTextWriter)
            Dim values As Integer() = CType([Enum].GetValues(_type), Integer())
            Dim names As String() = [Enum].GetNames(_type)

            Array.Sort(names, values)

            MyBase.StartTag(writer)
            With writer
                For x As Integer = 0 To values.Length - 1
                    If _mask = -1 OrElse (values(x) And _mask) > 0 Then
                        If Common.IsFlag(values(x)) Then
                            .Write("<option value=""")
                            .Write(values(x))
                            .Write("""")
                            If Me.IsSelected(values(x)) Then .Write(" selected=""selected""")
                            .Write(">")
                            .Write(Format.NormalSpacing(names(x)))
                            .Write("</option>")
                        End If
                    End If
                Next
            End With
            MyBase.EndTag(writer)
        End Sub
    End Class
End Namespace