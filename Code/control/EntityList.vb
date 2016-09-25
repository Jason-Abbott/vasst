Namespace Controls
    Public Class EntityList
        Inherits AMP.Controls.HtmlControl

        Dim _includeChildren As Boolean = True

#Region " Properties "

        Public Property IncludeChildren() As Boolean
            Get
                Return _includeChildren
            End Get
            Set(ByVal Value As Boolean)
                _includeChildren = Value
            End Set
        End Property

#End Region

        Public Enum Display
            Checkbox
            DropDown
            List
        End Enum

        '---COMMENT---------------------------------------------------------------
        '	render entity types and optionally child types in form
        '
        '	Date:		Name:	Description:
        '	2/15/05	    JEA		Creation
        '-------------------------------------------------------------------------
        Protected Overrides Sub Render(ByVal writer As System.Web.UI.HtmlTextWriter)
            Me.RenderNodes(writer, GetType(AMP.Site.Entity))
        End Sub

        Private Sub RenderNodes(ByVal writer As HtmlTextWriter, ByVal type As System.Type)
            Dim values As Integer() = CType([Enum].GetValues(type), Integer())
            Dim names As String() = [Enum].GetNames(type)

            Array.Sort(names, values)

            With writer
                .Write("<div id=""")
                .Write(type.FullName)
                .Write(""">")
                For x As Integer = 0 To values.Length - 1
                    If Common.IsFlag(values(x)) Then
                        .Write("<div><input id=""")
                        .Write(type.FullName)
                        .Write("_")
                        .Write(values(x))
                        .Write(""" type=""checkbox"" class=""cbx"" value=""")
                        .Write(values(x))
                        .Write(""">")
                        .Write(Format.NormalSpacing(names(x)))

                        Select Case names(x)
                            Case AMP.Site.Entity.Asset.ToString
                                Me.RenderNodes(writer, GetType(AMP.Asset.Types))
                            Case AMP.Site.Entity.Software.ToString
                                Me.RenderNodes(writer, GetType(AMP.Software.Types))
                        End Select

                        .Write("</div>")
                    End If
                Next
                .Write("</div>")
            End With
        End Sub
    End Class
End Namespace