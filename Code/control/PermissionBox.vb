Namespace Controls
    Public Class PermissionBox
        Inherits AMP.Controls.HtmlControl

        '---COMMENT---------------------------------------------------------------
        '	render a list of permissions
        '
        '	Date:		Name:	Description:
        '	1/29/05 	JEA		Creation
        '-------------------------------------------------------------------------
        Protected Overrides Sub Render(ByVal writer As HtmlTextWriter)
            Dim type As Type = GetType(AMP.Site.Permission)
            Dim values As Integer() = CType([Enum].GetValues(type), Integer())
            Dim names As String() = [Enum].GetNames(type)

            Array.Sort(names, values)

            With writer
                .Write("<div id=""permissionsBox"">")
                For x As Integer = 0 To values.Length - 1
                    .Write("<div id=""permission:")
                    .Write(values(x))
                    .Write(""">")
                    .Write(Format.NormalSpacing(names(x)))
                    .Write("</div>")
                Next
                .Write("</div>")
            End With
        End Sub
    End Class
End Namespace
