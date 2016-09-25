Namespace Controls
    Public Class RoleTree
        Inherits AMP.Controls.HtmlControl

        Private _toggle As New AMP.Controls.Toggle
        Private _maxRecursion As Integer
        Private _recursion As Integer

        Private Sub RoleTree_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
            _toggle.Height = 11
            _toggle.InitialState = Draw.ToggleState.Closed
        End Sub

        '---COMMENT---------------------------------------------------------------
        '	write roles with inheritance and permissions as tree
        '
        '	Date:		Name:	Description:
        '	1/29/05	    JEA		Creation
        '-------------------------------------------------------------------------
        Protected Overrides Sub Render(ByVal writer As HtmlTextWriter)
            Dim nodeID As String

            _maxRecursion = WebSite.Roles.Count

            With writer
                .Write("<div id=""tree"">")
                For Each r As AMP.Role In WebSite.Roles
                    _recursion = 0
                    nodeID = "role:" & CInt(r.ID)
                    .Write("<div id=""")
                    .Write(nodeID)
                    .Write("""><div class=""roleLabel"">")
                    _toggle.Write(writer)
                    .Write(r.Name)
                    .Write("</div>")
                    ' directly assigned permissions
                    Permissions(writer, r, nodeID)
                    ' inherited roles
                    Inheritance(writer, r, nodeID)
                    .Write("</div>")
                Next
                .Write("</div>")
            End With
        End Sub

        '---COMMENT---------------------------------------------------------------
        '	write inherited roles for given role
        '
        '	Date:		Name:	Description:
        '	1/29/05	    JEA		Creation
        '-------------------------------------------------------------------------
        Private Sub Inheritance(ByVal writer As HtmlTextWriter, ByVal role As AMP.Role, _
            ByVal nodeID As String)

            Dim childNodeID As String

            If role.InheritFrom.Count > 0 AndAlso _recursion <= _maxRecursion Then
                _recursion += 1
                With writer
                    .Write("<div class=""inherit"">")
                    For Each r As AMP.Role In role.InheritFrom
                        childNodeID = nodeID & "_inherit:" & CInt(r.ID)
                        .Write("<div id=""")
                        .Write(childNodeID)
                        .Write("""><div class=""roleLabel"">")
                        _toggle.Write(writer)
                        .Write("From ")
                        .Write(r.Name)
                        .Write("</div>")
                        ' assigned permissions
                        Permissions(writer, r, childNodeID)
                        ' inherited roles
                        Inheritance(writer, r, childNodeID)
                        .Write("</div>")
                    Next
                    .Write("</div>")
                End With
            End If

        End Sub

        '---COMMENT---------------------------------------------------------------
        '	write permissions for given role
        '
        '	Date:		Name:	Description:
        '	1/29/05	    JEA		Creation
        '-------------------------------------------------------------------------
        Private Sub Permissions(ByVal writer As HtmlTextWriter, ByVal role As AMP.Role, _
            ByVal nodeID As String)

            If role.Permissions.Length > 0 Then
                With writer
                    .Write("<div class=""permissions"">")
                    For Each p As AMP.Site.Permission In role.Permissions
                        .Write("<div id=""")
                        .Write(nodeID)
                        .Write("_permission:")
                        .Write(CInt(p))
                        .Write(""">")
                        .Write(Format.NormalSpacing(p.ToString))
                        .Write("</div>")
                    Next
                    .Write("</div>")
                End With
            End If
        End Sub


    End Class
End Namespace