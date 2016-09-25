Imports System.Xml
Imports System.Web.UI
Imports System.Web.Caching
Imports System.Web.HttpContext
Imports System.Configuration.ConfigurationSettings

Namespace Controls
    Public Class Menu
        Inherits AMP.Controls.HtmlControl

        Private _fileName As String
        Private _depth As Integer
        Private _orientation As Orientations = Orientations.Horizontal
        Private _nodes As XmlNodeList

        Public Enum Orientations
            Horizontal
            Vertical
        End Enum

#Region " Properties "

        Public WriteOnly Property File() As String
            Set(ByVal Value As String)
                _fileName = Value
            End Set
        End Property

        Public WriteOnly Property DisplayDepth() As Integer
            Set(ByVal Value As Integer)
                _depth = Value
            End Set
        End Property

        Public WriteOnly Property Orientation() As Orientations
            Set(ByVal Value As Orientations)
                _orientation = Value
            End Set
        End Property

#End Region

        '---COMMENT---------------------------------------------------------------
        '   load menu as xml nodes from cache or file
        ' 
        '   Date:       Name:   Description:
        '	11/15/04	JEA     Created
        '   12/20/04    JEA     Get file path from config
        '   1/27/05     JEA     Use file property if given
        '-------------------------------------------------------------------------
        Private Sub Menu_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
            Dim xml As New AMP.Data.Xml
            If _fileName Is Nothing Then _fileName = AppSettings("MenuStore")
            _nodes = xml.GetNodes(_fileName, "siteMap/siteMapNode")
        End Sub

        '---COMMENT---------------------------------------------------------------
        '   render nodes as menu
        ' 
        '   Date:       Name:   Description:
        '	11/15/04	JEA     Created
        '-------------------------------------------------------------------------
        Protected Overrides Sub Render(ByVal writer As HtmlTextWriter)
            If Not _nodes Is Nothing Then
                Me.RenderChildNodes(_nodes, "menu", writer)
            End If
        End Sub

        '---COMMENT---------------------------------------------------------------
        '   recursively render out child nodes
        ' 
        '   Date:       Name:   Description:
        '	11/16/04	JEA     Created
        '-------------------------------------------------------------------------
        Private Sub RenderChildNodes(ByVal parentNode As XmlNodeList, ByVal parentID As String, ByVal writer As HtmlTextWriter)
            Dim x As Integer = 1
            Dim thisNode As String
            Dim root As Boolean = (parentID = "menu")

            With writer
                .Write("<div id=""")
                .Write(parentID)
                .Write(""" class=""menu"">")
                For Each node As XmlNode In parentNode
                    thisNode = ""
                    If node.HasChildNodes Then thisNode = parentID & "_" & x
                    .Write("<div class=""menuItem ")
                    If Not node.Attributes("newSection") Is Nothing Then
                        If Boolean.Parse(node.Attributes("newSection").Value) Then
                            .Write("section ")
                        End If
                    End If
                    .Write(IIf(root, "root", "sub"))
                    .Write("Item""")
                    If True OrElse node.HasChildNodes Then
                        'thisNode = parentID & "_" & x
                        .Write(" onmouseover=""Menu.On(this, '")
                        .Write(thisNode)
                        .Write("');"" onmouseout=""Menu.Off(this, '")
                        .Write(thisNode)
                        .Write("');""")
                    End If
                    If node.Attributes("url").Value <> String.Empty Then
                        ' make item clickable
                        .Write(" onclick=""location.href='")
                        .Write(HttpUtility.HtmlEncode(node.Attributes("url").Value.Replace("~", Global.BasePath)))
                        .Write("';""")
                    End If
                    .Write(">")
                    .Write(node.Attributes("title").Value)
                    .Write("</div>")

                    If node.HasChildNodes Then Me.RenderChildNodes(node.ChildNodes, thisNode, writer)
                    x += 1
                Next
                .Write("</div>")

            End With
        End Sub
    End Class
End Namespace