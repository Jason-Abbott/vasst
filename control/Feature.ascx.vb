Imports System.Xml
Imports System.Web.UI
Imports System.Web.Caching
Imports System.Web.HttpContext
Imports System.Configuration.ConfigurationSettings

Namespace Controls
    Public Class Feature
        Inherits System.Web.UI.UserControl

        Private _content As String
        Private _image As String
        Private _link As String
        Private _startDate As DateTime
        Private _endDate As DateTime

        Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
            Me.GetFeature()
        End Sub

        '---COMMENT---------------------------------------------------------------
        '	fill control with content
        '
        '	Date:		Name:	Description:
        '	12/17/04	JEA		Creation
        '-------------------------------------------------------------------------
        Public Sub GetFeature()
            Dim filename As String = AppSettings("FeaturesStore")
            Dim nodes As XmlNodeList
            Dim cacheKey As String = filename

            nodes = DirectCast(Current.Cache.Item(cacheKey), XmlNodeList)

            If nodes Is Nothing Then
                ' retrieve content
                Dim xml As New AMP.Data.Xml
                Dim fullPath As String = Current.Request.MapPath(filename)
                If System.IO.File.Exists(fullPath) Then
                    nodes = xml.GetNodes(filename, "features/feature")

                    ' put content in cache
                    Dim dependsOn As CacheDependency = New CacheDependency(fullPath)
                    Current.Cache.Insert(cacheKey, nodes, dependsOn)
                End If
            End If

            If nodes.Count > 0 Then
                Dim node As XmlNode
                Dim random As New System.Random
                Dim index As Integer = random.Next(0, nodes.Count)
                random = Nothing

                With nodes(index)
                    node = .SelectSingleNode("href")
                    If Not node Is Nothing Then _link = node.InnerText

                    node = .SelectSingleNode("image")
                    If Not node Is Nothing Then _image = node.InnerText

                    node = .SelectSingleNode("text")
                    If Not node Is Nothing Then _content = node.InnerText

                    If Not .Attributes("StartDate") Is Nothing Then
                        _startDate = Date.Parse(.Attributes("StartDate").Value)
                    End If
                    If Not .Attributes("EndDate") Is Nothing Then
                        _endDate = Date.Parse(.Attributes("EndDate").Value)
                    End If
                End With
            End If
        End Sub

        '---COMMENT---------------------------------------------------------------
        '	render content within control
        '
        '	Date:		Name:	Description:
        '	12/17/04	JEA		Creation
        '-------------------------------------------------------------------------
        Protected Overrides Sub Render(ByVal writer As System.Web.UI.HtmlTextWriter)
            If _content <> Nothing OrElse _image <> Nothing Then
                With writer
                    .Write("<div id=""feature"" class=""topRight"">")
                    If _image <> Nothing Then
                        If _link <> Nothing Then
                            .Write("<a href=""")
                            .Write(_link)
                            .Write(""">")
                        End If
                        .Write("<img class=""feature"" src=""")
                        .Write(Global.BasePath)
                        .Write("/images/feature/")
                        .Write(_image)
                        .Write(""" alt="""" />")
                        If _link <> Nothing Then .Write("</a>")
                        .Write("<br/>")
                    End If
                    If _content <> Nothing Then
                        .Write(_content)
                    End If
                    .Write("</div>")
                End With
            End If
        End Sub

    End Class
End Namespace
