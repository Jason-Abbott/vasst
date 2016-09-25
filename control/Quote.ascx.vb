Imports System.Xml
Imports System.Web.UI
Imports System.Web.Caching
Imports System.Web.HttpContext
Imports System.Configuration.ConfigurationSettings

Namespace Controls
    Public Class Quote
        Inherits System.Web.UI.UserControl

        Private _quote As String
        Private _authorName As String
        Private _authorFrom As String
        Private _date As String

        Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
            Me.GetQuote()
        End Sub

        '---COMMENT---------------------------------------------------------------
        '	fill control with content
        '
        '	Date:		Name:	Description:
        '	11/14/04	JEA		Creation
        '   12/1/04     JEA     Use node list instead of dataset
        '-------------------------------------------------------------------------
        Public Sub GetQuote()
            Dim filename As String = AppSettings("QuotesStore")
            Dim nodes As XmlNodeList
            Dim cacheKey As String = filename

            nodes = DirectCast(Current.Cache.Item(cacheKey), XmlNodeList)

            If nodes Is Nothing Then
                ' retrieve content
                Dim xml As New AMP.Data.Xml
                Dim fullPath As String = Current.Request.MapPath(filename)
                If System.IO.File.Exists(fullPath) Then
                    nodes = xml.GetNodes(filename, "quotes/quote")

                    ' put content in cache
                    Dim dependsOn As CacheDependency = New CacheDependency(fullPath)
                    Current.Cache.Insert(cacheKey, nodes, dependsOn)
                End If
            End If

            If nodes.Count > 0 Then
                Dim random As New System.Random
                Dim index As Integer = random.Next(0, nodes.Count)
                random = Nothing

                With nodes(index)
                    _quote = .InnerText
                    If Not .Attributes("saidBy") Is Nothing Then
                        _authorName = .Attributes("saidBy").Value
                    End If
                    If Not .Attributes("from") Is Nothing Then
                        _authorFrom = .Attributes("from").Value
                    End If
                    If Not .Attributes("saidOn") Is Nothing Then
                        _date = .Attributes("saidOn").Value
                    End If
                End With
            End If
        End Sub

        '---COMMENT---------------------------------------------------------------
        '	render quote within control
        '
        '	Date:		Name:	Description:
        '	11/14/04	JEA		Creation
        '   12/1/04     JEA     Add other attributes
        '-------------------------------------------------------------------------
        Protected Overrides Sub Render(ByVal writer As System.Web.UI.HtmlTextWriter)
            If _quote <> Nothing Then
                Dim image As New AMP.Controls.Image

                writer.Write("<div class=""quote"">")
                With image
                    .Transparency = True
                    .CssClass = "open"
                    .Src = "~/images/quote-open.png"
                    .RenderControl(writer)
                    .CssClass = "close"
                    .Src = "~/images/quote-close.png"
                    .RenderControl(writer)
                End With
                With writer
                    .Write("<div class=""border""><div class=""body"">")
                    .Write(_quote)
                    If _authorName <> Nothing Then
                        .Write("<div class=""author"">&ndash; ")
                        .Write(_authorName)
                        If _authorFrom <> Nothing Then
                            .Write("<div class=""from"">")
                            .Write(_authorFrom)
                            .Write("</div>")
                        End If
                        .Write("</div>")
                    End If
                    .Write("</div></div></div>")
                End With
            End If
        End Sub

    End Class
End Namespace
