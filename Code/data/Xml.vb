Imports System.xml
Imports System.Web.Caching
Imports System.Web.HttpContext

Namespace Data
    Public Class Xml
        Private _cache As Boolean = True

        Public WriteOnly Property Cache() As Boolean
            Set(ByVal Value As Boolean)
                _cache = Value
            End Set
        End Property

        '---COMMENT---------------------------------------------------------------
        '   load node list from cache or file system
        ' 
        '   Date:       Name:   Description:
        '	11/16/04	JEA     Created
        '-------------------------------------------------------------------------
        Public Function GetNodes(ByVal fileName As String, ByVal xPath As String) As XmlNodeList
            Dim nodes As XmlNodeList
            Dim cacheKey As String = fileName

            nodes = DirectCast(Current.Cache.Item(cacheKey), XmlNodeList)

            If nodes Is Nothing And _cache Then
                Dim xml As New XmlDocument
                fileName = Current.Server.MapPath(fileName)
                xml.Load(fileName)
                nodes = xml.SelectNodes(xPath)

                ' put content in cache
                Dim dependsOn As CacheDependency = New CacheDependency(fileName)
                Current.Cache.Insert(cacheKey, nodes, dependsOn)
            End If

            Return nodes

        End Function
    End Class
End Namespace