Imports AMP.Common
Imports System.Text.RegularExpressions
Imports System.Configuration.ConfigurationSettings

Namespace Pages
    Public Class Search
        Inherits AMP.Page

        Private _searchText As String

#Region " Controls "

        Protected tbText As TextBox
        Protected rptResults As Repeater
        Protected pnlTip As Panel
        Protected lblMatches As Label
        Protected lblCount As Label
        Protected pnlSearch As Panel
        Protected btnNext As AMP.Controls.Button
        Protected btnPrevious As AMP.Controls.Button

#End Region

        '---COMMENT---------------------------------------------------------------
        '	load appropriate search results
        '
        '	Date:		Name:	Description:
        '	11/2/04	    JEA		Creation
        '   1/5/05      JEA     Get search text from query string instead of form
        '   3/1/05      JEA     Support paged results
        '-------------------------------------------------------------------------
        Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
            Dim matches As ArrayList
            Dim page As Integer = 1
            Dim pageSize As Integer = CType(AppSettings("SearchPageSize"), Integer)
            Dim type As Asset.Types
            Dim category As AMP.Category
            Dim typeName As String
            Dim sortBy As AssetCollection.SortBy
            Dim sortDirection As AssetCollection.SortDirection = AssetCollection.SortDirection.Descending
            Dim filter As String
            Dim total As Integer

            With Me
                .StyleSheet.Add("list")
                .StyleSheet.Add("form")
                .Title = Me.Say("Title.FreeResources")
            End With

            ' get desired page
            If Request.QueryString("page") <> Nothing Then
                page = CInt(Request.QueryString("page"))
            End If

            ' get results per page
            If Request.QueryString("per") <> Nothing Then
                pageSize = CInt(Request.QueryString("per"))
            End If

            ' get type (project, script, review, etc.)
            If Request.QueryString("type") <> Nothing Then
                type = CType(Integer.Parse(Request.QueryString("type")), Asset.Types)
                typeName = type.ToString
            Else
                typeName = "Resource"
            End If

            ' get category
            If Request.QueryString("category") <> Nothing Then
                category = WebSite.Categories.WithName(Request.QueryString("category"))
            End If

            ' get search text
            _searchText = Request.QueryString("text")

            ' get sort preference
            If Request.QueryString("sort") <> Nothing Then
                Select Case Request.QueryString("sort")
                    Case "date"
                        sortBy = AssetCollection.SortBy.VersionDate
                        filter = "Newest"
                    Case "popularity"
                        sortBy = AssetCollection.SortBy.Popularity
                        filter = "Most Popular"
                    Case "rank"
                        sortBy = AssetCollection.SortBy.Rank
                        filter = "Best Ranking"
                    Case "name"
                        sortBy = AssetCollection.SortBy.Name
                        sortDirection = AssetCollection.SortDirection.Ascending
                    Case "author"
                        sortBy = AssetCollection.SortBy.Author
                        sortDirection = AssetCollection.SortDirection.Ascending
                    Case "description"
                        sortBy = AssetCollection.SortBy.Description
                        sortDirection = AssetCollection.SortDirection.Ascending
                    Case Else
                        sortBy = AssetCollection.SortBy.SubmitDate
                End Select
            Else
                sortBy = AssetCollection.SortBy.SubmitDate
            End If

            ' get matching results
            matches = WebSite.Assets.Search(page, pageSize, _searchText, Nothing, _
                category, type, sortBy, sortDirection, total)

            ' build ui
            If matches.Count = 1 Then
                ' for single match send directly to detail page
                Dim asset As AMP.Asset = DirectCast(matches(0), AMP.Asset)
                Server.Transfer(String.Format("~/resource.aspx?id={0}", asset.ID), False)

            ElseIf matches.Count > 0 Then
                ' display list of matches
                Dim text As String
                Dim flexibility As Integer = CType(AppSettings("PageSizeFlexible"), Integer)
                Dim startIndex As Integer = ((page - 1) * pageSize) + 1
                Dim endIndex As Integer = (startIndex + pageSize) - 1

                If endIndex > total Then endIndex = total
                If _searchText <> Nothing Then text = String.Format("with ""{0}""", _searchText)

                Profile.SearchResults = matches
                rptResults.DataSource = matches
                rptResults.DataBind()
                rptResults.Visible = True

                lblMatches.Text = String.Format("{0} {1}s {2}", filter, typeName, text)
                lblCount.Visible = True

                If total > pageSize + flexibility Then
                    ' display paged list
                    Dim url As String = Request.Url.PathAndQuery

                    lblCount.Text = String.Format("{0}&mdash;{1} of {2}", _
                        startIndex, endIndex, total)

                    ' show tip on second page only
                    If page = 2 Then pnlTip.Visible = True

                    ' buttons
                    If page > 1 Then
                        Dim previousPage As String = String.Format("page={0}", page - 1)
                        If url.IndexOf("page") = -1 Then
                            btnPrevious.Url = String.Format("{0}&{1}", url, previousPage)
                        Else
                            Dim re As New Regex("page=\d+")
                            btnPrevious.Url = re.Replace(url, previousPage)
                        End If
                        btnPrevious.Visible = True
                    End If
                    If endIndex < total Then
                        Dim nextPage As String = String.Format("page={0}", page + 1)
                        If url.IndexOf("page") = -1 Then
                            btnNext.Url = String.Format("{0}&{1}", url, nextPage)
                        Else
                            Dim re As New Regex("page=\d+")
                            btnNext.Url = re.Replace(url, nextPage)
                        End If
                        btnNext.Visible = True
                    End If


                    'If matches.Count < pageSize Then
                    '    Context.Items.Add("search", matches)
                    'Else
                    '    ' show paging
                    'End If
                Else
                    ' display list with no paging
                    lblCount.Text = String.Format("{0} matches", Format.WordForNumber(total))

                End If
            Else
                ' no matches
                ' TODO: send to search help page
                Profile.SearchResults = Nothing
                lblMatches.Text = "No Resources Found"
            End If
            lblMatches.Visible = True
        End Sub

#Region " Build Links "

        Protected Function ViewLink(ByVal data As Object) As String
            Dim asset As AMP.Asset = DirectCast(data, AMP.Asset)
            Return String.Format("location.href='{0}'", asset.ViewURL)
        End Function

        Protected Function EditLink(ByVal data As Object) As String
            Dim asset As AMP.Asset = DirectCast(data, AMP.Asset)
            Return String.Format("location.href='resource-edit.aspx?id={0}'", _
                asset.ID.ToString)
        End Function

#End Region

        '---COMMENT---------------------------------------------------------------
        '	encapsulate search term in tags for highlighting
        '
        '	Date:		Name:	Description:
        '	11/12/04    JEA		Creation
        '   12/8/04     JEA     Use regular expressions
        '-------------------------------------------------------------------------
        Protected Function Highlight(ByVal text As String) As String
            If _searchText <> Nothing Then
                Dim regex As New Regex("(" & _searchText & ")", RegexOptions.IgnoreCase)
                Return regex.Replace(text, "<span class=""highlight"">$1</span>")
            Else
                Return text
            End If
        End Function

    End Class
End Namespace
