Imports AMP.Site
Imports System.Net

Namespace Pages
    Public Class ResourceEdit
        Inherits AMP.Page

#Region " Controls "

        Protected ampSectionList As AMP.Controls.EnumList
        Protected fldTitle As AMP.Controls.Field
        Protected tbDescription As TextBox
        Protected fsSections As HtmlControls.HtmlContainerControl
        ' file specific
        Protected pnlFiles As Panel
        Protected fldFile As AMP.Controls.Upload
        Protected ampSoftware As AMP.Controls.SoftwareList
        Protected fldRenderedURL As AMP.Controls.Field
        Protected fldMediaURL As AMP.Controls.Field
        Protected lbCategories As AMP.Controls.ListBox
        Protected fsPlugins As HtmlControls.HtmlContainerControl
        Protected lbPlugins As AMP.controls.ListBox
        ' link specific
        Protected pnlLinks As Panel
        Protected fldLink As AMP.Controls.Field
        Protected fldScrape As AMP.Controls.Field
        Protected ampTypeList As AMP.Controls.EnumList

#End Region

        Private Sub Page_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Init
            ampSectionList.Type = GetType(AMP.Site.Section)
            fldFile.AllowedTypes = AMP.File.Types.Resource
            Me.RequireAuthentication = True
        End Sub

        Private Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
            Me.StyleSheet.Add("form")
            Me.StyleSheet.Add("asset")
            Me.ScriptFile.Add("broker/common")

            Dim asset As AMP.Asset

            If Request.QueryString("id") <> Nothing Then
                asset = WebSite.Assets(Request.QueryString("id"))
            End If

            If asset Is Nothing OrElse Not asset.CanEdit Then Me.SendBack()

            With asset
                If Me.IsPostBack Then
                    Dim success As Boolean = True

                    If (.Type And asset.Types.File) > 0 Then
                        ' file
                        .File.RenderedUrl = fldRenderedURL.Value
                        .File.RequiredUrl = fldMediaURL.Value
                        .File.Software = ampSoftware.Software
                        .File.SoftwareVersion = ampSoftware.Version

                        If fldFile.Uploaded Then
                            .File.Size = fldFile.File.Size
                            .File.Name = fldFile.File.Name
                            .VersionDate = DateTime.Now
                            .Version += 1
                            .File.Approve()
                        End If
                    Else
                        ' link
                        If .Link.Url <> fldLink.Value.Replace("http://", Nothing) Then
                            ' verify new link
                            Dim web As New WebClient
                            Try
                                Dim reponse As Byte() = web.DownloadData(fldLink.Value)
                                .Link.Url = fldLink.Value
                            Catch ex As System.Net.WebException
                                Profile.Message = Me.Say("Error.BadUrl")
                                success = False
                            End Try
                        End If
                        If success Then .Link.Scrape = fldScrape.Checked
                    End If

                    If success Then
                        .Title = fldTitle.Value
                        .Description = tbDescription.Text
                        Me.SaveClassification(asset)
                        WebSite.Assets.ClearSearchCache()
                        WebSite.Save()
                        Response.Redirect(String.Format("~/resource.aspx?id={0}", .ID), False)
                        Return
                    End If
                End If

                Dim user As AMP.Person = Profile.User

                ' display asset
                If (.Type And asset.Types.File) > 0 Then
                    ' file
                    pnlFiles.Visible = True

                    fldRenderedURL.Value &= .File.RenderedUrl
                    fldMediaURL.Value &= .File.RequiredUrl

                    ampSoftware.Software = .File.Software
                    ampSoftware.Version = .File.SoftwareVersion

                    Dim plugins As ArrayList = WebSite.Publishers.Plugins
                    If plugins.Count > 0 Then
                        lbPlugins.DataSource = plugins
                        lbPlugins.DataBind()
                        lbPlugins.Selected = .File.Plugins.IdArray
                    Else
                        lbPlugins.Visible = False
                        fsPlugins.Visible = False
                    End If
                    lbCategories.DataSource = WebSite.Categories.ForAssetType(.Type)
                Else
                    ' link
                    pnlLinks.Visible = True
                    fldLink.Value &= .Link.Url
                    fsPlugins.Visible = False
                    lbPlugins.Visible = False
                    ampTypeList.Type = GetType(AMP.Asset.Types)
                    ampTypeList.Selected = .Type
                    ampTypeList.Mask = asset.Types.Link
                    fldScrape.Checked = .Link.Scrape
                    fldScrape.Visible = user.HasPermission(Permission.EditAnyAsset)
                    lbCategories.DataSource = WebSite.Categories.ForAssetType(asset.Types.Link)
                End If

                Me.Title = .Title
                fldTitle.Value = .Title

                lbCategories.DataBind()
                lbCategories.Selected = .Categories.IdArray

                tbDescription.Text = .Description

                fsSections.Visible = user.HasPermission(Permission.EditAnyAsset)
                ampSectionList.Selected = .Section
            End With
        End Sub

        '---COMMENT---------------------------------------------------------------
        '	save asset classification
        '
        '	Date:		Name:	Description:
        '	2/11/05  	JEA		Copied from share.aspx.vb
        '-------------------------------------------------------------------------
        Private Function SaveClassification(ByVal asset As AMP.Asset) As Boolean
            With asset
                ' build category collection
                Dim cc As New AMP.CategoryCollection
                For Each c As String In lbCategories.Selected
                    cc.Add(WebSite.Categories(c))
                Next
                .Categories = cc

                ' add plugins
                If lbPlugins.Posted Then
                    Dim sc As New AMP.SoftwareCollection
                    For Each s As String In lbPlugins.Selected
                        sc.Add(WebSite.Publishers.SoftwareWithID(s))
                    Next
                    .File.Plugins = sc
                End If

                If ampTypeList.Posted Then
                    .Type = CType(ampTypeList.Selected, AMP.Asset.Types)
                End If

                If ampSectionList.Posted Then .Section = ampSectionList.Selected
            End With
        End Function
    End Class
End Namespace