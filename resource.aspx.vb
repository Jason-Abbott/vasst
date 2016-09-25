Imports AMP.Asset
Imports AMP.Site
Imports System.Configuration.ConfigurationSettings

Namespace Pages
    Public Class Resource
        Inherits AMP.Page

        Private _asset As AMP.Asset

#Region " Controls "

        Protected fldAction As HtmlControls.HtmlInputHidden
        ' text
        Protected tbComment As TextBox
        ' labels
        Protected lblType As Label
        Protected lblName As Label
        Protected lblUser As Label
        Protected lblDate As Label
        Protected lblDescription As Label
        Protected lblViews As Label
        ' databound
        Protected pnlRateIt As Panel
        Protected ddlRating As DropDownList
        Protected rptComments As Repeater
        Protected rptCategories As Repeater
        Protected rptPlugins As Repeater
        ' amp
        Protected ampRating As AMP.Controls.Rating
        Protected btnVideo As AMP.Controls.Button
        Protected btnView As AMP.Controls.Button
        Protected btnApprove As AMP.Controls.Button
        Protected btnEdit As AMP.Controls.Button
        Protected btnDelete As AMP.Controls.Button
        ' web parts
        Protected SearchResults As AMP.Controls.SearchResults

#End Region

        Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
            If Profile.SearchResults Is Nothing Then
                SearchResults.Visible = False
                Me.Template.WebParts.Controls.Add(Me.Page.LoadControl("~/control/NewAssets.ascx"))
            End If

            If Request.QueryString("id") <> Nothing Then
                _asset = WebSite.Assets(Request.QueryString("id"))
                If _asset Is Nothing Then
                    Me.SendBack()
                ElseIf Profile.ResumeDownload.Equals(_asset.ID) Then
                    MyBase.ScriptBlock = Common.JSRedirect(_asset.ViewURL)
                    Profile.ResumeDownload = Nothing
                End If
            End If
        End Sub

        Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
            ' http://www.eggheadcafe.com/articles/20030830.asp
            Dim user As AMP.Person = Profile.User

            With Me
                .StyleSheet.Add("asset")
                .ScriptFile.Add("common/validation")
                .StyleSheet.Add("form")
                .ScriptFile.Add("asset")
                .Title = Me.Say("Title.FreeResources")
            End With

            If Page.IsPostBack Then
                Select Case fldAction.Value
                    Case "delete"
                        If _asset.CanDelete Then
                            _asset.Delete()
                            Log.Activity(AMP.Site.Activity.DeleteAsset, user.ID)
                            Profile.Message = Me.Say("Msg.AssetDeleted")
                            Response.Redirect("~/default.aspx", True)
                        Else
                            Profile.Message = Me.Say("Error.Permissions")
                        End If
                    Case Else
                        ' store form in session to handle login redirection
                        Profile.FormValues = Request.Form()
                End Select
            End If

            If (Not Profile.FormValues Is Nothing) AndAlso _
                Profile.FormValues(tbComment.UniqueID) <> Nothing Then
                Me.SaveRating(_asset, user)
            End If

            With _asset
                lblType.Text = .FullType
                lblName.Text = .DisplayTitle
                lblUser.Text = "By " & .AuthoredBy.DetailLink
                lblDate.Text = String.Format("{0:MMMM d, yyyy}", .VersionDate)
                lblDescription.Text = .Description
                lblViews.Text = .Views
                lblViews.Visible = (.Views <> Nothing)

                ' actions
                btnView.Visible = True
                btnView.Url = .ViewURL
                btnView.Text = Me.Say(.ViewAction)

                If _asset.CanEdit Then
                    btnEdit.Visible = True
                    btnEdit.Url = String.Format("resource-edit.aspx?id={0}", _asset.ID)
                End If

                If .Status = AMP.Site.Status.Pending AndAlso _
                    user.HasPermission(Permission.ApproveAsset) Then
                    ' show approve/deny buttons
                    btnApprove.Visible = True
                    btnApprove.OnClick = "Asset.Approve();"
                    btnDelete.Visible = True
                    btnDelete.OnClick = "Asset.Deny();"
                    btnDelete.Text = Me.Say("Action.DenyAsset")
                ElseIf _asset.CanDelete Then
                    btnDelete.Visible = True
                    btnDelete.OnClick = "Asset.Delete();"
                End If

                If (.Type And Types.File) > 0 AndAlso .File.RenderedUrl <> Nothing Then
                    ' display link to rendered video
                    btnVideo.Visible = True
                    btnVideo.Url = .File.RenderedLink
                End If

                ' categories
                If .Categories.Count > 0 Then
                    rptCategories.DataSource = .Categories
                    rptCategories.DataBind()
                Else
                    rptCategories.Visible = False
                End If

                ' plugins
                If .Plugins.Count > 0 Then
                    rptPlugins.DataSource = .Plugins
                    rptPlugins.DataBind()
                Else
                    rptPlugins.Visible = False
                End If

                ' new rating fields
                If Not (.AuthoredBy.ID.Equals(user.ID) OrElse .SubmittedBy.ID.Equals(user.ID)) Then
                    pnlRateIt.Visible = True
                    For x As Integer = 1 To 5
                        ddlRating.Items.Add(x.ToString)
                    Next
                    ddlRating.SelectedIndex = 4

                    Dim rating As AMP.Rating = .Ratings(Profile.User)
                    If Not rating Is Nothing Then tbComment.Text = rating.Comment
                Else
                    pnlRateIt.Visible = False
                End If

                ' existing ratings
                If .Ratings.Count > 0 Then
                    rptComments.DataSource = .Ratings
                    rptComments.DataBind()
                    ampRating.Rating = .Ratings.Average
                    ampRating.Visible = True
                Else
                    rptComments.Visible = False
                End If
            End With
        End Sub

        '---COMMENT---------------------------------------------------------------
        '	save comment and rating
        '
        '	Date:		Name:	Description:
        '	1/19/05	    JEA		Creation
        '-------------------------------------------------------------------------
        Private Sub SaveRating(ByVal asset As AMP.Asset, ByVal user As AMP.Person)
            If Profile.Authenticated Then
                ' save comment and rating
                Dim exists As Boolean = True
                Dim rating As AMP.Rating = asset.Ratings.ByPerson(user)

                If rating Is Nothing Then
                    rating = New AMP.Rating
                    exists = False
                End If

                With rating
                    .Person = user
                    .Date = DateTime.Now
                    .Value = CSng(Profile.FormValues(ddlRating.UniqueID))
                    .Comment = Profile.FormValues(tbComment.UniqueID)
                End With
                Profile.FormValues = Nothing

                If Not exists Then asset.Ratings.Add(rating)
            Else
                Me.SendToLogin()
            End If
        End Sub
    End Class
End Namespace