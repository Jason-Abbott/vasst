Imports AMP.Site
Imports System.IO
Imports System.Net
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Configuration.ConfigurationSettings

Namespace Pages
    Public Class Share
        Inherits AMP.Page

#Region " Controls "

        Protected h2Step As HtmlControls.HtmlContainerControl

        Protected fsStep1 As HtmlControls.HtmlContainerControl
        Protected fsStep2 As HtmlControls.HtmlContainerControl
        Protected fsStep3 As HtmlControls.HtmlContainerControl
        Protected ampContent As AMP.Controls.ContentWebPart

        ' buttons
        Protected btnPrevious As AMP.Controls.Button
        Protected btnNext As AMP.Controls.Button

        ' step 1
        Protected fldUpload As AMP.Controls.Upload
        Protected fldWebSite As AMP.Controls.Field
        Protected fldTerms As AMP.controls.Field

        ' step 2
        Protected fldTitle As AMP.Controls.Field
        Protected ampSoftware As AMP.Controls.SoftwareList
        Protected fldActivate As AMP.Controls.Field
        Protected fldRenderedURL As AMP.Controls.Field
        Protected fldMediaURL As AMP.Controls.Field
        Protected tbDescription As TextBox
        Protected fsOptionalLinks As HtmlControls.HtmlContainerControl
        Protected pnlSoftware As Panel

        ' step 3
        Protected ampTypeList As AMP.Controls.EnumList
        Protected lbCategories As AMP.Controls.ListBox
        Protected lbPlugins As AMP.Controls.ListBox
        Protected ampSectionList As AMP.Controls.EnumList
        Protected fsPlugins As HtmlControls.HtmlContainerControl
        Protected pnlTypes As Panel
        Protected fsSections As HtmlControls.HtmlContainerControl

        ' step 4
        Protected ampSignInUp As AMP.Controls.SignInUp

#End Region

        Private Const _http As String = "http://"
        Private _asset As AMP.Asset

        '---COMMENT---------------------------------------------------------------
        '	process wizard steps
        '
        '	Date:		Name:	Description:
        '	12/22/04  	JEA		Creation
        '   2/11/05     JEA     Resume paused contribution
        '-------------------------------------------------------------------------
        Private Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
            Dim totalSteps As Integer = CInt(IIf(Profile.Authenticated, 3, 4))
            Dim posted As Boolean = False
            Dim lastStep As Integer = 0
            Dim contribution As AMP.AssetContribution = Profile.Contribution
            _asset = contribution.Asset

            Me.Title = "Share a Resource"
            Me.StyleSheet.Add("form")
            Me.StyleSheet.Add("share")

            ampContent.Title = Me.Say("Title.Tips")

            If Page.IsPostBack Then
                lastStep = CInt(Request.Form("step"))
                posted = True
            ElseIf Request.QueryString("step") <> Nothing Then
                lastStep = CInt(Request.QueryString("step")) - 1
            Else
                lastStep = contribution.FinishedStep
            End If

            If lastStep > contribution.FinishedStep Then contribution.FinishedStep = lastStep

            btnPrevious.Url = String.Format("share.aspx?step={0}", lastStep)

            ' TODO: check when contribution was started if resuming

            Select Case lastStep
                Case 0
                    ' start page
                    Me.FirstStep()
                Case 1
                    ' file or link submitted
                    If posted Then
                        Dim success As Boolean = False

                        If fldUpload.Uploaded Then
                            success = Me.SaveFile(fldUpload.File, fldUpload.Content)
                        Else
                            success = Me.LinkSubmit
                        End If

                        If Not success Then
                            ' show step 1 again
                            Me.FirstStep()
                            Return
                        End If
                    End If

                    If _asset.Link Is Nothing Then
                        ' file
                        ampSoftware.Software = _asset.File.Software
                        ampSoftware.Version = _asset.File.SoftwareVersion
                    Else
                        ' link
                        fsOptionalLinks.Visible = False
                        pnlSoftware.Visible = False
                    End If

                    ampContent.File = "help/ShareStep2.txt"

                    ' status
                    If Profile.User.HasPermission(Permission.ApproveAsset) Then
                        fldActivate.Visible = True
                        fldActivate.Checked = True
                    End If

                    fldTitle.Value = _asset.Title
                    tbDescription.Text = _asset.Description

                    fsStep2.Visible = True
                    btnNext.Visible = True
                    btnPrevious.Visible = True

                Case 2
                    ' description entered
                    Dim file As Boolean = Not _asset.File Is Nothing
                    Dim type As Integer = CInt(IIf(file, Asset.Types.File, Asset.Types.Link))

                    If posted Then
                        Me.SaveDetail()
                        _asset.InferType()
                        _asset.InferSection()
                    End If

                    ' categories
                    lbCategories.DataSource = WebSite.Categories.ForAssetType(type)
                    lbCategories.DataBind()

                    ' plugins
                    If file Then
                        Dim plugins As ArrayList = WebSite.Publishers.Plugins
                        If plugins.Count > 0 Then
                            lbPlugins.DataSource = plugins
                            lbPlugins.DataBind()
                            fsPlugins.Visible = True
                        End If
                    End If

                    ' type
                    If _asset.Type = Nothing Then
                        ampTypeList.Type = GetType(AMP.Asset.Types)
                        ampTypeList.Mask = type
                        pnlTypes.Visible = True
                    End If

                    ' section
                    If Profile.User.HasPermission(Permission.ChangeAssetSection) OrElse _
                        _asset.Section = 0 Then

                        ampSectionList.Selected = _asset.Section
                        ampSectionList.Type = GetType(AMP.Site.Section)
                        fsSections.Visible = True
                    End If

                    ' page setup
                    fsStep3.Visible = True
                    If totalSteps = 3 Then
                        btnNext.Text = Me.Say("Action.Finish")
                    Else
                        ' login will be displayed so write test cookie
                        Profile.WriteTestCookie()
                    End If

                    ampContent.File = "help/ShareStep3.txt"
                    btnNext.Visible = True
                    btnPrevious.Visible = True

                Case 3
                    ' classification entered
                    If posted Then
                        Me.SaveClassification()
                        If totalSteps = 3 Then Me.SaveAsset()
                    End If

                    ampSignInUp.Visible = True
                    btnNext.Text = Me.Say("Action.Finish")
                    btnNext.Visible = True
                    btnPrevious.Visible = True
                    ampContent.File = "help/SignInOrUp.html"
            End Select

            h2Step.InnerText = String.Format("Step {0} of {1}", lastStep + 1, totalSteps)

            Profile.Contribution = contribution
        End Sub

        '---COMMENT---------------------------------------------------------------
        '	setup page for first step
        '
        '	Date:		Name:	Description:
        '	2/11/05  	JEA		Grouped logic
        '   2/24/05     JEA     Setup fields
        '-------------------------------------------------------------------------
        Private Sub FirstStep()
            ampContent.File = "help/ShareStep1.txt"
            Me.ScriptFile.Add("validation/share")
            Profile.Contribution.FinishedStep = 0
            fldTerms.Label = "I accept the <a style=""cursor: pointer;"" href=""#"" onclick=""javascript:DOM.Show('TermsConditions');"">Terms & Conditions</a> for uploaded files"
            fldUpload.MaxFileSize = CInt(AppSettings("MaxFileUploadKB"))
            fsStep1.Visible = True
            btnNext.Visible = True
        End Sub

        '---COMMENT---------------------------------------------------------------
        '	handle login step in pre-render so control has time to load
        '
        '	Date:		Name:	Description:
        '	1/10/05  	JEA		Creation
        '-------------------------------------------------------------------------
        Private Sub Page_PreRender(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.PreRender
            If Page.IsPostBack AndAlso Profile.Contribution.FinishedStep = 4 Then
                ' attempted login or registration
                If ampSignInUp.Succeeded Then
                    ' update user reference
                    Profile.Contribution.Asset.SubmittedBy = Profile.User
                    Me.SaveAsset()
                Else
                    ampSignInUp.Visible = True
                    btnNext.Text = "Finish"
                    btnNext.Visible = True
                End If
            End If
        End Sub

        '---COMMENT---------------------------------------------------------------
        '	copy asset from user profile to central store and redirect
        '
        '	Date:		Name:	Description:
        '	1/10/05  	JEA		Creation
        '   2/16/05     JEA     Add submit date and alternate message
        '-------------------------------------------------------------------------
        Private Sub SaveAsset()
            Dim message As String = "ShareThanks"

            With Profile.Contribution
                .Asset.SubmittedBy = Profile.User
                .Asset.SubmitDate = Date.Now
                If .Asset.AuthoredBy Is Nothing Then .Asset.AuthoredBy = Profile.User
                If (.Asset.Status And Status.Approved) > 0 Then message = "AssetSaved"
                If .Save() Then
                    Profile.Contribution = Nothing
                    Profile.Message = DirectCast(Me.Page, AMP.Page).Say("Msg." & message)
                    Response.Redirect("default.aspx", True)
                End If
            End With
        End Sub

        '---COMMENT---------------------------------------------------------------
        '	save asset classification
        '
        '	Date:		Name:	Description:
        '	1/5/05  	JEA		Creation
        '-------------------------------------------------------------------------
        Private Function SaveClassification() As Boolean
            With _asset
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

        '---COMMENT---------------------------------------------------------------
        '	save asset details
        '
        '	Date:		Name:	Description:
        '	12/30/04  	JEA		Creation
        '   1/5/05      JEA     Retrieval logic moved to control
        '   2/16/05     JEA     Set status
        '-------------------------------------------------------------------------
        Private Function SaveDetail() As Boolean
            With _asset
                .Title = fldTitle.Value
                .Description = tbDescription.Text
                If fldActivate.Checked AndAlso _
                    Profile.User.HasPermission(Permission.ApproveAsset) Then .Approve()

                If Not .File Is Nothing Then
                    .File.Software = ampSoftware.Software
                    .File.SoftwareVersion = ampSoftware.Version
                    .File.RenderedUrl = fldRenderedURL.Value
                    .File.RequiredUrl = fldMediaURL.Value
                End If
            End With
        End Function

        '---COMMENT---------------------------------------------------------------
        '	process posted file
        '
        '	Date:		Name:	Description:
        '	12/23/04  	JEA		Creation
        '   12/29/04    JEA     Infer software type
        '   1/11/05     JEA     Read Vegas file content for name and description
        '   2/11/05     JEA     Check if file could be saved
        '   2/24/05     JEA     Use upload control
        '-------------------------------------------------------------------------
        Private Function SaveFile(ByVal file As AMP.File, ByVal content As String()) As Boolean
            Dim title As String
            Dim description As String

            If file.Type = file.Types.Vegas AndAlso Not content Is Nothing Then
                ' attempt to get description from file content
                Dim index As Integer = Array.IndexOf(content, "ICMT^")
                If index >= 0 Then description = content(index + 1)
                index = Array.IndexOf(content, "INAM")
                If index >= 0 Then title = content(index + 1)
                If title = "IART" Then title = Nothing
            End If

            With _asset
                .File = file
                .SubmitDate = Now
                .SubmittedBy = Profile.User
                .Status = AMP.Site.Status.Pending
                .InferTitle(title)
                .Description = description
            End With

            Profile.Contribution.FinishedStep = 1
            Return True
        End Function

        '---COMMENT---------------------------------------------------------------
        '	process submitted link
        '
        '	Date:		Name:	Description:
        '	12/23/04  	JEA		Creation
        '   1/11/05     JEA     Use webclient to get link detail
        '-------------------------------------------------------------------------
        Private Function LinkSubmit() As Boolean
            If fldWebSite.Value <> "" Then
                Dim web As New WebClient
                Dim response As String
                Dim title As String

                Try
                    response = Encoding.Default.GetString(web.DownloadData(fldWebSite.Value))
                Catch e As System.Net.WebException
                    Profile.Message = DirectCast(Me.Page, AMP.Page).Say("Error.BadUrl")
                    Return False
                End Try

                ' use page title as asset name
                Dim re As New Regex("<title>([\s\S]*)<\/title>", RegexOptions.IgnoreCase Or RegexOptions.Multiline)
                Dim m As Match = re.Match(response)
                If m.Success Then title = m.Groups(1).Value.Replace(Environment.NewLine, "")

                Dim link As New AMP.Link
                link.Url = fldWebSite.Value

                With _asset
                    .Link = link
                    .SubmitDate = Now
                    .SubmittedBy = Profile.User
                    .Status = AMP.Site.Status.Pending
                    .InferTitle(title)
                End With

                Profile.Contribution.FinishedStep = 1
                Return True
            End If
            Return False
        End Function
    End Class
End Namespace
