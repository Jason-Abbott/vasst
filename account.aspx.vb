Imports AMP.Site
Imports System.IO
Imports System.Configuration.ConfigurationSettings

Namespace Pages
    Public Class Account
        Inherits AMP.Page

#Region " Controls "

        Protected ampSectionList As AMP.Controls.EnumCheckbox
        Protected lblRole As Label
        ' name
        Protected tbFirstName As TextBox
        Protected tbLastName As TextBox
        Protected fldScreenName As AMP.Controls.Field
        ' credentials
        Protected fldEmail As AMP.Controls.Field
        Protected fldOldEmail As HtmlControls.HtmlInputHidden
        Protected fldConfirmationCode As AMP.Controls.Field
        Protected fldConfirmation As HtmlControls.HtmlInputHidden
        Protected fldPassword As AMP.Controls.Field
        Protected fldPrivacy As AMP.Controls.Field
        Protected fldPicture As AMP.Controls.Upload
        Protected fldWebSite As AMP.Controls.Field
        Protected tbDescription As TextBox

#End Region

        Public Sub New()
            Me.RequireAuthentication = True
        End Sub

        Private Sub Page_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Init
            Page.EnableViewState = True
            ampSectionList.Type = GetType(AMP.Site.Section)
            fldPicture.Folder = AppSettings("UserImageFolder")
            fldPicture.MaxFileSize = CType(AppSettings("MaxPhotoUploadKB"), Integer)
            fldPicture.AllowedTypes = AMP.File.Types.Image
        End Sub

        Private Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
            Me.Title = "Your Account"
            Me.StyleSheet.Add("account")
            Me.StyleSheet.Add("form")
            Me.ScriptFile.Add("account")
            Me.ScriptFile.Add("validation/common")
            Me.ScriptFile.Add("validation/account")
            Me.ScriptFile.Add("broker/common")
            Me.ScriptFile.Add("broker/account")
        End Sub

        '---COMMENT---------------------------------------------------------------
        '	load account data
        '
        '	Date:		Name:	Description:
        '	1/5/05  	JEA		Creation
        '-------------------------------------------------------------------------
        Private Sub Page_PreRender(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.PreRender 'Handles MyBase.Load
            Dim person As AMP.Person
            Dim key As String = "personID"

            If Me.IsPostBack Then
                ' save changes to person
                person = WebSite.Persons(ViewState.Item(key).ToString)
                If Me.Save(person) Then
                    Response.Redirect(String.Format("~/person.aspx?id={0}", person.ID), False)
                    Return
                End If
            End If

            If person Is Nothing Then
                ' load person from querystring
                If Request.QueryString("id") <> Nothing AndAlso _
                   Profile.User.HasPermission(Permission.EditAnyUser) Then
                    person = WebSite.Persons(Request.QueryString("id"))
                End If

                If person Is Nothing AndAlso _
                   Profile.User.HasPermission(Permission.EditMyself) Then
                    person = Profile.User
                End If
            End If

            If person Is Nothing Then Me.SendBack()

            ViewState.Add(key, person.ID)

            fldConfirmationCode.MaxLength = CType(AppSettings("ValidationCodeLength"), Integer)
            With person
                lblRole.Text = .Roles.ToString
                tbFirstName.Text = .FirstName
                tbLastName.Text = .LastName
                fldScreenName.Value = .NickName
                fldEmail.Value = .Email
                fldOldEmail.Value = .Email
                fldPrivacy.Checked = .PrivateEmail
                fldWebSite.Value &= .WebSite
                tbDescription.Text = .Description
                ampSectionList.Selected = .Section
            End With
        End Sub

        Private Function EmailChanged() As Boolean
            Return fldEmail.Value <> fldOldEmail.Value
        End Function

        Private Function PasswordChanged() As Boolean
            Return fldPassword.Value <> ""
        End Function

        '---COMMENT---------------------------------------------------------------
        '	save account changes
        '
        '	Date:		Name:	Description:
        '	1/5/05  	JEA		Creation
        '   2/28/05     JEA     Add valiation
        '-------------------------------------------------------------------------
        Private Function Save(ByVal person As AMP.Person) As Boolean
            If fldConfirmationCode.Value <> fldConfirmation.Value Then
                ' confirmation code doesn't match
                Profile.Message = Me.Say("Error.ConfirmationCode")
                Return False
            End If

            If EmailChanged() AndAlso WebSite.Persons.EmailUsed(fldEmail.Value) Then
                ' signing up with existing e-mail
                Profile.Message = Me.Say("Error.ExistingEmail")
                Return False
            End If

            With person
                If fldPicture.Uploaded Then
                    If .ImageFile <> Nothing Then
                        ' delete old picture
                        Dim picture As New FileInfo(String.Format("{0}{1}\{2}", _
                            HttpRuntime.AppDomainAppPath, AppSettings("UserImageFolder"), .ImageFile))
                        Try
                            If picture.Exists Then picture.Delete()
                        Catch e As UnauthorizedAccessException
                            Log.Error(e, Log.ErrorType.FileSystem, Profile.User)
                        End Try
                    End If
                    .ImageFile = fldPicture.File.Name
                End If
                .FirstName = tbFirstName.Text
                .LastName = tbLastName.Text
                .NickName = fldScreenName.Value
                .Email = fldEmail.Value
                .Description = tbDescription.Text
                .PrivateEmail = fldPrivacy.Checked
                .Section = ampSectionList.Selected
                .WebSite = fldWebSite.Value
                If Me.PasswordChanged Then .Password = Security.Encrypt(fldPassword.Value)
                Response.Cookies.Add(New HttpCookie("section", .Section.ToString))
            End With

            Profile.Message = Me.Say("Msg.AccountUpdated")
            WebSite.Save()
            Return True
        End Function
    End Class
End Namespace