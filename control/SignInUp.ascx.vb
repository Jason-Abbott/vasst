Imports System.IO
Imports System.Net
Imports System.Configuration.ConfigurationSettings

Namespace Controls
    Public Class SignInUp
        Inherits System.Web.UI.UserControl

#Region " Controls "

        Protected lgdTitle As HtmlContainerControl
        Protected fldClientTime As HtmlControls.HtmlInputHidden
        Protected fldServerTime As HtmlControls.HtmlInputHidden
        Protected fldSignup As HtmlControls.HtmlInputHidden
        Protected fldTipType As HtmlControls.HtmlInputHidden
        Protected fldConfirmation As HtmlControls.HtmlInputHidden
        Protected lblWarning As Label

        ' login
        Protected fldEmail As AMP.Controls.Field
        Protected fldPassword As AMP.Controls.Field

        ' register
        Protected tbFirstName As TextBox
        Protected tbLastName As TextBox
        Protected fldScreenName As AMP.Controls.Field
        Protected fldConfirmationCode As AMP.Controls.Field
        Protected fldPrivacy As AMP.Controls.Field
        Protected fldWebSite As AMP.Controls.Field

#End Region

        Private _redirect As Boolean = True
        Private _succeeded As Boolean = False
        Private _title As String = "Sign in or Sign up"
        Private _hostPage As AMP.Page

#Region " Properties "

        Public Property TipType() As String
            Get
                Return fldTipType.Value
            End Get
            Set(ByVal Value As String)
                fldTipType.Value = Value
            End Set
        End Property

        Public Property Title() As String
            Get
                Return _title
            End Get
            Set(ByVal Value As String)
                _title = Value
            End Set
        End Property

        Public Property Redirect() As Boolean
            Get
                Return _redirect
            End Get
            Set(ByVal Value As Boolean)
                _redirect = Value
            End Set
        End Property

        Public Property Succeeded() As Boolean
            Get
                Return _succeeded
            End Get
            Set(ByVal Value As Boolean)
                _succeeded = Value
            End Set
        End Property

#End Region

        '---COMMENT---------------------------------------------------------------
        '	process login or registration
        '
        '	Date:		Name:	Description:
        '	1/5/05  	JEA		Creation
        '   1/24/05     JEA     Add cookie check
        '-------------------------------------------------------------------------
        Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
            _hostPage = DirectCast(Page, AMP.Page)

            If Me.Visible Then
                'If Not Profile.SupportsCookies Then
                '    Response.Redirect(AppSettings("NoCookiePage"), False)
                'End If

                _hostPage.StyleSheet.Add("signInUp")
                _hostPage.ScriptFile.Add("signInUp")
                _hostPage.ScriptFile.Add("broker/common")
                _hostPage.ScriptFile.Add("broker/account")
                _hostPage.ScriptFile.Add("validation/account")
                fldServerTime.Value = DateTime.Now.ToString

                lblWarning.Visible = (Request.Browser.Browser = "Opera" AndAlso _
                    Request.Browser.MajorVersion < 8)
            End If

            If Page.IsPostBack AndAlso fldEmail.Value <> Nothing Then
                ' process form
                If fldSignup.Value = "true" Then
                    ' registration

                    If fldConfirmationCode.Value <> fldConfirmation.Value Then
                        ' confirmation code doesn't match
                        Profile.Message = _hostPage.Say("Error.ConfirmationCode")
                        Return
                    End If

                    If WebSite.Persons.EmailUsed(fldEmail.Value) Then
                        ' signing up with existing e-mail
                        Profile.Message = _hostPage.Say("Error.ExistingEmail")
                        Return
                    End If

                    If _hostPage.Security.Register(tbFirstName.Text, tbLastName.Text, _
                        fldScreenName.Value, fldEmail.Value, fldPassword.Value, _
                        fldWebSite.Value, fldPrivacy.Checked) Then

                        Profile.TimeOffset = Me.ClientTimeOffset
                        Me.Succeeded = True
                    Else
                        Profile.Message = _hostPage.Say("Error.Registering")
                    End If
                Else
                    ' login
                    If _hostPage.Security.Authenticate(fldEmail.Value, fldPassword.Value) Then
                        Profile.TimeOffset = Me.ClientTimeOffset
                        Me.Succeeded = True
                    Else
                        Profile.Message = _hostPage.Say("Error.BadCredentials")
                    End If
                End If

                If Me.Succeeded AndAlso Me.Redirect Then
                    Dim sendTo As String = Profile.DestinationPage
                    If sendTo Is Nothing Then sendTo = "default.aspx"
                    Profile.Message = String.Format("{0} {1}", _
                        _hostPage.Say("Msg.Signin"), Profile.User.DisplayName)
                    Response.Redirect(sendTo, True)
                End If

            ElseIf Request.QueryString("action") = "out" AndAlso Profile.Authenticated Then
                _hostPage.Profile.Clear()
                Profile.Message = _hostPage.Say("Msg.Signout")
            End If
        End Sub

        Private Sub Page_PreRender(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.PreRender
            lgdTitle.InnerText = _title
            fldConfirmationCode.MaxLength = CInt(AppSettings("ValidationCodeLength"))
        End Sub

        '---COMMENT---------------------------------------------------------------
        '	get the time difference between what js wrote and server wrote
        '
        '	Date:		Name:	Description:
        '	1/5/05  	JEA		Creation
        '   2/28/05     JEA     Try/Catch block
        '-------------------------------------------------------------------------
        Private Function ClientTimeOffset() As TimeSpan
            ' TODO: check without using try/catch
            Try
                Dim clientTime As DateTime = DateTime.Parse(fldClientTime.Value)
                Dim serverTime As DateTime = DateTime.Parse(fldServerTime.Value)
                Return TimeSpan.FromSeconds( _
                    DateDiff(DateInterval.Second, serverTime, clientTime))
            Catch
                Return TimeSpan.FromSeconds(0)
            End Try
        End Function

        '---COMMENT---------------------------------------------------------------
        '	convenience function to call locale object
        '
        '	Date:		Name:	Description:
        '	1/13/05  	JEA		Creation
        '-------------------------------------------------------------------------
        Protected Function Say(ByVal resx As String) As String
            Return _hostPage.Say(resx)
        End Function

    End Class
End Namespace
