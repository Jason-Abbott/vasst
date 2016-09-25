Imports System.Text
Imports System.Web.Mail
Imports System.Text.RegularExpressions
Imports System.Configuration.ConfigurationSettings

Public Class Email
    Private _from As String
    Private _say As AMP.Locale
    Private _siteName As String
    Private Const _schema As String = "http://schemas.microsoft.com/cdo/configuration/"

    Public Sub New()
        _from = String.Format("VASST ({0})", AppSettings("MailFrom"))
        _say = New AMP.Locale
        _siteName = _say("SiteName")
    End Sub

#Region " Properties "

    Public Property From() As String
        Get
            Return _from
        End Get
        Set(ByVal Value As String)
            _from = Value
        End Set
    End Property

#End Region

#Region " Enumerations "

    Private Enum SendUsing
        Pickup = 1
        Port = 2
    End Enum

    Private Enum Authentication
        Anonymous
        Basic
        Ntlm
    End Enum

#End Region

    '---COMMENT---------------------------------------------------------------
    '	send email about critical error
    '
    '	Date:		Name:	Description:
    '	1/18/05	    JEA 	Creation
    '   2/11/05     JEA     Display times this error has occurred
    '-------------------------------------------------------------------------
    Public Sub [Error](ByVal ex As Exception, ByVal user As AMP.Person, ByVal clientIP As String, _
        ByVal machine As String, ByVal process As String, ByVal sendTo As String(), ByVal occurrences As Integer)

        Dim mail As New MailMessage
        Dim file As New AMP.data.File
        Dim template As String = AppSettings("ErrorTemplate")
        Dim recipients As New StringBuilder
        Dim body As String = file.Load(template)
        Dim occurrenceText As String = String.Empty
        Dim solutionText As String = String.Empty
        Dim folder As String = String.Empty

        If occurrences > 0 Then
            occurrenceText = String.Format("<p>This error has occurred {0} times since the last notification.  All occurrences have been logged.</p>", occurrences)
        End If

        Select Case True
            Case ex.Message.IndexOf("GDI") <> -1
                folder = AppSettings("GeneratedImageFolder")
            Case ex.Message.IndexOf("updateable query") <> -1
                folder = String.Format("{0} and {1}", _
                    AppSettings("DataFolder"), AppSettings("MailDataFolder"))
            Case ex.Message.IndexOf("Access to the path") <> -1
                If ex.Message.IndexOf("person") <> -1 Then
                    folder = AppSettings("UserImageFolder")
                Else
                    Dim re As New Regex(String.Format("{0}(\w+)[\\/]", _
                        HttpRuntime.AppDomainAppPath.Replace("\", "\\")), _
                        RegexOptions.IgnoreCase)
                    Dim m As Match = re.Match(ex.Message)
                    folder = m.Groups(1).Value
                End If
        End Select

        If folder <> String.Empty Then
            solutionText = file.Load(AppSettings("PermissionsTemplate"))
            solutionText = solutionText.Replace("<hostName>", AppSettings("HostName"))
            solutionText = solutionText.Replace("<folder>", folder)
        End If

        body = body.Replace("<userName>", user.FullName)
        body = body.Replace("<dateTime>", DateTime.Now.ToString)
        body = body.Replace("<userIP>", clientIP)
        body = body.Replace("<server>", machine)
        body = body.Replace("<process>", process)
        body = body.Replace("<message>", ex.Message)
        body = body.Replace("<stack>", ex.StackTrace.Replace(vbCrLf, "<br/>"))
        body = body.Replace("<occurrences>", occurrenceText)
        body = body.Replace("<solution>", solutionText)

        With recipients
            For Each address As String In sendTo
                If .Length > 0 Then .Append(",")
                .Append(address.Trim)
            Next
        End With

        With mail
            .BodyFormat = DirectCast(IIf(template.EndsWith("htm"), MailFormat.Html, MailFormat.Text), MailFormat)
            .Priority = MailPriority.High
            .From = _from
            .To = recipients.ToString
            .Subject = String.Format(_say("Subject.CriticalError"), _siteName)
            .Body = body.ToString
        End With

        Me.Send(mail)
    End Sub

#Region " Sign in/up e-mails "

    '---COMMENT---------------------------------------------------------------
    '	send confirmation code to user
    '
    '	Date:		Name:	Description:
    '	1/7/05	    JEA 	Creation
    '-------------------------------------------------------------------------
    Public Sub Confirmation(ByVal name As String, ByVal email As String, ByVal code As String)
        Dim mail As New MailMessage
        Dim body As String = Me.Template("ConfirmationTemplate")

        body = body.Replace("<name>", name)
        body = body.Replace("<code>", code)

        With mail
            .BodyFormat = MailFormat.Text
            .From = _from
            .To = String.Format("{0} ({1})", name, email)
            .Subject = String.Format(_say("Subject.Confirmation"), _siteName)
            .Body = body
        End With

        Me.Send(mail)
    End Sub

    '---COMMENT---------------------------------------------------------------
    '	send password to user
    '
    '	Date:		Name:	Description:
    '	1/7/05	    JEA 	Creation
    '-------------------------------------------------------------------------
    Public Sub Password(ByVal name As String, ByVal email As String, ByVal password As String)
        Dim mail As New MailMessage
        Dim body As String = Me.Template("PasswordTemplate")

        body = body.Replace("<name>", name)
        body = body.Replace("<password>", password)

        With mail
            .BodyFormat = MailFormat.Text
            .From = _from
            .To = String.Format("{0} ({1})", name, email)
            .Subject = String.Format(_say("Subject.NewPassword"), _siteName)
            .Body = body
        End With

        Me.Send(mail)
    End Sub

    Public Sub Password(ByVal person As AMP.Person, ByVal password As String)
        Me.Password(person.FullName, person.Email, password)
    End Sub

#End Region

#Region " Contest e-mails "

    '---COMMENT---------------------------------------------------------------
    '	send e-mail notifying admins of new contest entry
    '
    '	Date:		Name:	Description:
    '	3/1/05	    JEA 	Creation
    '-------------------------------------------------------------------------
    Public Sub ContestEntry(ByVal entry As AMP.ContestEntry)
        Dim sendTo As New StringBuilder
        Dim recipients As ArrayList = WebSite.Persons.WithPermission(Site.Permission.ApproveEntry)
        With sendTo
            For Each p As Person In recipients
                .Append(p.FullName)
                .Append("(")
                .Append(p.Email)
                .Append(");")
            Next
        End With
        Me.EntryMail(entry, Me.Template("ContestEntryTemplate"), sendTo.ToString)
    End Sub

    '---COMMENT---------------------------------------------------------------
    '	send e-mail approving or denying contest entry
    '
    '	Date:		Name:	Description:
    '	2/27/05	    JEA 	Creation
    '-------------------------------------------------------------------------
    Public Sub EntryApproval(ByVal entry As AMP.ContestEntry)
        Dim body As String = Me.Template("EntryApproveTemplate")
        body = body.Replace("<contestUrl>", entry.Contest.FullDetailUrl)
        Me.EntryMail(entry, body)
    End Sub

    Public Sub EntryDenial(ByVal entry As AMP.ContestEntry)
        Me.EntryMail(entry, Me.Template("EntryDenyTemplate"))
    End Sub

    Private Sub EntryMail(ByVal entry As AMP.ContestEntry, ByVal body As String, _
        ByVal sendTo As String)

        Dim mail As New MailMessage
        Dim user As AMP.Person = entry.Contestant

        body = body.Replace("<name>", user.FullName)
        body = body.Replace("<contestTitle>", entry.Contest.Title)
        body = body.Replace("<entryTitle>", entry.Title)

        With mail
            .BodyFormat = MailFormat.Text
            .From = _from
            .To = sendTo
            .Subject = String.Format(_say("Subject.ContestEntry"), _siteName, entry.Title)
            .Body = body
        End With

        'Me.Send(mail)
    End Sub

    Private Sub EntryMail(ByVal entry As AMP.ContestEntry, ByVal body As String)
        Dim user As AMP.Person = entry.Contestant
        Me.EntryMail(entry, body, String.Format("{0} ({1})", user.FullName, user.Email))
    End Sub

#End Region

#Region " Asset e-mails "

    '---COMMENT---------------------------------------------------------------
    '	send e-mail approving or denying asset
    '
    '	Date:		Name:	Description:
    '	2/18/05	    JEA 	Creation
    '-------------------------------------------------------------------------
    Public Sub AssetApproval(ByVal asset As AMP.Asset)
        Dim body As String = Me.Template("AssetApproveTemplate")
        body = body.Replace("<resourceUrl>", asset.FullDetailUrl)
        Me.AssetMail(asset, body)
    End Sub

    Public Sub AssetDenial(ByVal asset As AMP.Asset)
        Me.AssetMail(asset, Me.Template("AssetDenyTemplate"))
    End Sub

    Private Sub AssetMail(ByVal asset As AMP.Asset, ByVal body As String)
        Dim mail As New MailMessage
        Dim user As AMP.Person = asset.SubmittedBy

        body = body.Replace("<name>", user.FullName)
        body = body.Replace("<resource>", asset.Title)

        With mail
            .BodyFormat = MailFormat.Text
            .From = _from
            .To = String.Format("{0} ({1})", user.FullName, user.Email)
            .Subject = String.Format(_say("Subject.AssetContribution"), _siteName, asset.Title)
            .Body = body
        End With

        Me.Send(mail)
    End Sub

#End Region

    '---COMMENT---------------------------------------------------------------
    '	retrieve content of template file as string
    '
    '	Date:		Name:	Description:
    '	2/18/05	    JEA 	Abstracted
    '-------------------------------------------------------------------------
    Private Function Template(ByVal name As String) As String
        Dim file As New AMP.Data.File
        Return file.Load(AppSettings(name))
    End Function

    '---COMMENT---------------------------------------------------------------
    '	Send e-mail through specified mail server
    '   use smtp login when needed
    '
    '	Date:		Name:	Description:
    '	10/1/04	    JEA 	Creation
    '   2/18/05     JEA     Check for debug address
    '-------------------------------------------------------------------------
    Private Sub Send(ByVal mail As MailMessage)
        Dim userName As String = AppSettings("MailUser")
        Dim redirectTo As String = AppSettings("RedirectMailTo")

        If userName <> Nothing Then
            ' server must require authentication
            With mail.Fields
                .Add(_schema & "sendusername", userName)
                .Add(_schema & "sendpassword", AppSettings("MailPassword"))
                .Add(_schema & "smtpauthenticate", Authentication.Basic)
                '.Add(_schema & "sendusing", SendUsing.Port)
                '.Add(_schema & "smtpserverport", 25)
                '.Add(_schema & "smtpusessl", False)
                '.Add(_schema & "smtpconnectiontimeout", 60)
            End With
        End If

        ' send only to debug address if one specified
        If redirectTo <> Nothing Then
            mail.Body &= String.Format("{0}{0}(originally sent to {1})", Environment.NewLine, mail.To)
            mail.To = redirectTo
        End If

        SmtpMail.SmtpServer = AppSettings("MailServer")
        Try
            SmtpMail.Send(mail)
        Catch e As Exception
            BugOut("Error sending mail through server {0}: {1}", AppSettings("MailServer"), e.Message)
        End Try
    End Sub
End Class