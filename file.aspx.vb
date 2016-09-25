Imports AMP.Site
Imports System.IO
Imports System.Configuration.ConfigurationSettings

Namespace Pages
    Public Class File
        Inherits System.Web.UI.Page

        Private _entry As AMP.ContestEntry
        Private _asset As AMP.Asset

        Private Enum Download
            None
            Asset
            ContestEntry
        End Enum

        '---COMMENT---------------------------------------------------------------
        '	move file to temporary path for download
        '
        '	Date:		Name:	Description:
        '	1/15/05     JEA		Creation
        '   2/23/05     JEA     Also process contest entry downloads
        '-------------------------------------------------------------------------
        Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
            Dim id As String = Request.QueryString("id")
            Dim type As Download = Download.None
            Profile.ResumeDownload = Nothing    ' cancel any paused download

            If id <> Nothing Then
                Dim file As FileInfo
                Dim name As String
                Dim folder As String = AppSettings("UploadFolder")
                _asset = WebSite.Assets(id)
                _entry = WebSite.Contests.EntryWithID(id)

                If Not _asset Is Nothing AndAlso (_asset.Type And _asset.Types.File) > 0 Then
                    ' asset exists and has file
                    type = Download.Asset
                    name = _asset.File.Name
                    If _asset.Status = Status.Approved Then folder = AppSettings("ResourceFolder")
                ElseIf Not _entry Is Nothing Then
                    ' id is for contest entry
                    type = Download.ContestEntry
                    name = _entry.File.Name
                    If _entry.Status = Status.Approved Then folder = AppSettings("ContestFolder")
                End If

                If Not Profile.Authenticated Then
                    ' send to signin first
                    Dim sendTo As String
                    Select Case type
                        Case Download.Asset, Download.ContestEntry
                            Profile.DestinationPage = Me.GetURL(type)
                            Profile.ResumeDownload = New Guid(id)
                            Profile.WriteTestCookie()
                            sendTo = AppSettings("LoginPage")
                        Case Else
                            sendTo = Request.UrlReferrer.PathAndQuery()
                    End Select
                    Response.Redirect(sendTo, False)
                    Return
                End If

                BugOut(HttpRuntime.AppDomainAppPath)
                BugOut("{0}{1}\{2}", HttpRuntime.AppDomainAppPath, folder, name)

                file = New FileInfo(String.Format("{0}{1}\{2}", _
                    HttpRuntime.AppDomainAppPath, folder, name))

                If Not file Is Nothing AndAlso file.Exists Then
                    If type = Download.Asset Then _asset.File.Downloads += 1
                    With Response
                        .RedirectLocation = file.Name
                        .ContentType = "application/octet-stream"
                        .AppendHeader("Content-Disposition", "attachment;filename=" & file.Name)
                        .AppendHeader("Content-Length", file.Length.ToString)
                        .WriteFile(file.FullName)
                        .Flush()
                    End With
                    Return
                Else
                    Dim say As New AMP.Locale
                    AMP.Log.Error(String.Format("File {0} does not exist", file.FullName), _
                        Log.ErrorType.Custom, Profile.User)
                    Profile.Message = say("Error.FileNotFound")
                End If
            End If

            ' no file if we made it here
            AMP.Log.Error(String.Format("No asset or entry for GUID ""{0}""", id), _
                Log.ErrorType.Custom, Profile.User)

            ' return to original page if unable to download
            Response.Redirect(Me.GetURL(type), False)
        End Sub

        '---COMMENT---------------------------------------------------------------
        '	get URL to redirect back to after authentication
        '
        '	Date:		Name:	Description:
        '	3/3/05      JEA		Creation
        '-------------------------------------------------------------------------
        Private Function GetURL(ByVal type As Download) As String
            Select Case type
                Case Download.Asset
                    Return String.Format("~/resource.aspx?id={0}", _
                        _asset.ID.ToString)
                Case Download.ContestEntry
                    Return String.Format("~/contest.aspx?id={0}", _
                        _entry.Contest.ID.ToString)
                Case Else
                    Return "~/default.aspx"
            End Select
        End Function
    End Class
End Namespace