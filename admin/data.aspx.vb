Imports AMP.Site
Imports System.IO
Imports System.Configuration.ConfigurationSettings

Namespace Pages.Administration
    Public Class Data
        Inherits System.Web.UI.Page

        '---COMMENT---------------------------------------------------------------
        '	download latest data file
        '
        '	Date:		Name:	Description:
        '	2/20/05     JEA		Creation
        '   3/3/05      JEA     Add Content-Length header
        '-------------------------------------------------------------------------
        Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
            If Not Profile.Authenticated Then
                ' send to signin first
                Profile.DestinationPage = Request.Url.PathAndQuery
                Profile.WriteTestCookie()
                Response.Redirect(AppSettings("LoginPage"), False)
            ElseIf Not Profile.User.HasPermission(Permission.DownloadDataFile) Then
                ' no permission for this
                Dim say As New AMP.Locale
                Profile.Message = say("Error.Permissions")
                Response.Redirect("~/", True)
            Else
                ' download data file--matches logic in global.asax.vb
                Dim file As New AMP.Data.File
                Dim type As System.Type = System.Type.GetType("AMP.Site")
                Dim path As FileInfo = file.Newest(AppSettings("DataFolder"), type)

                With Response
                    .RedirectLocation = path.Name
                    .ContentType = "application/octet-stream"
                    .AppendHeader("Content-Disposition", "attachment;filename=" & path.Name)
                    .AppendHeader("Content-Length", path.Length.ToString)
                    .WriteFile(path.FullName)
                    .Flush()
                End With
            End If
        End Sub

    End Class
End Namespace