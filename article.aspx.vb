Imports System.IO
Imports System.Net
Imports System.Text
Imports System.Configuration.ConfigurationSettings

Namespace Pages
    Public Class Article
        Inherits AMP.Page

        Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init

            If Not Profile.Authenticated Then
                ' send to signin first
                Profile.DestinationPage = Request.Url.PathAndQuery
                Profile.WriteTestCookie()
                response.Redirect(AppSettings("LoginPage"), False)
            Else
                Dim assetID As String = Request.QueryString("id")
                If assetID <> Nothing Then
                    Dim asset As AMP.Asset = WebSite.Assets(assetID)
                    If Not asset Is Nothing AndAlso (asset.Type And asset.Types.Link) > 0 Then
                        ' asset exists and has link
                        ' TODO: fix scraping and condition
                        If False AndAlso asset.Link.Scrape Then
                            ' display link content in page
                            Dim response As String = asset.Link.Content
                            If response <> Nothing Then
                                asset.Link.Views += 1
                                Me.pnlBody.Controls.Add(New LiteralControl(response))
                                Me.Title = asset.Title
                                Return
                            End If
                        Else
                            ' redirect to asset URL
                            asset.Link.Views += 1
                            response.Redirect(asset.Link.FullUrl, False)
                            Return
                        End If
                    Else
                        AMP.Log.Error(String.Format("No asset for GUID ""{0}""", assetID), _
                            Log.ErrorType.Custom, Profile.User)
                    End If
                Else
                    AMP.Log.Error(String.Format("No asset ID specified for download"), _
                        Log.ErrorType.Custom, Profile.User)
                End If
                ' return to resource page if unable to get url
                response.Redirect(String.Format("~/resource.aspx?id={0}", assetID), False)
            End If
        End Sub

    End Class
End Namespace