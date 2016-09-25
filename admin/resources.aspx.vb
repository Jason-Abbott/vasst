Imports AMP.Site

Namespace Pages.Administration
    Public Class Resources
        Inherits AMP.AdminPage

#Region " Controls "

        Protected rptUnapproved As Repeater
        Protected fldAction As HtmlControls.HtmlInputHidden
        Protected fldAssetID As HtmlControls.HtmlInputHidden

#End Region

        Private Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
            Me.Title = "Resources"
            Me.StyleSheet.Add("list")
            Me.ScriptFile.Add("asset")

            If Page.IsPostBack And Profile.User.HasPermission(Permission.ApproveAsset) Then
                Dim asset As AMP.Asset = WebSite.Assets(fldAssetID.Value)
                If Not asset Is Nothing Then
                    Dim email As New AMP.Email

                    Select Case Request.Form(fldAction.UniqueID)
                        Case "approve"
                            If asset.Approve() Then
                                Log.Activity(Activity.ApproveSubmittedAsset, Profile.PersonID)
                                email.AssetApproval(asset)
                                Profile.Message = String.Format("""{0}"" has been approved", asset.Title)
                            End If
                        Case "deny"
                            Profile.Message = String.Format("""{0}"" has been deleted", asset.Title)
                            asset.Delete()
                            Log.Activity(Activity.DenySubmittedAsset, Profile.PersonID)
                            email.AssetDenial(asset)
                    End Select
                End If
            End If

            rptUnapproved.DataSource = WebSite.Assets.Unapproved
            rptUnapproved.DataBind()
        End Sub

#Region " Build Links "

        Protected Function ViewLink(ByVal data As Object) As String
            Dim asset As AMP.Asset = DirectCast(data, AMP.Asset)
            Return String.Format("location.href='{0}'", asset.ViewURL)
        End Function

        Protected Function ApproveLink(ByVal data As Object) As String
            Dim asset As AMP.Asset = DirectCast(data, AMP.Asset)
            Return String.Format("Asset.Approve('{0}')", asset.ID)
        End Function

        Protected Function DenyLink(ByVal data As Object) As String
            Dim asset As AMP.Asset = DirectCast(data, AMP.Asset)
            Return String.Format("Asset.Deny('{0}')", asset.ID)
        End Function

        Protected Function EditLink(ByVal data As Object) As String
            Dim asset As AMP.Asset = DirectCast(data, AMP.Asset)
            Return String.Format("location.href='../resource-edit.aspx?id={0}'", _
                asset.ID.ToString)
        End Function

#End Region

    End Class
End Namespace
