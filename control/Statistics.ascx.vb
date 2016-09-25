Imports AMP.Site
Imports System.IO
Imports System.Configuration.ConfigurationSettings

Namespace Controls
    Public Class Statistics
        Inherits AMP.Controls.WebPart

#Region " Controls "

        Protected lblAssetCount As Label
        Protected lblUserCount As Label
        Protected lblProductCount As Label
        Protected lblAppStart As Label
        ' server
        Protected lblIIS As Label
        Protected lblOS As Label
        Protected lblServerName As Label
        Protected lblServerUser As Label
        Protected lblAspNetVersion As Label
        ' data file
        Protected lblDataSize As Label
        Protected lblDataSaved As Label
        Protected lnkDataFile As HyperLink
        ' today
        Protected lblSales As Label
        Protected lblVisitsToday As Label

#End Region

        Private Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
            If Profile.User.HasPermission(Permission.ViewSiteStatistics) Then
                Me.Title = Me.Page.Say("Title.SiteStatistics")
                Me.Minimizable = True

                Me.Page.ScriptFile.Add("statistics")
                Me.Page.ScriptFile.Add("broker/common")

                lblAssetCount.Text = String.Format("{0:N0}", WebSite.Assets.Count)
                lblUserCount.Text = String.Format("{0:N0}", WebSite.Persons.Count)
                lblProductCount.Text = String.Format("{0:N0}", WebSite.Catalog.Count)
                lblAppStart.Text = String.Format("{0:g}", Global.ApplicationStart)
                ' server
                lblIIS.Text = Request.ServerVariables("SERVER_SOFTWARE").Replace("Microsoft-", Nothing)
                lblOS.Text = Environment.OSVersion.ToString().Replace("Microsoft ", Nothing)
                lblServerName.Text = Environment.MachineName
                lblServerUser.Text = Environment.UserName
                lblAspNetVersion.Text = Environment.Version.ToString
                ' data file
                lnkDataFile.Text = Global.ActiveDataFile.Name
                lnkDataFile.NavigateUrl = String.Format("{0}/admin/data.aspx", Global.BasePath)
                lblDataSize.Text = String.Format("{0:N0} bytes", Global.ActiveDataFile.Length)
                Dim lastSave As Date = Global.LastSave
                If lastSave = Nothing Then lastSave = Global.ApplicationStart
                lblDataSaved.Text = String.Format("{0:g}", lastSave)
                ' today
                lblSales.Text = "Unknown"
                lblVisitsToday.Text = String.Format("{0:N0}", Log.VisitsToday)
            Else
                Me.Visible = False
            End If
        End Sub
    End Class
End Namespace