Imports AMP.Site
Imports System.text
Imports System.Configuration.ConfigurationSettings

Namespace Pages
    Public Class Person
        Inherits AMP.Page

#Region " Controls "

        Protected lblDisplayName As Label
        Protected lblRealName As Label
        Protected lblRole As Label
        Protected lblEmail As Label
        Protected lblWebSite As Label
        Protected imgPerson As AMP.controls.Image
        Protected lblDescription As Label

        Protected lblResources As Label
        Protected pnlResources As Panel
        Protected rptResources As Repeater

        Protected btnEdit As AMP.Controls.Button

#End Region

        '---COMMENT---------------------------------------------------------------
        '	display person details
        '
        '	Date:		Name:	Description:
        '	11/12/04    JEA		Creation
        '   1/24/05     JEA     Show all assets from person
        '   2/9/05      JEA     Edit button
        '-------------------------------------------------------------------------
        Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
            Me.StyleSheet.Add("person")

            If Request.QueryString("id") <> Nothing Then
                Dim person As AMP.Person = WebSite.Persons(Request.QueryString("id"))
                If Not person Is Nothing Then
                    With person
                        lblDisplayName.Text = .DisplayName
                        If .NickName <> Nothing Then lblRealName.Text = .FullName

                        If Profile.User.HasPermission(Permission.ViewUserDetails) _
                            OrElse Not .PrivateEmail Then lblEmail.Text = .EmailLink

                        If person.CanEdit Then
                            btnEdit.Visible = True
                            btnEdit.Url = String.Format("account.aspx?id={0}", person.ID)
                            lblRole.Text = .Roles.ToString
                        Else
                            lblRole.Visible = False
                        End If

                        lblWebSite.Text = .WebSiteLink
                        If .ImageFile <> Nothing Then
                            Dim draw As New AMP.Draw
                            imgPerson.Visible = True
                            imgPerson.Image = draw.Resize( _
                                String.Format("{0}/images/person/{1}", _
                                HttpRuntime.AppDomainAppPath, .ImageFile), _
                                CInt(AppSettings("MaxPersonalImageWidth")), Nothing)
                        End If
                        lblDescription.Text = .Description
                        Me.Title = .DisplayName

                        Dim assets As ArrayList = WebSite.Assets.FromUser(person)

                        If assets.Count > 0 Then
                            Dim label As New StringBuilder
                            With label
                                .Append(Format.WordForNumber(assets.Count, True))
                                .Append(" resource")
                                If assets.Count > 1 Then .Append("s")
                                .Append(" from <nobr>")
                                .Append(person.DisplayName)
                                .Append("</nobr>")
                                lblResources.Text = .ToString
                            End With
                            lblResources.Visible = True
                            pnlResources.Visible = True
                            rptResources.DataSource = assets
                            rptResources.DataBind()

                            If assets.Count > 15 Then pnlResources.Style.Add("height", "200px")
                        End If
                    End With
                    Return
                End If
            End If

            Me.SendBack()
        End Sub

    End Class
End Namespace
