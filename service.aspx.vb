Imports AMP.Common
Imports System.Net
Imports System.Text
Imports System.Configuration.ConfigurationSettings

Namespace Pages
    Public Class Service
        Inherits System.Web.UI.Page

        Private _answer As String

        '---COMMENT---------------------------------------------------------------
        '	respond to out-of-band calls
        '
        '	Date:		Name:	Description:
        '	1/7/05  	JEA		Creation
        '   1/29/05     JEA     Role stuff
        '   2/15/05     JEA     Categories
        '-------------------------------------------------------------------------
        Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
            Select Case Request.QueryString("method")
                Case "EmailCheck"
                    Dim email As String = Request.QueryString("address")
                    Dim person As AMP.Person = WebSite.Persons.WithEmail(email)
                    _answer = "answer={{Errors:[{0}],Exists:{1},Disabled:{2}}};"

                    If person Is Nothing Then
                        _answer = String.Format(_answer, Nothing, "false", "false")
                    Else
                        _answer = String.Format(_answer, Nothing, "true", _
                            (person.Status = AMP.Site.Status.Disabled).ToString.ToLower)
                    End If

                Case "ResetPassword"
                    Dim email As String = Request.QueryString("address")
                    Dim person As AMP.Person = WebSite.Persons.WithEmail(email)
                    Dim security As New AMP.Security
                    _answer = "answer={{Errors:[{0}],Exists:{1},Sent:{2}}};"

                    If person Is Nothing Then
                        _answer = String.Format(_answer, Nothing, "false", "false")
                    Else
                        _answer = String.Format(_answer, Nothing, "true", _
                            security.ResetPassword(person).ToString.ToLower)
                    End If

                Case "SendConfirmation"
                    Dim email As String = Request.QueryString("address")
                    Dim code As String = AMP.Security.RandomText( _
                        CInt(AppSettings("ValidationCodeLength")))
                    Dim mail As New AMP.Email
                    _answer = "answer={{Errors:[{0}],Sent:{1},Code:""{2}""}};"

                    mail.Confirmation("VASST Guest", email, code)
                    _answer = String.Format(_answer, Nothing, "true", code)

                Case "RolePermissionAdd"
                    Dim role As AMP.Role = WebSite.Roles(Request.QueryString("rID"))
                    _answer = "answer={{Errors:[{0}]}};"

                    If role Is Nothing Then
                        _answer = String.Format(_answer, "No role with that ID")
                        Return
                    End If

                    If role.AddPermission(Request.QueryString("pID")) Then
                        _answer = String.Format(_answer, "")
                        WebSite.Save()
                    Else
                        _answer = String.Format(_answer, "The role already has that permission")
                    End If

                Case "RolePermissionRemove"
                    Dim role As AMP.Role = WebSite.Roles(Request.QueryString("rID"))
                    _answer = "answer={{Errors:[{0}]}};"

                    If role Is Nothing Then
                        _answer = String.Format(_answer, "No role with that ID")
                        Return
                    End If

                    If role.RemovePermission(Request.QueryString("pID")) Then
                        _answer = String.Format(_answer, "")
                        WebSite.Save()
                    Else
                        _answer = String.Format(_answer, "The role does not have that permission")
                    End If

                Case "UrlCheck"
                    Dim web As New WebClient
                    Try
                        web.DownloadData(Request.QueryString("url"))
                        _answer = "true"
                    Catch ex As System.Net.WebException
                        _answer = "false"
                    End Try

                Case "CategoryAdd"
                    Dim category As New AMP.Category
                    _answer = "answer={{Errors:[{0}],ID:""{1}"",New:true}};"

                    With category
                        .Name = AMP.Security.SafeString(Request.QueryString("name"), 50)
                        .Section = StringToEnum(Request.QueryString("sections"))
                        .ForEntity = StringToEnum(Request.QueryString("entities"))
                        .ForAssetType = StringToEnum(Request.QueryString("assets"))
                        .ForSoftwareType = StringToEnum(Request.QueryString("software"))
                        _answer = String.Format(_answer, Nothing, .ID)
                    End With
                    WebSite.Categories.Add(category)
                    WebSite.Save()

                Case "CategorySave"
                    Dim category As AMP.Category = WebSite.Categories(Request.QueryString("id"))
                    _answer = "answer={{Errors:[{0}],ID:""{1}"",New:false}};"

                    If category Is Nothing Then
                        _answer = String.Format(_answer, "No category with that ID", Nothing)
                    Else
                        With category
                            .Name = AMP.Security.SafeString(Request.QueryString("name"), 50)
                            .Section = StringToEnum(Request.QueryString("sections"))
                            .ForEntity = StringToEnum(Request.QueryString("entities"))
                            .ForAssetType = StringToEnum(Request.QueryString("assets"))
                            .ForSoftwareType = StringToEnum(Request.QueryString("software"))
                            _answer = String.Format(_answer, Nothing, .ID)
                        End With
                        WebSite.Save()
                    End If

                Case "CategoryLoad"
                        Dim category As AMP.Category = WebSite.Categories(Request.QueryString("id"))
                    _answer = "answer={{Errors:[{0}],Sections:[{1}],Entities:[{2}],Assets:[{3}],Software:[{4}]}};"

                        If category Is Nothing Then
                            _answer = String.Format(_answer, "No category with that ID")
                        Else
                            _answer = String.Format(_answer, Nothing, _
                                EnumToString(category.Section, GetType(AMP.Site.Section)), _
                                EnumToString(category.ForEntity, GetType(AMP.Site.Entity)), _
                                EnumToString(category.ForAssetType, GetType(AMP.Asset.Types)), _
                                EnumToString(category.ForSoftwareType, GetType(AMP.Software.Types)))
                        End If

                Case "CategoryDelete"
                    Dim category As AMP.Category = WebSite.Categories(Request.QueryString("id"))
                    _answer = "answer={{Errors:[{0}],ID:""{1}""}};"

                    If category Is Nothing Then
                        _answer = String.Format(_answer, "No category with that ID", Nothing)
                    Else
                        WebSite.Categories.Remove(category)
                        WebSite.Save()
                        _answer = String.Format(_answer, Nothing, category.ID)
                    End If

                Case "GetStatistics"
                    Dim answer As New StringBuilder
                    Dim lastSave As Date = Global.LastSave
                    If lastSave = Nothing Then lastSave = Global.ApplicationStart
                    With answer
                        .Append("answer={{Errors:[]")
                        .Append(",Visits:""{0}""")
                        .Append(",Users:""{1}""")
                        .Append(",DataFile:{{Name:""{2}""")
                        .Append(",Size:""{3}""")
                        .Append(",Saved:""{4}""}}")
                        .Append("}};")
                    End With
                    _answer = String.Format(answer.ToString, _
                        String.Format("{0:N0}", Log.VisitsToday), _
                        String.Format("{0:N0}", WebSite.Persons.Count), _
                        Global.ActiveDataFile.Name, _
                        String.Format("{0:N0} bytes", Global.ActiveDataFile.Length), _
                        String.Format("{0:g}", lastSave))
            End Select
        End Sub

        Private Sub Page_PreRender(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.PreRender
            Response.ContentType = "text/plain"
            Me.Controls.Add(New LiteralControl(_answer))
        End Sub

    End Class
End Namespace
