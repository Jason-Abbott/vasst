Imports AMP.Site
Imports AMP.LegacyDataSet
Imports AMP.Data.Serialization
Imports System.Data
Imports System.Data.OleDb
Imports System.Web.Caching
Imports System.Web.HttpContext
Imports System.Configuration.ConfigurationSettings
Imports System.IO
Imports System.Runtime.Serialization
Imports System.Runtime.Serialization.Formatters.Binary
Imports AMP.Global

Namespace Legacy
    Public Class Import
        Private _ds As AMP.LegacyDataSet
        Private _packageMap As New Hashtable
        Private _pluginMap As New Hashtable
        Private _personMap As New Hashtable
        Private _publisherMap As New Hashtable
        Private _softwareMap As New Hashtable
        Private _assetMap As New Hashtable
        Private _contestMap As New Hashtable
        Private _categoryMap As New Hashtable
        Private _keepUsers As Integer()

        Public Enum ItemType
            Project = 1
            Tutorial = 2
            Forum = 3
            Plugin = 4
            Publisher = 5
            Script = 6
            Review = 7
            Preset = 8
        End Enum

        Public Sub Collections()
            _ds = Me.GetDataSet

            ' map multiple plugin packages to single publisher entity
            _packageMap.Add(1, 11)
            _packageMap.Add(2, 11)
            _packageMap.Add(3, 12)
            _packageMap.Add(6, 13)
            _packageMap.Add(7, 14)

            ' map multiple plugins to single software entity
            _pluginMap.Add(1, 11)
            _pluginMap.Add(2, 11)
            _pluginMap.Add(3, 11)
            _pluginMap.Add(4, 12)
            _pluginMap.Add(5, 12)
            _pluginMap.Add(6, 13)
            _pluginMap.Add(7, 13)
            _pluginMap.Add(8, 13)
            _pluginMap.Add(9, 13)
            _pluginMap.Add(10, 13)
            _pluginMap.Add(16, 14)
            _pluginMap.Add(17, 15)
            _pluginMap.Add(18, 15)
            _pluginMap.Add(19, 16)
            _pluginMap.Add(20, 17)
            _pluginMap.Add(21, 18)

            'BugOut("Importing roles ...")
            'BugTab()
            '.Roles = Me.MakeRoles
            'BugOut("{0} roles", .Roles.Count)
            'BugUntab()

            BugOut("Importing persons ...")
            BugTab()
            BugOut("Added {0} people", Me.LoadPersons())
            BugUntab()

            BugOut("Importing categories ...")
            BugTab()
            BugOut("Added {0} categories", Me.LoadCategories())
            BugUntab()

            BugOut("Importing publishers ...")
            BugTab()
            BugOut("{0} publishers", Me.LoadPublishers)
            BugUntab()

            BugOut("Importing assets ...")
            BugTab()
            BugOut("{0} assets", Me.LoadAssets)
            BugUntab()

            BugOut("Importing contests ...")
            BugTab()
            BugOut("{0} contests", Me.LoadContests)
            BugUntab()

            ' importing saves immediately
            WebSite.UseSaveBuffer = False
            WebSite.Save()

            _ds.Dispose()
            _ds = Nothing
        End Sub

#Region " Roles "

        '---COMMENT---------------------------------------------------------------
        '	build roles and permissions for persons
        '
        '	Date:		Name:	Description:
        '	12/7/04     JEA		Creation
        '   1/3/04      JEA     Add anonymouse role
        '-------------------------------------------------------------------------
        Private Function MakeRoles() As AMP.RoleCollection
            Dim anonymous As New AMP.Role
            With anonymous
                .ID = AMP.Site.Role.Anonymous
                .Name = "Anonymous User"
                .Permissions = New AMP.Site.Permission() {}
            End With

            Dim unverified As New AMP.Role
            With unverified
                .ID = WebSite.Role.UnverifiedGuest
                .Name = "Unverified Guest"
                .Permissions = New AMP.Site.Permission() { _
                    Permission.PurchaseProduct, _
                    Permission.EditMyself}
            End With

            Dim verified As New AMP.Role
            With verified
                .ID = WebSite.Role.VerifiedGuest
                .Name = "Verified Guest"
                .InheritFrom.Add(unverified)
                .Permissions = New AMP.Site.Permission() { _
                    Permission.AddAsset, _
                    Permission.DownloadAsset, _
                    Permission.EditMyAsset, _
                    Permission.EnterContest, _
                    Permission.LinkToAsset, _
                    Permission.PurchaseProduct, _
                    Permission.VoteInContest}
            End With

            Dim manager As New AMP.Role
            With manager
                .ID = WebSite.Role.Manager
                .Name = "Manager"
                .InheritFrom.Add(verified)
                .Permissions = New Permission() { _
                    Permission.AddProduct, _
                    Permission.ApproveAsset, _
                    Permission.RejectAsset, _
                    Permission.ViewUserDetails, _
                    Permission.ChangeAssetSection}
            End With

            Dim editor As New AMP.Role
            With editor
                .ID = WebSite.Role.Editor
                .Name = "Editor"
                .InheritFrom.Add(verified)
                .Permissions = New Permission() { _
                    Permission.AddContent, _
                    Permission.AddQuote, _
                    Permission.AddFeaturedItem, _
                    Permission.DeleteContent, _
                    Permission.DeleteFeaturedItem, _
                    Permission.DeleteQuote, _
                    Permission.EditContent, _
                    Permission.EditContest, _
                    Permission.EditQuote, _
                    Permission.EditAnyAsset}
            End With

            Dim admin As New AMP.Role
            With admin
                .ID = WebSite.Role.Administrator
                .Name = "Administrator"
                .InheritFrom.Add(manager)
                .InheritFrom.Add(editor)
                .Permissions = New AMP.Site.Permission() { _
                    Permission.AddMenuItem, _
                    Permission.AddUser, _
                    Permission.DeleteProduct, _
                    Permission.DeleteUser, _
                    Permission.EditAnyUser}
            End With

            Dim rc As New RoleCollection

            rc.Add(admin)
            rc.Add(editor)
            rc.Add(manager)
            rc.Add(verified)
            rc.Add(unverified)
            rc.Add(anonymous)
            Return rc
        End Function

#End Region

#Region " Persons "

        '---COMMENT---------------------------------------------------------------
        '	import from database to person objects
        '
        '	Date:		Name:	Description:
        '	11/29/04    JEA		Creation
        '   12/7/04     JEA     Process permissions
        '   1/4/05      JEA     Use GUID for ID and map to old ID and get image
        '   1/15/05     JEA     Stronger user filter
        '   1/27/05     JEA     Only add new people
        '-------------------------------------------------------------------------
        Private Function LoadPersons() As Integer
            Dim p As AMP.Person
            Dim added As Integer = 0

            For Each row As tblUsersRow In _ds.tblUsers
                Dim lastLogin As DateTime

                If row.Item("dtLastLogin").Equals(DBNull.Value) Then
                    If row.lUserID <> 0 AndAlso row.lUserID <> 2522 Then
                        lastLogin = row.dtDateRegistered
                    End If
                Else
                    lastLogin = row.dtLastLogin
                End If

                If Array.IndexOf(_keepUsers, row.lUserID) > 0 OrElse _
                    ((row.lUserID <> 0 AndAlso row.lUserID <> 2522) AndAlso _
                    DateDiff(DateInterval.Month, lastLogin, DateTime.Now) < 7 AndAlso _
                    (row.lStatusID = 2 OrElse _
                    (row.lStatusID = 1 AndAlso DateDiff(DateInterval.Day, row.dtDateRegistered, DateTime.Now) < 15))) Then

                    p = WebSite.Persons.WithEmail(row.vsEmail)
                    If p Is Nothing Then
                        ' create new person
                        p = New AMP.Person
                        p.FirstName = row.vsFirstName
                        p.LastName = row.vsLastName

                        If Not row.Item("vsAboutMe").Equals(DBNull.Value) Then
                            p.Description = row.vsAboutMe
                        End If
                        If Not row.Item("vsScreenName").Equals(DBNull.Value) Then
                            p.NickName = row.vsScreenName
                        End If
                        p.Email = row.vsEmail
                        p.Password = AMP.Security.Encrypt(row.vspassword)
                        p.RegisteredOn = row.dtDateRegistered
                        If Not row.Item("dtLastLogin").Equals(DBNull.Value) Then
                            p.LastLogin = row.dtLastLogin
                        End If
                        p.Status = CType(row.lStatusID, AMP.Site.Status)

                        If p.Status = Status.Pending Then
                            ' only add validation code if not validated yet
                            If Not row.Item("sValidationCode").Equals(DBNull.Value) Then
                                p.ConfirmationCode = row.sValidationCode
                            End If
                        End If
                        If Not row.Item("vsHomePageURL").Equals(DBNull.Value) Then
                            p.WebSite = row.vsHomePageURL
                        End If

                        If Not row.Item("lSiteID").Equals(DBNull.Value) Then
                            Select Case row.lSiteID
                                Case 1  ' Vegas
                                    p.Section = Section.Vegas
                                Case 2  ' Ulead
                                    p.Section = Section.Ulead
                                Case 3  ' Adobe
                                    p.Section = Section.Adobe
                            End Select
                        End If
                        p.PrivateEmail = row.bPrivateEmail

                        Select Case row.lUserTypeID
                            Case 1  ' anonymous
                                p.Roles.Add(WebSite.Roles(AMP.Site.Role.Anonymous))
                            Case 2  ' unverified user
                                p.Roles.Add(WebSite.Roles(AMP.Site.Role.UnverifiedGuest))
                            Case 3  ' verified user
                                p.Roles.Add(WebSite.Roles(AMP.Site.Role.VerifiedGuest))
                            Case 4  ' admin
                                p.Roles.Add(WebSite.Roles(AMP.Site.Role.Administrator))
                        End Select

                        ' get image
                        p.ImageFile = Me.FindImage(row.lUserID)

                        WebSite.Persons.Add(p)
                        added += 1
                    Else
                        ' update existing person
                        If p.LastLogin < lastLogin Then p.LastLogin = lastLogin
                    End If
                    _personMap.Add(row.lUserID, p.ID)
                End If
            Next
            Return added
        End Function

        '---COMMENT---------------------------------------------------------------
        '	retrieve the name of the image for this user, if any
        '
        '	Date:		Name:	Description:
        '	1/4/04      JEA		Creation
        '-------------------------------------------------------------------------
        Private Function FindImage(ByVal userID As Integer) As String
            Dim fileName As String = String.Format("{0}/images/person/person_{1:00000}.jpg", _
                HttpRuntime.AppDomainAppPath, userID)
            Dim file As New FileInfo(fileName)

            If Not file.Exists Then
                file = New FileInfo(fileName.Replace("jpg", "gif"))
                If Not file.Exists Then Return String.Empty
            End If

            Return file.Name
        End Function

#End Region

#Region " Categories "

        '---COMMENT---------------------------------------------------------------
        '	import from database to category objects
        '
        '	Date:		Name:	Description:
        '	11/29/04    JEA		Creation
        '   1/27/05     JEA     Only add new categories
        '-------------------------------------------------------------------------
        Private Function LoadCategories() As Integer
            Dim c As AMP.Category
            Dim added As Integer = 0

            For Each row As tblCategoriesRow In _ds.tblCategories
                c = WebSite.Categories.WithName(row.vsCategoryName)
                If c Is Nothing Then
                    c = New AMP.Category
                    c.Name = row.vsCategoryName

                    For Each r As tblCategoryItemTypesRow In row.GettblCategoryItemTypesRows
                        c.ForAssetType = (c.ForAssetType Or Me.GetAssetType(r.lItemTypeID))
                        c.ForEntity = (c.ForEntity Or Me.GetEntityEnum(r.lItemTypeID))
                        c.ForSoftwareType = (c.ForSoftwareType Or Me.GetSoftwareType(r.lItemTypeID))
                    Next

                    c.Section = Section.All
                    WebSite.Categories.Add(c)
                    added += 1
                End If
                _categoryMap.Add(row.lCategoryID, c.ID)
            Next
            Return added
        End Function

        Private Function GetSoftwareType(ByVal itemType As Integer) As Software.Types
            If itemType = 4 Then
                Return Software.Types.Plugin
            Else
                Return 0
            End If
        End Function

        Private Function GetAssetType(ByVal itemType As Integer) As Asset.Types
            Select Case itemType
                Case 1
                    Return Asset.Types.Project
                Case 2
                    Return Asset.Types.Tutorial
                Case 6
                    Return Asset.Types.Script
                Case 7
                    Return Asset.Types.Review
                Case Else
                    Return 0
            End Select
        End Function

        Private Function GetEntityEnum(ByVal itemType As Integer) As Site.Entity
            Select Case itemType
                Case 1, 2, 6, 7
                    Return AMP.Site.Entity.Asset
                Case 4
                    Return AMP.Site.Entity.Software
                Case 5
                    Return AMP.Site.Entity.Publisher
                Case Else
                    Return 0
            End Select
        End Function

#End Region

#Region " Publishers "

        Private Function LoadPublishers() As Integer
            Dim p As AMP.Publisher
            Dim added As Integer = 0

            ' software
            For Each row As tblPublishersRow In _ds.tblPublishers
                If row.lPublisherID < 5 AndAlso row.vsPublisherName <> "Unspecified" Then

                    p = WebSite.Publishers.WithName(row.vsPublisherName)
                    If p Is Nothing Then
                        ' create new publisher
                        p = New AMP.Publisher

                        p.Name = row.vsPublisherName

                        If Not row("vsPublisherURL").Equals(DBNull.Value) Then
                            p.WebSite = row.vsPublisherURL
                        End If

                        If Not row("vsPublisherLogo").Equals(DBNull.Value) Then
                            p.LogoFile = row.vsPublisherLogo
                        End If

                        p.Section = Section.All
                        p.ForAssetType = Me.LoadAssetTypes(row.GettblPublisherItemTypesRows)

                        WebSite.Publishers.Add(p)
                        added += 1
                    End If
                    _publisherMap.Add(row.lPublisherID, p.ID)
                    Me.LoadSoftware(row.GettblSoftwareRows, p)
                End If
            Next

            ' plugins
            For Each row As tblPluginPackagesRow In _ds.tblPluginPackages
                If _publisherMap.ContainsKey(_packageMap(row.lPluginPackageID)) Then
                    ' package already added as publisher
                    ' add additional plugins for package as software entities
                    p = WebSite.Publishers(_publisherMap(_packageMap(row.lPluginPackageID)).toString)

                    'Dim software As AMP.SoftwareCollection = Me.LoadPlugins(row.GettblPluginsRows, p)
                    'With p.Software
                    '    For Each s As AMP.Software In software
                    '        If .WithName(s.Name) Is Nothing Then .Add(s)
                    '    Next
                    'End With
                Else
                    ' create new publisher for plugin package
                    p = New AMP.Publisher

                    Select Case row.lPluginPackageID
                        Case 1, 2
                            p.Name = "DubugMode"
                            p.WebSite = "debugmode.com"
                            p.Section = Section.Vegas
                        Case 3
                            p.Name = "Pixélan"
                            p.WebSite = "pixelan.com"
                            p.Section = Section.Vegas
                        Case 6
                            p.Name = "Scott Moore"
                            p.WebSite = "www.endor.demon.co.uk"
                            p.Section = Section.Vegas
                        Case 7
                            p.Name = "Boris"
                            p.WebSite = "borisfx.com"
                            p.Section = Section.All
                    End Select

                    If WebSite.Publishers.WithName(p.Name) Is Nothing Then
                        ' add new publisher
                        p.ForAssetType = Asset.Types.Project Or Asset.Types.Script _
                            Or Asset.Types.Preset
                        'Me.LoadPlugins(row.GettblPluginsRows, p)
                        WebSite.Publishers.Add(p)
                        added += 1
                    Else
                        ' publisher already exists
                        p = WebSite.Publishers.WithName(p.Name)
                    End If

                    _publisherMap.Add(_packageMap(row.lPluginPackageID), p.ID)
                End If

                Me.LoadPlugins(row.GettblPluginsRows, p)
            Next

            Return added
        End Function

        '---COMMENT---------------------------------------------------------------
        '	get software entity for file
        '
        '	Date:		Name:	Description:
        '	11/29/04	JEA		Creation
        '   12/26/04    JEA     Make members of publisher entity
        '   1/27/05     JEA     Check if software exists for publisher
        '-------------------------------------------------------------------------
        Private Sub LoadSoftware(ByVal rows As tblSoftwareRow(), _
            ByVal publisher As AMP.Publisher)

            Dim s As AMP.Software
            Dim sc As New AMP.SoftwareCollection

            For Each row As tblSoftwareRow In rows
                If row.vsSoftwareName <> "Unspecified" Then
                    s = publisher.Software.WithName(row.vsSoftwareName)
                    If s Is Nothing Then
                        s = New AMP.Software
                        s.Name = row.vsSoftwareName
                        BugOut("Created software for {0}", row.lSoftwareID)
                        s.Type = Software.Types.Application
                        s.Free = False
                        s.Versions = Me.LoadVersions(row.GettblSoftwareVersionsRows)
                        s.Publisher = publisher
                        Select Case s.Name.ToLower
                            Case "vegas"
                                s.Extensions.Add("veg")
                                s.Extensions.Add("jpg")
                                s.Extensions.Add("wmv")
                                s.Extensions.Add("wav")
                                s.Extensions.Add("avi")
                                s.Extensions.Add("png")
                                s.Extensions.Add("bmp")
                                s.Extensions.Add("pca")
                                s.Extensions.Add("swf")
                                s.Extensions.Add("mp3")
                                s.Extensions.Add("jpeg")
                                s.Extensions.Add("gif")
                                s.Extensions.Add("psd")
                                s.Extensions.Add("mpg")
                                s.Extensions.Add("mpeg")
                                s.Extensions.Add("m2t")
                                s.Extensions.Add("tif")
                                s.Extensions.Add("tiff")
                                s.Extensions.Add("mpe")
                                s.Extensions.Add("tga")
                                s.Extensions.Add("js")
                                s.Extensions.Add("sfpreset")

                            Case "cool 3d"
                                s.Extensions.Add("c3d")
                                s.Extensions.Add("3ds")
                                s.Extensions.Add("png")
                                s.Extensions.Add("emf")
                                s.Extensions.Add("wmf")
                                s.Extensions.Add("mp3")
                                s.Extensions.Add("avi")
                                s.Extensions.Add("mov")
                                s.Extensions.Add("aiff")
                                s.Extensions.Add("wav")
                                s.Extensions.Add("au")

                            Case "media studio pro"
                                s.Extensions.Add("msp")
                                s.Extensions.Add("mp3")
                                s.Extensions.Add("avi")
                                s.Extensions.Add("mov")
                                s.Extensions.Add("pcm")
                                s.Extensions.Add("swf")
                                s.Extensions.Add("c3d")
                                s.Extensions.Add("gif")
                                s.Extensions.Add("flc")
                                s.Extensions.Add("bmp")
                                s.Extensions.Add("img")
                                s.Extensions.Add("jpg")
                                s.Extensions.Add("png")
                                s.Extensions.Add("tif")
                                s.Extensions.Add("wmf")

                            Case "premiere"
                                s.Extensions.Add("avi")
                                s.Extensions.Add("psd")
                                s.Extensions.Add("ai")
                                s.Extensions.Add("eps")
                                s.Extensions.Add("mpg")
                                s.Extensions.Add("mpeg")
                                s.Extensions.Add("mpe")
                                s.Extensions.Add("asf")
                                s.Extensions.Add("mov")
                                s.Extensions.Add("moov")
                                s.Extensions.Add("dlx")
                                s.Extensions.Add("wmv")
                                s.Extensions.Add("wma")
                                s.Extensions.Add("gif")
                                s.Extensions.Add("flc")
                                s.Extensions.Add("fli")
                                s.Extensions.Add("bmp")
                                s.Extensions.Add("rle")
                                s.Extensions.Add("dib")
                                s.Extensions.Add("flm")
                                s.Extensions.Add("pcx")
                                s.Extensions.Add("tif")
                                s.Extensions.Add("pic")
                                s.Extensions.Add("pct")
                                s.Extensions.Add("tga")
                                s.Extensions.Add("icb")
                                s.Extensions.Add("vst")
                                s.Extensions.Add("vda")
                                s.Extensions.Add("aif")
                                s.Extensions.Add("wav")
                                s.Extensions.Add("mp3")
                                s.Extensions.Add("pbl")
                                s.Extensions.Add("psq")
                                s.Extensions.Add("ppj")
                                s.Extensions.Add("ptl")
                                s.Extensions.Add("prtl")

                            Case "excel"
                                s.Extensions.Add("xls")
                                s.Extensions.Add("csv")

                                Dim versions As String() = {"2.1", "3.0", "4.0", "5.0", "95", "97", "XP", "2003"}
                                Dim v As AMP.Version
                                Dim vc As New VersionCollection

                                For x As Integer = 0 To versions.Length - 1
                                    v = New AMP.Version
                                    v.Number = versions(x)
                                    v.IconFile = "excel"
                                    vc.Add(v)
                                Next

                                s.Versions = vc
                        End Select
                        s.Extensions.Sort()
                        publisher.Software.Add(s)
                    End If
                    _softwareMap.Add(row.lSoftwareID, s.ID)

                End If
            Next
        End Sub

        '---COMMENT---------------------------------------------------------------
        '	get plugins as software entities
        '
        '	Date:		Name:	Description:
        '	12/26/04	JEA		Creation
        '   12/28/04    JEA     Use version collection instead of hashtable
        '   1/4/05      JEA     handle GUID id
        '-------------------------------------------------------------------------
        Private Sub LoadPlugins(ByVal rows As tblPluginsRow(), ByVal publisher As AMP.Publisher)
            Dim s As AMP.Software

            For Each row As tblPluginsRow In rows
                If Not _softwareMap.ContainsKey(_pluginMap(row.lPluginID)) Then
                    s = New AMP.Software

                    Dim versions As New AMP.VersionCollection

                    Select Case row.lPluginID
                        Case 1, 2, 3
                            s.Name = "Wax"
                            s.Url = "debugmode.com/wax/"
                            versions.Add(New AMP.Version("1.01", Nothing))
                            versions.Add(New AMP.Version("1.01", Nothing))
                            versions.Add(New AMP.Version("2.0", Nothing))
                            versions.Add(New AMP.Version("2.0c", Nothing))
                        Case 4, 5
                            s.Name = "WinMorph"
                            s.Url = "debugmode.com/winmorph/"
                            versions.Add(New AMP.Version("2.01", Nothing))
                            versions.Add(New AMP.Version("3.01", Nothing))
                        Case 6, 7, 8, 9, 10
                            s.Name = "SpiceFILTERS"
                            s.Url = "pixelan.com/sfilters/details.htm"
                        Case 16
                            s.Name = "Cubes"
                            s.Url = "www.endor.demon.co.uk"
                        Case 17, 18
                            s.Name = "SpiceMASTER"
                            s.Url = "pixelan.com/sm25/intro.htm"
                            versions.Add(New AMP.Version("1.2", Nothing))
                            versions.Add(New AMP.Version("2.0", Nothing))
                            versions.Add(New AMP.Version("2.5", Nothing))
                        Case 19
                            s.Name = "Luminance"
                            s.Url = "www.endor.demon.co.uk"
                        Case 20
                            s.Name = "Red"
                            s.Url = "borisfx.com/products/RED/index_3gl.php"
                            versions.Add(New AMP.Version("2.1", Nothing))
                            versions.Add(New AMP.Version("2.5", Nothing))
                            versions.Add(New AMP.Version("3GL", Nothing))
                        Case 21
                            s.Name = "Graffiti"
                            s.Url = "borisfx.com/products/GRAFFITI/"
                            versions.Add(New AMP.Version("2.0", Nothing))
                            versions.Add(New AMP.Version("3.0", Nothing))
                    End Select

                    If s.Name <> Nothing Then
                        If publisher.Software.WithName(s.Name) Is Nothing Then
                            s.Free = row.tblPluginPackagesRow.bFree
                            s.Publisher = publisher
                            s.Type = Software.Types.VideoPlugin
                            s.Versions = versions
                            s.Versions.Sort()
                            publisher.Software.Add(s)
                        Else
                            s = publisher.Software.WithName(s.Name)
                        End If
                        _softwareMap.Add(_pluginMap(row.lPluginID), s.ID)
                    End If

                End If
            Next
        End Sub

        '---COMMENT---------------------------------------------------------------
        '	get versions collection for software
        '
        '	Date:		Name:	Description:
        '	12/26/04	JEA		Creation
        '   12/28/04    JEA     Use version collection instead of hashtable
        '-------------------------------------------------------------------------
        Private Function LoadVersions(ByVal rows As tblSoftwareVersionsRow()) As AMP.VersionCollection
            Dim versions As New AMP.VersionCollection
            For Each row As tblSoftwareVersionsRow In rows
                If row.vsVersionText <> "(any)" Then
                    Dim icon As String
                    If Not row("vsIcon").Equals(DBNull.Value) Then icon = row.vsIcon
                    versions.Add(New AMP.Version(row.vsVersionText, icon))
                End If
            Next
            versions.Sort()
            Return versions
        End Function

        Private Function LoadAssetTypes(ByVal rows As tblPublisherItemTypesRow()) As Integer
            Dim forType As Integer = 0

            For Each row As tblPublisherItemTypesRow In rows
                forType = (forType Or Me.GetAssetType(row.lItemTypeID))
            Next
            Return forType
        End Function

#End Region

#Region " Assets "

        '---COMMENT---------------------------------------------------------------
        '	build asset collection from typed data set
        '
        '	Date:		Name:	Description:
        '	11/29/04	JEA		Creation
        '   1/15/05     JEA     Added stronger filtering
        '   1/23/05     JEA     Handle presets
        '   2/10/05     JEA     Check type as well as title in duplicate check
        '-------------------------------------------------------------------------
        Public Function LoadAssets() As Integer
            Dim a As AMP.Asset
            Dim added As Integer = 0
            Dim preset As Boolean
            Dim newAsset As Boolean
            Dim name As String

            ' transfer assets to collection
            For Each row As AssetItemsRow In _ds.AssetItems
                preset = False
                newAsset = False

                If preset Then
                    With row.GettblProjectsRows(0)
                        name = .vsFriendlyName.Replace("_", "").Trim
                        a = WebSite.Assets.WithTitle(name)

                        If a Is Nothing Then
                            a = WebSite.Assets.WithTitle(Format.NormalSpacing(name))
                        End If

                        If a Is Nothing OrElse a.Type <> Asset.Types.Preset Then
                            ' new asset
                            a = New AMP.Asset
                            newAsset = True
                            a.Title = name
                            a.Type = Asset.Types.Preset
                            If Not .Item("vsDescription").Equals(DBNull.Value) Then
                                a.Description = .vsDescription
                            End If
                            If .tblUsersRow Is Nothing Then
                                BugOut("No user row for preset {0}", a.Title)
                            End If
                            a.SubmittedBy = LoadPerson(.tblUsersRow)
                            a.SubmitDate = .dtSubmitDate
                            a.Status = CType(.lStatusID, AMP.Site.Status)

                            ' file
                            Dim file As New AMP.File
                            file.Name = .vsFileName
                            file.Path = .vsPath.Replace("files", "")
                            If Not .Item("vsRequiredMediaURL").Equals(DBNull.Value) Then
                                file.RequiredUrl = .vsRequiredMediaURL
                            End If
                            file.Type = AMP.File.Types.Preset
                            file.Plugins = LoadPlugins(.GettblProjectPluginsRows)
                            file.Software = LoadSoftware(.tblSoftwareVersionsRow, _
                                file.SoftwareVersion, a.Section)
                            a.File = file

                            ' author
                            a.AuthoredBy = a.SubmittedBy
                        End If

                        a.File.Downloads = .lDownloads

                        If .Item("lVersionCount").Equals(DBNull.Value) Then
                            a.Version = 1
                        Else
                            a.Version = .lVersionCount
                        End If

                        If .Item("dtVersionDate").Equals(DBNull.Value) Then
                            a.VersionDate = a.SubmitDate
                        Else
                            a.VersionDate = .dtVersionDate
                        End If

                    End With
                Else
                    Select Case CType(row.lItemTypeID, Import.ItemType)
                        Case ItemType.Project
                            With row.GettblProjectsRows(0)
                                name = .vsFriendlyName.Replace("_", " ").Trim
                                a = WebSite.Assets.WithTitle(name)

                                If a Is Nothing Then
                                    a = WebSite.Assets.WithTitle(Format.NormalSpacing(name))
                                End If

                                If a Is Nothing OrElse a.Type <> Asset.Types.Project Then
                                    ' new asset
                                    a = New AMP.Asset
                                    newAsset = True
                                    a.Title = name
                                    a.Type = Asset.Types.Project
                                    If Not .Item("vsDescription").Equals(DBNull.Value) Then
                                        a.Description = .vsDescription
                                    End If
                                    If .tblUsersRow Is Nothing Then
                                        BugOut("No user row for project {0}", a.Title)
                                    End If
                                    a.SubmittedBy = LoadPerson(.tblUsersRow)
                                    a.SubmitDate = .dtSubmitDate
                                    a.Status = CType(.lStatusID, AMP.Site.Status)

                                    ' file
                                    Dim file As New AMP.File
                                    file.Name = .vsFileName
                                    file.Path = .vsPath.Replace("files", "")
                                    If Not .Item("vsRenderedURL").Equals(DBNull.Value) Then
                                        file.RenderedUrl = .vsRenderedURL
                                    End If
                                    If Not .Item("vsRequiredMediaURL").Equals(DBNull.Value) Then
                                        file.RequiredUrl = .vsRequiredMediaURL
                                    End If
                                    file.Type = CType(.lFormatID, AMP.File.Types)
                                    file.Plugins = LoadPlugins(.GettblProjectPluginsRows)
                                    file.Software = LoadSoftware(.tblSoftwareVersionsRow, _
                                        file.SoftwareVersion, a.Section)
                                    a.File = file

                                    ' author
                                    a.AuthoredBy = a.SubmittedBy
                                End If

                                a.File.Downloads = .lDownloads

                                If .Item("lVersionCount").Equals(DBNull.Value) Then
                                    a.Version = 1
                                Else
                                    a.Version = .lVersionCount
                                End If

                                If .Item("dtVersionDate").Equals(DBNull.Value) Then
                                    a.VersionDate = a.SubmitDate
                                Else
                                    a.VersionDate = .dtVersionDate
                                End If

                            End With

                        Case ItemType.Script
                            With row.GettblScriptsRows(0)
                                name = .vsFriendlyName.Replace("_", "").Trim
                                a = WebSite.Assets.WithTitle(name)

                                If a Is Nothing Then
                                    a = WebSite.Assets.WithTitle(Format.NormalSpacing(name))
                                End If

                                If a Is Nothing OrElse a.Type <> Asset.Types.Script Then
                                    ' new asset
                                    a = New AMP.Asset
                                    newAsset = True
                                    a.Title = name
                                    a.Type = Asset.Types.Script
                                    If Not .Item("vsDescription").Equals(DBNull.Value) Then
                                        a.Description = .vsDescription
                                    End If
                                    If .tblUsersRow Is Nothing Then
                                        BugOut("No user row for script {0}", a.Title)
                                    End If
                                    a.SubmittedBy = LoadPerson(.tblUsersRow)
                                    a.SubmitDate = .dtSubmitDate
                                    a.Status = CType(.lStatusID, AMP.Site.Status)

                                    ' file
                                    Dim file As New AMP.File
                                    file.Name = .vsFileName
                                    file.Path = .vsPath.Replace("files", "")
                                    file.RequiredUrl = .vsRequiredMediaURL
                                    file.Type = CType(.lFormatID, File.Types)
                                    file.Software = LoadSoftware(.tblSoftwareVersionsRow, _
                                        file.SoftwareVersion, a.Section)
                                    a.File = file

                                    ' author
                                    a.AuthoredBy = a.SubmittedBy
                                End If

                                a.File.Downloads = .lDownloads

                                If .Item("lVersionCount").Equals(DBNull.Value) Then
                                    a.Version = 1
                                Else
                                    a.Version = .lVersionCount
                                End If

                                If .Item("dtVersionDate").Equals(DBNull.Value) Then
                                    a.VersionDate = a.SubmitDate
                                Else
                                    a.VersionDate = .dtVersionDate
                                End If
                            End With

                        Case ItemType.Tutorial
                            With row.GettblTutorialsRows(0)
                                name = .vsTutorialName.Replace("_", "").Trim
                                a = WebSite.Assets.WithTitle(name)

                                If a Is Nothing Then
                                    a = WebSite.Assets.WithTitle(Format.NormalSpacing(name))
                                End If

                                If a Is Nothing OrElse a.Type <> Asset.Types.Tutorial Then
                                    ' new asset
                                    a = New AMP.Asset
                                    newAsset = True
                                    a.Title = name
                                    a.Type = Asset.Types.Tutorial
                                    If Not .Item("vsDescription").Equals(DBNull.Value) Then
                                        a.Description = .vsDescription
                                    End If
                                    If .tblUsersRowByUsersToTutorialsSubmitter Is Nothing Then
                                        BugOut("No user row for tutorial {0}", a.Title)
                                    End If
                                    a.SubmittedBy = LoadPerson(.tblUsersRowByUsersToTutorialsSubmitter)
                                    a.SubmitDate = .dtSubmitDate
                                    a.Version = 1
                                    a.VersionDate = a.SubmitDate
                                    a.Status = CType(.lStatusID, AMP.Site.Status)

                                    ' link
                                    Dim link As New AMP.Link
                                    link.Url = .vsTutorialURL
                                    a.Link = link

                                    ' author
                                    If .tblUsersRowByUsersToTutorialsAuthor Is Nothing Then
                                        BugOut("No author row for tutorial {0}", a.Title)
                                    End If
                                    a.AuthoredBy = LoadPerson(.tblUsersRowByUsersToTutorialsAuthor)
                                    If a.AuthoredBy Is Nothing Then a.AuthoredBy = a.SubmittedBy
                                End If
                            End With

                        Case ItemType.Review
                            With row.GettblReviewsRows(0)
                                name = .vsReviewName.Replace("_", "").Trim
                                a = WebSite.Assets.WithTitle(name)

                                If a Is Nothing Then
                                    a = WebSite.Assets.WithTitle(Format.NormalSpacing(name))
                                End If

                                If a Is Nothing OrElse a.Type <> Asset.Types.Review Then
                                    ' new asset
                                    a = New AMP.Asset
                                    newAsset = True
                                    a.Title = name
                                    a.Type = Asset.Types.Review
                                    If Not .Item("vsDescription").Equals(DBNull.Value) Then
                                        a.Description = .vsDescription
                                    End If
                                    a.SubmittedBy = LoadPerson(.tblUsersRowByUsersToReviewsSubmitter)
                                    a.SubmitDate = .dtSubmitDate
                                    a.Version = 1
                                    a.VersionDate = a.SubmitDate
                                    a.Status = CType(.lStatusID, AMP.Site.Status)

                                    ' link
                                    Dim link As New AMP.Link
                                    link.Url = .vsReviewURL
                                    a.Link = link

                                    ' author
                                    Dim author As New AMP.Person

                                    If .Item("vsAuthorName").Equals(DBNull.Value) Then
                                        author = a.SubmittedBy
                                    Else
                                        author.FirstName = .vsAuthorName
                                    End If

                                    a.AuthoredBy = author
                                End If
                            End With
                    End Select
                End If

                _assetMap.Add(row.AssetID, a.ID)

                If newAsset Then
                    BugCheck(a.Section = 0 AndAlso a.Status = Status.Approved, """{0}"" asset ({1}) has no section", a.Title, row.AssetID)
                    If row.AssetID = 96 AndAlso a.Section = 0 Then a.Section = Section.Vegas

                    If a.Status = Status.Approved OrElse DateDiff(DateInterval.Month, a.SubmitDate, DateTime.Now) < 2 Then
                        ' only add approved or recently submitted assets
                        added += 1

                        a.Ratings = LoadRatings(row.GettblRankingsRows)

                        ' item site info
                        For Each s As tblItemSitesRow In row.GettblItemSitesRows
                            Select Case s.lSiteID
                                Case 1  ' Vegas
                                    a.Section = a.Section Or Section.Vegas
                                Case 2  ' Ulead
                                    a.Section = a.Section Or Section.Ulead
                                Case 3  ' Adobe
                                    a.Section = a.Section Or Section.Adobe
                                Case Else
                                    BugOut("Asset {0} has no section for site {1}", row.AssetID, s.lSiteID)
                            End Select
                        Next

                        ' categories
                        For Each r As tblItemCategoriesRow In row.GettblItemCategoriesRows
                            a.Categories.Add(WebSite.Categories( _
                                _categoryMap(r.tblCategoriesRow.lCategoryID).ToString))
                            If r.tblCategoriesRow.lCategoryID = 49 Then preset = True
                        Next
                        a.Categories.Sort()

                        WebSite.Assets.Add(a)
                    End If
                End If
            Next

            Return added
        End Function

        '---COMMENT---------------------------------------------------------------
        '	build rating collection from data row array
        '
        '	Date:		Name:	Description:
        '	11/23/04	JEA		Creation
        '-------------------------------------------------------------------------
        Private Function LoadRatings(ByVal rows As tblRankingsRow()) As AMP.RatingCollection
            Dim r As AMP.Rating
            Dim rc As New AMP.RatingCollection

            For Each row As tblRankingsRow In rows
                r = New AMP.Rating
                If row.tblUsersRow Is Nothing Then
                    BugOut("No user row for rating")
                End If
                r.Person = LoadPerson(row.tblUsersRow)
                r.Value = row.lRank
                r.Date = row.dtRankDate
                r.Comment = row.vsComment
                rc.Add(r)
            Next
            Return rc
        End Function

        '---COMMENT---------------------------------------------------------------
        '	get software for asset from publisher collection
        '
        '	Date:		Name:	Description:
        '	12/26/04	JEA		Creation
        '   12/28/04    JEA     convert ID 9, graphic, to actual software
        '-------------------------------------------------------------------------
        Private Function LoadSoftware(ByVal row As tblSoftwareVersionsRow, _
            ByRef version As AMP.Version, ByVal section As Integer) As AMP.Software

            Dim s As AMP.Software

            Dim softwareID As Integer = row.lSoftwareID
            If softwareID = 9 Then
                If (section And WebSite.Section.Ulead) > 0 Then
                    softwareID = 5
                    'BugOut("Converted graphic to Cool 3D")
                ElseIf (section And WebSite.Section.Sony) > 0 Then
                    softwareID = 1
                    'BugOut("Converted graphic to Vegas")
                Else
                    softwareID = 1
                    'BugOut("No conversion for {0}", section.ToString)
                End If
            End If

            Try
                s = WebSite.Publishers.SoftwareWithID(_softwareMap(softwareID).ToString)
                version = s.Versions(row.vsVersionText)
                If version Is Nothing Then version = s.Versions.Latest
            Catch
                BugOut("No software for {0}", softwareID)
            End Try

            Return s
        End Function

        '---COMMENT---------------------------------------------------------------
        '	build plugin collection from data row array
        '
        '	Date:		Name:	Description:
        '	11/29/04	JEA		Creation
        '-------------------------------------------------------------------------
        Private Function LoadPlugins(ByVal rows As tblProjectPluginsRow()) As AMP.SoftwareCollection
            Dim s As AMP.Software
            Dim sc As New AMP.SoftwareCollection

            For Each row As tblProjectPluginsRow In rows
                s = WebSite.Publishers.SoftwareWithID(_softwareMap(_pluginMap(row.lPluginID)).ToString)
                'If s Is Nothing Then s = WebSite.Publishers.SoftwareWithName(row.
                sc.Add(s)
            Next

            Return sc
        End Function

        '---COMMENT---------------------------------------------------------------
        '	build person entity from data row
        '
        '	Date:		Name:	Description:
        '	11/23/04	JEA		Creation
        '   11/30/04    JEA     Get from PersonCollection
        '   1/4/05      JEA     Check if user id is mapped to person entity
        '-------------------------------------------------------------------------
        Private Function LoadPerson(ByVal row As tblUsersRow) As AMP.Person
            Dim person As AMP.Person
            If (Not row Is Nothing) AndAlso _personMap.ContainsKey(row.lUserID) Then
                person = WebSite.Persons(_personMap(row.lUserID).ToString)
            End If
            Return person
        End Function

#End Region

#Region " Contests "

        Private Function LoadContests() As Integer
            Dim c As AMP.Contest
            Dim added As Integer = 0

            For Each row As tblContestsRow In _ds.tblContests
                c = WebSite.Contests.WithTitle(row.vsContestName)
                If c Is Nothing Then
                    c = New AMP.Contest

                    c.Title = row.vsContestName
                    c.Description = row.vsDescription
                    c.Start = row.dtStartDate
                    c.Finish = row.dtEndDate
                    c.FinishVote = row.dtVoteByDate
                    c.VotesAllowed = row.lVotesAllowed
                    c.WeightFactor = row.lWeightFactor
                    c.WinnerCount = row.lWinners

                    If Not row.Item("lSiteID").Equals(DBNull.Value) Then
                        Select Case row.lSiteID
                            Case 1  ' Vegas
                                c.Section = Section.Vegas
                            Case 2  ' Ulead
                                c.Section = Section.Ulead
                            Case 3  ' Adobe
                                c.Section = Section.Adobe
                        End Select
                    End If

                    c.Entry = LoadEntries(row.GettblContestItemsRows)

                    WebSite.Contests.Add(c)
                    added += 1
                End If
                _contestMap.Add(row.lContestID, c.ID)
                ' if there are active contests then we need to add new entries
                ' and votes to the existing contest
            Next
            Return added
        End Function

        Private Function LoadEntries(ByVal rows As tblContestItemsRow()) As ContestEntryCollection
            Dim e As AMP.ContestEntry
            Dim ec As New AMP.ContestEntryCollection

            For Each row As tblContestItemsRow In rows
                e = New AMP.ContestEntry
                e.Asset = WebSite.Assets(_assetMap(row.AssetItemsRowParent.AssetID).ToString)
                e.Date = row.dtAddedDate
                e.Votes = LoadVotes(row.GettblContestVotesRows)
                ec.Add(e)
            Next
            Return ec
        End Function

        Private Function LoadVotes(ByVal rows As tblContestVotesRow()) As ContestVoteCollection
            Dim v As AMP.ContestVote
            Dim vc As New AMP.ContestVoteCollection
            For Each row As tblContestVotesRow In rows
                v = New AMP.ContestVote
                v.Person = WebSite.Persons(_personMap(row.lVoterID).ToString)
                v.Date = row.dtDateVoted
                v.Rank = row.lRank
                vc.Add(v)
            Next
            Return vc
        End Function

#End Region

#Region " Data Load "

        '---COMMENT---------------------------------------------------------------
        '	load data from storage when expired from cache
        '
        '	Date:		Name:	Description:
        '	11/7/04		JEA		Creation
        '-------------------------------------------------------------------------
        Private Function GetDataSet() As AMP.LegacyDataSet
            Dim asset As LegacyDataSet.AssetItemsRow
            Dim tableDefinitions As New Hashtable
            Dim ds As New AMP.LegacyDataSet
            Dim db As AMP.Data.Jet

            ds.EnforceConstraints = False

            BugOut("Loading legacy DataSet")
            BugTab()

            ' load asset data
            With tableDefinitions
                .Add("tblCategories", "*")
                .Add("tblCategoryItemTypes", "*")
                '.Add("tblComputedRankings", "*")
                '.Add("tblComputedVotePoints", "*")
                .Add("tblContestItems", "*")
                .Add("tblContests", "*")
                .Add("tblContestVotes", "*")
                .Add("tblFormats", "*")
                '.Add("tblForumHosts", "*")
                '.Add("tblForums", "*")
                .Add("tblItemCategories", "*")
                .Add("tblItemSites", "*")
                .Add("tblItemTypes", "*")
                '.Add("tblLog", "*")
                '.Add("tblLoggedActivities", "*")
                .Add("tblPluginPackages", "*")
                .Add("tblPlugins", "*")
                .Add("tblPluginTypes", "*")
                .Add("tblProjectPlugins", "*")
                .Add("tblProjects", "*")
                .Add("tblPublisherItemTypes", "*")
                .Add("tblPublishers", "*")
                '.Add("tblQueries", "*")
                .Add("tblRankings", "*")
                .Add("tblReviews", "*")
                .Add("tblScripts", "*")
                '.Add("tblSettings", "*")
                .Add("tblSite", "*")
                .Add("tblSoftware", "*")
                .Add("tblSoftwareVersions", "*")
                '.Add("tblSortValues", "*")
                .Add("tblStatusCodes", "*")
                '.Add("tblTemporaryFolders", "*")
                .Add("tblTutorials", "*")
                '.Add("tblTutorialSites", "*")
                .Add("tblUsers", "*")
                .Add("tblUserTypes", "*")
            End With

            db = New AMP.Data.Jet(AppSettings("LegacyStore"))
            db.FillDataset(DirectCast(ds, DataSet), tableDefinitions)
            db = Nothing

            BugOut("Adding custom rows ...")
            BugTab()

            ' add needed row to tblSite
            BugOut("tblSite")
            Dim site As LegacyDataSet.tblSiteRow
            site = ds.tblSite.NewtblSiteRow
            site.lSiteID = 0
            site.vsSiteName = "Unspecified"
            ds.tblSite.AddtblSiteRow(site)

            ' add needed row to tblFormats
            BugOut("tblFormats")
            Dim format As LegacyDataSet.tblFormatsRow
            format = ds.tblFormats.NewtblFormatsRow
            format.lFormatID = 0
            format.vsFormatDescription = "Unspecified"
            ds.tblFormats.AddtblFormatsRow(format)

            ' add needed row to tblPublisher
            BugOut("tblPublishers")
            Dim pub As LegacyDataSet.tblPublishersRow
            pub = ds.tblPublishers.NewtblPublishersRow
            pub.lPublisherID = 0
            pub.vsPublisherName = "Unspecified"
            ds.tblPublishers.AddtblPublishersRow(pub)

            ' add needed row to tblSoftware
            BugOut("tblSoftware")
            Dim soft As LegacyDataSet.tblSoftwareRow
            soft = ds.tblSoftware.NewtblSoftwareRow
            soft.lSoftwareID = 3
            soft.vsSoftwareName = "Unspecified"
            soft.lPublisherID = 1
            ds.tblSoftware.AddtblSoftwareRow(soft)
            soft = ds.tblSoftware.NewtblSoftwareRow
            soft.lSoftwareID = 2
            soft.vsSoftwareName = "Unspecified"
            soft.lPublisherID = 1
            ds.tblSoftware.AddtblSoftwareRow(soft)

            ' add needed row to tblUser
            BugOut("tblUsers")
            Dim user As LegacyDataSet.tblUsersRow
            user = ds.tblUsers.NewtblUsersRow
            user.lUserID = 0
            user.vsFirstName = "Unspecified"
            ds.tblUsers.AddtblUsersRow(user)
            user = ds.tblUsers.NewtblUsersRow
            user.lUserID = 2522
            user.vsFirstName = "Unspecified"
            ds.tblUsers.AddtblUsersRow(user)

            ds.AcceptChanges()

            BugUntab()
            BugOut("Deleting bad rows ...")
            BugTab()

            ' delete forum rankings
            BugOut("tblRankings")
            Dim dv As New DataView(ds.tblRankings)
            dv.RowFilter = "lItemTypeID = 3"
            For x As Integer = dv.Count - 1 To 0 Step -1
                dv(x).Delete()
            Next

            ' delete rankings from removed users
            dv = New DataView(ds.tblRankings)
            dv.RowFilter = "lUserID = 5037"
            For x As Integer = dv.Count - 1 To 0 Step -1
                dv(x).Delete()
            Next

            ' delete invalid projects
            BugOut("tblProjects")
            dv = New DataView(ds.tblProjects)
            dv.RowFilter = "lSoftwareVersionID = 0"
            For x As Integer = dv.Count - 1 To 0 Step -1
                dv(x).Delete()
            Next

            ' delete orphans and obsolete types from site
            BugOut("tblItemSites")
            dv = New DataView(ds.tblItemSites)
            dv.RowFilter = "lItemTypeID IN (3,4,5)" _
                & " OR (lItemTypeID = 1 AND lItemID IN (79,116,194,357))" _
                & " OR (lItemTypeID = 2 AND lItemID IN (116))" _
                & " OR (lItemTypeID = 6 AND lItemID in (357))"
            For x As Integer = dv.Count - 1 To 0 Step -1
                dv(x).Delete()
            Next

            ' delete orphans and obsolete types from categories
            BugOut("tblItemCategories")
            dv = New DataView(ds.tblItemCategories)
            dv.RowFilter = "lItemTypeID IN (3,4,5)" _
                & " OR (lItemTypeID = 1 AND lItemID IN (85,103,119,120,121))" _
                & " OR (lItemTypeID = 2 AND lItemID IN (66))" _
                & " OR (lItemTypeID = 6 AND lItemID in (357))"
            For x As Integer = dv.Count - 1 To 0 Step -1
                dv(x).Delete()
            Next

            ' delete orphan categories
            BugOut("tblCategoryItemTypes")
            dv = New DataView(ds.tblCategoryItemTypes)
            dv.RowFilter = "lCategoryID IN (36,37)"
            For x As Integer = dv.Count - 1 To 0 Step -1
                dv(x).Delete()
            Next

            ' delete orphan votes
            BugOut("tblContestVotes")
            dv = New DataView(ds.tblContestVotes)
            dv.RowFilter = "lVoterID IN (54, 5037)"
            For x As Integer = dv.Count - 1 To 0 Step -1
                dv(x).Delete()
            Next

            ' delete orphan contest items
            BugOut("tblContestItems")
            dv = New DataView(ds.tblContestItems)
            dv.RowFilter = "lItemID = 16 AND lItemTypeID = 1"
            For x As Integer = dv.Count - 1 To 0 Step -1
                dv(x).Delete()
            Next

            ' delete unused plugin packages
            BugOut("tblPluginPackages")
            dv = New DataView(ds.tblPluginPackages)
            dv.RowFilter = "lPluginPackageID = 5"
            For x As Integer = dv.Count - 1 To 0 Step -1
                dv(x).Delete()
            Next

            ds.AcceptChanges()

            BugUntab()
            BugOut("Building Asset table ...")
            BugTab()

            ' build asset table
            ds.AssetItems.BeginLoadData()
            ds.AssetItems.Clear()
            ' projects
            BugOut("projects")
            For Each row As tblProjectsRow In ds.tblProjects
                asset = ds.AssetItems.NewAssetItemsRow
                asset.lItemID = row.lProjectID
                asset.lItemTypeID = ItemType.Project
                ds.AssetItems.AddAssetItemsRow(asset)
                row.AssetID = asset.AssetID
            Next
            ' scripts
            BugOut("scripts")
            For Each row As tblScriptsRow In ds.tblScripts
                asset = ds.AssetItems.NewAssetItemsRow
                asset.lItemID = row.lScriptID
                asset.lItemTypeID = ItemType.Script
                ds.AssetItems.AddAssetItemsRow(asset)
                row.AssetID = asset.AssetID
            Next
            ' tutorials
            BugOut("tutorials")
            For Each row As tblTutorialsRow In ds.tblTutorials
                asset = ds.AssetItems.NewAssetItemsRow
                asset.lItemID = row.lTutorialID
                asset.lItemTypeID = ItemType.Tutorial
                ds.AssetItems.AddAssetItemsRow(asset)
                row.AssetID = asset.AssetID
            Next
            ' reviews
            BugOut("reviews")
            For Each row As tblReviewsRow In ds.tblReviews
                asset = ds.AssetItems.NewAssetItemsRow
                asset.lItemID = row.lReviewID
                asset.lItemTypeID = ItemType.Review
                ds.AssetItems.AddAssetItemsRow(asset)
                row.AssetID = asset.AssetID
            Next
            ds.AssetItems.EndLoadData()

            ds.AcceptChanges()

            BugUntab()

            Try
                BugOut("Enabling constraints")
                ds.EnforceConstraints = True
            Catch
                BugOut("Unable to enable constraints")
                BugTab()
                Dim rowsInError As DataRow()
                Dim col As DataColumn
                For Each dt As DataTable In ds.Tables
                    If dt.HasErrors Then
                        BugOut("Errors in {0}", dt.TableName)
                        BugTab()

                        ' Get an array of all rows with errors.
                        rowsInError = dt.GetErrors()
                        For x As Integer = 0 To rowsInError.Length - 1
                            BugOut(rowsInError(x).RowError)
                        Next
                        BugUntab()
                    End If
                Next
                BugUntab()
            Finally
                BugUntab()
            End Try

            ' get array of all users that must be kept
            Dim distinct As New Hashtable

            BugOut("Getting all users that must be kept")
            BugTab()

            BugOut("from projects")
            dv = New DataView(ds.tblProjects)
            dv.RowFilter = "lStatusID = 2"
            Me.Distinct(dv, "lUserID", distinct)

            BugOut("from scripts")
            dv = New DataView(ds.tblScripts)
            dv.RowFilter = "lStatusID = 2"
            Me.Distinct(dv, "lUserID", distinct)

            BugOut("from reviews")
            dv = New DataView(ds.tblReviews)
            dv.RowFilter = "lStatusID = 2"
            Me.Distinct(dv, "lAuthorID", distinct)

            BugOut("from tutorials")
            dv = New DataView(ds.tblTutorials)
            dv.RowFilter = "lStatusID = 2"
            Me.Distinct(dv, "lAuthorID", distinct)

            BugOut("from contests")
            dv = New DataView(ds.tblContestVotes)
            Me.Distinct(dv, "lVoterID", distinct)

            BugOut("from rankings")
            dv = New DataView(ds.tblRankings)
            Me.Distinct(dv, "lUserID", distinct)
            BugUntab()

            ReDim _keepUsers(distinct.Count)
            Dim y As Integer = 0
            For Each key As Integer In distinct.Keys
                _keepUsers(y) = key
                y += 1
            Next
            BugOut("Sorting users array")
            Array.Sort(_keepUsers)

            Return ds
        End Function

        '---COMMENT---------------------------------------------------------------
        '	add distinct values from column to arraylist
        '
        '	Date:		Name:	Description:
        '	11/7/04		JEA		Creation
        '-------------------------------------------------------------------------
        Private Sub Distinct(ByVal dv As DataView, ByVal column As String, ByRef distinct As Hashtable)
            Dim value As Object

            For x As Integer = 0 To dv.Count - 1
                value = dv(x).Row.Item(column)
                If CInt(value) <> 0 AndAlso CInt(value) <> 2522 AndAlso _
                    Not distinct.ContainsKey(value) Then

                    distinct.Add(value, value)
                End If
            Next
        End Sub

#End Region

    End Class
End Namespace
