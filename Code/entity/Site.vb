Imports System.Threading
Imports System.Web.HttpContext
Imports AMP.Data.Serialization
Imports System.Runtime.Serialization
Imports System.Configuration.ConfigurationSettings

<Serializable()> _
Public Class Site
    Implements ICloneable, IDeserializationCallback

    ' don't serialize fields that could be in an odd state
    <NonSerialized()> Private _saveDelay As TimeSpan
    <NonSerialized()> Private _bufferSave As Boolean = True
    <NonSerialized()> Private _saveState As SaveState = SaveState.None

    Private _assets As New AMP.AssetCollection
    Private _persons As New AMP.PersonCollection
    Private _contests As New AMP.ContestCollection
    Private _roles As New AMP.RoleCollection
    Private _catalog As New AMP.Catalog
    Private _tours As New AMP.TourCollection
    Private _publishers As New AMP.PublisherCollection
    Private _vendors As New AMP.VendorCollection
    Private _categories As New AMP.CategoryCollection

#Region " Enumerations "

    <Flags()> _
    Private Enum SaveState
        None = &H0
        Sleeping = &H1
        Writing = &H2
    End Enum

    <Flags()> _
    Public Enum Entity
        Asset = &H1
        Person = &H2
        Tour = &H4
        Contest = &H8
        Product = &H10
        Software = &H20
        Publisher = &H40
    End Enum

    <Flags()> _
    Public Enum Section
        All = Adobe Or Sony Or Ulead
        Sony = Vegas Or DvdArchitect
        Ulead = Cool3D Or MediaStudio
        None = &H0
        Vegas = &H1
        DvdArchitect = &H2
        Acid = &H4
        SoundForge = &H8
        Cool3D = &H10
        MediaStudio = &H20
        Adobe = &H100
        HDV = &H200
    End Enum

    Public Enum Status
        Pending = 1
        Approved = 2
        Rejected = 3
        Disabled = 4
        Anonymous = 5
    End Enum

    Public Enum Role
        Anonymous
        UnverifiedGuest
        VerifiedGuest
        Editor
        Manager
        Administrator
    End Enum

    Public Enum Activity
        ' site
        StartedApplication = 29
        ' user
        Login = 0
        Register = 1
        FailedLogin = 2
        EmailedPassword = 3
        AutoLoginFromCookie = 4
        ResentValidationCode = 5
        ValidateRegistration = 6
        EnteredBadValidationCode = 7
        UnauthorizedAccessAttempt = 8
        TriedRegisteringWithExistingEmail = 9
        TriedDisabledAccountLogin = 28
        Unsubscribed = 31
        ' asset
        RankAsset = 10
        EditAsset = 11
        DenySubmittedAsset = 12
        ApproveSubmittedAsset = 13
        DeleteAsset = 30
        ' file
        FileUpload = 14
        DeleteFile = 15
        FileDownload = 16
        AttemptedLargeFileUpload = 17
        ' link
        ViewLink = 18
        ' contest
        Vote = 19
        ApproveContestEntry = 32
        DenyContestEntry = 33
        EnterAsset = 20
        SaveContest = 21
        CreateContest = 22
        ' database
        BackupDatabase = 23
        CompactDatabase = 24
        ' accounts
        EditUserAccount = 25
        CreateUserAccount = 26
        DisableUserAccount = 27
    End Enum

    Public Enum Permission
        ' site
        ViewSiteStatistics = 35
        DownloadDataFile = 36
        EditCategories = 39
        EditRoles = 40
        ' user
        AddUser = 0
        EditAnyUser = 1
        DeleteUser = 2
        EditMyself = 3
        ViewUserDetails = 4
        ' product
        AddProduct = 5
        EditProduct = 6
        DeleteProduct = 7
        PurchaseProduct = 8
        ' assets
        AddAsset = 9
        ApproveAsset = 14
        ChangeAssetSection = 16
        DownloadAsset = 12
        DeleteAnyAsset = 34
        DeleteMyAsset = 33
        EditMyAsset = 10
        EditAnyAsset = 11
        LinkToAsset = 13
        RejectAsset = 15
        ' contests
        AddContest = 17
        EditContest = 18
        EnterContest = 19
        VoteInContest = 20
        ApproveEntry = 37
        RejectEntry = 38
        ' menu
        AddMenuItem = 21
        EditMenuItem = 22
        DeleteMenuItem = 23
        ' feature
        AddFeaturedItem = 24
        EditFeaturedItem = 25
        DeleteFeaturedItem = 26
        ' quote
        AddQuote = 27
        EditQuote = 28
        DeleteQuote = 29
        ' general content
        AddContent = 30
        EditContent = 31
        DeleteContent = 32
    End Enum

#End Region

#Region " Properties "

    Public Property Categories() As AMP.CategoryCollection
        Get
            _categories.Sort()
            Return _categories
        End Get
        Set(ByVal Value As AMP.CategoryCollection)
            _categories = Value
        End Set
    End Property

    Public Property Vendors() As AMP.VendorCollection
        Get
            Return _vendors
        End Get
        Set(ByVal Value As AMP.VendorCollection)
            _vendors = Value
        End Set
    End Property

    Public Property Publishers() As AMP.PublisherCollection
        Get
            Return _publishers
        End Get
        Set(ByVal Value As AMP.PublisherCollection)
            _publishers = Value
        End Set
    End Property

    Public Property Tours() As AMP.TourCollection
        Get
            Return _tours
        End Get
        Set(ByVal Value As AMP.TourCollection)
            _tours = Value
        End Set
    End Property

    Public Property Catalog() As AMP.Catalog
        Get
            Return _catalog
        End Get
        Set(ByVal Value As AMP.Catalog)
            _catalog = Value
        End Set
    End Property

    Public Property UseSaveBuffer() As Boolean
        Get
            Return _bufferSave
        End Get
        Set(ByVal Value As Boolean)
            _bufferSave = Value
        End Set
    End Property

    Public Property Roles() As AMP.RoleCollection
        Get
            Return _roles
        End Get
        Set(ByVal Value As AMP.RoleCollection)
            _roles = Value
        End Set
    End Property

    Public Property Assets() As AMP.AssetCollection
        Get
            Return _assets
        End Get
        Set(ByVal Value As AMP.AssetCollection)
            _assets = Value
        End Set
    End Property

    Public Property Persons() As AMP.PersonCollection
        Get
            Return _persons
        End Get
        Set(ByVal Value As AMP.PersonCollection)
            _persons = Value
        End Set
    End Property

    Public Property Contests() As AMP.ContestCollection
        Get
            Return _contests
        End Get
        Set(ByVal Value As AMP.ContestCollection)
            _contests = Value
        End Set
    End Property

#End Region

    Public Sub New()
        Me.Initialize()
    End Sub

    Private Sub Initialize()
        _saveDelay = TimeSpan.FromMinutes(CDbl(AppSettings("SaveDelayMinutes")))
        _bufferSave = True
    End Sub

    '---COMMENT---------------------------------------------------------------
    '	remove all entities for user then user
    '
    '	Date:		Name:	Description:
    '	2/23/05 	JEA		stubbed
    '-------------------------------------------------------------------------
    Public Function RemoveUser(ByVal user As AMP.Person) As Boolean
        For Each c As AMP.Contest In _contests
            c.RemoveUserVotes(user)
        Next
        ' TODO: remove assets and such then website.save()
        _persons.Remove(user)
    End Function

    '---COMMENT---------------------------------------------------------------
    '	remove or archive expired entities
    '
    '	Date:		Name:	Description:
    '	3/8/05   	JEA		stubbed
    '-------------------------------------------------------------------------
    Private Sub Cleanup()

    End Sub

#Region " Serialization "

    '---COMMENT---------------------------------------------------------------
    '	serialize collections and persist to disk
    '
    '	Date:		Name:	Description:
    '	12/18/04	JEA		Creation
    '-------------------------------------------------------------------------
    Public Sub Save()
        BugOut("WebSite.Save() called")
        BugTab()

        If _saveState = SaveState.None Then
            BugOut("Queueing new thread to save")
            ThreadPool.QueueUserWorkItem(AddressOf Me.Flush)
        ElseIf Not (_bufferSave OrElse _saveState = SaveState.Writing) Then
            ' if immediate save requested and existing thread not already writing
            BugOut("Queuing another thread to save immediately")
            ThreadPool.QueueUserWorkItem(AddressOf Me.Flush)
        Else
            BugOut("Existing thread will satisfy save request")
        End If

        BugUntab()
    End Sub

    '---COMMENT---------------------------------------------------------------
    '	flush save calls to disk
    '
    '	Date:		Name:	Description:
    '	12/18/04	JEA		Creation
    '-------------------------------------------------------------------------
    Private Sub Flush(ByVal state As Object)
        If _bufferSave Then
            BugOut("Putting WebSite.Flush() thread to sleep")
            _saveState = SaveState.Sleeping
            Thread.Sleep(_saveDelay)
            BugOut("WebSite.Flush() is awake")
        Else
            ' always default back to buffered saves
            BugOut("Buffer disabled so flushing immediately after call to Flush()")
            _bufferSave = True
        End If

        Dim file As New AMP.Data.File
        _saveState = SaveState.Writing
        file.Save(AppSettings("DataFolder"), Me)
        _saveState = SaveState.None
        Global.LastSave = DateTime.Now
    End Sub

    Public Function Clone() As Object Implements System.ICloneable.Clone
        Return AMP.Data.Serialization.Clone(Me)
    End Function

    '---COMMENT---------------------------------------------------------------
    '	properly initialize certain values on deserialization
    '
    '	Date:		Name:	Description:
    '	12/18/04	JEA		Creation
    '-------------------------------------------------------------------------
    Public Sub OnDeserialization(ByVal sender As Object) Implements IDeserializationCallback.OnDeserialization
        BugOut("Executing AMP.WebSite.OnDeserialization")
        Me.Initialize()
    End Sub

#End Region

End Class
