Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Runtime.Serialization
Imports AMP.Data
Imports AMP.Site

<Serializable()> _
Public Class Asset
    Implements IComparable
    'Implements IDeserializationCallback

    Private _id As Guid
    Private _title As String
    Private _file As AMP.File
    Private _link As AMP.Link
    Private _description As String
    Private _version As Integer
    Private _type As Asset.Types
    Private _status As AMP.Site.Status
    Private _submitDate As DateTime
    Private _versionDate As DateTime
    Private _ratings As AMP.RatingCollection
    Private _submittedBy As AMP.Person
    Private _authoredBy As AMP.Person
    Private _categories As New AMP.CategoryCollection
    Private _section As Integer = 0     ' bitmask
    <NonSerialized()> Private _views As String
    <NonSerialized()> Private _viewURL As String

#Region " Enumerations "

    <Flags()> _
    Public Enum Types
        File = Project Or Script Or Preset
        Link = Tutorial Or Review
        Project = &H1
        Tutorial = &H2
        Script = &H4
        Review = &H8
        Preset = &H10   ' 16
    End Enum

#End Region

#Region " Properties "

    '---COMMENT---------------------------------------------------------------
    '	bitmask of site.section values
    '
    '	Date:		Name:	Description:
    '	12/10/04	JEA		Creation
    '-------------------------------------------------------------------------
    Public Property Section() As Integer
        Get
            Return _section
        End Get
        Set(ByVal Value As Integer)
            _section = Value
        End Set
    End Property

    Public ReadOnly Property SectionName() As String
        Get
            If Common.IsFlag(_section) Then
                Dim section As AMP.Site.Section = CType(_section, AMP.Site.Section)
                Return section.ToString
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public Property Categories() As AMP.CategoryCollection
        Get
            Return _categories
        End Get
        Set(ByVal Value As AMP.CategoryCollection)
            _categories = Value
        End Set
    End Property

    Public Property Link() As AMP.Link
        Get
            _views = Nothing
            Return _link
        End Get
        Set(ByVal Value As AMP.Link)
            _views = Nothing
            _link = Value
        End Set
    End Property

    Public Property File() As AMP.File
        Get
            _views = Nothing
            Return _file
        End Get
        Set(ByVal Value As AMP.File)
            _views = Nothing
            _file = Value
        End Set
    End Property

    Public Property Type() As Asset.Types
        Get
            Return _type
        End Get
        Set(ByVal Value As Asset.Types)
            _type = Value
        End Set
    End Property

    '---COMMENT---------------------------------------------------------------
    '   return type and section information
    '   e.g. "A Script File for Vegas 5.0b"
    '
    '	Date:		Name:	Description:
    '	1/14/05	    JEA		Creation
    '   2/15/05     JEA     Handle no section
    '-------------------------------------------------------------------------
    Public ReadOnly Property FullType() As String
        Get
            Dim type As String = Me.Type.ToString
            Dim section As String

            If Me.File Is Nothing Then
                ' link
                section = Me.SectionName
            Else
                ' file
                With Me.File
                    section = String.Format("{0} {1}", .Software.Name, .SoftwareVersion.Number)
                End With
            End If
            If section = Nothing Then
                Return String.Format("A {0}", type)
            Else
                Return String.Format("A {0} for {1}", type, section)
            End If
        End Get
    End Property

    Public Property Status() As AMP.Site.Status
        Get
            Return _status
        End Get
        Set(ByVal Value As AMP.Site.Status)
            _status = Value
        End Set
    End Property

    Public ReadOnly Property ID() As Guid
        Get
            Return _id
        End Get
    End Property

    Public Property Description() As String
        Get
            Return _description
        End Get
        Set(ByVal Value As String)
            _description = Security.SafeString(Value, 1500)
        End Set
    End Property

    Public Property VersionDate() As DateTime
        Get
            Return _versionDate
        End Get
        Set(ByVal Value As DateTime)
            _versionDate = Value
        End Set
    End Property

    Public Property Ratings() As AMP.RatingCollection
        Get
            If _ratings Is Nothing Then _ratings = New AMP.RatingCollection
            Return _ratings
        End Get
        Set(ByVal Value As AMP.RatingCollection)
            _ratings = Value
        End Set
    End Property

    Public Property Version() As Integer
        Get
            Return _version
        End Get
        Set(ByVal Value As Integer)
            _version = Value
        End Set
    End Property

    Public Property AuthoredBy() As AMP.Person
        Get
            Return _authoredBy
        End Get
        Set(ByVal Value As AMP.Person)
            _authoredBy = Value
        End Set
    End Property

    Public Property SubmittedBy() As AMP.Person
        Get
            Return _submittedBy
        End Get
        Set(ByVal Value As AMP.Person)
            _submittedBy = Value
        End Set
    End Property

    Public Property SubmitDate() As DateTime
        Get
            Return _submitDate
        End Get
        Set(ByVal Value As DateTime)
            _submitDate = Value
            If _versionDate = Nothing Then _versionDate = Value
        End Set
    End Property

    Public Property Title() As String
        Get
            Return _title
        End Get
        Set(ByVal Value As String)
            Value = Security.SafeString(Value, 100)
            'If Value.IndexOf("_") <> -1 Then
            '    _originalName = Value
            '    _name = Value.Replace("_", " ")
            'Else
            _title = Value
            'End If
        End Set
    End Property

    'Public ReadOnly Property OriginalName() As String
    '    Get
    '        Return _originalName
    '    End Get
    'End Property

    '---COMMENT---------------------------------------------------------------
    '   fix names that have long unspaced words
    '
    '	Date:		Name:	Description:
    '	1/14/05	    JEA		Creation
    '-------------------------------------------------------------------------
    Public ReadOnly Property DisplayTitle() As String
        Get
            Dim title As String = _title
            Dim re As New Regex("\S{10,}")
            If re.IsMatch(title) Then title = Format.NormalSpacing(title)
            Return title
        End Get
    End Property

#End Region

    Public Sub New()
        _id = Guid.NewGuid
    End Sub

    '---COMMENT---------------------------------------------------------------
    '   approve this asset
    '
    '	Date:		Name:	Description:
    '	2/11/05	    JEA		Creation
    '   3/1/05      JEA     Clear search cache
    '-------------------------------------------------------------------------
    Public Function Approve() As Boolean
        If Not File Is Nothing Then
            If Not _file.Approve Then
                Profile.Message = "Failed to approve"   ' Error.CouldNotApprove
                Return False
            End If
        End If
        _status = Status.Approved
        WebSite.Assets.ClearSearchCache()
        WebSite.Save()
        Return True
    End Function

    '---COMMENT---------------------------------------------------------------
    '   delete this asset
    '
    '	Date:		Name:	Description:
    '	2/11/05	    JEA		Creation
    '-------------------------------------------------------------------------
    Public Sub Delete()
        If (_type And Types.File) > 0 Then _file.Delete()
        WebSite.Assets.Remove(Me)
        WebSite.Save()
    End Sub

#Region " Links "

    '---COMMENT---------------------------------------------------------------
    '   url to view detail page for asset
    '
    '	Date:		Name:	Description:
    '	1/18/05	    JEA		Creation
    '-------------------------------------------------------------------------
    Public Function DetailURL() As String
        Return String.Format("{0}/resource.aspx?id={1}", Global.BasePath, Me.ID)
    End Function

    Public Function DetailLink() As String
        Return String.Format("<a href=""{0}"" title=""{1}"">{2}</a>", _
            Me.DetailURL, HttpUtility.HtmlEncode(Me.FullType), HttpUtility.HtmlEncode(Me.Title))
    End Function

    Public Function FullDetailUrl() As String
        Return String.Format("http://{0}{1}", _
            HttpContext.Current.Request.ServerVariables("SERVER_NAME"), Me.DetailURL)
    End Function

    '---COMMENT---------------------------------------------------------------
    '   return url used to view or download this asset
    '
    '	Date:		Name:	Description:
    '	1/18/05	    JEA		Creation
    '-------------------------------------------------------------------------
    Public Function ViewURL() As String
        If _viewURL Is Nothing Then
            Dim link As String = String.Format("{0}/{{0}}.aspx?id={1}", Global.BasePath, Me.ID)
            If (Me.Type And Types.File) > 0 Then
                _viewURL = String.Format(link, "file")
            Else
                _viewURL = String.Format(link, "article")
            End If
        End If
        Return _viewURL
    End Function

    '---COMMENT---------------------------------------------------------------
    '   return html link to view or download asset
    '
    '	Date:		Name:	Description:
    '	2/23/05	    JEA		Creation
    '-------------------------------------------------------------------------
    Public Function ViewLink() As String
        Return String.Format("<a href=""{0}"" title=""{1}"">{2}</a>", _
            Me.ViewURL, HttpUtility.HtmlEncode(Me.FullType), HttpUtility.HtmlEncode(Me.Title))
    End Function

#End Region

    '---COMMENT---------------------------------------------------------------
    '   verb for viewing this asset
    '
    '	Date:		Name:	Description:
    '	1/18/05	    JEA		Creation
    '-------------------------------------------------------------------------
    Public Function ViewAction() As String
        If (Me.Type And Me.Types.File) > 0 Then
            Return "Action.Download"
        Else
            Return "Action.Read"
        End If
    End Function

    '---COMMENT---------------------------------------------------------------
    '   string describing number of times asset has been viewed
    '
    '	Date:		Name:	Description:
    '	1/14/05	    JEA		Creation
    '-------------------------------------------------------------------------
    Public Function Views() As String
        If _views = Nothing Then
            Dim count As Integer
            Dim type As String

            If (Me.Type And Types.File) > 0 Then
                type = "download"
                count = Me.File.Downloads
            Else
                type = "view"
                count = Me.Link.Views
            End If

            If count > 0 Then
                If count > 1 Then type = type & "s"
                _views = String.Format("{0:N0} {1}", count, type)
            End If
        End If
        Return _views
    End Function

#Region " Permission Checks "

    '---COMMENT---------------------------------------------------------------
    '   can this asset be edited by current user
    '
    '	Date:		Name:	Description:
    '	1/3/05	    JEA		Creation
    '-------------------------------------------------------------------------
    Public Function CanEdit() As Boolean
        Dim user As AMP.Person = Profile.User
        Return ((Me.SubmittedBy.ID.Equals(user.ID) AndAlso _
            user.HasPermission(Permission.EditMyAsset)) OrElse _
            user.HasPermission(Permission.EditAnyAsset))
    End Function

    '---COMMENT---------------------------------------------------------------
    '   can this asset be deleted by current user
    '
    '	Date:		Name:	Description:
    '	1/3/05	    JEA		Creation
    '-------------------------------------------------------------------------
    Public Function CanDelete() As Boolean
        Dim user As AMP.Person = Profile.User
        Return ((Me.SubmittedBy.ID.Equals(user.ID) AndAlso _
            user.HasPermission(Permission.DeleteMyAsset)) OrElse _
            user.HasPermission(Permission.DeleteAnyAsset))
    End Function

#End Region

#Region " Inferences "

    '---COMMENT---------------------------------------------------------------
    '   auto-generate a friendly name for the asset
    '   calling method should check if name is already set
    '
    '	Date:		Name:	Description:
    '	1/2/05	    JEA		Creation
    '-------------------------------------------------------------------------
    Public Function InferTitle(ByVal title As String) As Boolean
        If title = Nothing Then
            Dim tempTitle As String
            Dim startAt As Integer

            If Not _file Is Nothing Then
                tempTitle = _file.Name
                startAt = tempTitle.LastIndexOf("\")
            ElseIf Not _link Is Nothing Then
                tempTitle = _link.Url
                startAt = tempTitle.LastIndexOf("/")
            Else
                Return False
            End If

            If startAt < 0 Then
                startAt = 0
            Else
                startAt += 1
            End If

            'startAt = CInt(IIf(startAt < 0, 0, startAt + 1))
            tempTitle = tempTitle.Substring(startAt, tempTitle.LastIndexOf(".") - startAt)

            Me.Title = Format.NormalSpacing(tempTitle)
        Else
            Me.Title = title
        End If

        Return True
    End Function

    '---COMMENT---------------------------------------------------------------
    '	infer asset type
    '
    '	Date:		Name:	Description:
    '	1/2/05	    JEA		Creation
    '-------------------------------------------------------------------------
    Public Function InferType() As Boolean
        If Not _file Is Nothing Then
            Select Case _file.Type
                Case File.Types.Vegas, File.Types.Cool3D, File.Types.MediaStudioPro
                    _type = Types.Project
                    Return True
                Case File.Types.Script
                    _type = Types.Script
                    Return True
                Case File.Types.Preset
                    _type = Types.Preset
                    Return True
            End Select
            'ElseIf Not _link Is Nothing Then
            '    _type = Types.Review
            '    Return True
        End If
        Return False
    End Function

    '---COMMENT---------------------------------------------------------------
    '	infer asset section
    '
    '	Date:		Name:	Description:
    '	1/28/05	    JEA		Creation
    '-------------------------------------------------------------------------
    Public Function InferSection() As Boolean
        If Common.IsFlag(Profile.User.Section) Then
            ' if user has only one section, then assume same for asset
            Me.Section = Profile.User.Section
            Return True
        Else
            If Not _file Is Nothing Then
                ' attempt to infer from file type
                Select Case _file.Type
                    Case File.Types.Vegas, File.Types.Preset, File.Types.Script
                        Me.Section = AMP.Site.Section.Vegas
                        Return True
                    Case File.Types.MediaStudioPro
                        Me.Section = AMP.Site.Section.MediaStudio
                        Return True
                    Case File.Types.DvdArchitect
                        Me.Section = AMP.Site.Section.DvdArchitect
                        Return True
                    Case File.Types.Cool3D
                        Me.Section = AMP.Site.Section.Cool3D
                        Return True
                End Select
            ElseIf Not _link Is Nothing Then
                ' attempt to infer from title
                Select Case True
                    Case Me.Title.IndexOf("Vegas") <> -1
                        Me.Section = AMP.Site.Section.Vegas
                        Return True
                    Case Me.Title.IndexOf("HDV") <> -1
                        Me.Section = AMP.Site.Section.HDV
                        Return True
                End Select
            End If
        End If
        Return False
    End Function

#End Region

    '---COMMENT---------------------------------------------------------------
    '	default compare sorts from oldest to newest
    '
    '	Date:		Name:	Description:
    '	11/23/04	JEA		Creation
    '-------------------------------------------------------------------------
    Public Function CompareTo(ByVal entity As Object) As Integer Implements System.IComparable.CompareTo
        Dim a As AMP.Asset = DirectCast(entity, AMP.Asset)
        Dim compare As Integer
        compare = Date.Compare(a.SubmitDate, Me.SubmitDate)
        If compare = 0 Then compare = String.Compare(Me.Title, a.Title)
        Return compare
    End Function

    '---COMMENT---------------------------------------------------------------
    '	get plugins from file member, if any
    '
    '	Date:		Name:	Description:
    '	1/14/05 	JEA		Creation
    '-------------------------------------------------------------------------
    Public Function Plugins() As ArrayList
        Dim matches As New ArrayList
        If Not Me.File Is Nothing Then
            If Not Me.File.Plugins Is Nothing Then
                For Each p As AMP.Software In Me.File.Plugins
                    matches.Add(p)
                Next
            End If
        End If
        Return matches
    End Function

    '---COMMENT---------------------------------------------------------------
    '	get icon for this asset, if any
    '
    '	Date:		Name:	Description:
    '	1/24/05 	JEA		Creation
    '-------------------------------------------------------------------------
    Public Function IconFile() As String
        If (Me.Type And Types.File) > 0 Then
            Return Me.File.SoftwareVersion.IconFile
        Else
            Return "blank.gif"
        End If
    End Function

#Region " Serialization "

    Public Function Clone() As AMP.Asset
        Return DirectCast(Serialization.Clone(Me), AMP.Asset)
    End Function

#End Region

    'Public Sub OnDeserialization(ByVal sender As Object) Implements System.Runtime.Serialization.IDeserializationCallback.OnDeserialization
    '    ' TODO: remove this
    '    If _type = Nothing Then
    '        If Not _file Is Nothing Then
    '            _type = Types.File
    '        ElseIf Not _link Is Nothing Then
    '            _type = Types.Link
    '        End If
    '    End If
    'End Sub
End Class
