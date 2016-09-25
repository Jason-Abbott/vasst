Imports AMP.Global
Imports System.IO
Imports System.Runtime.Serialization
Imports System.Configuration.ConfigurationSettings

<Serializable()> _
Public Class File
    'Implements IDeserializationCallback

    Private _name As String
    Private _type As File.Types
    Private _size As Single
    Private _downloads As Integer
    Private _requiredUrl As String
    Private _renderedUrl As String
    Private _plugins As New AMP.SoftwareCollection
    Private _software As AMP.Software
    Private _softwareVersion As New AMP.Version

#Region " Enumerations "

    <Flags()> _
    Public Enum Types
        Unknown = 0
        Any = Media Or Project Or Acrobat Or Excel Or Script Or Text Or Flash Or Preset
        Media = Image Or Video Or Flash Or Compressed
        Project = Vegas Or DvdArchitect Or Cool3D Or MediaStudioPro
        Resource = Project Or Compressed Or Script Or Preset
        Image = &H1
        Video = &H2
        Acrobat = &H4
        Vegas = &H8
        DvdArchitect = &H10
        Cool3D = &H20
        Compressed = &H40
        Excel = &H80
        Script = &H100
        Text = &H200
        MediaStudioPro = &H400
        Flash = &H800
        Preset = &H1000
    End Enum

#End Region

#Region " Properties "

    Public Property RenderedUrl() As String
        Get
            Return _renderedUrl
        End Get
        Set(ByVal Value As String)
            _renderedUrl = Security.SafeString(Value, 150).Replace("http://", Nothing)
        End Set
    End Property

    Public ReadOnly Property RenderedLink() As String
        Get
            Return String.Format("http://{0}", _renderedUrl)
        End Get
    End Property

    Public Property Plugins() As AMP.SoftwareCollection
        Get
            Return _plugins
        End Get
        Set(ByVal Value As AMP.SoftwareCollection)
            _plugins = Value
        End Set
    End Property

    Public Property Software() As AMP.Software
        Get
            Return _software
        End Get
        Set(ByVal Value As AMP.Software)
            _software = Value
        End Set
    End Property

    Public Property SoftwareVersion() As AMP.Version
        Get
            Return _softwareVersion
        End Get
        Set(ByVal Value As AMP.Version)
            _softwareVersion = Value
        End Set
    End Property

    Public Property Type() As AMP.File.Types
        Get
            Return _type
        End Get
        Set(ByVal Value As AMP.File.Types)
            _type = Value
        End Set
    End Property

    Public Property Size() As Single
        Get
            Return _size
        End Get
        Set(ByVal Value As Single)
            _size = Value
        End Set
    End Property

    Public Property Downloads() As Integer
        Get
            Return _downloads
        End Get
        Set(ByVal Value As Integer)
            _downloads = Value
        End Set
    End Property

    Public Property RequiredUrl() As String
        Get
            Return _requiredUrl
        End Get
        Set(ByVal Value As String)
            _requiredUrl = Security.SafeString(Value, 150).Replace("http://", Nothing)
        End Set
    End Property

    Public Property Name() As String
        Get
            Return _name
        End Get
        Set(ByVal Value As String)
            _name = Security.SafeString(Value, 100)
            If Value <> Nothing Then
                If Me.Type = Nothing Then Me.Type = Me.InferType(_name)
                If Me.Software Is Nothing Then Me.InferSoftware(_name)
            End If
        End Set
    End Property

    Public ReadOnly Property Extension() As String
        Get
            Return Me.Name.Substring(Me.Name.LastIndexOf(".") + 1).ToLower
        End Get
    End Property

#End Region

    '---COMMENT---------------------------------------------------------------
    '	move physical file to approved location
    '
    '	Date:		Name:	Description:
    '	2/16/05     JEA		Creation
    '   2/27/05     JEA     Added overload to customize target folder
    '-------------------------------------------------------------------------
    Public Function Approve(ByVal targetFolder As String) As Boolean
        Dim file As FileInfo = Me.Info(AppSettings("UploadFolder"))
        Dim fs As New Data.File
        Dim newFile As New FileInfo(String.Format("{0}{1}\{2}", _
            HttpRuntime.AppDomainAppPath, targetFolder, _name))

        newFile = fs.AvailableName(newFile)

        Try
            file.MoveTo(newFile.FullName)
            _name = newFile.Name
        Catch e As UnauthorizedAccessException
            Log.Error(e, Log.ErrorType.FileSystem, Profile.User)
            Return False
        Catch e As Exception
            Log.Error(e, Log.ErrorType.Unknown, Profile.User)
            Return False
        End Try

        Return True
    End Function

    Public Function Approve() As Boolean
        Return Me.Approve(AppSettings("ResourceFolder"))
    End Function

    '---COMMENT---------------------------------------------------------------
    '	delete physical file from file system
    '
    '	Date:		Name:	Description:
    '	2/11/05     JEA		Creation
    '   2/27/05     JEA     Added overload to customize source folder
    '-------------------------------------------------------------------------
    Public Function Delete(ByVal folder As String) As Boolean
        Dim file As FileInfo = Me.Info(folder)
        If file.Exists Then
            file.Delete()
            Return True
        Else
            Log.Error(String.Format("Unable to delete {0}", file.FullName), _
                Log.ErrorType.Custom, Profile.User)
            Return False
        End If
    End Function

    Public Function Delete() As Boolean
        Return Me.Delete(AppSettings("UploadFolder"))
    End Function

    Private Function Info(ByVal folder As String) As FileInfo
        Return New FileInfo(String.Format("{0}{1}\{2}", _
            HttpRuntime.AppDomainAppPath, folder, _name))
    End Function

    '---COMMENT---------------------------------------------------------------
    '	Infer file type from extension
    '
    '	Date:		Name:	Description:
    '	12/21/04    JEA		Creation
    '   2/27/05     JEA     Made public
    '-------------------------------------------------------------------------
    Public Shared Function InferType(ByVal name As String) As File.Types
        Dim extension As String = name.Substring(name.LastIndexOf(".") + 1).ToLower
        Select Case extension
            Case "pdf"
                Return File.Types.Acrobat
            Case "zip"
                Return File.Types.Compressed
            Case "c3d"
                Return File.Types.Cool3D
            Case "dar"
                Return File.Types.DvdArchitect
            Case "xls"
                Return File.Types.Excel
            Case "swf"
                Return File.Types.Flash
            Case "jpg", "jpeg", "png", "gif", "bmp"
                Return File.Types.Image
            Case "msp"
                Return File.Types.MediaStudioPro
            Case "js"
                Return File.Types.Script
            Case "txt"
                Return File.Types.Text
            Case "veg"
                Return File.Types.Vegas
            Case "avi", "wmv", "m2t", "mpg", "mpeg", "qt", "mov", "mvp"
                Return File.Types.Video
            Case "sfpreset"
                Return File.Types.Preset
            Case Else
                Return File.Types.Unknown
        End Select
    End Function

    '---COMMENT---------------------------------------------------------------
    '	Infer software from extension
    '
    '	Date:		Name:	Description:
    '	12/30/04    JEA		Creation
    '-------------------------------------------------------------------------
    Private Sub InferSoftware(ByVal name As String)
        Dim extension As String = name.Substring(name.LastIndexOf(".") + 1).ToLower
        Me.Software = WebSite.Publishers.SoftwareForExtension(extension)
        If Not Me.Software Is Nothing Then
            ' default version to latest
            Me.SoftwareVersion = Me.Software.Versions.Latest
        End If
    End Sub

    'Public Sub OnDeserialization(ByVal sender As Object) Implements System.Runtime.Serialization.IDeserializationCallback.OnDeserialization
    '    If _plugins Is Nothing Then _plugins = New AMP.SoftwareCollection
    'End Sub
End Class
