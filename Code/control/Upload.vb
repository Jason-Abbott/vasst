Imports System.IO
Imports System.Configuration.ConfigurationSettings

Namespace Controls
    Public Class Upload
        Inherits AMP.Controls.HtmlControl
        Implements IPostBackDataHandler

        Private _note As String
        Private _content As String()
        Private _parse As Boolean = False
        Private _allowedTypes As Integer = File.Types.Any
        Private _uploaded As Boolean = False
        Private _folder As String
        Private _inline As Boolean = False
        Private _maxFileKB As Integer = CInt(AppSettings("MaxFileUploadKB"))
        Private _resx As String
        Private _label As String
        Private _file As AMP.File

#Region " Properties "

        Public WriteOnly Property AllowedTypes() As Integer
            Set(ByVal Value As Integer)
                _allowedTypes = Value
            End Set
        End Property


        Public WriteOnly Property Parse() As Boolean
            Set(ByVal Value As Boolean)
                _parse = Value
            End Set
        End Property

        Public ReadOnly Property Content() As String()
            Get
                Return _content
            End Get
        End Property

        Public ReadOnly Property Uploaded() As Boolean
            Get
                Return _uploaded
            End Get
        End Property

        Public WriteOnly Property Folder() As String
            Set(ByVal Value As String)
                _folder = Value
            End Set
        End Property

        Public ReadOnly Property File() As AMP.File
            Get
                Return _file
            End Get
        End Property

        Public WriteOnly Property MaxFileSize() As Integer
            Set(ByVal Value As Integer)
                _maxFileKB = Value
            End Set
        End Property

        Public WriteOnly Property Note() As String
            Set(ByVal Value As String)
                _note = Value
            End Set
        End Property

        Public WriteOnly Property Inline() As Boolean
            Set(ByVal Value As Boolean)
                _inline = Value
            End Set
        End Property

        Public WriteOnly Property Resx() As String
            Set(ByVal Value As String)
                _resx = Value
            End Set
        End Property

#End Region

        '---COMMENT---------------------------------------------------------------
        '	save file
        '
        '	Date:		Name:	Description:
        '	2/24/05 	JEA		Creation
        '-------------------------------------------------------------------------
        Public Function LoadPostData(ByVal key As String, ByVal posted As System.Collections.Specialized.NameValueCollection) As Boolean Implements IPostBackDataHandler.LoadPostData
            Dim postedFile As HttpPostedFile = Context.Request.Files(Me.UniqueID)

            If (Not postedFile Is Nothing) AndAlso postedFile.ContentLength > 0 Then
                _uploaded = Me.Save(postedFile)
            End If
            Return False
        End Function

        Public Sub RaisePostDataChangedEvent() Implements IPostBackDataHandler.RaisePostDataChangedEvent
            ' this is called if LoadPostData returns true
        End Sub

        '---COMMENT---------------------------------------------------------------
        '	sanity checks and setup validation
        '
        '	Date:		Name:	Description:
        '	2/24/05 	JEA		Creation
        '-------------------------------------------------------------------------
        Private Sub Upload_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Init
            If _resx = Nothing Then
                Throw New Exception("No string resource supplied for labeled field")
            End If

            _label = Me.Page.Say(String.Format("Label.{0}", _resx))

            If _label = Nothing Then
                Throw New Exception(String.Format("No label resource found for ""{0}""", _resx))
            End If

            Me.Page.Form.Enctype = "multipart/form-data"
        End Sub

        '---COMMENT---------------------------------------------------------------
        '	register validation; script created during page pre-render
        '
        '	Date:		Name:	Description:
        '	2/26/05 	JEA		Creation
        '-------------------------------------------------------------------------
        Private Sub Upload_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
            If Me.Visible Then
                ' validation
                Dim alert As String = Me.Page.Say(String.Format("Validate.{0}", _resx))
                If alert = Nothing Then alert = _label
                Me.RegisterValidation("File", alert)
            End If
        End Sub

        Private Sub Upload_PreRender(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.PreRender
            Page.RegisterRequiresPostBack(Me)
        End Sub

        '---COMMENT---------------------------------------------------------------
        '	write labeled control with validation and notes
        '
        '	Date:		Name:	Description:
        '	2/24/05 	JEA		Creation
        '-------------------------------------------------------------------------
        Protected Overrides Sub Render(ByVal writer As System.Web.UI.HtmlTextWriter)
            If _note = Nothing Then _note = Me.Page.Say(String.Format("Note.{0}", _resx))
            Me.Attributes.Add("class", "file")
            Me.RenderLabel(_label, writer)
            writer.Write("<input type=""file"" name=""")
            writer.Write(Me.UniqueID)
            writer.Write("""")
            Me.RenderAttributes(writer)
            writer.Write(">")
            Me.RenderNote(_note, _inline, writer)
        End Sub

        '---COMMENT---------------------------------------------------------------
        '	saved posted file to disk and create file entity
        '
        '	Date:		Name:	Description:
        '	2/24/05 	JEA		Creation
        '-------------------------------------------------------------------------
        Private Function Save(ByVal posted As HttpPostedFile) As Boolean
            If posted.ContentLength > (_maxFileKB * 1000) Then
                Profile.Message = String.Format(Me.Page.Say("Error.LargeFile"), _maxFileKB)
                Return False
            End If

            If (_allowedTypes And AMP.File.InferType(posted.FileName)) = 0 Then
                Profile.Message = String.Format(Me.Page.Say("Error.FileType"), posted.FileName)
                Return False
            End If

            _file = New AMP.File
            Dim fileData As New AMP.Data.File

            _file.Name = fileData.Save(_folder, posted)
            _content = fileData.Content

            If File.Name = Nothing Then
                ' unable to save file
                Profile.Message = Me.Page.Say("Error.FileSave")
                Return False
            End If

            _file.Size = posted.ContentLength
            Return True
        End Function
    End Class
End Namespace