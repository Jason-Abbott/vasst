Imports System
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Web.Caching
Imports System.Threading
Imports System.Web.HttpContext
Imports System.Runtime.Serialization
Imports System.Runtime.Serialization.Formatters.Binary
Imports System.Configuration.ConfigurationSettings

Namespace Data
    Public Class File

        Private _root As String = HttpRuntime.AppDomainAppPath
        Private _content As String()

        Public ReadOnly Property Content() As String()
            Get
                Return _content
            End Get
        End Property

#Region " Save() "

        '---COMMENT---------------------------------------------------------------
        '	serialize and save object to disk
        '
        '	Date:		Name:	Description:
        '	12/1/04	    JEA		Creation
        '   12/14/04    JEA     Create temp file
        '   2/12/05     JEA     Only save one file
        '-------------------------------------------------------------------------
        Public Sub Save(ByVal folder As String, ByVal entity As Object)
            Dim file As FileInfo = Me.DataFile(entity.GetType)
            Dim stream As FileStream

            Try
                BugOut("Saving {0}", file.FullName)
                Dim bf As New BinaryFormatter
                stream = file.OpenWrite
                bf.AssemblyFormat = Formatters.FormatterAssemblyStyle.Simple
                bf.Serialize(stream, entity)
                Me.Cleanup(entity.GetType)
                Global.ActiveDataFile = file

            Catch e As UnauthorizedAccessException
                Log.Error(e, Log.ErrorType.FileSystem, Nothing)
                BugOut("Failed to save with {0}", e.Message)
            Catch e As Exception
                BugOut("Failed to save with {0}", e.Message)
            Finally
                If Not stream Is Nothing Then stream.Close()
            End Try
        End Sub

        '---COMMENT---------------------------------------------------------------
        '   save posted file to disk
        ' 
        '   Date:       Name:   Description:
        '	12/23/04    JEA     Created
        '   1/11/05     JEA     Get array of strings in binary file
        '   2/27/05     JEA     Added overload to handle custom target folder
        '-------------------------------------------------------------------------
        Public Function Save(ByVal folder As String, ByVal posted As HttpPostedFile) As String
            Dim fileName As String = posted.FileName
            If folder = Nothing Then folder = AppSettings("UploadFolder")
            fileName = fileName.Substring(fileName.LastIndexOf("\") + 1).Replace(" ", "_")
            Dim file As New FileInfo(String.Format("{0}/{1}/{2}", _root, folder, fileName))

            file = Me.AvailableName(file)
            If file.Extension = ".veg" Then Me.LoadContent(posted.InputStream, 4)

            Try
                posted.SaveAs(file.FullName)
                Return file.Name
            Catch ex As UnauthorizedAccessException
                Log.Error(ex, Log.ErrorType.FileSystem, Profile.User)
                Return Nothing
            End Try
        End Function

        Public Function Save(ByVal posted As HttpPostedFile) As String
            Return Me.Save(Nothing, posted)
        End Function

        '---COMMENT---------------------------------------------------------------
        '   if name taken, append numbers until unique name found
        ' 
        '   Date:       Name:   Description:
        '	2/16/05     JEA     Abstracted
        '-------------------------------------------------------------------------
        Public Function AvailableName(ByVal file As FileInfo) As FileInfo
            Dim tryName As String = file.Name
            Dim newName As String
            Dim x As Integer = 2

            Do While file.Exists
                ' find a file name that doesn't already exist
                newName = String.Format("{0}_{1:00}{2}", _
                    Path.GetFileNameWithoutExtension(tryName), _
                    x, file.Extension)
                file = New FileInfo(file.FullName.Replace(file.Name, newName))
                x += 1
            Loop
            Return file
        End Function

        '---COMMENT---------------------------------------------------------------
        '	get an array of the printable bytes within a byte array
        '   only consider minimum length of consecutive bytes
        '
        '	Date:		Name:	Description:
        '	1/11/05	    JEA		Creation
        '   2/18/05     JEA     Do not close reader since it also closes stream
        '-------------------------------------------------------------------------
        Public Sub LoadContent(ByVal file As Stream, ByVal minLength As Integer)
            Dim br As New BinaryReader(file)
            Dim bytes As Byte() = br.ReadBytes(CInt(file.Length))
            'br.Close()

            Dim text As New ArrayList
            Dim phrase As New StringBuilder
            Dim nullCount As Integer = 0
            Dim re As New Regex("^\W|^list\/", RegexOptions.IgnoreCase)
            Dim c As Integer

            For x As Integer = 0 To bytes.Length - 1
                c = bytes(x)

                If c > 31 AndAlso c < 128 Then
                    phrase.Append(Chr(c))
                    nullCount = 0
                Else
                    If c = 0 AndAlso nullCount = 0 Then
                        ' allow single null byte between printable bytes
                        nullCount += 1
                    Else
                        ' terminate phrase
                        If phrase.Length >= minLength AndAlso Not re.IsMatch(phrase.ToString) Then
                            'If phrase.Length >= minLength Then
                            text.Add(phrase.ToString)
                        End If
                        phrase.Length = 0
                        nullCount = 0
                    End If
                End If
            Next

            _content = CType(text.ToArray(GetType(String)), String())
        End Sub

#End Region

        '---COMMENT---------------------------------------------------------------
        '	remove excess entity files, if any
        '
        '	Date:		Name:	Description:
        '	12/18/04	JEA		Creation
        '   2/12/05     JEA     Abstracted for type
        '-------------------------------------------------------------------------
        Private Sub Cleanup(ByVal type As System.Type)
            Dim maxTempFiles As Integer = CInt(AppSettings("TempDataCount"))
            Dim path As New DirectoryInfo(String.Format("{0}{1}", _
                _root, AppSettings("DataFolder")))
            Dim files As FileInfo() = path.GetFiles(Me.Pattern(type))
            Dim filesToDelete As Integer = files.Length - maxTempFiles

            BugOut("Found {0} files to delete", filesToDelete)

            If filesToDelete > 0 Then
                BugTab()
                For x As Integer = 0 To filesToDelete - 1
                    BugOut("Deleting file {0}", files(x).FullName)
                    files(x).Delete()
                Next
                BugUntab()
            End If
        End Sub

#Region " Newest() "

        '---COMMENT---------------------------------------------------------------
        '	get newest file in folder
        '
        '	Date:		Name:	Description:
        '	12/18/04	JEA		Creation
        '-------------------------------------------------------------------------
        Public Function Newest(ByVal folder As String, ByVal pattern As String) As FileInfo
            Dim path As New DirectoryInfo(String.Format("{0}{1}", _root, folder))
            Dim files As FileInfo() = path.GetFiles(pattern)
            ' last file is newest
            Return files(files.Length - 1)
        End Function

        Public Function Newest(ByVal folder As String, ByVal type As System.Type) As FileInfo
            Return Me.Newest(folder, Me.Pattern(type))
        End Function

        Public Function Newest(ByVal folder As String) As FileInfo
            Return Me.Newest(folder, "*.*")
        End Function

#End Region

        '---COMMENT---------------------------------------------------------------
        '	pattern for serialized entity file names
        '
        '	Date:		Name:	Description:
        '	2/12/05 	JEA		Creation
        '-------------------------------------------------------------------------
        Private Function Pattern(ByVal type As System.Type) As String
            Return String.Format("{0}_*.dat", type.FullName)
        End Function

        '---COMMENT---------------------------------------------------------------
        '	create data file for given entity type
        '
        '	Date:		Name:	Description:
        '	12/18/04	JEA		Creation
        '   2/12/05     JEA     Abstracted to name by type
        '-------------------------------------------------------------------------
        Private Function DataFile(ByVal type As System.Type) As FileInfo
            Return New FileInfo(String.Format("{0}{1}/{2}_{3}.dat", _
                _root, AppSettings("DataFolder"), type.FullName, Now.Ticks))
        End Function

#Region " Load() "

        '---COMMENT---------------------------------------------------------------
        '	load serialized object
        '
        '	Date:		Name:	Description:
        '	12/1/04	    JEA		Creation
        '   12/7/04     JEA     Remove caching from this layer
        '   2/12/05     JEA     Pass fileinfo instead of name
        '-------------------------------------------------------------------------
        Public Function Load(ByVal file As FileInfo, ByVal type As System.Type, _
            ByRef schemaChange As Boolean) As Object

            Dim entity As Object

            schemaChange = False

            If file.Exists Then
                Dim stream As FileStream = file.OpenRead
                Dim bf As New BinaryFormatter
                bf.Context = New StreamingContext(StreamingContextStates.Persistence)
                Try
                    entity = bf.Deserialize(stream)
                Catch ex As SerializationException
                    ' standad deserialization didn't work so attempt schema migration
                    BugOut("Normal deserialization failed with ""{0}""", ex.Message)
                    stream.Seek(0, SeekOrigin.Begin)
                    Try
                        bf.SurrogateSelector = New AMP.Data.Migration(type.Assembly)
                        entity = bf.Deserialize(stream)
                        schemaChange = True
                    Catch ex2 As SerializationException
                        BugOut("Surrogate deserialization failed with ""{0}""", ex2.Message)
                    End Try
                Catch ex As Exception
                    BugOut("Normal deserialization failed with ""{0}""", ex.Message)
                Finally
                    stream.Close()
                    BugUntab()
                End Try
            Else
                BugOut("{0} does not exist", file.Name)
                Dim ex As New Exception(String.Format("{0} does not exist", file.FullName))
                Log.Error(ex, Log.ErrorType.Custom)
            End If

            Return entity
        End Function

        Public Function Load(ByVal fileName As String, ByVal type As System.Type, _
            ByRef schemachange As Boolean) As Object

            Return Me.Load(New FileInfo(fileName), type, schemachange)
        End Function

        '---COMMENT---------------------------------------------------------------
        '	load file content as string
        '
        '	Date:		Name:	Description:
        '	12/1/04	    JEA		Creation (matches logic in Controls.Content.Fill())
        '   2/12/05     JEA     Add caching
        '-------------------------------------------------------------------------
        Public Function Load(ByVal fileName As String) As String
            Dim content As String
            Dim key As String = fileName

            'content = DirectCast(Current.Cache.Item(key), String)

            If content = Nothing Then
                fileName = fileName.Replace("~", _root)

                If System.IO.File.Exists(fileName) Then
                    Dim file As New System.IO.StreamReader(fileName)
                    content = file.ReadToEnd()
                    file.Close()

                    'Dim dependsOn As CacheDependency = New CacheDependency(fileName)
                    'Current.Cache.Insert(key, content, dependsOn)
                Else
                    Log.Error(String.Format("{0} not found", fileName), _
                        Log.ErrorType.Custom, Profile.User)
                End If
            End If

            Return content
        End Function

#End Region

    End Class
End Namespace
