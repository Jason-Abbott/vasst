Imports System.IO
Imports System.Web.UI
Imports System.Web.Caching
Imports System.Web.HttpContext
Imports System.Configuration.ConfigurationSettings

Namespace Controls
    Public Class Content
        Inherits AMP.Controls.HtmlControl

        Private _fileName As String
        Private _startDate As DateTime
        Private _endDate As DateTime
        Private _content As String
        Private _needsFormatted As Boolean = False
        Private _daysOld As Integer
        Private _maxFiles As Integer

#Region " Properties "

        Public WriteOnly Property MaxFiles() As Integer
            Set(ByVal Value As Integer)
                _maxFiles = Value
            End Set
        End Property

        Public WriteOnly Property DaysOld() As Integer
            Set(ByVal Value As Integer)
                _daysOld = Value
            End Set
        End Property

        Public WriteOnly Property StartDate() As DateTime
            Set(ByVal Value As DateTime)
                _startDate = Value
            End Set
        End Property

        Public WriteOnly Property Enddate() As DateTime
            Set(ByVal Value As DateTime)
                _endDate = Value
            End Set
        End Property

        Public WriteOnly Property File() As String
            Set(ByVal Value As String)
                _fileName = Value
            End Set
        End Property

#End Region

        Public Sub New()
            Me.CssClass = "content"
        End Sub

        '---COMMENT---------------------------------------------------------------
        '	retrieve content
        '
        '	Date:		Name:	Description:
        '	11/1/04	    JEA		Creation
        '-------------------------------------------------------------------------
        Protected Overrides Sub OnLoad(ByVal e As System.EventArgs)
            Me.Fill()
        End Sub

        '---COMMENT---------------------------------------------------------------
        '	fill control with content
        '
        '	Date:		Name:	Description:
        '	11/10/04	    JEA		Creation
        '-------------------------------------------------------------------------
        Public Sub Fill()
            _fileName = String.Format("{0}{1}\{2}", HttpRuntime.AppDomainAppPath, AppSettings("ContentFolder"), _fileName)

            If (_startDate = Nothing OrElse _startDate < Now) AndAlso _
               (_endDate = Nothing OrElse _endDate > Now) Then
                ' retrieve content

                _needsFormatted = _fileName.EndsWith("txt")
                _content = DirectCast(Current.Cache.Item(_fileName), String)
                
                If _content = Nothing Then
                    ' read from file system if not in cache
                    If System.IO.File.Exists(_fileName) Then
                        Dim file As New System.IO.StreamReader(_fileName)
                        _content = file.ReadToEnd()
                        file.Close()

                        ' put content in cache
                        Dim dependsOn As CacheDependency = New CacheDependency(_fileName)
                        Current.Cache.Insert(_fileName, _content, dependsOn)
                    Else
                        Log.Error(String.Format("{0} not found", _fileName), _
                            Log.ErrorType.Custom, Profile.User)
                    End If
                End If
            End If
        End Sub

        '---COMMENT---------------------------------------------------------------
        '	render file content within control
        '
        '	Date:		Name:	Description:
        '	10/31/04	JEA		Creation
        '-------------------------------------------------------------------------
        Protected Overrides Sub Render(ByVal writer As System.Web.UI.HtmlTextWriter)
            If _content <> Nothing Then
                If _needsFormatted Then _content = Format.ToHtml(_content)
                With writer
                    .Write("<div")
                    Me.RenderCss(writer)
                    .Write(">")
                    .Write(_content)
                    .Write("</div>")
                End With
            Else
                AMP.Log.Error(String.Format("No content in {0}", _fileName), _
                    Log.ErrorType.Custom, Profile.User)
            End If
        End Sub
    End Class
End Namespace

