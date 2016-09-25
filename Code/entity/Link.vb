Imports System.IO
Imports System.Net
Imports System.Text
Imports System.Text.RegularExpressions

<Serializable()> _
Public Class Link
    Private _url As String
    Private _views As Integer
    Private _scrape As Boolean = False
    Private _failures As Integer = 0

#Region " Properties "

    Public ReadOnly Property FullUrl() As String
        Get
            Return String.Format("http://{0}", _url)
        End Get
    End Property

    Public Property Failures() As Integer
        Get
            Return _failures
        End Get
        Set(ByVal Value As Integer)
            _failures = Value
        End Set
    End Property

    Public Property Url() As String
        Get
            Return _url
        End Get
        Set(ByVal Value As String)
            _url = Security.SafeString(Value, 150).Replace("http://", Nothing)
        End Set
    End Property

    Public Property Views() As Integer
        Get
            Return _views
        End Get
        Set(ByVal Value As Integer)
            _views = Value
        End Set
    End Property

    Public Property Scrape() As Boolean
        Get
            Return _scrape
        End Get
        Set(ByVal Value As Boolean)
            _scrape = Value
        End Set
    End Property

#End Region

    '---COMMENT---------------------------------------------------------------
    '	get content of the link
    '
    '	Date:		Name:	Description:
    '	1/23/05		JEA		Creation
    '-------------------------------------------------------------------------
    Public Function Content() As String
        Dim web As New WebClient
        Dim response As String

        Try
            Dim urlPath As String = Me.FullUrl.Substring(0, Me.FullUrl.LastIndexOf("/") + 1)
            response = Encoding.Default.GetString(web.DownloadData(Me.FullUrl))
            Dim re As New Regex("<body.*>([\s\S]*)<\/body>")
            Dim m As Match = re.Match(response)
            If m.Success Then
                response = m.Groups(1).Value
                response = re.Replace(response, "( src\s*=\s*[\""\'])(?!http)", "$1" & urlPath, _
                    RegexOptions.Multiline Or RegexOptions.IgnoreCase)

                Me.Failures = 0
            End If

        Catch ex As System.Net.WebException
            Me.Failures += 1
            AMP.Log.Error(ex, Log.ErrorType.Url, Profile.User)
        End Try

        Return response
    End Function
End Class
