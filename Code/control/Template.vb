Imports System
Imports System.Text
Imports System.Web
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Web.UI.HtmlControls

Public Class Template
    Inherits System.Web.UI.UserControl

#Region " Controls "

    Protected phTitle As PlaceHolder
    Protected phScriptFiles As PlaceHolder
    Protected phStyleSheet As PlaceHolder
    Protected phViewState As PlaceHolder
    Protected phScriptBlock As PlaceHolder
    Protected phMessage As PlaceHolder
    Public frmMain As AMP.Controls.Form

#End Region

    Private _styleSheet As New ArrayList
    Private _scriptFile As New ArrayList
    Private _scriptBlock As New StringBuilder

#Region " Properties "

    '---COMMENT---------------------------------------------------------------
    '	displayed in browser title bar
    '
    '	Date:		Name:	Description:
    '	11/22/04	JEA		Creation
    '-------------------------------------------------------------------------
    Public WriteOnly Property Title() As String
        Set(ByVal Value As String)
            phTitle.Controls.Add(New LiteralControl(String.Format(" : {0}", Value)))
        End Set
    End Property

    '---COMMENT---------------------------------------------------------------
    '	message to display, usually from session on repost
    '
    '	Date:		Name:	Description:
    '	11/22/04	JEA		Creation
    '-------------------------------------------------------------------------
    Public Overridable WriteOnly Property Message() As String
        Set(ByVal Value As String)
            If Value <> Nothing Then
                phMessage.Controls.Add(New LiteralControl(String.Format( _
                    "<div id=""message"" class=""topRight"">{0}</div>", Value)))
            End If
        End Set
    End Property

    '---COMMENT---------------------------------------------------------------
    '	list of javascript files rendered in head
    '
    '	Date:		Name:	Description:
    '	11/22/04	JEA		Creation
    '   2/22/05     JEA     Expose arraylist directly
    '-------------------------------------------------------------------------
    Public Property ScriptFile() As ArrayList
        Get
            Return _scriptFile
        End Get
        Set(ByVal Value As ArrayList)
            _scriptFile = Value
        End Set
    End Property

    '---COMMENT---------------------------------------------------------------
    '	raw javascript rendered in XHTML compliant block
    '
    '	Date:		Name:	Description:
    '	11/22/04	JEA		Creation
    '   2/22/05     JEA     Expose stringbuilder directly
    '-------------------------------------------------------------------------
    Public Property ScriptBlock() As StringBuilder
        Get
            Return _scriptBlock
        End Get
        Set(ByVal Value As StringBuilder)
            _scriptBlock = Value
        End Set
    End Property

    '---COMMENT---------------------------------------------------------------
    '	list of stylesheets rendered in head
    '
    '	Date:		Name:	Description:
    '	11/22/04	JEA		Creation
    '   2/22/05     JEA     Expose arraylist directly
    '-------------------------------------------------------------------------
    Public Property StyleSheet() As ArrayList
        Get
            Return _styleSheet
        End Get
        Set(ByVal Value As ArrayList)
            _styleSheet = Value
        End Set
    End Property

    '---COMMENT---------------------------------------------------------------
    '	state value for custom form control
    '
    '	Date:		Name:	Description:
    '	11/22/04	JEA		Creation
    '-------------------------------------------------------------------------
    Public WriteOnly Property StateValue() As Control
        Set(ByVal Value As Control)
            phViewState.Controls.Add(Value)
        End Set
    End Property

#End Region

#Region " Unimplemented Properties "

    '---COMMENT---------------------------------------------------------------
    '	various controls used by different templates
    '
    '	Date:		Name:	Description:
    '	1/27/05 	JEA		Creation
    '-------------------------------------------------------------------------
    Public Overridable WriteOnly Property Headline() As WebControls.PlaceHolder
        Set(ByVal Value As WebControls.PlaceHolder)
            Throw New NotImplementedException
        End Set
    End Property

    Public Overridable Property Body() As WebControls.PlaceHolder
        Get
            Throw New NotImplementedException
        End Get
        Set(ByVal Value As WebControls.PlaceHolder)
            Throw New NotImplementedException
        End Set
    End Property

    Public Overridable WriteOnly Property LeftColumn() As WebControls.PlaceHolder
        Set(ByVal Value As WebControls.PlaceHolder)
            Throw New NotImplementedException
        End Set
    End Property

    Public Overridable Property WebParts() As WebControls.PlaceHolder
        Get
            Throw New NotImplementedException
        End Get
        Set(ByVal Value As WebControls.PlaceHolder)
            Throw New NotImplementedException
        End Set
    End Property

#End Region

    '---COMMENT---------------------------------------------------------------
    '	render scripts and styles in XHTML compliant manner
    '
    '	Date:		Name:	Description:
    '	11/22/04	JEA		Creation
    '-------------------------------------------------------------------------
    Private Sub Page_PreRender(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.PreRender
        If frmMain.Validation.Count > 0 Then
            ' render script block for validation
            _scriptFile.Add("validation/common")
            _scriptBlock.Append(frmMain.ValidationScriptBlock)
        End If

        If _scriptBlock.Length > 0 Then
            With _scriptBlock
                ' inserts in reverse order
                .Insert(0, Environment.NewLine)
                .Insert(0, "//<![CDATA[")
                .Insert(0, Environment.NewLine)
                .Insert(0, "<script type=""text/javascript"">")
                .Append(Environment.NewLine)
                .Append("//]]>")
                .Append(Environment.NewLine)
                .Append("</script>")
            End With
            phScriptBlock.Controls.Add(New LiteralControl(_scriptBlock.ToString))
        End If

        With phStyleSheet.Controls
            Dim tag As String = String.Format( _
                "<link rel=""stylesheet"" type=""text/css"" href=""{0}/style/{{0}}.css"" />{1}", _
                Global.BasePath, Environment.NewLine)
            For x As Integer = 0 To _styleSheet.Count - 1
                .Add(New LiteralControl(String.Format(tag, _styleSheet(x))))
            Next
        End With

        With phScriptFiles.Controls
            Dim tag As String = String.Format( _
                "<script type=""text/javascript"" src=""{0}/script/{{0}}.js""></script>{1}", _
                Global.BasePath, Environment.NewLine)
            For x As Integer = 0 To _scriptFile.Count - 1
                .Add(New LiteralControl(String.Format(tag, _scriptFile(x))))
            Next
        End With
    End Sub

#Region " Copy Controls "

    '---COMMENT---------------------------------------------------------------
    '	copy control content to template controls
    '
    '	Date:		Name:	Description:
    '	11/22/04	JEA		Creation
    '-------------------------------------------------------------------------
    Public Sub CopyToBody(ByVal source As Control)
        Me.CopyControlChildren(source, Me.Body)
    End Sub

    Public Sub CopyToWebPart(ByVal source As Control)
        Me.CopyControlChildren(source, Me.WebParts)
    End Sub

    Private Sub CopyControlChildren(ByVal source As Control, ByVal target As Control)
        Dim count As Integer = source.Controls.Count
        For x As Integer = 0 To count - 1
            target.Controls.Add(source.Controls(0))
        Next
    End Sub

#End Region

    Protected Function Say(ByVal resx As String) As String
        Return DirectCast(Me.Page, AMP.Page).Say(resx)
    End Function
End Class
