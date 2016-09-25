Option Strict On

Imports System.Web.UI
Imports System.Configuration.ConfigurationSettings

Namespace Controls

    '---COMMENT---------------------------------------------------------------
    '	Control to encapsulate form fields and labels in a standard table format
    '
    '	Date:		Name:	Description:
    '	9/13/04	    JEA		Creation
    '-------------------------------------------------------------------------
    Public Class TableForm
        Inherits HtmlContainerControl

        Private _say As New AMP.Locale
        Private _cssClass As String = "formSection"
        Private _cssGroupClass As String = "formGroup"
        Private _cssTitleClass As String = "title"
        Private _height As Integer
        Private _width As Integer
        Private _columns As Integer = 2
        Private _title As String = ""
        Private _group As String
        Private _twinCount As Integer = 0
        Private _twinOrder As Integer
        Private _padding As Integer = 8

#Region "Properties"

        Public WriteOnly Property Padding() As Integer
            Set(ByVal Value As Integer)
                _padding = Value
            End Set
        End Property

        '---COMMENT---------------------------------------------------------------
        '	Order of this control in twin sequence
        '
        '	Date:		Name:	Description:
        '	9/30/04	    JEA		Creation
        '-------------------------------------------------------------------------
        Public WriteOnly Property TwinOrder() As Integer
            Set(ByVal Value As Integer)
                _twinOrder = Value
            End Set
        End Property

        '---COMMENT---------------------------------------------------------------
        '	Number of sequential controls of same type and group name
        '
        '	Date:		Name:	Description:
        '	9/30/04	    JEA		Creation
        '-------------------------------------------------------------------------
        Public WriteOnly Property TwinCount() As Integer
            Set(ByVal Value As Integer)
                _twinCount = Value
            End Set
        End Property

        '---COMMENT---------------------------------------------------------------
        '	A common group name enables form tables to be visually grouped
        '
        '	Date:		Name:	Description:
        '	9/13/04	    JEA		Creation
        '-------------------------------------------------------------------------
        Public Property Group() As String
            Set(ByVal Value As String)
                _group = Value
            End Set
            Get
                Return _group
            End Get
        End Property

        Public WriteOnly Property cssClass() As String
            Set(ByVal Value As String)
                _cssClass = Value
            End Set
        End Property

        Public WriteOnly Property cssGroupClass() As String
            Set(ByVal Value As String)
                _cssGroupClass = Value
            End Set
        End Property

        Public WriteOnly Property Title() As String
            Set(ByVal Value As String)
                _title = Value
            End Set
        End Property

        Public WriteOnly Property cssTitleClass() As String
            Set(ByVal Value As String)
                _cssTitleClass = Value
            End Set
        End Property

        Public Property Height() As Integer
            Get
                Return _height
            End Get
            Set(ByVal Value As Integer)
                _height = Value
            End Set
        End Property

        Public Property Width() As Integer
            Get
                Return _width
            End Get
            Set(ByVal Value As Integer)
                _width = Value
            End Set
        End Property

        Public WriteOnly Property Columns() As Integer
            Set(ByVal Value As Integer)
                _columns = Value
            End Set
        End Property

#End Region

        '---COMMENT---------------------------------------------------------------
        '	compute grouping for all controls if not already initialized
        '
        '	Date:		Name:	Description:
        '	9/29/04	    JEA		Creation
        '-------------------------------------------------------------------------
        Private Sub TableForm_PreRender(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.PreRender
            If _group <> Nothing AndAlso _twinOrder = Nothing Then ComputeGroups("TableForm")
        End Sub

        '---COMMENT---------------------------------------------------------------
        '	Group controls of given type having the same group name
        '
        '	Date:		Name:	Description:
        '	9/30/04	    JEA		Creation
        '-------------------------------------------------------------------------
        Private Sub ComputeGroups(ByVal typeName As String)
            Dim height As Integer
            Dim sibling As Control
            Dim lastTwin As New TableForm
            Dim thisTwin As TableForm
            Dim twins As New ArrayList      ' hold references to twins until total count known
            Dim twinOrder As Integer = 1
            Dim controlWidth As Integer = 0
            Dim noWidthGiven As Integer = 0
            Dim firstControl As Boolean = True

            For Each sibling In Me.Parent.Controls
                If sibling.GetType.ToString.EndsWith(typeName) Then
                    thisTwin = DirectCast(sibling, TableForm)

                    If thisTwin.Group <> lastTwin.Group AndAlso Not firstControl Then
                        ' must have moved onto new group, finish previous group
                        ApplyGroupValues(twins, noWidthGiven, controlWidth, height)

                        ' reset values
                        height = Nothing
                        twinOrder = 1
                        noWidthGiven = 0
                        controlWidth = 0
                        twins.Clear()
                    End If

                    If thisTwin.Group <> Nothing Then
                        thisTwin.TwinOrder = twinOrder
                        twinOrder += 1
                        If height = Nothing Then height = thisTwin.Height
                        twins.Add(thisTwin)
                        If thisTwin.Width = Nothing Then
                            noWidthGiven += 1
                        Else
                            controlWidth += thisTwin.Width
                        End If
                    End If

                    lastTwin = thisTwin
                    firstControl = False
                End If
            Next

            ' finalize group, if any
            If twins.Count > 0 Then ApplyGroupValues(twins, noWidthGiven, controlWidth, height)
        End Sub

        '---COMMENT---------------------------------------------------------------
        '	Apply computed group values to relevant controls
        '
        '	Date:		Name:	Description:
        '	9/30/04	    JEA		Creation
        '-------------------------------------------------------------------------
        Private Sub ApplyGroupValues(ByRef twins As ArrayList, _
            ByVal noWidthGiven As Integer, ByVal controlWidth As Integer, ByVal height As Integer)

            Dim computedWidth As Integer
            Dim twinCount As Integer = twins.Count
            Try
                If noWidthGiven > 0 Then
                    ' try to compute width for those unspecified
                    Dim pageWidth As Integer = CInt(AppSettings.Item("UseableWidth"))
                    Dim extraWidth As Integer = pageWidth - controlWidth
                    extraWidth -= twins.Count * _padding
                    If extraWidth > 0 Then computedWidth = CInt(extraWidth / noWidthGiven)
                End If
            Catch
                computedWidth = Nothing
            End Try

            For Each twin As TableForm In twins
                ' set total twin count and match all twin heights
                twin.TwinCount = twinCount
                twin.Height = height
                If computedWidth <> Nothing AndAlso twin.Width = Nothing Then twin.Width = computedWidth
            Next
        End Sub

        '---COMMENT---------------------------------------------------------------
        '	Render HTML to display a form sub-section
        '
        '	Date:		Name:	Description:
        '	9/13/04	    JEA		Creation
        '   9/21/04     JEA     Handle title as a resource key
        '   9/30/04     JEA     Add HTML to group control twins
        '-------------------------------------------------------------------------
        Protected Overrides Sub RenderBeginTag(ByVal writer As HtmlTextWriter)
            With writer
                If _twinCount > 0 AndAlso _twinOrder = 1 Then
                    .Write("<table cellspacing='0' cellpadding='0' border='0' class='")
                    .Write(_cssGroupClass)
                    .Write("'><tr><td valign='top'>")
                End If
                .Write("<table cellspacing='1' cellpadding='0' class='")
                .Write(_cssClass)
                .Write("' style='")
                If _height <> Nothing Then
                    .Write("height: ")
                    .Write(_height)
                    .Write("px; ")
                End If
                If _width <> Nothing Then
                    .Write("width: ")
                    .Write(_width)
                    .Write("px;")
                End If
                .Write("'><tr>")
                If _title <> "" Then
                    .Write("<td colspan='")
                    .Write(_columns)
                    .Write("' class='")
                    .Write(_cssTitleClass)
                    .Write("'>")
                    .Write(_say(_title))
                    '.Write(": I am ")
                    '.Write(_twinOrder)
                    '.Write(" of ")
                    '.Write(_twinCount)
                    '.Write(" twins")
                    .Write("</td>")
                End If
            End With
        End Sub

        Protected Overrides Sub RenderEndTag(ByVal writer As HtmlTextWriter)
            With writer
                .Write("</td></tr></table>")
                If _twinCount > 0 Then
                    If _twinOrder = _twinCount Then
                        ' last twin, render end tags
                        .Write("</td></tr></table>")
                    ElseIf _twinOrder < _twinCount Then
                        .Write("</td><td valign='top'")
                        If _twinOrder = _twinCount - 1 Then .Write(" align='right'")
                        .Write(">")
                    End If
                End If
            End With

        End Sub

    End Class
End Namespace