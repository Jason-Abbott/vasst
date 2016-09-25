Imports AMP.Global
Imports System.Collections.Specialized

Namespace Controls
    Public Class SoftwareList
        Inherits AMP.Controls.SelectList

        Private _software As AMP.Software
        Private _version As AMP.Version
        Private Const _delimit As Char = "|"c

#Region " Properties "

        Public Property Software() As AMP.Software
            Get
                Return _software
            End Get
            Set(ByVal Value As AMP.Software)
                _software = Value

            End Set
        End Property

        Public Property Version() As AMP.Version
            Get
                Return _version
            End Get
            Set(ByVal Value As AMP.Version)
                _version = Value
            End Set
        End Property

#End Region

        '---COMMENT---------------------------------------------------------------
        '	convert selection to software and version entities
        '
        '	Date:		Name:	Description:
        '	1/5/05  	JEA		Creation
        '   2/23/05     JEA     Move to earlier event
        '-------------------------------------------------------------------------
        Public Overrides Function LoadPostData(ByVal key As String, ByVal posted As NameValueCollection) As Boolean
            If posted(key) <> Nothing Then
                Dim value As String() = posted(key).Split(","c)
                Dim software As String() = value(0).Split(_delimit)
                Me.Software = WebSite.Publishers.SoftwareWithID(software(0))
                Me.Version = Me.Software.Versions(software(1))
            End If
            Return False
        End Function

        '---COMMENT---------------------------------------------------------------
        '	render list of software
        '
        '	Date:		Name:	Description:
        '	12/30/04	JEA		Creation
        '-------------------------------------------------------------------------
        Protected Overrides Sub Render(ByVal writer As System.Web.UI.HtmlTextWriter)
            If Me.ID = Nothing Then Throw New NullReferenceException("Must supply ID")

            Dim selection As Boolean = Not (_software Is Nothing OrElse _version Is Nothing)

            MyBase.StartTag(writer)
            With writer
                For Each s As Software In WebSite.Publishers.Applications
                    For Each v As AMP.Version In s.Versions
                        .Write("<option value=""")
                        .Write(s.ID)
                        .Write(_delimit)
                        .Write(v.Number)
                        .Write("""")
                        If selection AndAlso (s.ID.Equals(_software.ID) AndAlso _
                            v.Number = _version.Number) Then

                            .Write(" selected=""selected""")
                        End If
                        .Write(">")
                        .Write(s.FullName)
                        .Write(" ")
                        .Write(v.Number)
                        .Write("</option>")
                    Next
                Next
            End With
            MyBase.EndTag(writer)
        End Sub
    End Class
End Namespace