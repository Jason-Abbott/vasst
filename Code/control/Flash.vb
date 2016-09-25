Option Strict On

Imports System.Web.UI

Namespace Controls
    Public Class Flash
        Inherits AMP.Controls.HtmlControl

        Private _file As String
        Private _width As Integer
        Private _height As Integer

#Region " Properties "

        Public WriteOnly Property File() As String
            Set(ByVal Value As String)
                _file = Value
            End Set
        End Property

        Public WriteOnly Property Width() As Integer
            Set(ByVal Value As Integer)
                _width = Value
            End Set
        End Property

        Public WriteOnly Property Height() As Integer
            Set(ByVal Value As Integer)
                _height = Value
            End Set
        End Property

#End Region

        '---COMMENT---------------------------------------------------------------
        '	render file content within control
        '
        '	Date:		Name:	Description:
        '	11/10/04	JEA		Creation
        '-------------------------------------------------------------------------
        Protected Overrides Sub Render(ByVal writer As System.Web.UI.HtmlTextWriter)
            With writer
                .Write("<object data=""")
                .Write(_file)
                .Write(""" type=""application/x-shockwave-flash"" ")
                .Write("id=""")
                .Write(Me.ID)
                .Write(""" width=""")
                .Write(_width)
                .Write(""" height=""")
                .Write(_height)
                .Write("""><param name=""movie"" value=""")
                .Write(_file)
                .Write(""" /></object>")
            End With
        End Sub
    End Class
End Namespace

