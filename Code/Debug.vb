Imports System.Diagnostics

'---COMMENT---------------------------------------------------------------
'	a simpler interface to debug methods
'   use Conditional() attribute to hide methods from release code
'
'	Date:		Name:	Description:
'	12/18/04    JEA		Creation
'-------------------------------------------------------------------------
Public Class Debug

#Region " BugCheck() "

    <Conditional("DEBUG")> _
    Public Shared Sub BugCheck(ByVal condition As Boolean, ByVal text As String)
        If condition Then BugOut(text)
    End Sub

    <Conditional("DEBUG")> _
    Public Shared Sub BugCheck(ByVal condition As Boolean, ByVal text As String, _
        ByVal arg0 As Object)

        If condition Then BugOut(text, arg0)
    End Sub

    <Conditional("DEBUG")> _
    Public Shared Sub BugCheck(ByVal condition As Boolean, ByVal text As String, _
       ByVal arg0 As Object, ByVal arg1 As Object)

        If condition Then BugOut(text, arg0, arg1)
    End Sub

#End Region

    <Conditional("DEBUG")> _
    Public Shared Sub BugTab()
        Diagnostics.Debug.Indent()
    End Sub

    <Conditional("DEBUG")> _
    Public Shared Sub BugUntab()
        Diagnostics.Debug.Unindent()
    End Sub

#Region " BugOut() "

    <Conditional("DEBUG")> _
    Public Shared Sub BugOut(ByVal text As String)
        Diagnostics.Debug.WriteLine(text)
    End Sub

    <Conditional("DEBUG")> _
    Public Shared Sub BugOut(ByVal text As String, ByVal arg0 As Object)
        Diagnostics.Debug.WriteLine(String.Format(text, arg0))
    End Sub

    <Conditional("DEBUG")> _
    Public Shared Sub BugOut(ByVal text As String, ByVal arg0 As Object, ByVal arg1 As Object)
        Diagnostics.Debug.WriteLine(String.Format(text, arg0, arg1))
    End Sub

    <Conditional("DEBUG")> _
    Public Shared Sub BugOut(ByVal text As String, ByVal arg0 As Object, ByVal arg1 As Object, ByVal arg2 As Object)
        Diagnostics.Debug.WriteLine(String.Format(text, arg0, arg1, arg2))
    End Sub

    <Conditional("DEBUG")> _
    Public Shared Sub BugOut(ByVal text As String, ByVal arg0 As Object, ByVal arg1 As Object, ByVal arg2 As Object, ByVal arg3 As Object)
        Diagnostics.Debug.WriteLine(String.Format(text, arg0, arg1, arg2, arg3))
    End Sub

#End Region

End Class
