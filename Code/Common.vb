Imports System.Web.UI.WebControls
Imports System.text
Imports System.Text.RegularExpressions
Imports System.Web

Public Class Common

    '---COMMENT---------------------------------------------------------------
    '   build string used to redirect using JavaScript
    ' 
    '   Date:       Name:   Description:
    '	2/23/05	    JEA     Created
    '-------------------------------------------------------------------------
    Public Shared Function JSRedirect(ByVal url As String) As String
        Return String.Format("AddEvent(null, ""global"", function() {{ Global.Redirect(""{0}""); }} );", url)
    End Function

    '---COMMENT---------------------------------------------------------------
    '   get an array of the selected enumeration values
    ' 
    '   Date:       Name:   Description:
    '	2/15/05	    JEA     Created
    '-------------------------------------------------------------------------
    Public Shared Function EnumToString(ByVal selected As Integer, ByVal type As System.Type) As String
        Dim values As Integer() = CType([Enum].GetValues(type), Integer())
        Dim matches As New StringBuilder

        For x As Integer = 0 To values.Length - 1
            If Common.IsFlag(values(x)) AndAlso (selected And values(x)) > 0 Then
                If matches.Length > 0 Then matches.Append(",")
                matches.Append(values(x).ToString)
            End If
        Next

        Return matches.ToString
    End Function

    '---COMMENT---------------------------------------------------------------
    '   convert list of enum values to bitmask integer
    ' 
    '   Date:       Name:   Description:
    '	2/15/05	    JEA     Created
    '-------------------------------------------------------------------------
    Public Shared Function StringToEnum(ByVal selection As String) As Integer
        Dim mask As Integer = 0
        If selection <> Nothing Then
            Dim values As String() = selection.Split(","c)
            For x As Integer = 0 To values.Length - 1
                mask = (mask Or CInt(values(x)))
            Next
        End If
        Return mask
    End Function

    '---COMMENT---------------------------------------------------------------
    '   used for style sheets to append "odd" or "even" to class name
    ' 
    '   Date:       Name:   Description:
    '	3/31/04	    JEA     Created
    '-------------------------------------------------------------------------
    Public Shared Function SayOddEven(ByVal number As Integer) As String
        SayOddEven = IIf(number Mod 2 = 0, "Even", "Odd").ToString
    End Function

    '---COMMENT---------------------------------------------------------------
    '   like SQL ISNULL or COALESCE
    ' 
    '   Date:       Name:   Description:
    '	12/21/04	JEA     Created
    '-------------------------------------------------------------------------
    Public Shared Function ReplaceNull(ByVal value1 As Object, ByVal value2 As Object) As Object
        Return IIf(value1 Is Nothing, value2, value1)
    End Function

    '---COMMENT---------------------------------------------------------------
    '   used to find unique bitmask values, which are even powers of two
    ' 
    '   Date:       Name:   Description:
    '	1/13/05 	JEA     Created
    '-------------------------------------------------------------------------
    Public Shared Function IsFlag(ByVal x As Integer) As Boolean
        Dim exponent As Double = Math.Round(Math.Log(x) / Math.Log(2), 4)
        Return (exponent - Math.Floor(exponent) = 0)
    End Function

End Class