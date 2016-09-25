Option Strict On

Imports AMP.Common
Imports System.Text
Imports System.Text.RegularExpressions

Public Class Format

#Region " General Formatting Methods "

    '---COMMENT---------------------------------------------------------------
    '   fix spacing on camel-case or underscored text
    '
    '	Date:		Name:	Description:
    '	1/2/05      JEA     Created
    '-------------------------------------------------------------------------
    Public Shared Function NormalSpacing(ByVal rawText As String) As String
        Dim text As String = Regex.Replace(rawText, "([a-z])([A-Z])", "$1 $2")
        text = Regex.Replace(text, "^(3[dD])(\w)", "$1 $2")
        'text = Regex.Replace(text, "([a-zA-Z]*)(\d+)([a-zA-Z]+)", "$1 $2 $3")
        'text = Regex.Repl'ace(text, "\d+$", "")
        text = text.Replace("Dvd", "DVD")
        Return text.Replace("_", " ").Trim
    End Function

    '---COMMENT---------------------------------------------------------------
    '   HTML format plain text
    '
    '	Date:		Name:	Description:
    '	11/1/04     JEA     Created
    '-------------------------------------------------------------------------
    Public Shared Function ToHtml(ByVal rawText As String) As String
        'rawText = rawText.Replace(Environment.NewLine & Environment.NewLine, "<br/>")
        rawText = rawText.Replace(Environment.NewLine, "<br/>")
        Return rawText
    End Function

    '---COMMENT---------------------------------------------------------------
    '	format string as proper case
    '
    '	Date:		Name:	Description:
    '	9/25/02		JEA		Created at Albertsons
    '   3/31/04     JEA     Converted to stringbuilder
    '   4/13/04     JEA     Use regular expression for finer control
    '-------------------------------------------------------------------------
    Public Shared Function ProperCase(ByVal rawText As String) As String
        Dim formattedText As New StringBuilder
        Dim word As String

        If Not rawText = Nothing Then
            Dim rex As New Regex("[a-zA-Z]+")
            Dim mc As MatchCollection = rex.Matches(rawText.Trim)

            With formattedText
                For x As Integer = 0 To mc.Count - 1
                    word = mc(x).Value
                    If mc(x).Index > 0 Then .Append(rawText.Substring(mc(x).Index - 1, 1))
                    If word.Length > 2 Then
                        .Append(word.Substring(0, 1).ToUpper())
                        .Append(word.Substring(1).ToLower())
                    Else
                        Select Case word.ToLower
                            Case "jr"
                                .Append("Jr.")
                            Case "dr"
                                .Append("Dr.")
                            Case Else
                                .Append(word)
                        End Select
                    End If
                Next
                ProperCase = .ToString.Trim
            End With
        Else
            ProperCase = rawText
        End If
    End Function

    '---COMMENT---------------------------------------------------------------
    '	return word for number; depends on having string resources defined
    '
    '	Date:		Name:	Description:
    '	10/24/02	JEA		Created at Albertsons
    '	3/24/04 	JEA		Added capitalize option
    '   10/4/04     JEA     Use resource strings
    '   2/23/05     JEA     Added overload
    '-------------------------------------------------------------------------
    Public Shared Function WordForNumber(ByVal sayNumber As Integer, ByVal capitalize As Boolean) As String
        Dim say As New Locale
        Dim numberName As String
        Select Case sayNumber
            Case 0 : numberName = say("Number.Zero")
            Case 1 : numberName = say("Number.One")
            Case 2 : numberName = say("Number.Two")
            Case 3 : numberName = say("Number.Three")
            Case 4 : numberName = say("Number.Four")
            Case 5 : numberName = say("Number.Five")
            Case 6 : numberName = say("Number.Six")
            Case 7 : numberName = say("Number.Seven")
            Case 8 : numberName = say("Number.Eight")
            Case 9 : numberName = say("Number.Nine")
            Case 10 : numberName = say("Number.Ten")
            Case 11 : numberName = say("Number.Eleven")
            Case 12 : numberName = say("Number.Twelve")
            Case 13 : numberName = say("Number.Thirteen")
            Case 14 : numberName = say("Number.Fourteen")
            Case 15 : numberName = say("Number.Fifteen")
            Case Else : numberName = sayNumber.ToString
        End Select
        If numberName = Nothing Then numberName = sayNumber.ToString
        If Not capitalize Then numberName = numberName.ToLower
        Return numberName
    End Function

    Public Shared Function WordForNumber(ByVal sayNumber As Integer) As String
        Return Format.WordForNumber(sayNumber, False)
    End Function

    '---COMMENT---------------------------------------------------------------
    '	return rank for number
    '
    '	Date:		Name:	Description:
    '	3/5/05  	JEA		Created
    '-------------------------------------------------------------------------
    Public Shared Function RankForNumber(ByVal sayNumber As Integer, ByVal capitalize As Boolean) As String
        Dim numberName As String
        Select Case sayNumber
            Case 1 : numberName = "First"
            Case 2 : numberName = "Second"
            Case 3 : numberName = "Third"
            Case 4 : numberName = "Fourth"
            Case 5 : numberName = "Fifth"
            Case 6 : numberName = "Sixth"
            Case 7 : numberName = "Seventh"
            Case 8 : numberName = "Eighth"
            Case 9 : numberName = "Ninth"
            Case 10 : numberName = "Tenth"
            Case Else : numberName = sayNumber.ToString
        End Select
        If numberName = Nothing Then numberName = sayNumber.ToString
        If Not capitalize Then numberName = numberName.ToLower
        Return numberName
    End Function

    Public Shared Function RankForNumber(ByVal sayNumber As Integer) As String
        Return Format.RankForNumber(sayNumber, False)
    End Function

#End Region

#Region " Zip Code Functions "

    '---COMMENT---------------------------------------------------------------
    '	Formats a Zip + 4 properly
    '
    '	Date:		Name:	Description:
    '	5/2/02		JEA		Created at Albertsons
    '   3/24/04     JEA     Use format class
    '-------------------------------------------------------------------------
    Public Shared Function Zip(ByVal rawZip As String) As String
        Dim formatPattern As String
        Try
            rawZip = NumbersOnly(rawZip)
            Select Case rawZip.Length()
                Case 9 : formatPattern = "{0:#####-####}"
                Case 5 : formatPattern = "{0:#####}"
                Case 0 : Exit Try
                Case Else : formatPattern = "{0}"
            End Select
            rawZip = String.Format(formatPattern, Integer.Parse(rawZip))
        Finally
            Zip = rawZip
        End Try
    End Function

    '---COMMENT---------------------------------------------------------------
    '	Splits a 9 charachter zip code into two useable parts
    '
    '	Date:		Name:	Description:
    '	12/1/2003	ZNO		Creation
    '-------------------------------------------------------------------------
    Public Shared Sub SplitZip(ByVal rawZip As String, ByRef zip5 As String, ByRef zip4 As String)
        Try
            If rawZip.IndexOf("-") > 0 Then rawZip = NumbersOnly(rawZip)

            If rawZip.Length > 5 Then
                zip5 = rawZip.Substring(0, 5)
                zip4 = rawZip.Substring(5, (rawZip.Length - 5))
            Else
                zip5 = rawZip
                zip4 = ""
            End If
        Catch
            zip5 = rawZip
            zip4 = ""
        End Try
    End Sub

#End Region

    '---COMMENT---------------------------------------------------------------
    '	Formats a phone number properly
    '
    '	Date:		Name:	Description:
    '   8/21/00		JEA		Created at Albertsons
    '   3/24/04     JEA     Used .NET system objects
    '-------------------------------------------------------------------------
    Public Shared Function Phone(ByVal rawPhone As String) As String
        Dim formatPattern As String
        Try
            rawPhone = NumbersOnly(rawPhone)
            Select Case rawPhone.Length()
                Case 10 : formatPattern = "{0:(###) ###-####}"
                Case 7 : formatPattern = "{0:###-####}"
                Case 0 : Exit Try
                Case Else : formatPattern = "{0}"
            End Select
            rawPhone = String.Format(formatPattern, Int64.Parse(rawPhone))
        Finally
            Phone = rawPhone
        End Try
    End Function

    '---COMMENT---------------------------------------------------------------
    '   strip non-numbers from string
    ' 
    '   Date		Name:	Description:
    '	6/30/02		JEA		Creation at Albertsons
    '-------------------------------------------------------------------------
    Public Shared Function NumbersOnly(ByVal number As String) As String
        Dim regExp As New Regex("\D", RegexOptions.Multiline)
        Try
            number = Trim(regExp.Replace(number, ""))
        Catch
            number = ""
        End Try
        NumbersOnly = number
    End Function

End Class