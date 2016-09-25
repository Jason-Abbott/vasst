Imports System.Text.RegularExpressions

<Serializable()> _
Public Class CategoryCollection
    Inherits CollectionBase

#Region " Properties "

    Public Property Item(ByVal index As Integer) As AMP.Category
        Get
            Return DirectCast(MyBase.InnerList(index), AMP.Category)
        End Get
        Set(ByVal Value As AMP.Category)
            MyBase.InnerList(index) = Value
        End Set
    End Property

    Default Public ReadOnly Property WithID(ByVal id As Guid) As AMP.Category
        Get
            For Each item As AMP.Category In Me.InnerList
                If item.ID.Equals(id) Then
                    Return item
                    Exit For
                End If
            Next
            Return Nothing
        End Get
    End Property

    Default Public ReadOnly Property WithID(ByVal id As String) As AMP.Category
        Get
            Return Me.WithID(New Guid(id))
        End Get
    End Property

    Public ReadOnly Property WithName(ByVal name As String) As AMP.Category
        Get
            For Each item As AMP.Category In Me.InnerList
                If item.Name = name Then
                    Return item
                    Exit For
                End If
            Next
            Return Nothing
        End Get
    End Property

#End Region

    '---COMMENT---------------------------------------------------------------
    '	does given category exist in this collection
    '
    '	Date:		Name:	Description:
    '	1/14/05 	JEA		Creation
    '-------------------------------------------------------------------------
    Public Function Contains(ByVal category As AMP.Category) As Boolean
        For Each c As AMP.Category In Me.InnerList
            If c Is category Then Return True
        Next
        Return False
    End Function

    '---COMMENT---------------------------------------------------------------
    '	return string array of IDs typically for using with Controls.SelectList
    '
    '	Date:		Name:	Description:
    '	2/11/05 	JEA		Creation
    '-------------------------------------------------------------------------
    Public Function IdArray() As String()
        Dim list(Me.InnerList.Count - 1) As String
        Dim x As Integer = 0

        For Each c As AMP.Category In Me.InnerList
            list(x) = c.ID.ToString
            x = x + 1
        Next
        Return list
    End Function

    '---COMMENT---------------------------------------------------------------
    '	does category with text exist in this collection
    '
    '	Date:		Name:	Description:
    '	12/26/04	JEA		Creation
    '   2/126/05    JEA     Use regex
    '-------------------------------------------------------------------------
    Public Function HasText(ByVal text As String) As Boolean
        Dim re As New Regex(text, RegexOptions.IgnoreCase)
        For Each c As AMP.Category In Me.InnerList
            If re.IsMatch(c.Name) Then Return True
        Next
        Return False
    End Function

    Public Function ForEntity(ByVal mask As Integer) As ArrayList
        Dim matches As New ArrayList
        For Each c As AMP.Category In Me.InnerList
            If (c.ForEntity And mask) > 0 Then
                matches.Add(c)
            End If
        Next
        Return matches
    End Function

    Public Function ForAssetType(ByVal mask As Integer) As ArrayList
        Dim matches As New ArrayList
        For Each c As AMP.Category In Me.InnerList
            If (c.ForAssetType And mask) > 0 Then
                matches.Add(c)
            End If
        Next
        Return matches
    End Function

    '---COMMENT---------------------------------------------------------------
    '	basic list methods
    '
    '	Date:		Name:	Description:
    '	12/26/04	JEA		Creation
    '-------------------------------------------------------------------------
    Public Function Add(ByVal entity As AMP.Category) As Integer
        Return MyBase.InnerList.Add(entity)
    End Function

    Public Sub Remove(ByVal entity As AMP.Category)
        MyBase.InnerList.Remove(entity)
    End Sub

    Public Sub Remove(ByVal id As Guid)
        Me.Remove(Me.WithID(id))
    End Sub

    Public Sub Sort()
        MyBase.InnerList.Sort()
    End Sub

End Class
