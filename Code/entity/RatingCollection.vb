Imports System.Collections

<Serializable()> _
Public Class RatingCollection
    Inherits CollectionBase

    <NonSerialized()> Private _average As Single
    <NonSerialized()> Private _weightedAverage As Single

#Region " Properties "

    Public Property Item(ByVal index As Integer) As AMP.Rating
        Get
            Return DirectCast(MyBase.InnerList(index), AMP.Rating)
        End Get
        Set(ByVal Value As AMP.Rating)
            MyBase.InnerList(index) = Value
        End Set
    End Property

    '---COMMENT---------------------------------------------------------------
    '	get rating from given person
    '
    '	Date:		Name:	Description:
    '	1/18/05	    JEA		Creation
    '-------------------------------------------------------------------------
    Default Public ReadOnly Property ByPerson(ByVal person As AMP.Person) As AMP.Rating
        Get
            For Each rating As AMP.Rating In MyBase.InnerList
                If rating.Person Is person Then Return rating
            Next
            Return Nothing
        End Get
    End Property

#End Region

    '---COMMENT---------------------------------------------------------------
    '	return average rating for this collection
    '
    '	Date:		Name:	Description:
    '	1/14/05	    JEA		Creation
    '-------------------------------------------------------------------------
    Public Function Average() As Single
        If True OrElse _average = Nothing Then
            _weightedAverage = Nothing  ' erase weighted whenever average is updated
            If MyBase.InnerList.Count > 0 Then
                Dim total As Single = 0
                For Each r As AMP.Rating In MyBase.InnerList
                    total += r.Value
                Next
                _average = total / MyBase.InnerList.Count
            Else
                _average = 0
            End If
        End If
        Return _average
    End Function

    '---COMMENT---------------------------------------------------------------
    '	return average rating weighted by number of votes cast
    '   copy of formula used by code project
    '
    '	Date:		Name:	Description:
    '	1/14/05	    JEA		Creation
    '-------------------------------------------------------------------------
    Public Function WeightedAverage() As Single
        If _weightedAverage = Nothing Then
            If MyBase.InnerList.Count > 0 Then
                _weightedAverage = CSng(Me.Average * Math.Log10(MyBase.InnerList.Count))
            Else
                _weightedAverage = 0
            End If
        End If
        Return _weightedAverage
    End Function

    Public Function Add(ByVal entity As AMP.Rating) As Integer
        _average = Nothing
        Return MyBase.InnerList.Add(entity)
    End Function

    Public Sub Remove(ByVal entity As AMP.Rating)
        _average = Nothing
        MyBase.InnerList.Remove(entity)
    End Sub
End Class
