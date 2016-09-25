Imports AMP.ContestCollection

Namespace Compare

    Public Class ContestTitle
        Implements IComparer

        Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements System.Collections.IComparer.Compare
            Dim c1 As AMP.Contest = DirectCast(x, AMP.Contest)
            Dim c2 As AMP.Contest = DirectCast(y, AMP.Contest)
            Return String.Compare(c1.Title, c2.Title)
        End Function
    End Class

    Public Class ContestEntryVotes
        Implements IComparer

        Private _votesAllowed As Integer
        Private _weightFactor As Single

        Public Sub New(ByVal votesAllowed As Integer, ByVal weightFactor As Single)
            _votesAllowed = votesAllowed
            _weightFactor = weightFactor
        End Sub

        Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements System.Collections.IComparer.Compare
            Dim entry1 As AMP.ContestEntry = DirectCast(x, AMP.ContestEntry)
            Dim entry2 As AMP.ContestEntry = DirectCast(y, AMP.ContestEntry)
            Dim points1 As Integer = 0
            Dim points2 As Integer = 0

            For Each vote As AMP.ContestVote In entry1.Votes
                points1 += Contest.Points(vote.Rank, _weightFactor, _votesAllowed)
            Next

            For Each vote As AMP.ContestVote In entry2.Votes
                points2 += Contest.Points(vote.Rank, _weightFactor, _votesAllowed)
            Next

            If points1 = points2 Then
                If Not (entry1.Asset Is Nothing OrElse entry2.Asset Is Nothing) Then
                    Return String.Compare(entry1.Asset.Title, entry2.Asset.Title)
                Else
                    Return 0
                End If
            ElseIf points1 < points2 Then
                Return 1
            Else
                Return -1
            End If
        End Function
    End Class

    Public Class ContestEntryTitle
        Implements IComparer

        Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements System.Collections.IComparer.Compare
            Dim entry1 As AMP.ContestEntry = DirectCast(x, AMP.ContestEntry)
            Dim entry2 As AMP.ContestEntry = DirectCast(y, AMP.ContestEntry)

            If Not (entry1.Asset Is Nothing OrElse entry2.Asset Is Nothing) Then
                Return String.Compare(entry1.Asset.Title, entry2.Asset.Title)
            Else
                Return 0
            End If
        End Function
    End Class
End Namespace
