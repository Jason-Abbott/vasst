Imports System.Text
Imports System.Data.OleDb
Imports System.Configuration.ConfigurationSettings

Namespace Data
    Public Class Legacy

        '---COMMENT---------------------------------------------------------------
        '   unsubscribe address from legacy mailing list
        ' 
        '   Date:       Name:   Description:
        '	2/17/05     JEA     Created
        '-------------------------------------------------------------------------
        Public Function DisableEmail(ByVal address As String) As Boolean
            Dim db As New AMP.Data.Jet(AppSettings("MailListStore"))
            Dim command As New StringBuilder

            With command
                .Append("UPDATE emailer_tblMembers SET memberDisabled = ?, ")
                .Append("memberNotes = 'Self-disabled ")
                .Append(DateTime.Now)
                .Append("' WHERE memberEmail = ?")
            End With

            With db
                .Command.CommandText = command.ToString
                .AddParam("memberDisabled", True, OleDbType.Boolean)
                .AddParam("memberEmail", address, OleDbType.VarChar)
                Try
                    .ExecuteOnly()
                    Return True
                Catch ex As Exception
                    Log.Error(ex, Log.ErrorType.FileSystem)
                    'BugOut("Couldn't save activity because {0}", ex.Message)
                    Return False
                End Try
            End With
        End Function

    End Class
End Namespace