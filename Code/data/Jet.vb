Option Strict Off

Imports AMP.Common
Imports System.Configuration.ConfigurationSettings
Imports System.Data
Imports System.Data.OleDb
Imports System.Web.UI.WebControls

Namespace Data
    Public Class Jet
        Inherits AMP.Data.Base

        Private Const _connectionString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="

#Region " New() "

        '---COMMENT---------------------------------------------------------------
        '	overload to open alternate connection
        '
        '	Date:		Name:	Description:
        '	3/17/04		JEA		Creation
        '   1/18/04     JEA     Use HttpRuntime to handle running in thread pool
        '-------------------------------------------------------------------------
        Public Sub New(ByVal dataPath As String)
            Me.Command = New OleDbCommand
            Me.Connection = New OleDbConnection
            Me.Adapter = New OleDbDataAdapter
            Me.OpenData(String.Format("{0}{1}", _connectionString, _
                dataPath.Replace("~", HttpRuntime.AppDomainAppPath)))
        End Sub

#End Region

        '---COMMENT---------------------------------------------------------------
        '	open given connection and set default command type
        '
        '	Date:		Name:	Description:
        '	3/17/04		JEA		Creation
        '-------------------------------------------------------------------------
        Protected Overrides Sub OpenData(ByVal connectionString As String)

            With Me.Connection
                .ConnectionString = connectionString
                .Open()
            End With
            Command.CommandType = CommandType.Text
        End Sub

        '---COMMENT---------------------------------------------------------------
        '	populate a data set with data tables
        '
        '	Date:		Name:	Description:
        '	11/7/04		JEA		Creation
        '-------------------------------------------------------------------------
        Public Sub FillDataset(ByRef ds As DataSet, ByVal tableDefinitions As Hashtable)
            Dim da As New OleDbDataAdapter

            Command.Connection = Me.Connection
            da.SelectCommand = DirectCast(Me.Command, OleDbCommand)

            For Each table As DictionaryEntry In tableDefinitions
                ds.Tables(table.Key.ToString).BeginLoadData()
                Command.CommandText = "SELECT " & table.Value & " FROM [" & table.Key & "]"
                da.Fill(ds, table.Key)
                ds.Tables(table.Key).EndLoadData()
            Next

            Me.Finish()
        End Sub

        '---COMMENT---------------------------------------------------------------
        '	retrieve return value from command
        '   no need to create return parameter in advance
        '
        '	Date:		Name:	Description:
        '	9/16/04		JEA		Creation
        '   10/1/04     JEA     Throw exception to calling method
        '-------------------------------------------------------------------------
        Public Overloads Overrides Function GetReturn() As Integer
            Dim oledbParameter As New OleDbParameter
            Dim returnIndex As Integer
            Dim returnValue As Object

            oledbParameter.Direction = ParameterDirection.ReturnValue
            oledbParameter.OleDbType = OleDbType.Integer
            With Me.Command
                returnIndex = .Parameters.Count()
                .Parameters.Add(oledbParameter)
                .Connection = Me.Connection
                .ExecuteNonQuery()
                returnValue = .Parameters(returnIndex).value
            End With
            Return CInt(returnValue)
            Me.Finish()
        End Function

#Region " AddParam() "

        '---COMMENT---------------------------------------------------------------
        '	customized method of adding i/o parameters
        '
        '	Date:		Name:	Description:
        '	3/17/04		JEA		Creation
        '   3/24/04     JEA     added option for empty parameters
        '   3/25/2004   ZNO     added option for ignoring zero parameters
        '-------------------------------------------------------------------------
        Public Sub AddParam(ByVal name As String, ByVal value As Object, _
            ByVal type As OleDbType, ByVal direction As ParameterDirection)

            If Not ((Me.IgnoreEmptyParams AndAlso Not value Is Nothing) Or _
                (Me.IgnoreZeroParams AndAlso value.Equals(0))) Then

                Dim parameter As New OleDbParameter

                With parameter
                    .ParameterName = name
                    .Value = value
                    .Direction = direction
                    .OleDbType = type
                End With

                Command.Parameters.Add(parameter)
                parameter = Nothing
            End If
        End Sub

        Public Sub AddParam(ByVal name As String, ByVal value As Object, ByVal type As OleDbType)
            Me.AddParam(name, value, type, ParameterDirection.Input)
        End Sub

        Public Sub AddParam(ByVal name As String, ByVal type As OleDbType, _
            ByVal direction As ParameterDirection)

            Dim parameter As New OleDbParameter
            With parameter
                .ParameterName = name
                .Direction = direction
                .OleDbType = type
            End With
            Command.Parameters.Add(parameter)
            parameter = Nothing
        End Sub

#End Region

    End Class
End Namespace