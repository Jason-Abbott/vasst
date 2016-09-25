Option Strict Off

Imports AMP.Common
Imports System.Configuration.ConfigurationSettings
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web.UI.WebControls

Namespace Data
    Public Class Sql
        Inherits AMP.Data.Base

#Region " New() "

        '---COMMENT---------------------------------------------------------------
        '	open default connection based on application setting
        '
        '	Date:		Name:	Description:
        '	3/17/04		JEA		Creation
        '-------------------------------------------------------------------------
        Public Sub New()
            OpenData(AppSettings.Item("DefaultConnection"))
        End Sub

        '---COMMENT---------------------------------------------------------------
        '	overload to open alternate connection
        '
        '	Date:		Name:	Description:
        '	3/17/04		JEA		Creation
        '-------------------------------------------------------------------------
        Public Sub New(ByVal connectionString As String)
            OpenData(connectionString)
        End Sub

#End Region

        '---COMMENT---------------------------------------------------------------
        '	open given connection and set default command type
        '
        '	Date:		Name:	Description:
        '	3/17/04		JEA		Creation
        '-------------------------------------------------------------------------
        Protected Overrides Sub OpenData(ByVal connectionString As String)
            Me.Command = New SqlCommand
            Me.Connection = New SqlConnection
            Me.Adapter = New SqlDataAdapter

            With Me.Connection
                .ConnectionString = connectionString
                .Open()
            End With
            Me.Command.CommandType = CommandType.StoredProcedure
        End Sub

#Region " AddParam() "

        '---COMMENT---------------------------------------------------------------
        '	customized method of adding i/o parameters
        '
        '	Date:		Name:	Description:
        '	3/17/04		JEA		Creation
        '   3/24/04     JEA     added option for empty parameters
        '   3/25/2004   ZNO     added option for ignoring zero parameters
        '-------------------------------------------------------------------------
        Public Sub AddParam(ByVal paramName As String, ByVal paramValue As Object, _
            ByVal paramType As SqlDbType, ByVal paramDirection As ParameterDirection)

            If Not ((Me.IgnoreEmptyParams AndAlso Not paramValue Is Nothing) Or _
                (Me.IgnoreZeroParams AndAlso paramValue.Equals(0))) Then

                Dim parameter As New SqlParameter

                With parameter
                    .ParameterName = paramName
                    .Value = paramValue
                    .Direction = paramDirection
                    .SqlDbType = paramType
                End With

                Command.Parameters.Add(parameter)
                parameter = Nothing
            End If
        End Sub

        '---COMMENT---------------------------------------------------------------
        '	Name: 		AddParam() overload
        '	customized method of adding input parameters only
        '
        '	Date:		Name:	Description:
        '	3/19/04		JEA		Creation
        '-------------------------------------------------------------------------
        Public Sub AddParam(ByVal paramName As String, ByVal paramValue As Object, ByVal paramType As SqlDbType)
            Me.AddParam(paramName, paramValue, paramType, ParameterDirection.Input)
        End Sub

        '---COMMENT---------------------------------------------------------------
        '	customized method of adding Command parameters
        '   use for output parameters with size instead of input value
        '
        '	Date:		Name:	Description:
        '	3/17/04		JEA		Creation
        '   4/21/04     JEA     Changed signature to avoid ambiguity
        '-------------------------------------------------------------------------
        Public Sub AddParam(ByVal paramName As String, ByVal paramType As SqlDbType, _
            ByVal paramDirection As ParameterDirection, ByVal paramSize As Integer)

            Dim parameter As New SqlParameter

            With parameter
                .ParameterName = paramName
                .Direction = paramDirection
                .SqlDbType = paramType
                .Size = paramSize
            End With

            Command.Parameters.Add(parameter)
            parameter = Nothing
        End Sub

        '---COMMENT---------------------------------------------------------------
        '	customized method of adding Command parameters
        '   simplified version
        '
        '	Date:		Name:	Description:
        '	4/22/04		JEA		Creation
        '-------------------------------------------------------------------------
        Public Sub AddParam(ByVal paramName As String, ByVal paramType As SqlDbType, _
            ByVal paramDirection As ParameterDirection)

            Dim parameter As New SqlParameter
            With parameter
                .ParameterName = paramName
                .Direction = paramDirection
                .SqlDbType = paramType
            End With
            Command.Parameters.Add(parameter)
            parameter = Nothing
        End Sub

#End Region

        '---COMMENT---------------------------------------------------------------
        '	retrieve return value from command
        '   no need to create return parameter in advance
        '
        '	Date:		Name:	Description:
        '	9/16/04		JEA		Creation
        '   10/1/04     JEA     Throw exception to calling method
        '-------------------------------------------------------------------------
        Public Overloads Overrides Function GetReturn() As Integer
            Dim sqlParameter As New SqlParameter
            Dim returnIndex As Integer
            Dim returnValue As Object

            sqlParameter.Direction = ParameterDirection.ReturnValue
            sqlParameter.SqlDbType = SqlDbType.Int
            With Me.Command
                returnIndex = .Parameters.Count()
                .Parameters.Add(sqlParameter)
                .Connection = Me.Connection
                .ExecuteNonQuery()
                returnValue = .Parameters(returnIndex).value
            End With
            Return CInt(returnValue)
            Me.Finish()
        End Function

        '---COMMENT---------------------------------------------------------------
        '	return well formed query string for pasting into query analyzer
        '   would typically be used to debug errors
        '
        '	Date:		Name:	Description:
        '	3/25/04		JEA		Creation
        '   10/1/04     JEA     Put ticks around strings; still needs more conditions
        '-------------------------------------------------------------------------
        Public Function MakeSQLStatement() As String
            Dim sqlStatement As New System.Text.StringBuilder
            Dim sqlParameter As SqlParameter

            With sqlStatement
                .Append(Command.CommandText)
                .Append(" ")
                For Each sqlParameter In Command.Parameters
                    .Append(sqlParameter.ParameterName)
                    .Append("=")
                    If sqlParameter.SqlDbType = SqlDbType.VarChar Then .Append("'")
                    .Append(sqlParameter.Value)
                    If sqlParameter.SqlDbType = SqlDbType.VarChar Then .Append("'")
                    .Append(", ")
                Next
            End With

            MakeSQLStatement = sqlStatement.ToString.TrimEnd(",".ToCharArray)
        End Function
    End Class
End Namespace