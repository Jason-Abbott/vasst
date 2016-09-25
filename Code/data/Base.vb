Imports AMP.Common
Imports System.Configuration.ConfigurationSettings
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web.UI.WebControls

Namespace Data
    Public MustInherit Class Base

        Private _ignoreEmptyParams As Boolean = False
        Private _ignoreZeroParams As Boolean = False
        Private _reuseConnection As Boolean = False
        Public Command As IDbCommand
        Protected Connection As IDbConnection
        Protected Adapter As IDbDataAdapter

#Region " Properties "

        '---COMMENT---------------------------------------------------------------
        '	option to not create a parameter object for empty values
        '
        '	Date:		Name:	Description:
        '	3/17/04		JEA		Creation
        '-------------------------------------------------------------------------
        Public Property IgnoreEmptyParams() As Boolean
            Get
                Return _ignoreEmptyParams
            End Get
            Set(ByVal Value As Boolean)
                _ignoreEmptyParams = Value
            End Set
        End Property

        '---COMMENT---------------------------------------------------------------
        '	option to not create a parameter object for zero values
        '
        '	Date:		Name:	Description:
        '	3/25/04		ZNO		Creation
        '-------------------------------------------------------------------------
        Public Property IgnoreZeroParams() As Boolean
            Get
                Return _ignoreZeroParams
            End Get
            Set(ByVal Value As Boolean)
                _ignoreZeroParams = Value
            End Set
        End Property

        '---COMMENT---------------------------------------------------------------
        '	option to leave connection open when .Finish is called
        '
        '	Date:		Name:	Description:
        '	3/17/04		JEA		Creation
        '-------------------------------------------------------------------------
        Public Property ReuseConnection() As Boolean
            Get
                Return _reuseConnection
            End Get
            Set(ByVal Value As Boolean)
                _reuseConnection = Value
            End Set
        End Property

#End Region

#Region " FillList() "

        ''---COMMENT---------------------------------------------------------------
        ''	fill drop down using passed SQL and specified columns
        ''
        ''	Date:		Name:	Description:
        ''	3/18/04		JEA		Creation
        ''-------------------------------------------------------------------------
        'Public Sub FillList(ByRef listBox As ListBox, ByVal sqlStatement As String, _
        '                    ByVal valueField As String, ByVal textField As String)
        '    With listBox
        '        Me.MakeTextCommand(sqlStatement)
        '        .DataSource = Me.GetDataTable
        '        .DataTextField = textField
        '        .DataValueField = valueField
        '        .DataBind()
        '    End With
        'End Sub

        ''---COMMENT---------------------------------------------------------------
        ''	fill list box using passed SQL and specified columns
        ''
        ''	Date:		Name:	Description:
        ''	3/18/04		JEA		Creation
        ''-------------------------------------------------------------------------
        'Public Sub FillList(ByRef dropDownList As DropDownList, ByVal sqlStatement As String, _
        '                    ByVal valueField As String, ByVal textField As String)
        '    With dropDownList
        '        Me.MakeTextCommand(sqlStatement)
        '        .DataSource = Me.GetDataTable
        '        .DataTextField = textField
        '        .DataValueField = valueField
        '        .DataBind()
        '    End With
        'End Sub

        ''---COMMENT---------------------------------------------------------------
        ''	fill list box with passed SQL and default columns
        ''
        ''	Date:		Name:	Description:
        ''	3/24/04		ZNO		Creation
        ''-------------------------------------------------------------------------
        'Public Sub FillList(ByRef listBox As ListBox, ByVal sqlStatement As String)
        '    With listBox
        '        Me.MakeTextCommand(sqlStatement)
        '        .DataSource = Me.GetDataTable
        '        .DataBind()
        '    End With
        'End Sub

        ''---COMMENT---------------------------------------------------------------
        ''	fill drop down with passed sql and default columns
        ''
        ''	Date:		Name:	Description:
        ''	3/24/04		ZNO		Creation
        ''-------------------------------------------------------------------------
        'Public Sub FillList(ByRef dropDownList As DropDownList, ByVal sqlStatement As String)
        '    With dropDownList
        '        Me.MakeTextCommand(sqlStatement)
        '        .DataSource = Me.GetDataTable
        '        .DataBind()
        '    End With
        'End Sub

#End Region

#Region " GetDataTable () "

        '---COMMENT---------------------------------------------------------------
        '	fill a datatable with Command
        '
        '	Date:		Name:	Description:
        '	3/18/04		JEA		Creation
        '-------------------------------------------------------------------------
        Public Function GetDataTable() As DataTable
            Dim ds As New DataSet
            Me.Command.Connection = Me.Connection
            Me.Adapter.SelectCommand = Me.Command
            Me.Adapter.Fill(ds)
            Return ds.Tables(0)
            Me.Finish()
        End Function

        Public Function GetDataTable(ByVal sqlStatement As String) As DataTable
            Me.MakeTextCommand(sqlStatement)
            Return Me.GetDataTable()
        End Function

#End Region

#Region " GetDataSet() "

        '---COMMENT---------------------------------------------------------------
        '	fill a DataSet with Command
        '
        '	Date:		Name:	Description:
        '	3/22/04		ZNO		Creation - Based on JEA's
        '-------------------------------------------------------------------------
        Public Function GetDataSet() As DataSet
            Dim dataSet As New DataSet
            Command.Connection = Me.Connection
            Me.Adapter.SelectCommand = Me.Command
            Me.Adapter.Fill(dataSet)
            Return dataSet
            Me.Finish()
        End Function

        Public Function GetDataSet(ByVal sqlStatement As String) As DataSet
            Me.MakeTextCommand(sqlStatement)
            Return Me.GetDataSet()
        End Function

#End Region

#Region " GetReader() "

        '---COMMENT---------------------------------------------------------------
        '	get a reader object; must leave connection open so can't call .Finish()
        '   but set CommandBehavior to automatically close when reader is closed
        '
        '	Date:		Name:	Description:
        '	3/17/04		JEA		Creation
        '   9/29/04     JEA     Keep connection open if reuse specified
        '-------------------------------------------------------------------------
        Public Function GetReader() As IDataReader
            Dim behavior As CommandBehavior
            If ReuseConnection Then
                behavior = CommandBehavior.Default
            Else
                behavior = CommandBehavior.CloseConnection
            End If

            Command.Connection = Me.Connection
            Return Me.Command.ExecuteReader(behavior)
        End Function

        ' overload to handle dynamic SQL
        Public Function GetReader(ByVal sqlStatement As String) As IDataReader
            Me.MakeTextCommand(sqlStatement)
            Return Me.GetReader()
        End Function

#End Region

#Region " GetSingleValue() "

        '---COMMENT---------------------------------------------------------------
        '	get single select value
        '
        '	Date:		Name:	Description:
        '	2/18/05		JEA		Creation
        '-------------------------------------------------------------------------
        Public Function GetSingleValue() As Object
            Command.Connection = Me.Connection
            Dim value As Object = Me.Command.ExecuteScalar
            Me.Finish()
            Return value
        End Function

        Public Function GetSingleValue(ByVal sqlStatement As String) As Object
            Me.MakeTextCommand(sqlStatement)
            Return Me.GetSingleValue()
        End Function

#End Region

        '---COMMENT---------------------------------------------------------------
        '	override default procedure command type and set text to passed SQL
        '
        '	Date:		Name:	Description:
        '	3/17/04		JEA		Creation
        '-------------------------------------------------------------------------
        Protected Sub MakeTextCommand(ByVal sqlStatement As String)
            With Me.Command
                .CommandType = CommandType.Text
                .CommandText = sqlStatement
            End With
        End Sub

        '---COMMENT---------------------------------------------------------------
        '	execute a Command with no results or output/return value only
        '
        '	Date:		Name:	Description:
        '	3/17/04		JEA		Creation
        '-------------------------------------------------------------------------
        Public Sub ExecuteOnly()
            Me.Command.Connection = Me.Connection
            Me.Command.ExecuteNonQuery()
            Me.Finish()
        End Sub

        Public Sub ExecuteOnly(ByVal sqlStatement As String)
            Me.MakeTextCommand(sqlStatement)
            Me.ExecuteOnly()
        End Sub

#Region " Virtuals "

        Protected MustOverride Sub OpenData(ByVal connectionString As String)

        Public MustOverride Function GetReturn() As Integer
        Public Function GetReturn(ByVal sqlStatement As String) As Integer
            Me.MakeTextCommand(sqlStatement)
            Return Me.GetReturn()
        End Function

#End Region

        '---COMMENT---------------------------------------------------------------
        '	close current data objects
        '
        '	Date:		Name:	Description:
        '	3/17/04		JEA		Creation
        '-------------------------------------------------------------------------
        Public Sub Finish()
            If Me.ReuseConnection Then
                ' prepare to reuse command
                Me.Command.Parameters.Clear()
            Else
                ' free the ADO objects
                Me.Command.Dispose()
                Me.Connection.Close()
                Me.Connection.Dispose()
            End If
        End Sub
    End Class
End Namespace