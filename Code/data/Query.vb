Namespace Data
    Public Class Query

#Region " Join() "

        Public Shared Function Join(ByVal first As DataTable, ByVal second As DataTable, ByVal firstJoinColumn As DataColumn, ByVal secondJoinColumn As DataColumn) As DataTable
            Return Join(first, second, New DataColumn() {firstJoinColumn}, New DataColumn() {secondJoinColumn})
        End Function

        Public Shared Function Join(ByVal first As DataTable, ByVal second As DataTable, ByVal firstJoinColumn As String, ByVal secondJoinColumn As String) As DataTable
            Return Join(first, second, New DataColumn() {first.Columns(firstJoinColumn)}, New DataColumn() {second.Columns(secondJoinColumn)})
        End Function

        Public Shared Function Join(ByVal first As DataTable, ByVal second As DataTable, ByVal firstJoinColumns As DataColumn(), ByVal secondJoinColumns As DataColumn()) As DataTable
            ' Create Empty Table
            Dim oTable As DataTable = New DataTable("Join")

            ' Use a DataSet to leverage DataRelation
            Dim oDataSet As DataSet = New DataSet
            With oDataSet
                ' Add Copy of Tables
                .Tables.AddRange(New DataTable() {first.Copy, second.Copy})

                ' Identify Joining Columns from First
                Dim arrParentColumns(firstJoinColumns.Length - 1) As DataColumn
                For iCounter As Int32 = 0 To arrParentColumns.Length - 1
                    arrParentColumns(iCounter) = oDataSet.Tables(0).Columns(firstJoinColumns(iCounter).ColumnName)
                Next

                ' Identify Joining Columns from Second
                Dim arrChildColumns(secondJoinColumns.Length - 1) As DataColumn
                For iCounter As Int32 = 0 To arrChildColumns.Length - 1
                    arrChildColumns(iCounter) = oDataSet.Tables(1).Columns(secondJoinColumns(iCounter).ColumnName)
                Next

                ' Create DataRelation
                Dim oDataRelation As DataRelation = New DataRelation(String.Empty, arrParentColumns, arrChildColumns, False)
                .Relations.Add(oDataRelation)

                ' Create Columns for JOIN table
                For iCounter As Int32 = 0 To first.Columns.Count - 1
                    oTable.Columns.Add(first.Columns(iCounter).ColumnName, first.Columns(iCounter).DataType)
                Next

                For iCounter As Int32 = 0 To second.Columns.Count - 1
                    ' Beware Duplicates
                    If Not oTable.Columns.Contains(second.Columns(iCounter).ColumnName) Then
                        oTable.Columns.Add(second.Columns(iCounter).ColumnName, second.Columns(iCounter).DataType)
                    Else
                        oTable.Columns.Add(second.Columns(iCounter).ColumnName & "_Second", second.Columns(iCounter).DataType)
                    End If
                Next

                ' Loop through First table
                oTable.BeginLoadData()
                For Each oFirstTableDataRow As DataRow In oDataSet.Tables(0).Rows
                    ' Get "joined" rows
                    Dim childRows As DataRow() = oFirstTableDataRow.GetChildRows(oDataRelation)
                    If Not childRows Is Nothing AndAlso childRows.Length > 0 Then
                        Dim arrParentArray() As Object = oFirstTableDataRow.ItemArray
                        For Each oSecondTableDataRow As DataRow In childRows
                            Dim arrSecondArray() As Object = oSecondTableDataRow.ItemArray
                            Dim arrJoinArray(arrParentArray.Length + arrSecondArray.Length - 1) As Object
                            Array.Copy(arrParentArray, 0, arrJoinArray, 0, arrParentArray.Length)
                            Array.Copy(arrSecondArray, 0, arrJoinArray, arrParentArray.Length, arrSecondArray.Length - 1)
                            oTable.LoadDataRow(arrJoinArray, True)
                        Next
                    End If

                Next

                oTable.EndLoadData()
            End With
            Return oTable
        End Function

#End Region

        '---COMMENT---------------------------------------------------------------
        '	make new data table from existing with rows matching criteria
        '
        '	Date:		Name:	Description:
        '	11/13/04	JEA		Creation
        '-------------------------------------------------------------------------
        Public Function SelectFromTable(ByVal table As DataTable, ByVal count As Integer, _
            ByVal columns As String(), ByVal filter As String, ByVal sort As String) As DataTable

            Dim selected As New DataTable

            With table
                If columns(0) = "*" Then
                    ' add all columns
                    For x As Integer = 0 To .Columns.Count - 1
                        selected.Columns.Add(.Columns(x).ColumnName, .Columns(x).DataType)
                    Next
                Else
                    ' add only selected columns
                    For x As Integer = 0 To columns.Length - 1
                        ' create selected columns in new table
                        selected.Columns.Add(columns(x), table.Columns(columns(x)).DataType)
                    Next
                End If
            End With

            Dim dv As New DataView(table)
            If filter <> Nothing Then dv.RowFilter = filter
            If sort <> Nothing Then dv.Sort = sort
            If count = Nothing Then count = dv.Count

            selected.BeginLoadData()
            For x As Integer = 0 To count - 1
                selected.ImportRow(dv.Item(x).Row)
            Next
            selected.EndLoadData()

            Return selected
        End Function
    End Class
End Namespace