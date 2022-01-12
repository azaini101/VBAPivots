Attribute VB_Name = "Pivot Table Macro"
Sub Pivots()
    Dim Columns As Variant
    Dim Rows As Variant
    Dim Values As Variant
    Dim Filters As Variant
    
    Filters = Array("state")
    Columns = Array()
    Rows = Array("county")
    Values = Array("population")
    'Change column, sourceSheet, destinationSheet, Table Name accordingly
    CreatePivot 1, "data", "State Data", "Pivot1", Filters, Columns, Rows, Values
    
    Filters = Array("state")
    Columns = Array("county")
    Rows = Array("zip")
    Values = Array("population")
    'Change column, sourceSheet, destinationSheet, Table Name accordingly
    CreatePivot 5, "data", "State Data", "Pivot2", Filters, Columns, Rows, Values
    
    Filters = Array("county", "zip")
    Columns = Array("state")
    Rows = Array("longitude", "latitude")
    Values = Array("population")
    'Change column, sourceSheet, destinationSheet, Table Name accordingly
    CreatePivot 1, "data", "Additional Data", "Pivot3", Filters, Columns, Rows, Values
End Sub

Sub CreatePivot(pCol As Integer, sourceSheet As String, destinationSheet As String, tableTitle As String, ByRef Filters As Variant, ByRef ColumnsField As Variant, ByRef RowsField As Variant, ByRef Values As Variant)
'destinationSheet As String, ByRef Arr As Variant, ByRef Arr1 As Variant

    Dim PSheet As Worksheet
    Dim DSheet As Worksheet
    Dim PCache As PivotCache
    Dim PTable As PivotTable
    Dim PRange As Range
    Dim LastRow As Long
    Dim LastCol As Long
    
    On Error Resume Next
    Application.DisplayAlerts = False
    TestSheet destinationSheet
    ActiveSheet.Name = destinationSheet
    Application.DisplayAlerts = True
    Set PSheet = Worksheets(destinationSheet)
    Set DSheet = Worksheets(sourceSheet)
    
    LastRow = DSheet.Cells(Rows.Count, 1).End(xlUp).Row
    LastCol = DSheet.Cells(1, Columns.Count).End(xlToLeft).Column
    Set PRange = DSheet.Cells(1, 1).Resize(LastRow, LastCol)
    
    Set PCache = ActiveWorkbook.PivotCaches.Create _
    (SourceType:=xlDatabase, SourceData:=PRange). _
    CreatePivotTable(TableDestination:=PSheet.Cells(1, pCol), _
    tableName:=tableTitle)

    Set PTable = PCache.CreatePivotTable _
    (TableDestination:=PSheet.Cells(1, pCol), tableName:=tableTitle)
    
    AddFields tableTitle, RowsField, xlRowField
    AddFields tableTitle, ColumnsField, xlColumnField
    AddFields tableTitle, Filters, xlPageField
    AddFields tableTitle, Values, xlDataField, xlSum
    
End Sub


Sub TestSheet(destinationSheet As String)
    'Creates sheet if it doesn't already exist
    Set wsTest = Nothing
    On Error Resume Next
    Set wsTest = ActiveWorkbook.Worksheets(destinationSheet)
    On Error GoTo 0
     
    If wsTest Is Nothing Then
        Worksheets.Add.Name = destinationSheet
    End If

End Sub

Sub AddFields(tableTitle As String, ByRef Arr As Variant, Orient As String, Optional func As String)
    'Adds fields to empty pivot table based on arguments given
    Dim i As Integer
    i = 1
    For Each Item In Arr
        With ActiveSheet.PivotTables(tableTitle).PivotFields(Item)
            .Orientation = Orient
            If func <> "" Then
                .Function = func
            Else
                .Position = i
            End If
            i = i + 1
        End With
    Next Item
End Sub

