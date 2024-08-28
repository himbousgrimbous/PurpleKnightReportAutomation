' ConcatenateRow works, DelColumn works, RowToString works, DelQueries works


Option Explicit
Function ConcatenateRow(rowRange As Range, joinString As String) As String
    Dim x As Variant, temp As String

    temp = ""
    For Each x In rowRange
        temp = temp & x & joinString
        ' Debug.Print temp
    Next

    ConcatenateRow = Left(temp, Len(temp) - Len(joinString))
End Function


Sub DelColumns(wSheet As Worksheet)
    Dim crtRegion As Range
    Dim i As Long
    Dim matchFound As Variant
    Dim hdrName As Variant
    
    hdrName = Array("Ignored", "EventTimestamp", "Operation", "Max Age", "Min Age", "Min Length", "Complexity Enabled", "History", "ServicePrincipalName", "Enabled")  ' array of column names we don't want
    
    For i = 0 To UBound(hdrName)                                                ' Ubound returns the sizee of the array
        Set crtRegion = wSheet.Range("A1").CurrentRegion                               ' CurrentRegion selects the current active region, ie the table currently filled
        matchFound = Application.Match(hdrName(i), crtRegion.Rows(1), 0)        ' Application.Match(str_to_match, Range_to_look_in, type_of_match) - looks for a match with specfic header in first row of the current sheet
        
        If Not IsError(matchFound) Then
            ' Debug.Print VarType(matchFound)
            ' Debug.Print matchFound                      ' index on y axis of match
            ' Debug.Print crtRegion.Cells(1, matchFound)  ' match value
            crtRegion.Columns(matchFound).Delete
        End If
    Next
    Set crtRegion = Nothing
End Sub


Function RowToString(sourceWs As Worksheet) As String
    Dim selectedRegion As Range
    Dim selectedRow As Range
    Dim i As Long
    Dim lastCol As Long
    Dim cattedRow As String
    
    
    
    'If Not IsError(sourceWs) Then
        Set selectedRegion = sourceWs.Range("A1").CurrentRegion         ' find active region in source worksheet
                    
        lastCol = selectedRegion.Columns.Count                          ' find last column index
        ' Debug.Print selectedRegion.Rows.Count
        
        cattedRow = ""
        
        For i = 2 To selectedRegion.Rows.Count                          ' parse through rows of active region, index is 2 so that header is not included

            ' Debug.Print Cells(i, 1).Address
            ' Debug.Print Cells(i, lastCol).Address
            Set selectedRow = sourceWs.Range(Cells(i, 1).Address, Cells(i, lastCol).Address)    ' select row i from A to lastCol
             
            cattedRow = cattedRow & ConcatenateRow(selectedRow, "; ") & vbNewLine
            Debug.Print cattedRow
        Next
    
    'Else
        'Debug.Print "sourceWs me clc"
   'End If
   RowToString = cattedRow
End Function


Sub DelQueries()
    Dim qry As WorkbookQuery
    Dim cn As WorkbookConnection
    
    For Each cn In ThisWorkbook.Connections
        cn.Delete
    Next
    
    For Each qry In ThisWorkbook.Queries
            qry.Delete
    Next
End Sub


Sub formatClientSheet(newSheetName As String, oldSheet As Worksheet)
    Dim newClientSheet As Worksheet
    Dim selectedRegion As Range
    Dim newRegion As Range
    Dim cel As Range
    
    ThisWorkbook.Sheets.Add.Name = newSheetName                         'create new sheet
    Set newClientSheet = ThisWorkbook.Sheets(newSheetName)
    
    Set selectedRegion = oldSheet.Range("A1").CurrentRegion             ' find active region in old worksheet
    selectedRegion.Copy newClientSheet.Cells
    
    oldSheet.Delete
    
    Set newRegion = newClientSheet.Range("A1").CurrentRegion
    
    For Each cel In newRegion                                           ' add format
        cel.Borders.Color = RGB(192, 192, 192)
        cel.Interior.Color = RGB(255, 255, 255)
        cel.Font.Color = RGB(0, 0, 0)
        cel.WrapText = True
        cel.VerticalAlignment = xlCenter
        cel.EntireColumn.AutoFit
        If cel.ColumnWidth > 75 Then cel.ColumnWidth = 75
    Next cel
End Sub


Sub PastePageNameByID()
    Dim ws As Worksheet
    Dim targetSheet As Worksheet
    Dim idColumn As Range
    Dim cell As Range
    Dim selectedRegion As Range
    Dim lastRow As Long
    Dim resultStr As String
    Dim newSheetName As String
    
    newSheetName = InputBox("Saisir le nom du Client")
    
    ' Set the worksheet where IDs are located
    Set targetSheet = ThisWorkbook.Sheets("Indicators results")
    
    Set idColumn = targetSheet.Range("A1:A" & targetSheet.Cells(targetSheet.Rows.Count, "A").End(xlUp).Row)     ' Define the column where the IDs are located - parsing through whole column returns error because column is huge
    
    
    'Set ws = Worksheets("SI000155")
    ' Call DelColumns(ws)
    'resultStr = RowToString(ws)

    
    For Each cell In idColumn                                           ' Double for loop to parse through ID colum and page names - Targets only used cells
        For Each ws In Worksheets
            If cell.Value = ws.Name Then
                Call DelColumns(ws)
                resultStr = RowToString(ws)
                targetSheet.Cells(cell.Row, 10).Value = resultStr       ' 10 if "remediation" column is included, else use 9
                ws.Delete
                
                ' Debug.Print cell.Value
                ' Debug.Print targetSheet.Cells(cell.Row, 9).Address
                ' Debug.Print cell.Row
                ' Debug.Print cell.Value
            End If
        Next ws
    Next cell
    
    Call formatClientSheet(newSheetName, targetSheet)
    
    targetSheet.Columns(1).Delete
    Call DelQueries
    Debug.Print

End Sub



