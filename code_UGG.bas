Sub filternUndDelete()
    Dim ws As Worksheet
    Set ws = ActiveSheet ' or specify a worksheet: Set ws = ThisWorkbook.Worksheets("Sheet1")

    ' Set the column to filter
    Dim filterColumn As Long
    filterColumn = 23 ' column E

    ' Apply the filter
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, filterColumn).End(xlUp).Row
    ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column)).AutoFilter field:=filterColumn, Criteria1:="=*CANCELLED*"

    ' Delete the filtered rows
    Dim filteredRange As Range
    Set filteredRange = ws.AutoFilter.Range.Offset(1).Resize(ws.AutoFilter.Range.Rows.Count - 1)
    If Application.WorksheetFunction.Subtotal(103, filteredRange.Columns(1)) > 1 Then
        filteredRange.SpecialCells(xlCellTypeVisible).EntireRow.Delete
    End If

End Sub
Sub STEP1()
    'Get the active worksheet
    Dim ws As Worksheet
    Set ws = ActiveSheet

    'Determine the location to insert the new cell
    Dim insertRow As Long
    insertRow = 1

    'Make room for the new cell by shifting existing cells down
    ws.Range("A" & insertRow & ":A" & ws.Rows.Count).Insert Shift:=xlDown
End Sub
Sub STEP11()
    'Get the active worksheet
    Dim ws As Worksheet
    Set ws = ActiveSheet
    ws.Range("bo1").EntireColumn.Insert

End Sub
Sub STEP2()
    'Get the active worksheet
    Dim ws As Worksheet
    Set ws = ActiveSheet

    'Set the value of cell A1 "
    ws.Range("A2").Formula = "=E2"
End Sub
Sub STEP22()
    'Get the active worksheet
    Dim ws As Worksheet
    Set ws = ActiveSheet

    'Set the value of cell A1 "
    ws.Range("A2").Formula = "=J2"
End Sub
Sub STEP222()
    'Get the active worksheet
    Dim ws As Worksheet
    Set ws = ActiveSheet

    'Set the value of cell A1 "
    ws.Range("A2").Formula = "=F2&G2&H2"
End Sub
Sub STEP2BAULIST()
    'Get the active worksheet
    Dim ws As Worksheet
    Set ws = ActiveSheet

    'Set the value of cell A1 "
    ws.Range("A2").Formula = "==H2&J2&K2&L2"
End Sub
Sub STEP3()
    'Get a reference to the active sheet
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    'Get a reference to the second cell in the column
    Dim secondCell As Range
    Set secondCell = ws.Range("A2")
    
    'Get a reference to the last cell in the second column
    Dim lastCell As Range
    Set lastCell = ws.Cells(ws.Rows.Count, secondCell.Column + 1).End(xlUp)
    
    'Fill the column with the value of the second cell until the last cell in the second column
    secondCell.Resize(lastCell.Row - secondCell.Row + 1).FillDown
End Sub
Sub STEP44()
    'Get the active worksheet
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    'Set the value"
    ws.Range("bt2").Formula = "=Countif(A:A,a2)"
End Sub
Sub filtern()
    'Get a reference to the active worksheet
    Dim ws As Worksheet
    Set ws = ActiveSheet

    'Activate the filter in the range A1:D100
    ws.Range("A1:EA1").AutoFilter
End Sub
Sub Als_werte_hinzufügen()

'Get a reference to the active worksheet
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ws.Cells.Select
    ws.Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ws.Columns("A:A").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
End Sub
Sub CodeBAULISTUGG()
    STEP1
    STEP22
    STEP3
    filternUndDelete
  
    'Add additional subroutine calls here as needed
End Sub
Sub CodeKUNDENDATEN()
    STEP1
    STEP2
    STEP3
    STEP11
    filtern
    'Add additional subroutine calls here as needed
End Sub
Sub CodeUnits_total()
    STEP1
    STEP222
    STEP3
    STEP44
    filtern
    'Add additional subroutine calls here as needed
End Sub
