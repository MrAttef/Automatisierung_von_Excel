Sub filternUndDelete()
    Dim ws As Worksheet
    Set ws = ActiveSheet ' or specify a worksheet: Set ws = ThisWorkbook.Worksheets("Sheet1")

    ' Set the column to filter
    Dim filterColumn As Long
    filterColumn = 24 ' column E

    ' Apply the filter
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, filterColumn).End(xlUp).Row
    ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column)).AutoFilter field:=filterColumn, Criteria1:="=*DENIED*"

    ' Delete the filtered rows
    Dim filteredRange As Range
    Set filteredRange = ws.AutoFilter.Range.Offset(1).Resize(ws.AutoFilter.Range.Rows.Count - 1)
    If Application.WorksheetFunction.Subtotal(103, filteredRange.Columns(1)) > 1 Then
        filteredRange.SpecialCells(xlCellTypeVisible).EntireRow.Delete
    End If

    ' Turn off the filter
    ws.AutoFilterMode = False
End Sub
Sub filternund0zu1()
    Dim lastRow As Long
    Dim i As Long
    Dim ws As Worksheet
    
    ' Get a reference to the active worksheet
    Set ws = ActiveSheet
    
    ' Find the last row in column L
    lastRow = ws.Cells(Rows.Count, "L").End(xlUp).Row
    
    ' Loop through each row and change the value in column L to 1 if it is 0
    For i = 1 To lastRow
        If ws.Cells(i, "L").Value = 0 Then
            ws.Cells(i, "L").Value = 1
        End If
    Next i
End Sub
Sub filternund0zu00()

    'Declare variables
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    'Set the worksheet to be worked on
    Set ws = ActiveSheet
    
    'Get the last row in column G
    lastRow = ws.Cells(Rows.Count, "G").End(xlUp).Row
    
    'Loop through all cells in column G and replace "0" with "0.0"
    For i = 1 To lastRow
        If ws.Cells(i, "G").Value = 0 Then
            ws.Cells(i, "G").Value = "0.030"
        End If
    Next i
    
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
Sub STEP2()
    'Get the active worksheet
    Dim ws As Worksheet
    Set ws = ActiveSheet

    'Set the value of cell A1 "
    ws.Range("A2").Formula = "=F2&G2&H2&I2"
End Sub
Sub STEP22()
    'Get the active worksheet
    Dim ws As Worksheet
    Set ws = ActiveSheet

    'Set the value of cell A1 "
    ws.Range("A2").Formula = "=H2&J2&K2&L2"
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

Sub STEP4()
    'Get the active worksheet
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    'Set the value"
    ws.Range("bO2").Formula = "=VLOOKUP($A2,'[TM_Baulist.csv]TM_Baulist'!$1:$1048576,20,FALSE)"
    ws.Range("bP2").Formula = "=VLOOKUP($A2,'[TM_Baulist.csv]TM_Baulist'!$1:$1048576,23,FALSE)"
    ws.Range("bq2").Formula = "=VLOOKUP($A2,'[TM_Baulist.csv]TM_Baulist'!$1:$1048576,21,FALSE)"
    ws.Range("br2").Formula = "=VLOOKUP($A2,'[TM_Baulist.csv]TM_Baulist'!$1:$1048576,22,FALSE)"
End Sub
Sub STEP44()
    'Get the active worksheet
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    'Set the value"
    ws.Range("bt2").Formula = "=Countif(A:A,a2)"
End Sub
Sub STEP5()
    'Get reference to the active worksheet
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    'Copy the value in cell A1 to cells B1 through B5
    For i = 1 To ENDTABELLE
        ws.Range("A1").Copy Destination:=ws.Range("B" & i)
    Next i
End Sub
Sub BT()
    'Get a reference to the active sheet
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    'Get a reference to the second cell in the column
    Dim secondCell As Range
    Set secondCell = ws.Range("BT2")
    
    'Get a reference to the last used cell in the column
    Dim lastCell As Range
    Set lastCell = ws.Cells(ws.Rows.Count, secondCell.Column).End(xlUp)
    
    'Fill the column with the value of the second cell
    secondCell.Resize(lastCell.Row - secondCell.Row + 1).FillDown
End Sub
Sub BO()
    'Get a reference to the active sheet
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    'Get a reference to the second cell in the column
    Dim secondCell As Range
    Set secondCell = ws.Range("BO2")
    
    'Get a reference to the last used cell in the column
    Dim lastCell As Range
    Set lastCell = ws.Cells(ws.Rows.Count, secondCell.Column).End(xlUp)
    
    'Fill the column with the value of the second cell
    secondCell.Resize(lastCell.Row - secondCell.Row + 1).FillDown
End Sub

Sub BP()
    'Get a reference to the active sheet
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    'Get a reference to the second cell in the column
    Dim secondCell As Range
    Set secondCell = ws.Range("BP2")
    
    'Get a reference to the last used cell in the column
    Dim lastCell As Range
    Set lastCell = ws.Cells(ws.Rows.Count, secondCell.Column).End(xlUp)
    
    'Fill the column with the value of the second cell
    secondCell.Resize(lastCell.Row - secondCell.Row + 1).FillDown
End Sub
Sub BQ()
    'Get a reference to the active sheet
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    'Get a reference to the second cell in the column
    Dim secondCell As Range
    Set secondCell = ws.Range("BQ2")
    
    'Get a reference to the last used cell in the column
    Dim lastCell As Range
    Set lastCell = ws.Cells(ws.Rows.Count, secondCell.Column).End(xlUp)
    
    'Fill the column with the value of the second cell
    secondCell.Resize(lastCell.Row - secondCell.Row + 1).FillDown
End Sub
Sub BR()
    'Get a reference to the active sheet
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    'Get a reference to the second cell in the column
    Dim secondCell As Range
    Set secondCell = ws.Range("BR2")
    
    'Get a reference to the last used cell in the column
    Dim lastCell As Range
    Set lastCell = ws.Cells(ws.Rows.Count, secondCell.Column).End(xlUp)
    
    'Fill the column with the value of the second cell
    secondCell.Resize(lastCell.Row - secondCell.Row + 1).FillDown
End Sub
Sub filtern()
    'Get a reference to the active worksheet
    Dim ws As Worksheet
    Set ws = ActiveSheet

    'Activate the filter in the range A1:D100
    ws.Range("A1:EA1").AutoFilter
End Sub
Sub CodeBAULIST()
    STEP1
    STEP22
    STEP3
    filtern
    filternUndDelete
    filternund0zu1
    'Add additional subroutine calls here as needed
End Sub
Sub CodeKUNDENDATEN()
    STEP1
    STEP2
    STEP3
    STEP4
    STEP5
    BO
    BP
    BQ
    BR
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

