Attribute VB_Name = "Dynamic_reference"
Sub lastRow()
lastRow1 = Cells(Rows.Count, 1).End(xlUp).Row
End Sub

Sub lastColumn()
lastCol = Cells(1, Columns.Count).End(xlToLeft).Column
End Sub

Sub nextRow()
nextRow1 = Cells(Rows.Count, 1).End(xlUp).Row + 1
End Sub

Sub DynamicSorting()
Range("A1").Select
lastRow2 = ThisWorkbook.Sheets("sheet1").Cells(Rows.Count, 1).End(xlUp).Row

    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Add Key:=Range("B2:B" & lastRow2) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

    With ActiveWorkbook.Worksheets("Sheet1").Sort
        .SetRange Range("A1:D" & lastRow2)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
        
    End With
    
End Sub

Sub DynamicRef()

ActiveCell.Offset(1, 0).Select 'select the cell that is 1 row below the selected row
ActiveCell.FormulaR1C1 = "hello world"

End Sub

Sub aboutWith()
Range("c6").Value = 12
Range("c6").Font.Bold = True
Range("C6").Font.Italic = True

With Range("C6")
     .Value = 12
     .Font.Bold = True
     .Font.Italic = True
     
End With

End Sub

Sub IfStatement()
If Range("d6") >= 250 Then
    MsgBox "Good!"
End If
    MsgBox "Not good!"
End Sub
 


