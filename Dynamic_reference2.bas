Attribute VB_Name = "Dynamic_reference2"
Sub Exercise06a()


'get last row
    If Cells(Rows.Count, 1).End(xlUp).Row <> 1 Then
        lastRow1 = Cells(Rows.Count, 1).End(xlUp).Row
    Else
        lastRow1 = 2
    End If
    
'clear last report
Range("A2:b" & lastRow1).ClearContents



Range("a2") = "1"
Range("b2") = "Name1"

Range("a3") = "2"
Range("b3") = "Name2"

Range("a4") = "3"
Range("b4") = "Name3"


End Sub


Sub Exercise06b()

selRow = Selection.Row 'get the row number of the selected cell

If Cells(selRow, 1) = Empty Then Exit Sub

answer = MsgBox("Add 100 to current row sales?", vbYesNo)

If answer = vbYes Then
    'add 100 to current row
    Cells(selRow, 4) = Cells(selRow, 4) + 100
    
End If

End Sub

Sub myOffset()

Selection.Offset(3, 1) = Selection
Cells(1, 1).Offset(3, 1) = "Tomatos"
Range("A4").Offset(3, 1) = "Bananas"

End Sub
