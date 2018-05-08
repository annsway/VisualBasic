Attribute VB_Name = "User_defined_functions"
Function kgrams(lbs, Optional decimal_places)

If IsMissing(decimal_places) Then
    kgrams = lbs * 0.453592
Else
    kgrams = Round(lbs * 0.453592, decimal_places)

End If

End Function


Sub myLoop()
For x = 1 To 10
    Cells(x, 1) = x * 10
    
    If Cells(x, 1) > 50 Then
        Cells(x, 2) = True
        Cells(x, 2).Font.Bold = True
        
    Else: Cells(x, 2) = False
    
    End If
Next x
End Sub

Sub myFirstReport()
Dim dsheet As Worksheet
Dim rptsheet As Worksheet

Set dsheet = ThisWorkbook.Sheets("data")      'name the sheet for "Data"
Set rptsheet = ThisWorkbook.Sheets("report")  'name the sheet for output "Report"

'Clear contents in the output report sheet
rptlr = rptsheet.Cells(Rows.Count, 1).End(xlUp).Row
If rptlr <> 1 Then
    rptlr = rptlr
Else
    rptlr = 2
End If

rptsheet.Range("a2:c" & rptlr).ClearContents

'Search by user input threshold from the Row 2 up to the last row in the "Data" sheet
lastRow3 = dsheet.Cells(Rows.Count, 1).End(xlUp).Row

myInput1 = InputBox("How much money do they make?", "Custom Sales Report", "300") + 0
y = 2 'starting row

answer = MsgBox("Add a column for Title?", vbYesNo)

If answer = vbYes Then
    rptsheet.Cells(1, 3) = "Title"
    rptsheet.Cells(1, 3).Font.Bold = True
Else
    rptsheet.Cells(1, 3) = ""
End If


For x = 2 To lastRow3
    If dsheet.Cells(x, 4) > myInput1 Then
        rptsheet.Cells(y, 1) = dsheet.Cells(x, 1)   'name column
        rptsheet.Cells(y, 2) = dsheet.Cells(x, 4)   'sale amount column
        If answer = vbYes Then
            rptsheet.Cells(y, 3) = dsheet.Cells(x, 2)   'title column
        End If
                
        y = y + 1 'increment one row
         
    End If
Next x

rptsheet.Visible = True
rptsheet.Select

End Sub

Sub Exercise07()
Dim report As Worksheet
Dim list As Worksheet

Set report = ThisWorkbook.Sheets("report")
Set list = ThisWorkbook.Sheets("list")

report.Select

reportLR = report.Cells(Rows.Count, 1).End(xlUp).Row
If reportLR <> 1 Then
    reportLR = reportLR
Else
    reportLR = 2
End If

list.Range("a2:c" & reportLR).ClearContents

myInput2 = InputBox("Show only the following user's records: ", "Individual User Report", "")
t = 2

list.Cells(1, 1) = report.Cells(1, 1)
list.Cells(1, 2) = report.Cells(1, 2)
list.Cells(1, 3) = report.Cells(1, 3)

list.Cells(1, 1).Font.Bold = True
list.Cells(1, 2).Font.Bold = True
list.Cells(1, 3).Font.Bold = True


For Z = 2 To reportLR
    nextRow2 = list.Cells(Rows.Count, 1).End(xlUp).Row + 1
    If report.Cells(Z, 1) = myInput2 Then
        list.Cells(nextRow2, 1) = report.Cells(Z, 1)
        list.Cells(nextRow2, 2) = report.Cells(Z, 2)
        list.Cells(nextRow2, 3) = report.Cells(Z, 3)
    End If
        
        t = t + 1
Next Z

list.Visible = True
list.Select


End Sub


