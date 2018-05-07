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
lastRow3 = Cells(Rows.Count, 1).End(xlUp).Row

For x = 2 To lastRow3
    If Cells(x, 4) > 500 Then
        mymsg = mymsg & vbNewLine & Cells(x, 1)
    End If
Next x

MsgBox mymsg

End Sub
