Attribute VB_Name = "Cell_Object"
Sub cell_1()
Cells(1, 2) = 50
End Sub


Sub cell_11()
Cells(1, "x") = 20
End Sub

Sub cell_2()
Range("A2:C10").Cells(6) = 44 'Look for the 6th cell from left to right, up to down

End Sub

Sub cell_3()
Cells.Font.Name = "Arial"
Cells.Font.Size = 15
End Sub

Sub Exercise03()
Cells.Font.Name = "Arial"
ActiveWindow.Zoom = 145
Range("D:D").Style = "Currency"
End Sub
