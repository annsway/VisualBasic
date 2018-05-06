Attribute VB_Name = "Range_Object"
Option Explicit

Sub CreateHeaders()

Range("A1") = "ID"
Range("B1") = "First Name"
Range("C1") = "Last Name"

Range("A1:C1").Font.Bold = True


End Sub

Sub ShowValue()

MsgBox Range("A1").Value
MsgBox Range("A1").Text

End Sub

Sub Selection()
Range("tax_table").Select
End Sub


Sub CountCell()
MsgBox Range("tax_table").Count

End Sub

Sub Exercise02()
Range("A1") = "ID"
Range("B1") = "Name"
Range("A1:B1").Font.Bold = True

Range("A2") = 1
Range("A3") = 2
Range("A4") = 3

Range("B2:B4") = "Name"

Range("B2") = Range("B2") & 1
Range("B3") = Range("B3") & 2
Range("B4") = Range("B4") & 3

Range("C1").Select

'''''''

Range("C1") = "Sales"
Range("C2") = 10
Range("C3") = 13
Range("C4") = 21

Range("C1").Font.Bold = True
Range("C2:C4").NumberFormat = "0.00"

Range("C5").Formula = "=SUM(C2:C4)"

With Range("C4").Borders(xlEdgeBottom)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With

End Sub



Sub myAddress()
MsgBox Range("a1:c6").Address(0, 1) 'absolute reference for columns

End Sub


Sub myFormula()
Range("c9").Formula = "=sum(tax_table)"
Range("c9") = Range("c9").Value

End Sub


Sub myNumberFormat()
Range("tax_table").NumberFormat = "$0.00"

End Sub

Sub myFonts()
Range("tax_table").Font.Bold = True
Range("tax_table").Font.Italic = True
Range("c1:c3").Font.Underline = True

Range("tax_table").Font.Bold = False

End Sub













