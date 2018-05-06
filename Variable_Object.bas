Attribute VB_Name = "Variable_Object"
Private myPriVar As Date
Const myName As String = "Ann Zhou"


'Define a variable
Sub MyExample()
myVar = 50
MsgBox myVar 'Note: there're no quotes around myVar, otherwise it's text
Call myVarDec

MsgBox myDate 'myDate won't show up b/c it's cleared by system to save memories

End Sub

Sub myVarDec()
Dim myString As String
Dim myDouble As Double
Dim myDate As Date

myString = "Hello World"
myDouble = 123.4444
myPriVar = "12/31/2018"

End Sub

Sub myConcat()
myText = "I love " & myName
MsgBox myText & " on " & myPriVar

End Sub


Sub DateMath()
myDate = Date - 1
myTime = Now

End Sub

Sub getDaysOld()

Dim inputdob As Date ' declare variable to hold value returned by InputBox

inputdob = InputBox(Prompt:="Enter your DOB", Title:="getDaysOld", Default:=Date - 1)
DaysOld = Date - inputdob
HoursOld = DaysOld * 24

Text = "You are " & DaysOld & " days old, " & HoursOld & " hours old in total. "
MsgBox Text

End Sub

Sub CallingSub()
a = 10
b = 20
CalledSub y:=a, z:=b 'assigning values to the parameters of CalledSub

MsgBox "a is now " & a & " and b is " & b

End Sub

Sub CalledSub(ByRef y, ByVal z)
y = 100 'ByRef: if the value of y is changed (to 100), then it will return the new value (100), instead of the one passed in just now (10)
z = 200

End Sub
















