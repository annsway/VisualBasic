Module Module1
    Sub Main()
        Dim myNum As Double
        Dim myMod As Integer
        Dim myString As String = "Hello World"
        Dim myBool As Boolean = 5 > 9
        Dim myAns As Double = Nothing

        myNum = 5 / 2
        myMod = 13 Mod 5
        myAns = myNum * myMod

        'This line writes the text hello world to the console'
        Console.WriteLine(myNum.ToString() & " - " & myMod & " - " & myAns & " - " & myString & " - " & myBool)

        'This line pauses the program and lets the user read the text'
        Console.ReadLine()

    End Sub
End Module