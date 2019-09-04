
'Nested
Sub IIF_Example2()

Dim FinalResult As String

Dim Marks As Long

Marks = 98

FinalResult = IIf(Marks > 90, "Dist", IIf(Marks > 80, "First", IIf(Marks > 70, "Second", IIf(Marks > 60, "Third", "Fail"))))

MsgBox FinalResult

End Sub


'Not Nested
Sub IIF_Example()

Dim FinalResult As String

Dim Number1 As Long
Dim Number2 As Long

Number1 = 105
Number2 = 100

FinalResult = IIf(Number1 > Number2, "Number 1 is Greater than Number 2", "Number 1 is Less than Number 2")

MsgBox FinalResult

End Sub
