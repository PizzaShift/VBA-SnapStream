
      'Goal Seek can be super helpful for you to solve complex problems.
      'Learn more about goal seek from here before you use this code.
      '[How to use Goal Seek in Excel to Find a Value (Manual + VBA)](https://excelchamps.com/blog/goal-seek/)'
      Sub GoalSeekVBA()
      Dim Target As Long
      On Error GoTo Errorhandler
      Target = InputBox("Enter the required value", "Enter Value")
      Worksheets("Goal_Seek").Activate
      With ActiveSheet .Range("C7")
      .GoalSeek_ Goal:=Target, _
      ChangingCell:=Range("C2")
      End With
      Exit Sub
      Errorhandler: MsgBox("Sorry, value is not valid.")
      End Sub

