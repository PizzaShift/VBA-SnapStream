
      Sub HighlightLowerThanValues()
      Dim i As Integer
      i = InputBox("Enter Lower Than Value", "Enter Value")
      Selection.FormatConditions.Delete
      Selection.FormatConditions.Add Type:=xlCellValue,
      Operator:=xlLower, Formula1:=i
      Selection.FormatConditions(Selection.FormatConditions.Count).S
      tFirstPriority
      With Selection.FormatConditions(1)
      .Font.Color = RGB(0, 0, 0)
      .Interior.Color = RGB(217, 83, 79)
      End With
      End Sub

