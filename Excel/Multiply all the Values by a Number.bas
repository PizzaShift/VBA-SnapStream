
      Sub multiplyWithNumber()
      Dim rng As Range
      Dim c As Integer c = InputBox("Enter number to multiple",
      "Input Required")
      For Each rng In Selection
      If WorksheetFunction.IsNumber(rng) Then
      rng.Value = rng * c
      Else
      End If
      Next rng
      End Sub

