
      Sub addNumber()
      Dim rngAs Range
      DimiAs Integer
      i= InputBox("Enter number to multiple", "Input Required")
      For Each rng In Selection
      If WorksheetFunction.IsNumber(rng) Then
      rng.Value= rng+ i
      Else
      End If
      Next rng
      End Sub

