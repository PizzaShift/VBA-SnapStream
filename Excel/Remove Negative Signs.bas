
      Sub removeNegativeSign()
      Dim rngAs Range
      Selection.Value= Selection.Value
      For Each rngIn Selection
      If WorksheetFunction.IsNumber(rng)
      Then rng.Value= Abs(rng)
      End If
      Next rng
      End Sub

