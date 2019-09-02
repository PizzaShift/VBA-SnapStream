
      Sub RemoveWrapText()
      Cells.Select
      Selection.WrapText = False
      Cells.EntireRow.AutoFit
      Cells.EntireColumn.AutoFit
      End Sub

