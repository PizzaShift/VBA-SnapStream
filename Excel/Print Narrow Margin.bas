
      Sub printNarrowMargin()
      With ActiveSheet.PageSetup
      .LeftMargin= Application
      .InchesToPoints(0.25)
      .RightMargin= Application.InchesToPoints(0.25)
      .TopMargin= Application.InchesToPoints(0.75)
      .BottomMargin= Application.InchesToPoints(0.75)
      .HeaderMargin= Application.InchesToPoints(0.3)
      .FooterMargin= Application.InchesToPoints(0.3)
      End With
      ActiveWindow.SelectedSheets.PrintOutCopies:=1, Collate:=True,
      IgnorePrintAreas:=False
      End Sub

