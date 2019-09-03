
      Sub dateInHeader()
      With ActiveSheet.PageSetup
      .LeftHeader = ""
      .CenterHeader = "&D"
      .RightHeader = ""
      .LeftFooter = ""
      .CenterFooter = ""
      .RightFooter = ""
      End With
      ActiveWindow.View = xlNormalView
      End Sub

