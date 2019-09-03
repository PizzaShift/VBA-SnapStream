
      Sub DisablePageBreaks()
      Dim wbAs Workbook
      Dim wksAs Worksheet
      Application.ScreenUpdating= False
      For Each wbIn Application.Workbooks
      For Each ShtIn wb.WorksheetsSht.DisplayPageBreaks= False
      Next Sht
      Next wb
      Application.ScreenUpdating= True
      End Sub

