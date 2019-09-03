
      Sub printCustomSelection()
      Dim startpageAs Integer
      Dim endpageAs Integer
      startpage= InputBox("Please Enter Start Page number.", "Enter Value")
      If Not WorksheetFunction.IsNumber(startpage) Then
      MsgBox"Invalid Start Page number. Please try again.", "Error"
      Exit Sub
      End If
      endpage= InputBox("Please Enter End Page number.", "Enter Value")
      If Not WorksheetFunction.IsNumber(endpage) Then
      MsgBox"Invalid End Page number. Please try again.", "Error"
      Exit Sub
      End If
      Selection.PrintOutFrom:=startpage, To:=endpage, Copies:=1
      Collate:=True
      End Sub

