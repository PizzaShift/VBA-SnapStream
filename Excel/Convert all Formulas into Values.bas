
      Sub ConvertToValues()
      Dim MyRange As Range
      Dim MyCell As Range
      Select Case MsgBox("You Can't Undo This Action. " & "Save
      Workbook First?", vbYesNoCancel, "Alert")
      Case Is = vbYes
      ThisWorkbook.Save
      Case Is = vbCancel
      Exit Sub
      End Select
      Set MyRange = Selection
      For Each MyCell In MyRange
      If MyCell.HasFormula Then
      MyCell.Formula = MyCell.Value
      End If
      Next MyCell
      End Sub

