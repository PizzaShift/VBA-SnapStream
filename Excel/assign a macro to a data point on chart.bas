
      Private Sub Cht_Select(ByVal ElementID As Long, ByVal Arg1 As Long, ByVal Arg2 As Long)
      
      Dim vntData As Variant
      Dim vntXData As Variant
      
      If ElementID = 3 Then
      ' selected series
      If Arg2 = -1 Then
      ' selected whole series
      'MsgBox "Selected Series [" & Arg1 & "] " & Cht.SeriesCollection(Arg1).Name, vbInformation, "Series"
      If Arg1 = "3" Then
      Sheets("Worksheet 1").Select
      Else
      If Arg1 = "4" Then
      Sheets("Worksheet 2").Select
      Else
      If Arg1 = "5" Then
      Sheets("Worksheet 3").Select
      Else
      If Arg1 = "6" Then
      Sheets("Worksheet 4").Select
      Else
      If Arg1 = "7" Then
      Sheets("Worksheet 5").Select
      Else
      If Arg1 = "8" Then
      Sheets("Worksheet 6").Select
      Else
      If Arg1 = "9" Then
      Sheets("Worksheet 7").Select
      Else
      If Arg1 = "10" Then
      Sheets("Worksheet 8").Select
      Else
      If Arg1 = "11" Then
      Sheets("Worksheet 9").Select
      Else
      If Arg1 = "12" Then
      Sheets("Worksheet 10").Select
      End If
      End If
      End If
      End If
      End If
      End If
      End If
      End If
      End If
      End If
      Else
      ' selected data point
      vntData = Cht.SeriesCollection(Arg1).Values
      vntXData = Cht.SeriesCollection(Arg1).XValues
      'MsgBox "Selected DataPoint " & Arg2 & " from Series [" & Arg1 & "] " & Cht.SeriesCollection(Arg1).Name & vbLf & _
      vntXData(Arg2) & "," & vntData(Arg2), vbInformation, "Series"
      ' Sheets("worksheet 1").Select
      
      End If
      End If
      
      End Sub

