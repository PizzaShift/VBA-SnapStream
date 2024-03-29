      Sub CopyToExcel()
      Dim olItem As Outlook.MailItem
      Dim xlApp As Object
      Dim xlWB As Object
      Dim xlSheet As Object
      Dim vText As Variant
      Dim sText As String
      Dim vItem As Variant
      Dim i As Long
      Dim rCount As Long
      Dim bXStarted As Boolean
      Const strPath As String = "O:\\FSR\\_excel\\From_email_to_excel.xlsm"        
      
          On Error Resume Next
          Set olItem = ActiveExplorer.Selection
          Set xlApp = GetObject(, "Excel.Application")
          If Err <> 0 Then
              Application.StatusBar = "Please wait while Excel source is opened ... "
              Set xlApp = CreateObject("Excel.Application")
              bXStarted = True
          End If
          On Error GoTo 0
          'Open the workbook to input the data
          Set xlWB = xlApp.Workbooks.Open(strPath)
          Set xlSheet = xlWB.Sheets("sheet1")
      
          'Process the message record
          For Each olItem In Application.ActiveExplorer.Selection
          sText = olItem.Body
          vText = Split(sText, Chr(13))
          'Find the next empty line of the worksheet
          rCount = xlSheet.Range("A" & xlSheet.Rows.Count).End(-4162).Row
          rCount = rCount + 1
      
          'Check each line of text in the message body
          For i = UBound(vText) To 0 Step -1
      
              If InStr(1, vText(i), "Job Ref :") > 0 Then
                  vItem = Split(vText(i), Chr(58))
                  xlSheet.Range("A" & rCount) = Trim(vItem(1))
              End If
      
              If InStr(1, vText(i), "Serial No. :") > 0 Then
                  vItem = Split(vText(i), Chr(58))
                  xlSheet.Range("C" & rCount) = Trim(vItem(1))
              End If
      
              If InStr(1, vText(i), "Serial No. 1 :") > 0 Then
                  vItem = Split(vText(i), Chr(58))
                  xlSheet.Range("F" & rCount) = Trim(vItem(1))
              End If
      
              If InStr(1, vText(i), "Serial No. 2 :") > 0 Then
                  vItem = Split(vText(i), Chr(58))
                  xlSheet.Range("G" & rCount) = Trim(vItem(1))
              End If
      
              If InStr(1, vText(i), "Site Name :") > 0 Then
                  vItem = Split(vText(i), Chr(58))
                  xlSheet.Range("I" & rCount) = Trim(vItem(1))
              End If
      
              If InStr(1, vText(i), "Fault :") > 0 Then
                  vItem = Split(vText(i), Chr(58))
                  xlSheet.Range("K" & rCount) = Trim(vItem(1))
              
              End If
                    
              
          Next i
          xlWB.Save
          Next olItem
          If bXStarted Then
              xlApp.Quit
          End If
      
          Set xlApp = Nothing
          Set xlWB = Nothing
          Set xlSheet = Nothing
      End Sub

