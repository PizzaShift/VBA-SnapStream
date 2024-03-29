      '[Use RegEx to extract text from an Outlook email message](https://www.slipstick.com/developer/regex-parse-message-text/)'
      'You'll need to set a reference to the:
      ' "Microsoft VBScript Regular Expressions 5.5" 
      'library in Tools, References.
      '
      Sub GetValueUsingRegEx()
          Dim olMail As Outlook.MailItem
          Dim Reg1 As RegExp
          Dim M1 As MatchCollection
          Dim M As Match
          Dim strSubject As String
          Dim testSubject As String
               
          Set olMail = Application.ActiveExplorer().Selection(1)
           
          Set Reg1 = New RegExp
          
      For i = 1 To 3
      
      With Reg1
          Select Case i
          Case 1
              .Pattern = "(Order ID\\s[:]([\\w-\\s]*)\\s*)\\n"
              .Global = False
              
          Case 2
             .Pattern = "(Date[:]([\\w-\\s]*)\\s*)\\n"
             .Global = False
             
          Case 3
              .Pattern = "(([\\d]*\\.[\\d]*))\\s*\\n"
              .Global = False
          End Select
          
      End With
          
          
          If Reg1.test(olMail.Body) Then
           
              Set M1 = Reg1.Execute(olMail.Body)
              For Each M In M1
                  Debug.Print M.SubMatches(1)
                  strSubject = M.SubMatches(1)
                  
               strSubject = Replace(strSubject, Chr(13), "")
               testSubject = testSubject & "; " & Trim(strSubject)
               Debug.Print i & testSubject
               
               Next
          End If
                
      Next i
      
      Debug.Print olMail.Subject & testSubject
      olMail.Subject = olMail.Subject & testSubject
      olMail.Save
      
      Set Reg1 = Nothing
           
      End Sub

