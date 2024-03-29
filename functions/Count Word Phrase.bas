
      '1. You may treat the body as a regular string returned by the Body property of the MailItem 
      'class. However, the easiest way to count the number of words is using the Word object model. 
      'The Inspector class provides the WordEditorproperty which returns the Microsoft Word Document
      'Object Model of the  message being displayed. Here is what MSDN states: The WordEditor property
      'is only valid if the IsWordMail method returns True and the EditorType property is olEditorWord. The returned WordDocument object provides access to most of the Word object model.
      'The VBA: How to Count the Occurrences of a Word or Phrase article provides the source code in 
      'VBA for getting the job done. 
      
      '[HOW TO? Count words in an email body using VB/VBA?](https://social.msdn.microsoft.com/Forums/windows/en-US/fbdd7eed-9f84-486a-802c-f6a2685b92cf/how-to-count-words-in-an-email-body-using-vbvba?forum=outlookdev)'
      
      Sub CountWordPhrase()
      
      Dim x, Response, ExitResponse
      Dim y As Integer
      
      ' If an error occurs, continue the macro.
      On Error Resume Next
      
      AskAgain:
      
      ' Ask for the text to count.
      x = InputBox("Type the word you want to count and then click OK." _
      & Chr$(13) & Chr$(13) & _
      "NOTE: This macro will find a whole word only. If the text you typed " _
      & "is part of a larger string, it will also be found.")
      
      ' If text typed is blank or spaces, then ask to quit.
      If x = "" Or x = " " Then
          ExitResponse = MsgBox("You either clicked Cancel or you did " & _
          "not type a word. Do you want to quit?", vbYesNo)
          
          If ExitResponse = 6 Then
              End
          Else
              ' If answer No to quit, then ask for text to count again.
              GoTo AskAgain
          End If
      Else
      
          ' Search for and count occurrences of the text typed.
          With ActiveDocument.Content.find
              Do While .Execute(FindText:=x, Forward:=True, Format:=True, _
                 MatchWholeWord:=True) = True
                  
                 ' Display message in Word's Status Bar.
                 StatusBar = "Word is counting the occurrences of the text " & _
                 Chr$(34) & x & Chr$(34) & "."
                 
                 y = y + 1
              Loop
          End With
      
      ' Display Message Box with results.
      Response = MsgBox("The text " & Chr$(34) & x & Chr$(34) & " was found" _
      & Str$(y) & " times.", vbOKOnly)
      
      End If
      End Sub

