'You'll need to set a reference to the Microsoft VBScript Regular Expressions 5.5 library in Tools, References.
Sub GetValueUsingRegEx()
 ' Set reference to VB Script library
 ' Microsoft VBScript Regular Expressions 5.5
 
    Dim olMail As Outlook.MailItem
    Dim Reg1 As RegExp
    Dim M1 As MatchCollection
    Dim M As Match
        
    Set olMail = Application.ActiveExplorer().Selection(1)
   ' Debug.Print olMail.Body
    
    Set Reg1 = New RegExp
    
    ' \s* = invisible spaces
    ' \d* = match digits
    ' \w* = match alphanumeric
    
    With Reg1
        .Pattern = "Carrier Tracking ID\s*[:]+\s*(\w*)\s*"
        .Global = True
    End With
    If Reg1.test(olMail.Body) Then
    
        Set M1 = Reg1.Execute(olMail.Body)
        For Each M In M1
            ' M.SubMatches(1) is the (\w*) in the pattern
            ' use M.SubMatches(2) for the second one if you have two (\w*)
            Debug.Print M.SubMatches(1)
            
        Next
    End If
    
End Sub
