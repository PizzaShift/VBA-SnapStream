
Private Sub Command1_MouseDown() '(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim o2
    Set o2 = CreateObject("internetexplorer.application")
        o2.navigate "http://www.vbaccelerator.com/home/vb/code/Controls/S_Grid/S-Grid_Documentation.asp" 'IE navigates to a webpage
        o2.Visible = False 'hides IE
        While o2.busy: DoEvents: Wend
     
Set o = o2.Document.All.tags("A")


M = o.length: mySubmit = -1
For r = 0 To M - 1

If InStr(1, o.Item(r).innerhtml, "S-Grid Documentation.zip (5K)", vbTextCompare) Then
    o.Item(r).focus
    o.Item(r).Click: Exit For
End If
Next

mySleep 50 'wait five seconds
SendKeys "%S" 'Hits the save button
mySleep 50 'wait five seconds and starts typing in the path c:\test.zip to save the file
SendKeys "c"
SendKeys ":"
SendKeys "\"
SendKeys "t"
SendKeys "e"
SendKeys "s"
SendKeys "t"
SendKeys "."
SendKeys "z"
SendKeys "i"
SendKeys "p"

mySleep 10 'waits 1 second
SendKeys "%S" 'Hits the save button
SendKeys "%c" 'Hits the Close this dialog box when download completes

mySleep 20 'Waits 2 seconds and Msgbox appears: 'file exists - do you want to overwrite'
SendKeys "%y"

o2.Quit
Set o2 = Nothing
End Sub

Sub mySleep(ByVal deciSeconds As Long)
    t = Timer: While Timer - t < (deciSeconds / 10): DoEvents: Wend
End Sub 
