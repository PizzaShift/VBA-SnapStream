This sample searches for the link "S-Grid Documentation.zip (5K)". It will be found
in the window no. 25

The link will be clicked automatically, and the download starts.

Needs 1 Form and 1 Button :
===================

Private Sub Command1_Click()

    Dim o2
    Set o2 = CreateObject("internetexplorer.application")
        o2.navigate "http://www.vbaccelerator.com/home/vb/code/Controls/S_Grid/S-Grid_Documentation.asp"
        o2.Visible = True
        While o2.busy: DoEvents: Wend
     
AppActivate Me.Caption
Set o = o2.Document.All.tags("A")


M = o.length: mySubmit = -1
For r = 0 To M - 1: zz = ""

zz = zz & "Link Index : " & r & " of " & o.length - 1
zz = zz & String(3, vbCrLf)
'
zz = zz & "A . tabindex : " & o.Item(r).TabIndex
zz = zz & String(3, vbCrLf)
'
zz = zz & "A . tagname : " & o.Item(r).tagname
zz = zz & String(3, vbCrLf)
'
zz = zz & "A . href : " & o.Item(r).href
zz = zz & String(3, vbCrLf)
'
zz = zz & "A . type : " & o.Item(r).Type
zz = zz & String(3, vbCrLf)
'
zz = zz & "A . name : " & o.Item(r).Name
zz = zz & String(3, vbCrLf)
'
zz = zz & "A . innerhtml : " & o.Item(r).innerhtml
zz = zz & String(3, vbCrLf)
'
zz = zz & "A . outerhtml : " & o.Item(r).outerhtml
zz = zz & String(3, vbCrLf)
'
zz = zz & "A . rel : " & o.Item(r).rel
zz = zz & String(3, vbCrLf)
'
zz = zz & "A . rev : " & o.Item(r).rev
zz = zz & String(3, vbCrLf)
'
zz = zz & "A . id : " & o.Item(r).Id
zz = zz & String(3, vbCrLf)


MsgBox zz
If InStr(1, o.Item(r).innerhtml, "S-Grid Documentation.zip (5K)", vbTextCompare) Then
    MsgBox "= F O U N D =" & vbCrLf & o.Item(r).innerhtml
    o.Item(r).Click: Exit For
End If
Next


'        o2.Quit
'        Set o2 = Nothing
End Sub
