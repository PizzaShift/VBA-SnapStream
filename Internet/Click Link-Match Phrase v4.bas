Sub clickOnLink()

Dim IE As New InternetExplorer
Dim str As String
Dim Doc As HTMLDocument
Dim tagElements As Object
Dim element As HTMLObjectElement

str = "https://software-solutions-online.com/"

IE.navigate str
IE.Visible = True

Do
DoEvents
Loop Until IE.readyState = READYSTATE_COMPLETE

Set Doc = IE.document

'First get all the links

Set tagElements = Doc.all.tags("a")

' Loop through all the links
For Each element In tagElements

' Look for the link that contains the text Excel
If element.innerText = "Excel" Then

'Click on the link
element.Click

Do
DoEvents
Loop Until IE.readyState = READYSTATE_COMPLETE

Exit For
End If
Next
