
Sub followWebsiteLink()
‘From Tools —-> References activate
‘1. Microsoft HTML Object Library
‘2. Microsoft Internet Controls

Dim ie As InternetExplorer
Dim html As HTMLDocument
Dim Link As Object
Dim ElementCol As Object

Application.ScreenUpdating = False

Set ie = New InternetExplorer
ie.Visible = True

ie.navigate “https://www.google.co.in”

Do While ie.readyState <> READYSTATE_COMPLETE
Application.StatusBar = “Loading website…”
DoEvents
Loop

Set html = ie.document
Set ElementCol = html.getElementsByTagName(“a”)

For Each Link In ElementCol
If Link.innerHTML = “Advertising” Then
Link.Click
End If
Next Link

Set ie = Nothing
Application.StatusBar = “”
Application.ScreenUpdating = True
End Sub
