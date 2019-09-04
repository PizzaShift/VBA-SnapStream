Sub extractVancouverData()

Dim IE As New InternetExplorer
 Dim url As String
 Dim item As HTMLHtmlElement
 Dim Doc As HTMLDocument
 Dim tagElements As Object
 Dim element As Object
 Dim lastRow
 
 Application.ScreenUpdating = False
 Application.DisplayAlerts = False
 Application.EnableEvents = False
 Application.Calculation = xlCalculationManual
 
 url = "https://www.theweathernetwork.com/ca/weather/british-columbia/vancouver"

IE.navigate url
 
 IE.Visible = True
 
 Do
 DoEvents
 Loop Until IE.readyState = READYSTATE_COMPLETE

Set Doc = IE.document

lastRow = Sheet1.UsedRange.Rows.Count + 1
 
 Set tagElements = Doc.all.tags("p")
 For Each element In tagElements

  If InStr(element.innerText, "°C") &gt; 0 And InStr(element.className, "temperature") &gt; 0 Then
    Sheet1.Cells(lastRow, 1).Value = element.innerText
    ' Exit the for loop once you get the temperature to avoid unnecessary processing
    Exit For
  End If
 Next

IE.Quit
 Set IE = Nothing
 
 Application.ScreenUpdating = True
 Application.DisplayAlerts = True
 Application.EnableEvents = True
 Application.Calculation = xlCalculationAutomatic

End Sub
