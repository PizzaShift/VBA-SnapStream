Public Sub test()
    Dim HTMLDoc As HTMLDocument
    Dim oBrowser As InternetExplorer
    Dim oHTML_Element As IHTMLElement
    Dim sURL As String
       
    sURL = "http://forecast.weather.gov/zipcity.php"
    Set oBrowser = New InternetExplorer
    oBrowser.Visible = True
    oBrowser.Silent = True
    oBrowser.navigate sURL
    'oBrowser.FullScreen = True
    
    Do
    ' Wait till the Browser is loaded
    Loop Until oBrowser.readyState = READYSTATE_COMPLETE
    
    Set HTMLDoc = oBrowser.Document
    Do
    ' Wait till the document is loaded
    Loop Until HTMLDoc.readyState = "complete"
    
    HTMLDoc.All.inputstring.Value = "07042"
    HTMLDoc.getElementsByName("Go2")(0).Click
    
    Do
    ' Wait till the document is loaded
    Loop Until HTMLDoc.readyState = "complete"
    
    For Each oHTML_Element In HTMLDoc.getElementsByTagName("a")
        If oHTML_Element.innerText = "3 Day History" Then
            oHTML_Element.Click
            Do
            ' Wait till the document is loaded
            Loop Until HTMLDoc.readyState = "complete"
            Exit For
        End If
    Next

End Sub
