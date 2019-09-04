Sub ClickElement
'get temps
    'find 3 day
    Dim Element As HTMLLinkElement
    Dim IeDoc As Object

    'find 3 day
    For Each Element In HTMLDoc.Links
        If InStr(Element.innerText, "3 Day History") Then
        Call Element.Click
        Exit For
        End If
    Next Element
    
Do
''Wait till the Browser is loaded
Loop Until oBrowser.readyState = READYSTATE_COMPLETE 
End Sub
