Sub Weather()
Dim HTMLDoc As HTMLDocument
Dim oBrowser As InternetExplorer
Dim oHTML_Element As IHTMLElement
Dim sURL As String
Dim occ As ContentControl
Dim strToday
Dim strPath
Dim iRet As Integer
Dim strPrompt As String
Dim strTitle As String

On Error GoTo Err_Clear


sURL = "http://forecast.weather.gov/zipcity.php"

Set oBrowser = New InternetExplorer
oBrowser.Visible = True
oBrowser.Silent = True
oBrowser.navigate sURL
oBrowser.FullScreen = True


Do
' Wait till the Browser is loaded
Loop Until oBrowser.readyState = READYSTATE_COMPLETE

Set HTMLDoc = oBrowser.Document

'time to load
WaitSeconds (2)

HTMLDoc.all.inputstring.Value = Me.txtZip

'time to load
WaitSeconds (2)

For Each oHTML_Element In HTMLDoc.getElementsByName("Go2")
If oHTML_Element.Type = "submit" Then oHTML_Element.Click: Exit For
Next oHTML_Element

'Do
' 'Wait till the Browser is loaded
'Loop Until oBrowser.readyState = READYSTATE_COMPLETE

'Request seconds to wait for user input case by case
Dim strUser As String
'strUser = InputBox("Default page load time is 5 seconds. Change?", "Change Default Time?", 5)
strUser = "7"

    'check directory
    strPath = ActiveDocument.Path & "\Weather"
    If Len(Dir(strPath, vbDirectory)) = 0 Then
    MkDir strPath
    End If
    
    'image path
    Set occ = ActiveDocument.SelectContentControlsByTitle("ctrlCalendar").Item(1)
    strToday = strPath & "\Weather " & Format(Date, "mm-dd-yy") & ".jpg"
    strPath = strPath & "\Weather " & Format(occ.Range.Text, "mm-dd-yy") & ".jpg"
    
    'time to load
    WaitSeconds (5)
    
    'hide browser
    oBrowser.Visible = False
    
    '***SCREENSHOT TODAY'S FORECAST
    If Len(Dir(strToday, vbDirectory)) = 0 Then
        MsgBox "    Please highlight today's forecast area"
        
        Do
        'Wait till the Browser is loaded
        Loop Until oBrowser.readyState = READYSTATE_COMPLETE
        
        'show browser
        oBrowser.Visible = True
        
            'print screen
            keybd_event VK_SNAPSHOT, 1, 0, 0 'Print Screen key down
            keybd_event VK_SNAPSHOT, 1, VK_KEYUP, 0 'Print key Up - Screenshot to Clipboard
        
        'time to capture area
        WaitSeconds (strUser)
        
        'save clipboard to file
        SaveClipboardasJPEG (strToday)
        
        'hide browser
        oBrowser.Visible = False
        
        Else
        
    End If

    
    '***CHECK FOR FILE TO IMPORT
        'If Format(Date, "mm-dd-yy") <> Format(occ.Range.Text, "mm-dd-yy") Then
            If Len(Dir(strPath, vbDirectory)) = 0 Then
                MsgBox "       There is not a screenshot available for this date" & vbCr & _
                "         Please import weather for this date manually" & vbCr & _
                "                     Now closing weather import"
                Exit Sub
                Else
                
                '***IMPORT EXISTING SCREENSHOT
                MsgBox "     A screenshot is now being imported"
                   
                'insert image to document
                Set occ = ActiveDocument.SelectContentControlsByTitle("ctrlPicture").Item(1)
                
                If occ.Type = wdContentControlPicture Then
                    If occ.Range.InlineShapes.Count > 0 Then occ.Range.InlineShapes(1).Delete
                    Dim pPicFileName As String
                    pPicFileName = strPath
                        ActiveDocument.InlineShapes.AddPicture Filename:=pPicFileName, _
                        linktofile:=True, Range:=occ.Range
                        With occ.Range.InlineShapes(1)
                        .LockAspectRatio = msoCTrue
                        .width = 540
                        End With
                    End If
                End If
         'End If
                
    'IMPORT TEMP DATA
    strPrompt = "      Would you like to import temperature data?"
    strTitle = "Import Temperature Data"
    iRet = MsgBox(strPrompt, vbYesNo, strTitle)
    If iRet = vbYes Then
    
        'clear clipboard
        ClearClipboard
        
        Do
        'Wait till the Browser is loaded
        Loop Until oBrowser.readyState = READYSTATE_COMPLETE
        
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
        
        'wait for data to load
        WaitSeconds (3)
        
        MsgBox "      Please highlight temperture nearest to 6am"
        
        Do
        'Wait till the Browser is loaded
        Loop Until oBrowser.readyState = READYSTATE_COMPLETE
        
        'show browser
        oBrowser.Visible = True
           
        WaitSeconds (strUser)
        
        'AM data
        SendKeys "^c", True
            
        'time for data to load
        WaitSeconds (5)
        
        oBrowser.Visible = False
        
        Set occ = ActiveDocument.SelectContentControlsByTitle("ctrlAM").Item(1)
        occ.Range.PasteAndFormat wdFormatPlainText
        
        'get PM
        MsgBox "      Please highlight temperture nearest to 2pm"
        
        Do
        'Wait till the Browser is loaded
        Loop Until oBrowser.readyState = READYSTATE_COMPLETE
        
        oBrowser.Visible = True
           
        WaitSeconds (strUser)
        
        'PM data
        SendKeys "^c", True
            
        'time for data to load
        WaitSeconds (2)
        
        oBrowser.Visible = False
        
        Set occ = ActiveDocument.SelectContentControlsByTitle("ctrlPM").Item(1)
        occ.Range.PasteAndFormat wdFormatPlainText
        
    End If
    
    'release occ
    Set occ = Nothing
    
    'close browser
    oBrowser.Quit
    
    
    MsgBox "        Weather import complete"

Err_Clear:
If Err <> 0 Then
'MsgBox "Error: " & Err.Number & ". " & Err.Description, , "modDateTime.WaitSeconds"
Err.Clear
Resume Next
End If

'close browser
oBrowser.Quit

End Sub 
