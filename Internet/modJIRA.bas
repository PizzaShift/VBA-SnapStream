Public Sub extractJiraQuery(userName As String, password As String)
 
    Dim MyRequest As New WinHttpRequest
    Dim resultXml As MSXML2.DOMDocument, resultNode As IXMLDOMElement
    Dim nodeContainer As IXMLDOMElement
    Dim rowCount As Integer, colCount As Integer
    Dim fixVersionString As String
    Dim dumpRange As Range, tempValue As Variant
 
    Set dumpRange = Sheet2.Range("A2")
    Sheet2.Range("A2:AZ65536").Clear
     
    Application.ScreenUpdating = False
     
    Application.StatusBar = "JIRA: Preparing header..."
    MyRequest.Open "GET", _
                   "https://JIRA_HOST_NAME/sr/jira.issueviews:searchrequest-xml/temp/SearchRequest.xml?jqlQuery=" & modCommonFunction.URLEncode(ThisWorkbook.Names("jqlQuery").RefersToRange.value) & "&tempMax=1000"
 
    MyRequest.setRequestHeader "Authorization", "Basic " & modCommonFunction.EncodeBase64(userName & ":" & password)
 
    MyRequest.setRequestHeader "Content-Type", "application/xml"
 
    'Send Request.
    Application.StatusBar = "JIRA: Querying request to  JIRA..."
    MyRequest.Send
 
    Set resultXml = New MSXML2.DOMDocument
    resultXml.LoadXML MyRequest.ResponseText
 
    Application.StatusBar = "JIRA: Processing Response..."
    For Each nodeContainer In resultXml.ChildNodes(2).ChildNodes(0).ChildNodes
        fixVersionString = ""
        If nodeContainer.BaseName = "issue" Then
            Application.StatusBar = "JIRA: The total issues found: " & nodeContainer.Attributes(2).text
        End If
        If nodeContainer.BaseName = "item" Then
            For Each resultNode In nodeContainer.ChildNodes
                'Debug.Print resultNode.nodeName & " :: " & resultNode.text
 
                If resultNode.nodeName = "fixVersion" Then
                    fixVersionString = fixVersionString & resultNode.text & " | "
                    GoTo nextNode
                End If
 
                If resultNode.nodeName = "aggregatetimeoriginalestimate" Then
                    tempValue = GetOriginalEstimate(resultNode)
                    dumpRange.Offset(rowCount, 23).value = tempValue
                    If tempValue <> "" Then
                        dumpRange.Offset(rowCount, 25).value = CLng(tempValue) / (60 * 60)
                    End If
                End If
 
                If resultNode.nodeName = "customfields" Then
                    dumpRange.Offset(rowCount, 22).value = GetStoryPoint(resultNode)
                    GoTo nextNode
                End If
 
                If resultNode.nodeName = "timespent" Then
                    tempValue = GetTimeSpent(resultNode)
                    dumpRange.Offset(rowCount, 24).value = tempValue
                    If tempValue <> "" Then
                        dumpRange.Offset(rowCount, 26).value = CLng(tempValue) / (60 * 60)
                    End If
                End If
 
                dumpRange.Offset(rowCount, GetColumnValueByName(resultNode.nodeName)).value = resultNode.text
 
nextNode:
            Next resultNode
 
            dumpRange.Offset(rowCount, 14).value = fixVersionString    ' Fix Version
            rowCount = rowCount + 1
        End If
 
    Next nodeContainer
 
    Application.ScreenUpdating = True
     
    MsgBox "The data extractions is now complete.", vbInformation, "Process Status"
End Sub
