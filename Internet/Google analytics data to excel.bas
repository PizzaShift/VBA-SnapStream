Private Function getGAauthenticationToken(ByVal email As String, ByVal password As String)
    Dim authResponse As String
    Dim authTokenStart As Integer
    Dim URL As String
    Dim authtoken As String
 
    If email = "" Then
        getGAauthenticationToken = ""
        Exit Function
    End If
 
    If password = "" Then
        getGAauthenticationToken = "Input password"
        Exit Function
    End If
    password = modCommonFunctions.uriEncode(password)
 
    On Error GoTo errhandler
    Dim objhttp As Object
    Set objhttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    URL = "https://www.google.com/accounts/ClientLogin"
    objhttp.Open "POST", URL, False
    objhttp.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
    objhttp.setTimeouts 1000000, 1000000, 1000000, 1000000
    objhttp.send ("accountType=GOOGLE&Email=" & email & "&Passwd=" & password & "&service=analytics&Source=Excel based analytics. (c) Amol Pandey 2012")
    authResponse = objhttp.responseText
    If InStr(1, authResponse, "BadAuthentication") = 0 Then
        authTokenStart = InStr(1, authResponse, "Auth=") + 4
        authtoken = Right(authResponse, Len(authResponse) - authTokenStart)
        getGAauthenticationToken = authtoken
    Else
        getGAauthenticationToken = "Authentication failed"
    End If
    Exit Function
errhandler:
    getGAauthenticationToken = "Authentication failed"
End Function

Private Sub getDetailedGoogleAnalyticsData(authtoken As String, profileNumber As Long, startDate As Date, endDate As Date, metrics() As String, Optional dimensions As Variant)
 
    Dim httpUrl As String
    Dim request As MSXML2.ServerXMLHTTP60
    Dim requstTimeOut As Long
    Dim requestReponseText As String
    Dim startDateString As String, endDateString As String
    Dim item As Variant, subItem As Variant, responseXml As MSXML2.DOMDocument
    Dim rowCount As Integer
    'clear Range
    RawData.Range("F5:F6").ClearContents
    RawData.Range("H2").Resize(MAX_ROWS, 4).ClearContents
    rowCount = 0
 
    startDateString = Year(startDate) & "-" & Right("0" & Month(startDate), 2) & "-" & Right("0" & Day(startDate), 2)
    endDateString = Year(endDate) & "-" & Right("0" & Month(endDate), 2) & "-" & Right("0" & Day(endDate), 2)
 
    httpUrl = "https://www.google.com/analytics/feeds/data?ids=ga:" & profileNumber & "&start-date=" & startDateString & "&end-date=" & endDateString & "&max-results=10000&metrics="
 
    For Each item In metrics
        httpUrl = httpUrl & "ga:" & item & ","
    Next item
    httpUrl = Left(httpUrl, Len(httpUrl) - 1)
 
 
    'Aggreagted values
    Set request = New MSXML2.ServerXMLHTTP60
    requstTimeOut = 1000000
    request.Open "GET", httpUrl, False
    request.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
    request.setRequestHeader "Authorization", "GoogleLogin Auth=" & authtoken
    request.setRequestHeader "GData-Version", "2"
    request.setTimeouts requstTimeOut, requstTimeOut, requstTimeOut, requstTimeOut
    request.send ("")
    Set responseXml = request.responseXml
    For Each item In responseXml.ChildNodes(1).ChildNodes
        If item.nodeName = "dxp:aggregates" Then
            For Each subItem In item.ChildNodes
                If subItem.Attributes.Length > 0 Then
                    RawData.Range("F5").Offset(rowCount, 0).Value = subItem.Attributes(3).Value
                    rowCount = rowCount + 1
                End If
            Next subItem
        End If
    Next item
    If Not (IsError(dimensions)) Then
        httpUrl = httpUrl + "&dimensions="
        For Each item In dimensions
            httpUrl = httpUrl & "ga:" & item & ","
        Next item
        httpUrl = Left(httpUrl, Len(httpUrl) - 1)
        Set request = New MSXML2.ServerXMLHTTP60
        requstTimeOut = 1000000
        request.Open "GET", httpUrl, False
        request.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
        request.setRequestHeader "Authorization", "GoogleLogin Auth=" & authtoken
        request.setRequestHeader "GData-Version", "2"
        request.setTimeouts requstTimeOut, requstTimeOut, requstTimeOut, requstTimeOut
        request.send ("")
        Set responseXml = request.responseXml
        rowCount = 0
        For Each item In responseXml.ChildNodes(1).ChildNodes
            If item.nodeName = "entry" Then
                RawData.Range("H2").Offset(rowCount, 0).Value = item.ChildNodes(4).Attributes(1).Value    'Country
                RawData.Range("H2").Offset(rowCount, 1).Value = item.ChildNodes(5).Attributes(1).Value    'region
                RawData.Range("H2").Offset(rowCount, 2).Value = item.ChildNodes(6).Attributes(3).Value    'visits
                RawData.Range("H2").Offset(rowCount, 3).Value = item.ChildNodes(7).Attributes(3).Value    'pageviews
                rowCount = rowCount + 1
            End If
        Next item
    End If
End Sub
