
      Function DefineWord(wordToDefine As String) As String
      	' add a reference to "Microsoft WinHTTP Services". 
        ' Array to hold the response data.
          Dim d() As Byte
          Dim r As Research
      
      
          Dim myDefinition As String
          Dim PARSE_PASS_1 As String
          Dim PARSE_PASS_2 As String
          Dim PARSE_PASS_3 As String
          Dim END_OF_DEFINITION As String
      
          'These "constants" are for stripping out just the definitions from the JSON data
          PARSE_PASS_1 = Chr(34) & "webDefinitions" & Chr(34) & ":"
          PARSE_PASS_2 = Chr(34) & "entries" & Chr(34) & ":"
          PARSE_PASS_3 = "{" & Chr(34) & "type" & Chr(34) & ":" & Chr(34) & "text" & Chr(34) & "," & Chr(34) & "text" & Chr(34) & ":"
          END_OF_DEFINITION = "," & Chr(34) & "language" & Chr(34) & ":" & Chr(34) & "en" & Chr(34) & "}"
          Const SPLIT_DELIMITER = "|"
      
          ' Assemble an HTTP Request.
          Dim url As String
          Dim WinHttpReq As Variant
          Set WinHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
      
          'Get the definition from Google's online dictionary:
          url = "http://www.google.com/dictionary/json?callback=dict_api.callbacks.id100&q=" & wordToDefine & "&sl=en&tl=en&restrict=pr%2Cde&client=te"
          WinHttpReq.Open "GET", url, False
      
          ' Send the HTTP Request.
          WinHttpReq.Send
      
          'Print status to the immediate window
          Debug.Print WinHttpReq.Status & " - " & WinHttpReq.StatusText
      
          'Get the defintion
          myDefinition = StrConv(WinHttpReq.ResponseBody, vbUnicode)
      
          'Get to the meat of the definition
          myDefinition = Mid$(myDefinition, InStr(1, myDefinition, PARSE_PASS_1, vbTextCompare))
          myDefinition = Mid$(myDefinition, InStr(1, myDefinition, PARSE_PASS_2, vbTextCompare))
          myDefinition = Replace(myDefinition, PARSE_PASS_3, SPLIT_DELIMITER)
      
          'Split what's left of the string into an array
          Dim definitionArray As Variant
          definitionArray = Split(myDefinition, SPLIT_DELIMITER)
          Dim temp As String
          Dim newDefinition As String
          Dim iCount As Integer
      
          'Loop through the array, remove unwanted characters and create a single string containing all the definitions
          For iCount = 1 To UBound(definitionArray) 'item 0 will not contain the definition
              temp = definitionArray(iCount)
              temp = Replace(temp, END_OF_DEFINITION, SPLIT_DELIMITER)
              temp = Replace(temp, "\\x22", "")
              temp = Replace(temp, "\\x27", "")
              temp = Replace(temp, Chr$(34), "")
              temp = iCount & ".  " & Trim(temp)
              newDefinition = newDefinition & Mid$(temp, 1, InStr(1, temp, SPLIT_DELIMITER) - 1) & vbLf  'Hmmmm....vbLf doesn't put a carriage return in the cell. Not sure what the deal is there.
          Next iCount
      
          'Put list of definitions in the Immeidate window
          Debug.Print newDefinition
      
          'Return the value
          DefineWord = newDefinition
      
      End Function

