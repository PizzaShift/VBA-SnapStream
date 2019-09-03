Option Explicit
 
Public Function EncodeBase64(text As String) As String
    Dim arrData() As Byte
    arrData = StrConv(text, vbFromUnicode)
 
    Dim objXML As MSXML2.DOMDocument
    Dim objNode As MSXML2.IXMLDOMElement
 
    Set objXML = New MSXML2.DOMDocument
    Set objNode = objXML.createElement("b64")
 
    objNode.DataType = "bin.base64"
    objNode.nodeTypedValue = arrData
    EncodeBase64 = objNode.text
 
    Set objNode = Nothing
    Set objXML = Nothing
End Function
 
Public Function URLEncode( _
       StringVal As String, _
       Optional SpaceAsPlus As Boolean = False _
     ) As String
 
    Dim StringLen As Long: StringLen = Len(StringVal)
 
    If StringLen > 0 Then
        ReDim result(StringLen) As String
        Dim i As Long, CharCode As Integer
        Dim Char As String, Space As String
 
        If SpaceAsPlus Then Space = "+" Else Space = "%20"
 
        For i = 1 To StringLen
            Char = Mid$(StringVal, i, 1)
            CharCode = Asc(Char)
            Select Case CharCode
            Case 97 To 122, 65 To 90, 48 To 57, 45, 46, 95, 126
                result(i) = Char
            Case 32
                result(i) = Space
            Case 0 To 15
                result(i) = "%0" & Hex(CharCode)
            Case Else
                result(i) = "%" & Hex(CharCode)
            End Select
        Next i
        URLEncode = Join(result, "")
    End If
End Function
 
Function GetColumnValueByName(columnName As String) As Integer
    Dim iCount As Integer
    GetColumnValueByName = 100    'Default if not found
    For iCount = 0 To Sheet2.Range("A1").End(xlToRight).Column
        If Sheet2.Range("A1").Offset(0, iCount).value = columnName Then
            GetColumnValueByName = iCount
            Exit Function
        End If
    Next iCount
 
End Function
 
Function GetTimeSpent(node As IXMLDOMElement) As String
 
    GetTimeSpent = node.Attributes(0).text
 
End Function
 
Function GetOriginalEstimate(node As IXMLDOMElement) As String
 
    GetOriginalEstimate = node.Attributes(0).text
 
End Function
 
Function GetStoryPoint(node As IXMLDOMElement) As String
 
    Dim itm As IXMLDOMElement
 
    For Each itm In node.ChildNodes
        If (itm.ChildNodes(0).text = "Story Points") Then
            GetStoryPoint = itm.ChildNodes(1).text
            Exit Function
        End If
    Next itm
 
End Function
