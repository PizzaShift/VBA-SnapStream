Attribute VB_Name = "Sort"
' version 1.0
'
' sSort - sorts 1.strRange using , separated list of headers in 2.strParameters



' strRange wrapper
Sub sSortWrapper(strArguments)
On Error GoTo Error_Handler

    arrArguments = Split(strArguments, Chr(10))
    sSort arrArguments(0), arrArguments(1)
    
Error_Handler:
    ExcelBricks.lngGlobalErrNumber = Err.Number
    ExcelBricks.strGlobalErrDescription = Err.Description
End Sub


' sorts strRange using strParameters
Sub sSort(strRange, strParameters)
On Error GoTo Error_Handler

    Set objRange = ExcelBricks.fGetRange(strRange)
    With objRange.Worksheet.Sort
        .SortFields.Clear
        
        arrParameters = Split(strParameters, ",")
        For Each Parameter In arrParameters
            keyColumn = Application.WorksheetFunction.Match(Parameter, objRange.Rows("1:1"), False)
            .SortFields.Add Key:=objRange.Columns(keyColumn), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        Next
        
        .SetRange objRange
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
        
        .SortFields.Clear
    End With

    ' this may fail if the cell is a merged cell
    On Error Resume Next
    With objRange.Worksheet.Range("IV1")
        .Copy
        .PasteSpecial xlPasteAll
    End With
    On Error GoTo Error_Handler


Error_Handler:
    If Err.Number <> 0 Then
        strErrDescription = Err.Description
        lngErrNumber = Err.Number
    End If
    
    On Error Resume Next
    objRange.Worksheet.Sort.SortFields.Clear
    Set objRange = Nothing
    On Error GoTo 0

    If lngErrNumber <> 0 Then Err.Raise lngErrNumber, , strErrDescription
End Sub
