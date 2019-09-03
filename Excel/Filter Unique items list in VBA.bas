Sub uniqueItems()
 
    Dim repArray(10) As Integer
    Dim unqColl As Collection
    Dim unqArray() As Integer
    Dim iCount As Integer
 
    '--Allocate repArray
    repArray(0) = 2
    repArray(1) = 2
    repArray(2) = 2
    repArray(3) = 2
    repArray(4) = 2
    repArray(5) = 2
    repArray(6) = 2
    repArray(7) = 2
    repArray(8) = 2
    repArray(9) = 2
    repArray(10) = 20
 
    '--Filtering via Collection
    Set unqColl = New Collection
 
    On Error Resume Next
    For iCount = 0 To UBound(repArray)
        '--The identifying key has to be string
        unqColl.Add repArray(iCount), CStr(repArray(iCount))
        '--In case of refrential object variables can also use VarPtr for idnetifying unique address
    Next iCount
    Err.Clear
    On Error GoTo 0
 
    '--Resize the unique array based upon the unique items in the collection
    ReDim unqArray(unqColl.Count - 1) As Integer
     
    '--Populate the unique array
    For iCount = 0 To (unqColl.Count - 1)
        unqArray(iCount) = unqColl(iCount + 1)
    Next iCount
 
    '--unqArray Content
    'unqArray(0) = 2
    'unqArray(1) = 20
 
    '--Free up the space
    Set unqColl = Nothing
    Erase repArray
    Erase unqArray
 
End Sub
