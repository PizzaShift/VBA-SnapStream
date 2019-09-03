Public Sub FileLister(rootPath As String, exploreSubFolder As Boolean)
 
    Dim file As Scripting.file, folder As Scripting.folder, subfolder As Scripting.folder
 
    If fso Is Nothing Then
        Set fso = New FileSystemObject
    End If
 
    Set folder = fso.GetFolder(rootPath)
 
    If (folder.SubFolders.Count > 0) And exploreSubFolder Then
     
        For Each subfolder In folder.SubFolders
             
            nestLevel = nestLevel + 1
            ReDim Preserve navigated(nestLevel)
             
            If navigated(nestLevel) = 0 Then
                For Each file In folder.Files
                    'Record files in folder
                    anchorRange.Offset(printRow, nestLevel + 1).Value = file.Name
                    anchorRange.Offset(printRow, 0).Hyperlinks.Add Control.Range("A12").Offset(printRow, 0), file.path, , , "Link"
                    printRow = printRow + 1
                Next file
            End If
             
            anchorRange.Offset(printRow, nestLevel + 1).Value = subfolder.Name
            anchorRange.Offset(printRow, nestLevel + 1).Interior.Color = 10092543  'Light Yellow
            anchorRange.Offset(printRow, 0).Hyperlinks.Add Control.Range("A12").Offset(printRow, 0), subfolder.path, , , "Link"
            printRow = printRow + 1
 
            navigated(nestLevel) = 1
            FileLister subfolder.path, exploreSubFolder
            nestLevel = nestLevel - 1
             
        Next subfolder
         
    Else
     
        nestLevel = nestLevel + 1
        For Each file In folder.Files
            'Record files in folder
            anchorRange.Offset(printRow, nestLevel + 1).Value = file.Name
            anchorRange.Offset(printRow, 0).Hyperlinks.Add Control.Range("A12").Offset(printRow, 0), file.path, , , "Link"
            printRow = printRow + 1
        Next file
        For Each subfolder In folder.SubFolders
            anchorRange.Offset(printRow, nestLevel + 1).Value = subfolder.Name
            anchorRange.Offset(printRow, nestLevel + 1).Interior.Color = 10092543  'Light Yellow
            anchorRange.Offset(printRow, 0).Hyperlinks.Add Control.Range("A12").Offset(printRow, 0), subfolder.path, , , "Link"
            printRow = printRow + 1
        Next subfolder
        nestLevel = nestLevel - 1
         
        If exploreSubFolder Then
            navigated(nestLevel) = 0
        End If
         
    End If
End Sub
