Option Explicit

Private wdApp As Object
Private wdDoc As Object
Private bStarted As Boolean
Private strSavePath As String
Private strName As String
Const strPath As String = "D:\Path\Email Messages\" 'The root folder


Sub SaveSelectedMessagesAsPDF()
'Graham Mayor - https://www.gmayor.com - Last updated - 13 Jul 2019
'Select the messages to process and run this macro
Dim olMsg As Object
Dim strFname As String, strExt As String
Dim j As Long
    'Create the folder to store the messages if not present
    If CreateFolders(strPath) = False Then GoTo lbl_Exit
    'Open or Create a Word object
    On Error Resume Next
    Set wdApp = GetObject(, "Word.Application")
    If Err Then
        Set wdApp = CreateObject("Word.Application")
        bStarted = True
    End If
    On Error GoTo 0
    For Each olMsg In Application.ActiveExplorer.Selection
        SaveAsPDFfile olMsg
    Next olMsg
    MsgBox "Completed"
lbl_Exit:
    If bStarted Then wdApp.Quit
    Set olMsg = Nothing
    Set wdApp = Nothing
    Exit Sub
End Sub


Private Sub SaveAsPDFfile(olItem As MailItem)
'Graham Mayor - https://www.gmayor.com - Last updated - 13 Jul 2019
Dim olNS As NameSpace
Dim fso As Object, TmpFolder As Object
Dim tmpPath As String
Dim strFileName As String
Dim oRegex As Object


    Set olNS = Application.GetNamespace("MAPI")


    'Get the user's TempFolder to store the temporary file
    Set fso = CreateObject("Scripting.FileSystemObject")
    tmpPath = fso.GetSpecialFolder(2)


    'construct the filename for the temp mht-file
    strName = "email_temp.mht"
    tmpPath = tmpPath & "\" & strName


    'Save temporary file
    olItem.SaveAs tmpPath, 10


    'Open the temporary file in Word
    Set wdDoc = wdApp.Documents.Open(fileName:=tmpPath, _
                                     AddToRecentFiles:=False, _
                                     Visible:=False, _
                                     Format:=7)


    'Create a file name from the message subject
    strFileName = Format(olItem.ReceivedTime, "yyyymmdd hh.mm") & "-" & olItem.SenderName & "-" & olItem.Subject
    'Remove illegal filename characters
    Set oRegex = CreateObject("vbscript.regexp")
    oRegex.Global = True
    oRegex.Pattern = "[\/:*?""<>|]"
    strFileName = Trim(oRegex.Replace(strFileName, ""))
    strSavePath = strPath & strFileName
    CreateFolders strSavePath
    strFileName = strFileName & ".pdf"
    strFileName = FileNameUnique(strSavePath, strFileName, "pdf")
    strFileName = strSavePath & strFileName


    'Save the attachments
    SaveAttachments olItem, strSavePath


    'Save As pdf
    wdDoc.ExportAsFixedFormat OutputFilename:= _
                              strFileName, _
                              ExportFormat:=17, _
                              OpenAfterExport:=False, _
                              OptimizeFor:=0, _
                              Range:=0, _
                              From:=0, _
                              To:=0, _
                              Item:=0, _
                              IncludeDocProps:=True, _
                              KeepIRM:=True, _
                              CreateBookmarks:=0, _
                              DocStructureTags:=True, _
                              BitmapMissingFonts:=True, _
                              UseISO19005_1:=False


    ' close the document
    wdDoc.Close 0
lbl_Exit:
    'Cleanup
    Set olNS = Nothing
    Set olItem = Nothing
    Set wdDoc = Nothing
    Set oRegex = Nothing
    Exit Sub
End Sub


Private Sub SaveAttachments(olItem As MailItem, strSaveFldr As String)
'Graham Mayor - http://www.gmayor.com - Last updated - 26 May 2017
Dim olAttach As Attachment
Dim strFname As String
Dim strExt As String
Dim j As Long


    On Error Resume Next
    If olItem.Attachments.Count > 0 Then
        For j = 1 To olItem.Attachments.Count
            Set olAttach = olItem.Attachments(j)
            strFname = olAttach.fileName
            strExt = Right(strFname, Len(strFname) - InStrRev(strFname, Chr(46)))
            strFname = FileNameUnique(strSaveFldr, strFname, strExt)
            olAttach.SaveAsFile strSaveFldr & strFname
        Next j
    End If
lbl_Exit:
    Set olAttach = Nothing
    Set olItem = Nothing
    Exit Sub
End Sub




Private Function CreateFolders(strPath As String) As Boolean
'Graham Mayor - https://www.gmayor.com - Last updated - 13 Jul 2019
Dim strTempPath As String
Dim lngPath As Long
Dim VPath As Variant
    VPath = Split(strPath, "\")
    strPath = VPath(0) & "\"
    For lngPath = 1 To UBound(VPath)
        strPath = strPath & VPath(lngPath) & "\"
        On Error GoTo Err_Handler
        If Not FolderExists(strPath) Then MkDir strPath
    Next lngPath
    CreateFolders = True
lbl_Exit:
    Exit Function
Err_Handler:
    MsgBox "The path " & strPath & " is invalid!"
    CreateFolders = False
    Resume lbl_Exit
End Function


Private Function FileNameUnique(strPath As String, _
                                strFileName As String, _
                                strExtension As String) As String
'Graham Mayor - https://www.gmayor.com - Last updated - 13 Jul 2019
Dim lngF As Long
Dim lngName As Long
    lngF = 1
    lngName = Len(strFileName) - (Len(strExtension) + 1)
    strFileName = Left(strFileName, lngName)
    Do While FileExists(strPath & strFileName & Chr(46) & strExtension) = True
        strFileName = Left(strFileName, lngName) & "(" & lngF & ")"
        lngF = lngF + 1
    Loop
    FileNameUnique = strFileName & Chr(46) & strExtension
lbl_Exit:
    Exit Function
End Function


Private Function FolderExists(fldr) As Boolean
'Graham Mayor - https://www.gmayor.com - Last updated - 13 Jul 2019
Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If (fso.FolderExists(fldr)) Then
        FolderExists = True
    Else
        FolderExists = False
    End If
lbl_Exit:
    Set fso = Nothing
    Exit Function
End Function


Private Function FileExists(filespec) As Boolean
'Graham Mayor - https://www.gmayor.com - Last updated - 13 Jul 2019
Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(filespec) Then
        FileExists = True
    Else
        FileExists = False
    End If
lbl_Exit:
    Set fso = Nothing
    Exit Function
End Function
