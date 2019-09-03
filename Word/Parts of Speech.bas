
      Sub parts_of_speech()
      
      Set mySynInfo = Selection.Range.SynonymInfo
      If mySynInfo.MeaningCount <> 0 Then
          myList = mySynInfo.MeaningList
          myPos = mySynInfo.PartOfSpeechList
          For i = 1 To UBound(myPos)
      'wdAdjective, wdAdverb, wdConjunction, wdIdiom, wdInterjection, wdNoun, wdOther, wdPreposition, wdPronoun, and wdVerb.
              Select Case myPos(i)
                  Case wdAdjective
                       pos = "adjective"
                  Case wdNoun
                       pos = "noun"
                  Case wdAdverb
                       pos = "adverb"
                  Case wdVerb
                       pos = "verb"
                  Case wdConjunction
                       pos = "Conjunction"
                  Case wdIdiom
                      pos = "Idiom"
                  Case wdInterjection
                      pos = "Interjection"
                  Case wdPreposition
                      pos = "Preposition"
                  Case wdPronoun
                      pos = "Pronoun"
      
                  Case Else
                       pos = "other"
              End Select
              MsgBox myList(i) & " found as " & pos
          Next i
      Else
          MsgBox "There were no meanings found."
      End If
      
      End Sub
      
      '------------------------------------------------------------------------'
      Option Explicit
      Private mObjWord As Object
       
      Sub Antonyms()
      Dim i As Long
      Dim c As Range
      Dim sWord As String
      Dim arr
      For Each c In Selection
          sWord = c
          
          If GetMeanings(sWord, arr) Then
          For i = LBound(arr) To UBound(arr)
              c.Offset(0, i).Value = arr(i)
          Next
          
          End If
      Next c
      Set mObjWord = Nothing 'clears the word object when done
      End Sub
       
      Function GetMeanings(myWord As String, vMeanings)
      Dim objSynonymInfo As Object
      If mObjWord Is Nothing Then
          Set mObjWord = CreateObject("word.application")
      End If
      Set objSynonymInfo = mObjWord.SynonymInfo(myWord)
      vMeanings = objSynonymInfo.AntonymList
      GetMeanings = UBound(vMeanings) > 0
      End Function
      '------------------------------------------------------------------------'
      Sub test()
      Dim i As Long
      Dim c As Range
      Dim sWord As String
      Dim arr
      For Each c In Selection
          sWord = c
       
          If GetMeanings(sWord, arr) Then
          For i = LBound(arr) To UBound(arr)
              c.Offset(0, i).Value = arr(i)
          Next
       
          End If
      Next c
      Set mObjWord = Nothing 'clears the word object when done
      End Sub
       
      Function GetMeanings(myWord As String, vMeanings)
      Dim objSynonymInfo As Object
      If mObjWord Is Nothing Then
          Set mObjWord = CreateObject("word.application")
      End If
      Set objSynonymInfo = mObjWord.SynonymInfo(myWord)
      vMeanings = objSynonymInfo.antonymlist
      GetMeanings = UBound(vMeanings) > 0
      End Function

    'Another Solution
      Option Explicit
       
      Public Sub PartsOfSpeech()
       
        Dim mObjWord As Word.Application
        Dim mySynInfo As Word.SynonymInfo
        Dim myList As Variant
        Dim myPos As Variant
        Dim i As Integer
        Dim iMax As Integer
        Dim thisPos As String
        Dim oCell As Range
       
        Set mObjWord = CreateObject("Word.Application")
        
        iMax = 1
       
        For Each oCell In Selection
          oCell.Offset(0, 1).Resize(1, 99).ClearContents
          If oCell.Column = 1 And Not IsEmpty(oCell) Then
            Set mySynInfo = SynonymInfo(Word:=oCell.Value, LanguageID:=wdEnglishUS)
            oCell.Offset(0, 1) = "'(" & CStr(mySynInfo.MeaningCount) & ")"
            If mySynInfo.MeaningCount <> 0 Then
              myList = mySynInfo.MeaningList
              myPos = mySynInfo.PartOfSpeechList
              If i > iMax Then iMax = i
              For i = 1 To UBound(myPos)
                Select Case myPos(i)
                  Case wdAdjective
                    thisPos = "adjective"
                  Case wdNoun
                    thisPos = "noun"
                  Case wdAdverb
                    thisPos = "adverb"
                  Case wdVerb
                    thisPos = "verb"
                  Case wdConjunction
                    thisPos = "conjunction"
                  Case wdIdiom
                    thisPos = "idiom"
                  Case wdInterjection
                    thisPos = "interjection"
                  Case wdPreposition
                    thisPos = "preposition"
                  Case wdPronoun
                    thisPos = "pronoun"
                   Case Else
                    thisPos = "other"
                End Select
                oCell.Offset(0, i + 1) = myList(i) & " (" & thisPos & ")"
              Next i
            Else
              oCell.Offset(0, 2) = "No meanings found"
            End If
          End If
        Next oCell
        
        For i = 3 To iMax
          Columns(i).EntireColumn.AutoFit
        Next i
       
      End Sub

    'Solution Two"
      Public Sub PartsOfSpeech()
        Application.ScreenUpdating = False
        Dim mObjWord As Word.Application
        Dim mySynInfo As Word.SynonymInfo
        Dim myList As Variant, myPos As Variant
        Dim i As Long, j As Long
        Dim thisPos As String, oCell As Range
        Set mObjWord = CreateObject("Word.Application")
        For Each oCell In Selection
          oCell.Offset(0, 1).Resize(1, 99).ClearContents
          If oCell.Column = 1 And Not IsEmpty(oCell) Then
            Set mySynInfo = SynonymInfo(Word:=oCell.Value, LanguageID:=wdEnglishUS)
            If mySynInfo.MeaningCount <> 0 Then
              myList = mySynInfo.MeaningList
              myPos = mySynInfo.PartOfSpeechList
              For i = 1 To UBound(myPos)
                Select Case myPos(i)
                  Case wdAdjective
                    thisPos = "adjective": j = 2
                  Case wdNoun
                    thisPos = "noun": j = 3
                  Case wdAdverb
                    thisPos = "adverb": j = 4
                  Case wdVerb
                    thisPos = "verb": j = 4
                  Case wdConjunction
                    thisPos = "conjunction": j = 5
                  Case wdIdiom
                    thisPos = "idiom": j = 6
                  Case wdInterjection
                    thisPos = "interjection": j = 7
                  Case wdPreposition
                    thisPos = "preposition": j = 8
                  Case wdPronoun
                    thisPos = "pronoun": j = 9
                   Case Else
                    thisPos = "other": j = 10
                End Select
                oCell.Offset(0, j) = thisPos
              Next i
            Else
              oCell.Offset(0, 1) = "No meanings found"
            End If
          End If
        Next oCell
        Columns.EntireColumn.AutoFit
        mObjWord.Quit: Set mObjWord = Nothing
        Application.ScreenUpdating = True
      End Sub

