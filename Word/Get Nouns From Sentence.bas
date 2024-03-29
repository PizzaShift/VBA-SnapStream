
       Function is_noun(ByVal wrd As String)
        Dim s As Object, l As Variant
        is_noun = False
        Set s = SynonymInfo(wrd)
        Let l = s.PartOfSpeechList
        If s.MeaningCount <> 0 Then
            For i = LBound(l) To UBound(l)
                If l(i) = wdNoun Then
                    is_noun = True
                End If
            Next
        End If
      End Function
      
      '[excel - Identify and extract noun and modifier - Stack Overflow](https://stackoverflow.com/questions/35150671/identify-and-extract-noun-and-modifier)'
      
      Function getNoun(ByVal sentence As String)
         getNoun = ""
         Dim wrds() As String
         wrds = Split(sentence)
         For i = LBound(wrds) To UBound(wrds)
              If is_noun(wrds(i)) Then
                  getNoun = wrds(i)
                  If i < UBound(wrds) Then
                      getNoun = getNoun & " " & wrds(i + 1)
                  End If
                  Exit Function
              End If
          Next
      End Function

