
      Function FindPosition(Ref As Range) As Integer
      Dim Position As Integer
      Position = InStr(1, Ref, "@")
      FindPosition = Position
      End Function

