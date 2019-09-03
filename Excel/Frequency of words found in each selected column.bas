
      Sub FrequencyV2() 'Modified from: https://stackoverflow.com/questions/21858874/counting-the-frequencies-of-words-in-excel-strings
      'It determines the frequency of words found in each selected column.
      'Puts results in new worksheets.
      'Before running, select one or more columns but not the header rows.
          Dim rng As Range
          Dim row As Range
          Dim col As Range
          Dim cell As Range
          Dim ws As Worksheet
          Dim wsNumber As Long 'Used to put a number in the names of the newly created worksheets
          wsNumber = 1
          Set rng = Selection
          For Each col In rng.Columns
              Dim BigString As String, I As Long, J As Long, K As Long
              BigString = ""
              For Each cell In col.Cells
                  BigString = BigString & " " & cell.Value
              Next cell
              BigString = Trim(BigString)
              ary = Split(BigString, " ")
              Dim cl As Collection
              Set cl = New Collection
              For Each a In ary
                  On Error Resume Next 'This works because an error occurs if item already exists in the collection.
                  'Note that it's not case sensitive.  Differently capitalized items will be identified as already belonging to collection.
                  cl.Add a, CStr(a)
              Next a
              Set ws = Sheets.Add(After:=Sheets(Sheets.Count))
              ws.Name = "F" & CStr(wsNumber)
              wsNumber = wsNumber + 1
              Worksheets(ws.Name).Cells(1, "A").Value = col.Cells(1, 1).Offset(-1, 0).Value 'Copies the table header text for current column to new worksheet.
              For I = 1 To cl.Count
                  v = cl(I)
                  Worksheets(ws.Name).Cells(I + 1, "A").Value = v 'The +1 needed because header text takes up row 1.
                  J = 0
                  For Each a In ary
                      If LCase(a) = LCase(v) Then J = J + 1
                  Next a
                  Worksheets(ws.Name).Cells(I + 1, "B") = J 'The +1 needed because header text takes up row 1.
              Next I
          Next col
      End Sub

