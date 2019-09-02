
      Public Sub TwoDimArray()
      
          ' Declare a two dimensional array
          Dim arrMarks(0 To 3, 0 To 2) As String
      
          ' Fill the array with text made up of i and j values
          Dim i As Long, j As Long
          For i = LBound(arrMarks) To UBound(arrMarks)
              For j = LBound(arrMarks, 2) To UBound(arrMarks, 2)
                  arrMarks(i, j) = CStr(i) & ":" & CStr(j)
              Next j
          Next i
      
          ' Print the values in the array to the Immediate Window
          Debug.Print "i", "j", "Value"
          For i = LBound(arrMarks) To UBound(arrMarks)
              For j = LBound(arrMarks, 2) To UBound(arrMarks, 2)
                  Debug.Print i, j, arrMarks(i, j)
              Next j
          Next i
      
      End Sub

