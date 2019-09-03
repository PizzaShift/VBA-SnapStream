Sub Demo()
Dim M() As Variant
'This will intialize your array with "-"
M = [if(isnumber(ROW(A1:J10)),"-",99)]
'This will intialize your array with integer value 5
M = [5*TRANSPOSE(SIGN(ROW(A1:J10))*SIGN(COLUMN(A1:J10)))]
End Sub
