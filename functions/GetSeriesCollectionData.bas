'http://eng-shady-mohsen.blogspot.com/2014/ 
Sub GetSeriesCollectionData()

' Get the formula of the first data series
SeriesCollectionFormula = Charts("Chart1").SeriesCollection(1).Formula

' Some string processing...
StartPos = InStr(SeriesCollectionFormula, "(")     'Index of first "(" found in the string
EndPos = InStr(SeriesCollectionFormula, "!")    'Index of first "!" found in the string

' Extracting the name of source worksheet using Mid function
SourceWorkSheet = Mid(SeriesCollectionFormula, StartPos + 1, EndPos - StartPos - 1)

'Display the source work sheet in a message
MsgBox ("The name of source work sheet of this chart is : " + SourceWorkSheet)

End Sub
