'To read a binary file you have to know the structure of it. In my case I was trying to read binary file written by LabVIEW-based 'software which records measurement data. The file represent a one-dimensional array of clusters, the cluster has two elements: 'TimeDate stamp and double precision number.
'
'According to the following link http://www.ni.com/tutorial/7900/en/
'
'LabVIEW 7.0 or earlier used a 64-bit double (DBL) to represent time, yielding 15 digits of precision. The number of seconds 'between 1st Jan 1904 (the time stamp Epoch or year zero) to 1st Jan 2000 is 3027456000. Representing this as a DBL would use 10 'out of the 15 digits of precision. That leaves a very small resolution space to perform hardware timings using most of the 'resolution by simply going from 1904 to today.  Representing time as a DBL was not ideal since it did not meet industry 'requirements.
'
'In MS office the date reference is year 1900, while LabVIEW date reference is year 1904. So, in calculations we will compensate 'this date reference difference. Number of days difference is 1462 days.

'Function to convert binary to decimal
Function BinaryToDecimal(ByVal Binary As String) As Double
Dim BinaryNum As Double
Dim BitCount As Integer
For BitCount = 1 To Len(Binary)
BinaryNum = BinaryNum + (CDbl(Mid(Binary, Len(Binary) - BitCount + 1, 1)) * (2 ^ (BitCount - 1)))
Next BitCount
BinaryToDecimal = BinaryNum
End Function

 ' Function to convert 64-bit binary to double-precision float
Function BinaryStringToDouble(ByVal BinaryString As String) As Double

Dim i, Sign, Exponent, BitCounter As Integer
Dim Fraction, DoubleNo As Double

'Read number sign
Sign = (-1) ^ CLng(Mid(BinaryString, 1, 1))     'Most-significant bit

' Read exponent
Exponent = BinaryToDecimal(Mid(BinaryString, 2, 11))

' Read the fraction
Fraction = 0
BitCounter = 0
For i = 13 To Len(BinaryString)
BitCounter = BitCounter + 1
Fraction = Fraction + (2 ^ (-BitCounter)) * CDbl(Mid(BinaryString, i, 1))
Next i
BinaryStringToDouble = Sign * (1 + Fraction) * 2 ^ (Exponent - 1023)

End Function

 ' Function to convert LabView date-time-stamp to string date and time
Function DoubleToDateTime(ByVal LVDateTimeStamp As Double) As String    ' input LabVIEW DateTime stamp (64-bit double precision number)

Dim RefOffset As Double
Dim MSDateTimeStamp As Double
Dim MSDate As Double
Dim DateString As String
Dim DayElapsedTime_sec, Hours, Minutes, Seconds As Double

RefOffset = 126316800     'Reference offset in seconds 1462[days]*24[h/day]*60[Min/h]*60[sec/Min]

'Convert it to Microsoft DateTime stamp: number of seconds from 1-Jan-1900
MSDateTimeStamp = LVDateTimeStamp + RefOffset

MSDate = Application.WorksheetFunction.Floor(MSDateTimeStamp / 86400, 1)    ' number of days from 1900
' 86400: number of seconds per day

DateString = CStr(CDate(MSDate))

DayElapsedTime_sec = MSDateTimeStamp - MSDate * 86400+7200    'Egypt time = UTC time + 2 hours (7200 sec)

Hours = Application.WorksheetFunction.Floor(DayElapsedTime_sec / 3600, 1)

Minutes = Application.WorksheetFunction.Floor((DayElapsedTime_sec - Hours* 3600) / 60, 1)

Seconds = DayElapsedTime_sec - Hours * 3600 - Minutes * 60

DoubleToDateTime = DateString + " " + CStr(Hours) + ":" + CStr(Minutes) + ":" + CStr(Round(Seconds, 0))

End Function
