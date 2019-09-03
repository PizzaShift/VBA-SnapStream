Option Explicit
 
'The following VBA illustrates the use of ParamArray functionality illustrated by a
'functiuon returning the count of arguments passed to the function.
 
Function param_array_function(ParamArray Args() As Variant) As Integer
 
    Dim iCount As Integer
    iCount = UBound(Args)
    param_array_function = IIf(iCount < 0, 0, iCount + 1)
 
End Function
 
 
Sub Test_Proc()
 
Debug.Print "Args :" & param_array_function(1, 2, 3, 4, 5)
Debug.Print "Args :" & param_array_function("amol", "pandey", "excel", "vba")
Debug.Print "Args :" & param_array_function(1, 2, 3)
Debug.Print "Args :" & param_array_function(1, 2)
 
End Sub
