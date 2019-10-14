Sub UpdateCode_AndExecute()
  Dim StrCode As String
  Dim  NewUpdate as Boolean
  Dim BypassUpdate as Boolean
'The name of this sub (UpdateCode_AndExecute) is required to automaticly 
'execute the update when the code below is inserted into the project.
'UpdateDate = Workbooks("Toolkit.xlam").Sheets(" ").Range("Z1").Value '10/13/2019
'NewUpdate = Workbooks("Toolkit.xlam").Sheets(" ").Range("Z2").Value 'TRUE
'BypassUpdate = Workbooks("Toolkit.xlam").Sheets(" ").Range("Z3").Value 'FALSE
  'KeepChanges = Workbooks("Toolkit.xlam").Sheets(" ").Range("Z4").Value
  NewUpdate = FALSE
  BypassUpdate = FALSE

'If the update date is older than current day, and a new update is true, or bypass is true then run the update.
  'If UpdateDate = Workbooks("Toolkit.xlam").Sheets(" ").Range("Z1").Value < Date Then
    If NewUpdate = True Or BypassUpdate = True Then
    Workbooks("Toolkit.xlam").Sheets(" ").Range("Z4").Value = TRUE
    'List all Modules & Procedures currently in the project for reference if needed:
    'Call ListModules
    'Call ListProcedures
    'Get the new code to insert into the addin project within a new Module:
    
    
    '....Check complete, new updates have been made.
  else
    Workbooks("Toolkit.xlam").Sheets(" ").Range("Z4").Value = FALSE
    'Check complete, no new updates.
    End If 
  'End if
End Sub

