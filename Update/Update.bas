Sub UpdateCode_AndExecute()
  Dim StrCode As String
  Dim  NewUpdate as Boolean
  Dim BypassUpdate as Boolean
'The name of this sub (UpdateCode_AndExecute) is required to automaticly 
'execute the update when the code below is inserted into the project.
'UpdateDate = Workbooks("Toolkit.xlam").Sheets(" ").Range("Z1").Value '10/13/2019
'NewUpdate = Workbooks("Toolkit.xlam").Sheets(" ").Range("Z2").Value 'TRUE
'BypassUpdate = Workbooks("Toolkit.xlam").Sheets(" ").Range("Z3").Value 'FALSE
  NewUpdate = TRUE
  NewUpdate = FALSE

'If the update date is older than current day, and a new update is true, or bypass is true then run the update.
  'If UpdateDate = Workbooks("Toolkit.xlam").Sheets(" ").Range("Z1").Value < Date Then
    If NewUpdate = True Or BypassUpdate = True Then
    'List all Modules & Procedures currently in the project for reference if needed:
    'Call ListModules
    'Call ListProcedures
    'Get the new code to insert into the addin project within a new Module:
      Msgbox("New code has now been inserted and exceuted.")
    End If 
  'End if
End Sub

