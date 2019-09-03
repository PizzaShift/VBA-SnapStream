'Outlook  Automatically save and print attachments

'Description:
'This code enables events in Outlook and "watches" a particular folder. When a mail item arrives in the folder, if it has an 
'attachment it is saved to a specific directory. If the Saved attachment is an Excel file, it is printed out. 

'Discussion:
'This is useful for anyone who recieves regular reports via email. An Outlook rule can be set to move the mail to a particular 
'folder, then the attachment can be saved and/or printed, all with no user interaction. The example prints Excel attachments, 
'but could easily be adapted to choose Excel/Word/Powerpoint based on the file type.  
 
' How to use:
	
    'From Outlook, open the VBEditor (Alt+F11)
    'Add a reference to the "Microsoft Excel <your version number> Object Library fron Tools>References
    'Paste the code into the ThisOutlookSession module
    'Create an Outlook folder named "Temp" in your Personal folders (or amend the code: Set TargetFolderItems to eqaul an 
    'existing folder)
    'Create a directory "C:\Temp" (or amend the constant: FILE_PATH to eqaul an existing folder)
    'Save the project
    'Restart Outlook (or run the routine "Application_Startup")

'Test the code:

    'Move a mail item with some attachments into you target folder.
    'The attachments will be saved in your specified directory
    'Any Excel files will be printed
    
 '###############################################################################
 '### Module level Declarations
 'expose the items in the target folder to events
Option Explicit 
Dim WithEvents TargetFolderItems As Items 
 'set the string constant for the path to save attachments
Const FILE_PATH As String = "C:\Temp\" 
 
 '###############################################################################
 '### this is the Application_Startup event code in the ThisOutlookSession module
Private Sub Application_Startup() 
     'some startup code to set our "event-sensitive" items collection
    Dim ns As Outlook.NameSpace 
     '
    Set ns = Application.GetNamespace("MAPI") 
    Set TargetFolderItems = ns.Folders.Item( _ 
    "Personal Folders").Folders.Item("Temp").Items 
     
End Sub 
 
 '###############################################################################
 '### this is the ItemAdd event code
Sub TargetFolderItems_ItemAdd(ByVal Item As Object) 
     'when a new item is added to our "watched folder" we can process it
    Dim olAtt As Attachment 
    Dim i As Integer 
     
    If Item.Attachments.Count > 0 Then 
        For i = 1 To Item.Attachments.Count 
            Set olAtt = Item.Attachments(i) 
             'save the attachment
            olAtt.SaveAsFile FILE_PATH & olAtt.FileName 
             
             'if its an Excel file, pass the filepath to the print routine
            If UCase(Right(olAtt.FileName, 3)) = "XLS" Then 
                PrintAtt (FILE_PATH & olAtt.FileName) 
            End If 
        Next 
    End If 
     
    Set olAtt = Nothing 
     
End Sub 
 
 '###############################################################################
 '### this is the Application_Quit event code in the ThisOutlookSession module
Private Sub Application_Quit() 
     
    Dim ns As Outlook.NameSpace 
    Set TargetFolderItems = Nothing 
    Set ns = Nothing 
     
End Sub 
 
 '###############################################################################
 '### print routine
Sub PrintAtt(fFullPath As String) 
     
    Dim xlApp As Excel.Application 
    Dim wb As Excel.Workbook 
     
     'in the background, create an instance of xl then open, print, quit
    Set xlApp = New Excel.Application 
    Set wb = xlApp.Workbooks.Open(fFullPath) 
    wb.PrintOut 
    xlApp.Quit 
     
     'tidy up
    Set wb = Nothing 
    Set xlApp = Nothing 
     
End Sub 
