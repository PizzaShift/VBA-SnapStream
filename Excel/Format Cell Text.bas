Option Explicit 
  
Sub xlCellTextMgmt( _ 
    TargetCell As Range, _ 
    TargetWord As String, _ 
    Optional FirstOnly As Boolean = True, _ 
    Optional FontName As String, _ 
    Optional FontBold As Boolean, _ 
    Optional FontSize As Variant, _ 
    Optional FontColor As Variant) 
     '
     '****************************************************************************************
     '       Title       xlCellTextMgmt
     '       Target Application:  MS Excel
     '       Function:   reformats selected text within the target cell
     '       Limitations:  no explicit checks for acceptable values, e.g., does not
     '                       check to ensure that FontName is a currently supported
     '                       font
     '       Passed Values
     '           TargetCell [input,range]   the target cell containing the text to
     '                                       be reformatted
     '           TargetWord [input,string]  the words in the target cell text that are
     '                                       to be reformatted.  TargetWord can contain
     '                                       anything from a single character to several
     '                                       words or even the entire text of the target
     '                                       cell
     '           FirstOnly  [input,boolean] a TRUE/FALSE flag indicating if the
     '                                       reformatting is to be done on ONLY the 1st
     '                                       instance of the target word (True) or on ALL
     '                                       instances (False)   {Default = True}
     '           FontName   [input,string]  the name of the new font.  Omit if the font
     '                                       is to be left unchanged
     '           FontBold   [input,boolean] a TRUE/FALSE flag indicating if the target
     '                                       words should be BOLD.  True ==> Bold.  Omit
     '                                       if the text is to be left unchanged.
     '           FontSize   [input,variant] the size of the new font.  Omit if the size
     '                                       is to be left unchanged.
     '           FontColor  [input,variant] the color of the new font.  Can be one of
     '                                       the standard colors from the Excel palette or
     '                                       can be one of the standard "vbColors".
     '                                       Omit if the color is to be left unchanged.
     '
     '****************************************************************************************
     '
     '
    Dim Start As Long
      
    Start = 0 
    Do
         '
         '           find the start of TargetWord in TargetCell.Text
         '               if TargetWord not found, exit
         '
        Start = InStr(Start + 1, TargetCell.Text, TargetWord) 
        If Start < 1 Then Exit Sub
         '
         '           test for each font arguement, if present, apply appropriately
         '
        With TargetCell.Characters(Start, Len(TargetWord)).Font 
            If IsNull(FontName) = False Then .Name = FontName 
            If IsNull(FontBold) = False Then .Bold = FontBold 
            If IsNull(FontSize) = False Then .Size = FontSize 
            If IsNull(FontColor) = False Then .ColorIndex = FontColor 
        End With
         '
         '           if request was for ONLY the first instance of TargetWord, exit
         '               otherwise, loop back and see if there are more instances
         '
        If FirstOnly = True Then Exit Sub
    Loop
      
End Sub
