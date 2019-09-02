
      Sub ExtractComments()
      Dim ExComment As Comment
      Dim i As Integer
      Dim ws As Worksheet
      Dim CS As Worksheet
      Set CS = ActiveSheet
      If ActiveSheet.Comments.Count = 0 Then Exit Sub
      
      For Each ws In Worksheets
        If ws.Name = "Comments" Then i = 1
      Next ws
          
      If i = 0 Then
        Set ws = Worksheets.Add(After:=ActiveSheet)
        ws.Name = "Comments"
      Else: Set ws = Worksheets("Comments")
      End If
      
      For Each ExComment In CS.Comments
        ws.Range("A1").Value = "Comment In"
        ws.Range("B1").Value = "Comment By"
        ws.Range("C1").Value = "Comment"
        With ws.Range("A1:C1")
          .Font.Bold = True
          .Interior.Color = RGB(189, 215, 238)
          .Columns.ColumnWidth = 20
        End With
        If ws.Range("A2") = "" Then
          ws.Range("A2").Value = ExComment.Parent.Address
          ws.Range("B2").Value = Left(ExComment.Text, InStr(1, ExComment.Text, ":") - 1)
          ws.Range("C2").Value = Right(ExComment.Text, Len(ExComment.Text) - InStr(1, ExComment.Text, ":"))
        Else
          ws.Range("A1").End(xlDown).Offset(1, 0) = ExComment.Parent.Address
          ws.Range("B1").End(xlDown).Offset(1, 0) = Left(ExComment.Text, InStr(1, ExComment.Text, ":") - 1)
          ws.Range("C1").End(xlDown).Offset(1, 0) = Right(ExComment.Text, Len(ExComment.Text) - InStr(1, ExComment.Text, ":"))
        End If
      Next ExComment
      End Sub

