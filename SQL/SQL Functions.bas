
      Private sqlConn As Object
      
      Function OpenSQLConnection()
          If sqlConn Is Nothing Then _
              Set sqlConn = CreateObject("ADODB.Connection")
             
          If sqlConn.ConnectionString = Empty Then _
              sqlConn.ConnectionString = "Data Source='192.168.203.97';Initial Catalog='Study';User ID='accao_pass';Password='pass(accao)';"
          
          If sqlConn.State = adStateClosed Then
              sqlConn.Provider = "SQLOLEDB"
              sqlConn.Open
          End If
      
      End Function
      
      Function CloseSQLConnection() As Boolean
          If sqlConn Is Nothing Then
              CloseSQLConnection = True
          Else
              sqlConn.Close
          End If
      End Function
      
      Function GetFromSQL(ByVal sql As String, Optional closeConn As Boolean = True) As String
         
          OpenSQLConnection
          
          Dim cmd As Object
          Dim rs As Object
          
          Set cmd = CreateObject("ADODB.Command")
          cmd.CommandType = adCmdText
          
          cmd.ActiveConnection = sqlConn
          cmd.CommandText = sql
          
          Set rs = cmd.Execute
          
          GetFromSQL = ReadRecordSet(rs)
          
          If closeConn Then _
              CloseSQLConnection
          
          Set cmd = Nothing
          Set rs = Nothing
          
      End Function
      
      Private Function ReadRecordSet(ByRef rs As Object) As String
          ' Read the recordset to a string
          Dim s As String
          Dim f As Object
          
          If rs Is Nothing Then
              s = "NODATA"
          ElseIf rs.State = adStateClosed Then
              s = "NODATA"
          ElseIf rs.EOF Then
              s = "NODATA"
          Else
              While Not rs.EOF
                  For Each f In rs.Fields
                      If f.Attributes And adFldIsNullable Then
                          s = s + ";" + "NULL"
                      Else
                          s = s + ";" + CStr(f.Value)
                      End If
                  Next
                  
                  rs.MoveNext ' next record
              Wend
          End If
          ReadRecordSet = s
      End Function
      
      '**************************************
      ' Suite de funcoes utilizadas
      ' na formatacao de dados do CGL
      ' Resp.: Neylor Ohmaly Rodigues e Silva
      ' Criado: 2006-08-26 19:07:00
      '***************************************
      
      Function SendToSQL() As String
          Dim shp As Shape
          Dim btnOK As CommandButton
          
          ' Copy the parameter container to
          ' the calling worksheet
          '
          data.Shapes("grpGroup").Copy
          ActiveSheet.Paste
          Set shp = ActiveSheet.Shapes("grpGroup")
          
          ' Change the button properties
          '
          shp.Visible = msoTrue
          shp.GroupItems(3).Name = "btnOK"
          Set btnOK = shp.GroupItems(3).DrawingObject.Object
          
          ' Copy the button callback function
          ' to the calling worksheet
          '
          
      End Function

