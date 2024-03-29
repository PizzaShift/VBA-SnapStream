
      Sub InsertPivotTable()
      'Macro By ExcelChamps
      'https://excelchamps.com/blog/vba-to-create-pivot-table/'
      
      'Declare Variables
      Dim PSheet As Worksheet
      Dim DSheet As Worksheet
      Dim PCache As PivotCache
      Dim PTable As PivotTable
      Dim PRange As Range
      Dim LastRow As Long
      Dim LastCol As Long
      
      'Insert a New Blank Worksheet
      On Error Resume Next
      Application.DisplayAlerts = False
      Worksheets("PivotTable").Delete
      Sheets.Add Before:=ActiveSheet
      ActiveSheet.Name = "PivotTable"
      Application.DisplayAlerts = True
      Set PSheet = Worksheets("PivotTable")
      Set DSheet = Worksheets("Data")
      
      'Define Data Range
      LastRow = DSheet.Cells(Rows.Count, 1).End(xlUp).Row
      LastCol = DSheet.Cells(1, Columns.Count).End(xlToLeft).Column
      Set PRange = DSheet.Cells(1, 1).Resize(LastRow, LastCol)
      
      
      'Define Pivot Cache (Solution 1)
      Set PCache = ActiveWorkbook.PivotCaches.Create(xlDatabase, “T_GRIR”)
      
      'Define Pivot Cache (Solution 2)
      'Set PCache = ActiveWorkbook.PivotCaches.Create _
      ''(SourceType:=xlDatabase, SourceData:=PRange). _
      'CreatePivotTable(TableDestination:=PSheet.Cells(2, 2), _
      'TableName:="SalesPivotTable")
      
      'Insert Blank Pivot Table (Solution 1)
      Set PTable = PCache.CreatePivotTable _
      (TableDestination:=PSheet.Cells(1, 1), TableName:=”GRIRpt”)
      
      'Insert Blank Pivot Table (Solution 2)
      'Set PTable = PCache.CreatePivotTable _
      ''(TableDestination:=PSheet.Cells(1, 1), TableName:="SalesPivotTable")
      
      'Insert Row Fields
      With ActiveSheet.PivotTables("SalesPivotTable").PivotFields("Year")
      .Orientation = xlRowField
      .Position = 1
      End With
      With ActiveSheet.PivotTables("SalesPivotTable").PivotFields("Month")
      .Orientation = xlRowField
      .Position = 2
      End With
      
      'Insert Column Fields
      With ActiveSheet.PivotTables("SalesPivotTable").PivotFields("Zone")
      .Orientation = xlColumnField
      .Position = 1
      End With
      
      'Insert Data Field
      With ActiveSheet.PivotTables("SalesPivotTable")
      .PivotFields ("Amount")
      .Orientation = xlDataField
      .Function = xlSum
      .NumberFormat = "#,##0"
      .Name = "Revenue "
      End With
      
      'Format Pivot Table
      ActiveSheet.PivotTables("SalesPivotTable").ShowTableStyleRowStripes = True
      ActiveSheet.PivotTables("SalesPivotTable").TableStyle2 = "PivotStyleMedium9"
      
      End Sub

      '------------------------------------------------------------------------------------------------
      Sub InsertPivotTable()
      
      ‘Declare Variables
      Dim PSheet As Worksheet
      Dim DSheet As Worksheet
      Dim Pcache As PivotCache
      Dim PTable As PivotTable
      Dim PRange As Range
      Dim LastRow As Long
      Dim LastCol As Long
      
      ‘Insert a New Blank Worksheet
      On Error Resume Next
      Application.DisplayAlerts = False
      Worksheets(“PivotTable”).Delete
      Set PSheet = ActiveWorkbook.Worksheets.Add
      PSheet.Name = “PivotTable”
      Set DSheet = Worksheets(“Page 1″)
      
      ‘Define Data Range
      LastRow = DSheet.Cells(Rows.Count, 1).End(xlUp).Row
      LastCol = DSheet.Cells(1, Columns.Count).End(xlToLeft).Column
      Set PRange = DSheet.Cells(1, 1).Resize(LastRow, LastCol)
      
      ‘Define Pivot Cache
      Set Pcache = ActiveWorkbook.PivotCaches.Create _
      (SourceType:=xlDatabase, SourceData:=PRange). _
      CreatePivotTable(tabledestination:=PSheet.Cells(2, 2), TableName:=”SalesPivotTable”)
      
      ‘Insert Blank Pivot Table
      Set PTable = Pcache.CreatePivotTable _
      (tabledestination:=PSheet.Cells(1, 1), TableName:=”SalesPivotTable”)
      
      ‘Insert Row Fields
      With ActiveSheet.PivotTables(“SalesPivotTable”).PivotFields(“Name”)
      .Orientation = xlRowField
      .Position = 1
      End With
      
      ‘Insert Column Fields
      With ActiveSheet.PivotTables(“SalesPivotTable”).PivotFields(“Priority”)
      .Orientation = xlColumnField
      .Position = 1
      End With
      
      ‘Insert Data Field
      With ActiveSheet.PivotTables(“SalesPivotTable”).PivotFields(“Number”)
      .Orientation = xlDataField
      .Position = 1
      .Function = xlCount
      .NumberFormat = “#,##0”
      .Name = “Revenue ”
      End With
      
      ‘Format Pivot Table
      ActiveSheet.PivotTables(“SalesPivotTable”).ShowTableStyleRowStripes = True
      ActiveSheet.PivotTables(“SalesPivotTable”).TableStyle2 = “PivotStyleMedium9”
      
      End Sub

