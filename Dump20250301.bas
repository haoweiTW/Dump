Attribute VB_Name = "Dump"
Sub Dump()
Attribute Dump.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Copyright 2025 HaoWei
'
    Dim myDate As String
    myDate="2025/3/Day+1"
    
    Dim myCSV As String
    myCSV=ThisWorkbook.ActiveSheet.Name
    
    
    Columns("K:K").ColumnWidth = 26
    Columns("L:L").ColumnWidth = 12
    Columns("M:M").ColumnWidth = 12
    Columns("T:T").ColumnWidth = 20
    Columns("U:U").ColumnWidth = 20
    Columns("AR:AR").ColumnWidth = 12
    Columns("AW:AW").ColumnWidth = 12
    Columns("DE:DE").ColumnWidth = 26
    
    Columns("K:K").Select
    Selection.Font.Bold = True
    Columns("M:M").Select
    Selection.Font.Bold = True
    Columns("AR:AR").Select
    Selection.Font.Bold = True
    Columns("AW:AW").Select
    Selection.Font.Bold = True
    Selection.NumberFormatLocal = "#,##0_ "
    Columns("DE:DE").Select
    Selection.Font.Bold = True
    
    Columns("A:J").Select
    Selection.EntireColumn.Hidden = True
    Columns("L:L").Select
    Selection.EntireColumn.Hidden = True
    Columns("X:AM").Select
    Selection.EntireColumn.Hidden = True
    Columns("AU:AV").Select
    Selection.EntireColumn.Hidden = True
    
    Rows("1:1").Select
    Selection.AutoFilter
    ActiveWorkbook.Worksheets(myCSV).AutoFilter.Sort. _
        SortFields.Add2 Key:=Range("K1"), SortOn:=xlSortOnValues, Order:= _
        xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(myCSV).AutoFilter. _
        Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ActiveSheet.Range("$A$1:$DF$99999").AutoFilter Field:=13, Criteria1:= _
        myDate, Operator:=xlAnd
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 14
    ActiveWindow.ScrollColumn = 15
    ActiveWindow.ScrollColumn = 16
    ActiveWindow.ScrollColumn = 17
    ActiveWindow.ScrollColumn = 18
    ActiveSheet.Range("$A$1:$DF$99999").AutoFilter Field:=45, Criteria1:= _
        "=EUR/USD", Operator:=xlOr, Criteria2:="=USD/EUR"
    
    Columns("AR:AR").Select
    
    ActiveWindow.LargeScroll ToRight:=-1
End Sub
