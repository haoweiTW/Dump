Attribute VB_Name = "Dump"
Sub Dump()
Attribute Dump.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Copyright 2024 HaoWei
'
    Dim myDate As String
    myDate="3/Day+1/2024"
    
    Dim myCSV As String
    myCSV=ThisWorkbook.ActiveSheet.Name
    
    
    Columns("K:K").ColumnWidth = 26
    Columns("L:L").ColumnWidth = 12
    Columns("M:M").ColumnWidth = 12
    Columns("T:T").ColumnWidth = 20
    Columns("U:U").ColumnWidth = 20
    Columns("AR:AR").ColumnWidth = 12
    Columns("AW:AW").ColumnWidth = 12
    
    Columns("K:K").Select
    Selection.Font.Bold = True
    Columns("M:M").Select
    Selection.Font.Bold = True
    Columns("AR:AR").Select
    Selection.Font.Bold = True
    Columns("AW:AW").Select
    Selection.Font.Bold = True
    Selection.NumberFormatLocal = "#,##0_ "
    
    Columns("A:J").Select
    Selection.EntireColumn.Hidden = True
    Columns("L:L").Select
    Selection.EntireColumn.Hidden = True
    Columns("X:AM").Select
    Selection.EntireColumn.Hidden = True
    
    Rows("1:1").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$DF$99999").AutoFilter Field:=22, Criteria1:="=EUR" _
        , Operator:=xlOr, Criteria2:="=USD"
    ActiveSheet.Range("$A$1:$DF$99999").AutoFilter Field:=23, Criteria1:="=EUR" _
        , Operator:=xlOr, Criteria2:="=USD"
    
    Columns("U:U").Select
    ActiveWorkbook.Worksheets(myCSV).AutoFilter.Sort. _
        SortFields.Add2 Key:=Range("U1:U99999"), SortOn:=xlSortOnValues, Order:= _
        xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(myCSV).AutoFilter. _
        Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Columns("M:M").Select
    ActiveSheet.Range("$A$1:$DF$99999").AutoFilter Field:=13, Operator:= _
        xlFilterValues, Criteria2:=Array(2, myDate)
    
    Columns("AR:AR").Select
End Sub
