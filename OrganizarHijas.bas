Attribute VB_Name = "OrganizarHijas"
Sub Hijas()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.CutCopyMode = False
Application.Calculation = xlCalculationAutomatic

dd = Day(Date) 'Día de hoy
mm = Month(Date) 'Mes de hoy
aa = Year(Date) 'Año de hoy

f2 = "Complementación extranjeros " & dd & "." & mm & "." & aa
ruta2 = "\\Smedarch1\grupos\Equity_Sales\Alejandro\Complementacion\"

Sheets("Hijas").Select
r = Application.InputBox("Dato de reemplazo", "Reemplazar", Type:=1)
Range("B1").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Replace What:=r, Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
   
Sheets("Hijas").Copy
Range("B1") = "Folios"
Range("C:C").Insert Shift:=xlToRight
Range("C1") = "Ordenes"
Range("E:E").Insert Shift:=xlToRight
Range("E1") = "Name"
Range("A1").AutoFilter
    ActiveWorkbook.ActiveSheet.AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.ActiveSheet.AutoFilter.Sort.SortFields.Add2 Key:=Range _
        ("I:I"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.ActiveSheet.AutoFilter.Sort.SortFields.Add2 Key:=Range _
        ("D:D"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.ActiveSheet.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
Range("I:I").Replace What:=".BGa", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

ActiveWorkbook.SaveAs ruta2 & f2 & ".xlsx", FileFormat:=xlWorkbookDefault, CreateBackup:=False
        
End Sub
