Attribute VB_Name = "OrdPC"
Sub ordenesParaCargue()



Windows(Orders).Activate
Sheets("ORDENES").Select
Range("A1").AutoFilter
ActiveWorkbook.ActiveSheet.AutoFilter.Sort.SortFields.Clear
ActiveWorkbook.ActiveSheet.AutoFilter.Sort.SortFields.Add2 Key:=Range _
        ("N:N"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.ActiveSheet.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

Set a = Range("I1", Range("I1").End(xlDown))
Set b = Range("N1", Range("N1").End(xlDown))
Set c = Range("Q1", Range("Q1").End(xlDown))
Set d = Range("R1", Range("R1").End(xlDown))
Set e = Range("U1", Range("U1").End(xlDown))
Union(a, b, c, d, e).Copy
 

Workbooks.Open "\\Smedarch1\grupos\Equity_Sales\Alejandro\PlantillaOrdenesComplementacion.xlsm"
 
 
 
Range("A1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
       :=False, Transpose:=False
 

Application.Run "PlantillaOrdenesComplementacion.xlsm!ThisWorkbook.Complementacion"

ActiveWorkbook.Close SaveChanges = True

End Sub
