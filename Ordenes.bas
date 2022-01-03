Attribute VB_Name = "Ordenes"
Public Orders As String


Sub OrdenesC()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.CutCopyMode = False
Application.Calculation = xlCalculationAutomatic

Abrir = "C:\Users\Albedoya\Documents\Acumulado" & ".xlsm"
Orders = ActiveWorkbook.Name
fecha1 = Date

Call ordenesParaCargue
Call Hijas

Windows(Orders).Activate
Sheets("ORDENES").Select

NOrdenes = Range("C1", Range("C1").End(xlDown)).Count

If NOrdenes > 2 Then

Range("C2").Select
Range(Selection, Selection.End(xlDown)).Offset(0, 1).Value = Date
Range("K2").Select
Range(Selection, Selection.End(xlDown)).Select
  Selection.Replace What:=Array("CITI_NY", "SANTANDER_NY", "MOERUS"), Replacement:=Array("CITIBANK", "SANTANDER", "MOERUS CAP"), LookAt:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
Range("N:N").Replace What:=".BGa", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
Set a = Range("D1", Range("D1").End(xlDown))
Set b = Range("I1", Range("I1").End(xlDown))
Set c = Range("K1", Range("K1").End(xlDown))
Set d = Range("N1", Range("N1").End(xlDown))
Set e = Range("Q1", Range("Q1").End(xlDown))
Set f = Range("R1", Range("R1").End(xlDown))
Set g = Range("S1", Range("S1").End(xlDown))
Union(a, b, c, d, e, f, g).Copy

Else

Range("C2").Offset(0, 1).Value = Date
Range("K2").Select
Range(Selection, Selection.End(xlDown)).Select
  Selection.Replace What:=Array("CITI_NY", "SANTANDER_NY", "MOERUS"), Replacement:=Array("CITIBANK", "SANTANDER", "MOERUS CAP"), LookAt:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
Range("N:N").Replace What:=".BGa", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
Set a = Range("D1", Range("D1").End(xlDown))
Set b = Range("I1", Range("I1").End(xlDown))
Set c = Range("K1", Range("K1").End(xlDown))
Set d = Range("N1", Range("N1").End(xlDown))
Set e = Range("Q1", Range("Q1").End(xlDown))
Set f = Range("R1", Range("R1").End(xlDown))
Set g = Range("S1", Range("S1").End(xlDown))
Union(a, b, c, d, e, f, g).Copy

End If


'Application.Union(Range("D2", Selection.End(xlDown)), Range("K2", Selection.End(xlDown)), Range("N2", Selection.End(xlDown)), Range("Q2", Selection.End(xlDown))).Copy

Sheets.Add
ActiveSheet.Name = "THistory"
Range("A1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
       :=False, Transpose:=False
Range("A1").AutoFilter
Columns("B:B").Insert Shift:=xlToRight
    ActiveWorkbook.ActiveSheet.AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.ActiveSheet.AutoFilter.Sort.SortFields.Add2 Key:=Range _
        ("D:D"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.ActiveSheet.AutoFilter.Sort.SortFields.Add2 Key:=Range _
        ("E:E"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.ActiveSheet.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
Range("A2", Range("A2").End(xlDown)).Select
    With Selection
        .NumberFormat = "d-mmm"
    End With
TD = Application.InputBox("T + X", "Cumplimiento", Type:=2)
'Range("B1") = TD
Range("B2").FormulaR1C1 = "=RC[-1]+" & TD
Range("B2").Select

TNOrdenes = Range("A1", Range("A1").End(xlDown)).Count

If TNOrdenes > 2 Then

Selection.AutoFill Destination:=Range("B2:B" & Range("A" & Rows.Count).End(xlUp).Row)
Set a = Range("A2", Range("A2").End(xlDown))
Set b = Range("B2", Range("B2").End(xlDown))
Set c = Range("D2", Range("D2").End(xlDown))
Union(a, b, c).Copy

Else

Set a = Range("A2")
Set b = Range("B2")
Set c = Range("D2")
Union(a, b, c).Copy

End If







Workbooks.Open (Abrir)
Application.ScreenUpdating = False
Acum = ActiveWorkbook.Name
Sheets("ACUMULADO").Select

If TNOrdenes > 2 Then

ActiveSheet.AutoFilter.ShowAllData
Range("B1").Select
Selection.End(xlDown).Offset(1, 0).PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
Windows(Orders).Activate
Sheets("THistory").Select
Set a = Range("C2", Range("C2").End(xlDown))
Set b = Range("E2", Range("E2").End(xlDown))
Union(a, b).Copy
Windows(Acum).Activate
Sheets("ACUMULADO").Select
Range("E1").Select
Selection.End(xlDown).Offset(1, 0).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
Windows(Orders).Activate
Sheets("THistory").Select
Set a = Range("F2", Range("F2").End(xlDown))
Set b = Range("G2", Range("G2").End(xlDown))
Union(a, b).Copy
Windows(Acum).Activate
Sheets("ACUMULADO").Select
Range("H1").Select
Selection.End(xlDown).Offset(1, 0).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
Windows(Orders).Activate
Sheets("THistory").Select
Range("H2", Range("H2").End(xlDown)).Copy
Windows(Acum).Activate
Sheets("ACUMULADO").Select
Range("O1").Select
Selection.End(xlDown).Offset(1, 0).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
Else

ActiveSheet.AutoFilter.ShowAllData
Range("B1").Select
Selection.End(xlDown).Offset(1, 0).PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
Windows(Orders).Activate
Sheets("THistory").Select
Set a = Range("C2")
Set b = Range("E2")
Union(a, b).Copy
Windows(Acum).Activate
Sheets("ACUMULADO").Select
Range("E1").Select
Selection.End(xlDown).Offset(1, 0).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
Windows(Orders).Activate
Sheets("THistory").Select
Set a = Range("F2")
Set b = Range("G2")
Union(a, b).Copy
Windows(Acum).Activate
Sheets("ACUMULADO").Select
Range("H1").Select
Selection.End(xlDown).Offset(1, 0).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
Windows(Orders).Activate
Sheets("THistory").Select
Range("H2").Copy
Windows(Acum).Activate
Sheets("ACUMULADO").Select
Range("O1").Select
Selection.End(xlDown).Offset(1, 0).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

End If

        
        
Application.Run "Acumulado.xlsm!Acumul.Acumulado"

End Sub







