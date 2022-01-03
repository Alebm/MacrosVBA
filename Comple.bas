Attribute VB_Name = "Comple"
Global fecha


Sub Contar_Condicional()
'inicializo la variable resultado

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.CutCopyMode = False
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.CutCopyMode = False

Workbooks.Open Application.GetOpenFilename

ActiveSheet.Range("A1").Select
Range(Selection, Selection.End(xlDown)).Select
Range(Selection, Selection.End(xlToRight)).Copy

fecha = Format(Date, "d.m.yyyy")

Workbooks("Complementación extranjeros " & fecha & ".xlsx").Activate

ActiveSheet.Name = "Folios"
Sheets.Add
ActiveSheet.Name = "Ordenes"
Orders = ActiveWorkbook.Name
Sheets("Ordenes").Select
Range("A1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
Range("F1", Range("F1").End(xlDown)).Select
Selection.Replace What:="VENTA", Replacement:="S", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
Selection.Replace What:="COMPRA", Replacement:="B", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
Range("I1").Formula = "=CONCAT(E1, J1, F1)"
Range("I1").Select
Selection.AutoFill Destination:=Range("I1:I" & Range("A" & Rows.Count).End(xlUp).Row)

j = 1

Workbooks("Complementación extranjeros " & fecha & ".xlsx").Activate


Sheets("Folios").Select

If Range("D1") = "Sales Trader Notes" Then
    Range("E2").Formula = "=CONCAT(I2, D2, G2)"
ElseIf Range("A1") = "Sales Trader Notes" Then
    Range("E2").Formula = "=CONCAT(I2, A2, G2)"
Else
    MsgBox "No hay que concatenar"
    Exit Sub
    
End If

Range("E2").Select
Selection.AutoFill Destination:=Range("E2:E" & Range("B" & Rows.Count).End(xlUp).Row)
    


    Range("A1").Select
    Registros = Range(Selection, Selection.End(xlDown)).SpecialCells(xlCellTypeVisible).Count
'1
 Dim MySheet As Worksheet
 Dim MyRange As Range
 Dim UList As Collection
 Dim UListValue As Variant
 Dim i As Long
 
 Sheets("Ordenes").Select
 
  
  If ActiveSheet.AutoFilterMode Then
        'Range("A1").AutoFilter
    Else
        Range("A1").AutoFilter
    End If

'2
 'Set MySheet = ActiveSheet
 Set MySheet = Sheets("Ordenes")
 
 
'3
 If MySheet.AutoFilterMode = False Then
 Exit Sub
 End If
 
 
 
'4
 Set MyRange = Range(MySheet.AutoFilter.Range.Columns(9).Address)
 
'5
 Set UList = New Collection
 
'6
 On Error Resume Next
 For i = 1 To MyRange.Rows.Count
 UList.Add MyRange.Cells(i, 1), CStr(MyRange.Cells(i, 1))
 Next i
 On Error GoTo 0
 
'7
 For Each UListValue In UList
  
'8
 On Error Resume Next
 Application.DisplayAlerts = False
 Sheets(CStr(UListValue)).Delete
 Application.DisplayAlerts = True
 On Error GoTo 0
 
'9
 'MyRange.AutoFilter Field:=5, Criteria1:=UListValue
 Sheets("Folios").Select
   If ActiveSheet.AutoFilterMode Then
        'Range("A1").AutoFilter
    Else
        Range("A1").AutoFilter
    End If
    Set MySheet2 = Sheets("Folios")
    Set MyRange2 = Range(MySheet2.AutoFilter.Range.Columns(5).Address)
    MyRange2.AutoFilter Field:=5, Criteria1:=UListValue
 
'10
 'Filtro = MySheet.AutoFilter.Range.Copy
 'Worksheets.Add.Paste
 'Workbooks.Add
 'ActiveSheet.Paste Destination:=Range("A1")
 'ActiveSheet.Name = Left(UListValue, 30)
 
RESULTADO = 0
    
    busqueda = UListValue
    
    Workbooks("Complementación extranjeros " & fecha & ".xlsx").Activate
    Sheets("Ordenes").Select
    registros2 = Range("B1", Range("B1").End(xlDown)).SpecialCells(xlCellTypeVisible).Count
    
    
    Sheets("Folios").Select
    Range("E1").Offset(1, 0).Select
    Registros = Range(Selection, Selection.End(xlDown)).Row
    rango1 = Range("E1", Range("E1").End(xlDown)).Rows.Count
    RangoC = Rango + rango1
    
    busqueda2 = Sheets("Ordenes").Cells(j, 9)
    reemplazo = Sheets("Ordenes").Cells(j, 3)
    Sheets("Folios").Select
    For i = Registros To RangoC
        If Sheets("Folios").Cells(i, 5) = busqueda Then 'criterio qeu debe cumplir
            If busqueda = busqueda2 Then
            Sheets("Folios").Cells(i, 3) = reemplazo
            End If
        End If
    Next
    'Next

 j = j + 1
    
 '11
 Next UListValue
 
 If ActiveSheet.FilterMode Then ActiveSheet.ShowAllData
 
 Workbooks("Complementación extranjeros " & fecha & ".xlsx").Activate
    Range("B1").Select
    OrdL = Range(Selection, Selection.End(xlDown)).SpecialCells(xlCellTypeVisible).Count
    

 For i = 2 To OrdL
    If Cells(i, 3) <> 0 Then
        i = i + 1
    Else: Cells(i, 3) = 0
        MsgBox "falta orden"
        Exit Sub
 End If

Next

ActiveWorkbook.Save
Application.DisplayAlerts = False
Sheets("Ordenes").Delete

ActiveWorkbook.Save
       
Call Envio_Com

ActiveWorkbook.Close SaveChanges = True


     
End Sub

Sub Envio_Com()

Application.DisplayAlerts = False

Abrir = "C:\Users\Albedoya\Documents\Envio correos" & ".xlsm"

Workbooks.Open (Abrir)

Dim OutlookApp As Outlook.Application
Dim Correo As Outlook.MailItem
Dim archivo As String


Set OutlookApp = New Outlook.Application
Set Correo = OutlookApp.CreateItem(o)
'Crear el correo y mostrarlo
destino = Range("D12")
BCC = Range("E12")
mensaje = Range("A12")
cuerpo = Range("B12")
archivo = Range("C12")

With Correo
.Display
    .To = destino
    .BCC = BCC
    .Subject = mensaje
    '.Body = cuerpo & .Body
    .HTMLBody = cuerpo & .HTMLBody
    .Attachments.Add archivo
    .Send

End With
With Application
.EnableEvents = True
.ScreenUpdating = True
End With

ActiveWorkbook.Close SaveChanges = True

End Sub




