Attribute VB_Name = "EnviConf"
Dim fecha1 As String
Dim MySheet As Worksheet
Dim MyRange As Range
Dim UList As Collection
Global UListValue As Variant
'Global UListValue

Dim i As Long
Sub EnvioConf()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.CutCopyMode = False
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.CutCopyMode = False

fecha1 = Format(Date, "mm/dd/yyyy")


Workbooks.Open "C:\Users\Albedoya\Documents\Acumulado.xlsm"
Sheets("ACUMULADO").Select
   If ActiveSheet.AutoFilterMode Then
        'Range("A1").AutoFilter
    Else
        Range("A1").AutoFilter
    End If
    
    ActiveWorkbook.ActiveSheet.AutoFilter.ShowAllData
    ActiveWorkbook.ActiveSheet.AutoFilter.Sort.SortFields.Clear
    Worksheets("ACUMULADO").Range("A1").AutoFilter Field:=2, Operator:= _
        xlFilterValues, Criteria2:=Array(2, fecha1)


'1
 Dim MySheet As Worksheet
 Dim MyRange As Range
 Dim UList As Collection
 Dim UListValue As Variant
 Dim i As Long
 
'2
 
 Sheets("ACUMULADO").Activate
 Set MySheet = Sheets("ACUMULADO")
 
'3
 If MySheet.AutoFilterMode = False Then
 Exit Sub
 End If
 
 
'4
 Range("D1").Offset(1, 0).Select
 MyRange1 = Range(Selection, Selection.End(xlDown)).SpecialCells(xlCellTypeVisible).Row
 MyRange2 = Range(Selection, Selection.End(xlDown)).SpecialCells(xlCellTypeVisible).Rows.Count
 FMyRange = MyRange1 + MyRange2 - 1
 
 Set MyRange = Range(MySheet.AutoFilter.Range.Columns(4).Address)
 
'5
 Set UList = New Collection
 
'6
 On Error Resume Next
  For i = MyRange1 To FMyRange
 'For i = 2 To MyRange.Rows.Count
 UList.Add MyRange.Cells(i, 1), CStr(MyRange.Cells(i, 1))
 Next i
 On Error GoTo 0
 
'7
 For Each UListValue In UList
  
'8
 On Error Resume Next
 Application.DisplayAlerts = False
 'Sheets(CStr(UListValue)).Delete
 'Application.DisplayAlerts = True
 On Error GoTo 0

Application.DisplayAlerts = False

Abrir = "C:\Users\Albedoya\Documents\Envio correos" & ".xlsm"

Dim TD As Date

Dim OutlookApp As Outlook.Application
Dim Correo As Outlook.MailItem
Dim archivo As String


Workbooks.Open (Abrir)

fecha1 = Format(Date, "dd/mm/yyyy")
año = Format(fecha1, "yyyy")
mes = Format(fecha1, "mmmm")
fecha2 = Replace(fecha1, "/", ".")

UListValue = UListValue

Sheets("Hoja1").Range("C22").Value = UListValue
'Range("C22").Value = "CITIBANK"
Sheets("Hoja1").Range("c23").Value = año
Sheets("Hoja1").Range("C24").Value = mes
Sheets("Hoja1").Range("C25").Value = fecha2


fecha1f = Format(fecha1, "dddd")

Pregunta = MsgBox("Hay un dia festivo en COLOMBIA de hoy hasta el cumplimiento?", vbYesNoCancel + vbQuestion, "EXCELeINFO")

Select Case Pregunta
 Case Is = 6
 If fecha1f = "viernes" Or fecha1f = "jueves" Or fecha1f = "miércoles" Then
    TD = Format(Date + 6, "dd-mm-yyyy")
 Else
    TD = Format(Date + 4, "dd-mm-yyyy")
 End If
 Case Is = 7
  If fecha1f = "viernes" Or fecha1f = "jueves" Or fecha1f = "miércoles" Then
    TD = Format(Date + 5, "dd-mm-yyyy")
 Else
    TD = Format(Date + 3, "dd-mm-yyyy")
 End If
 Case Is = 2
   d = Application.InputBox(prompt:="Ingrese # de dias", Title:=Dias, Type:=1)
   TD = Format(Date + d, "dd-mm-yyyy")
 End Select
 
 
 If Format(TD, "dddd") = "sabado" Then
    TD = Format(TD + 2, "dd-mm-yyyy")
 ElseIf Format(TD, "dddd") = "domingo" Then
    TD = Format(TD + 1, "dd-mm-yyyy")
 End If

 Sheets("Hoja1").Range("A24").Value = TD
 
 Ini = Sheets("Correos").Range("A1").Row
 Fin = Sheets("Correos").Range("A1", Sheets("Correos").Range("A1").End(xlDown)).Count
 

For i = Ini To Fin
    If Sheets("Correos").Cells(i, 1) = UListValue Then
        BCCref = Sheets("Correos").Cells(i, 1).Offset(0, 1)
    End If
Next

Set OutlookApp = New Outlook.Application
Set Correo = OutlookApp.CreateItem(o)
'Crear el correo y mostrarlo
destino = BCCref
BCC = Sheets("Hoja1").Range("E21")
mensaje = Sheets("Hoja1").Range("A21")
cuerpo = Sheets("Hoja1").Range("B21")
archivo = Sheets("Hoja1").Range("C21")

With Correo
.Display
    .To = destino
    .BCC = BCC
    .Subject = mensaje
    '.Body = cuerpo & .Body
    .HTMLBody = cuerpo & .HTMLBody
    .Attachments.Add archivo
    '.Send

End With
With Application
.EnableEvents = True
.ScreenUpdating = True
End With

ActiveWorkbook.Close SaveChanges = True

Next


End Sub

Sub Envio_Con()

Application.DisplayAlerts = False

Abrir = "C:\Users\Albedoya\Documents\Envio correos" & ".xlsm"

Dim TD As Date

Dim OutlookApp As Outlook.Application
Dim Correo As Outlook.MailItem
Dim archivo As String


Workbooks.Open (Abrir)

fecha1 = Format(Date - 1, "dd/mm/yyyy")
año = Format(fecha1, "yyyy")
mes = Format(fecha1, "mmmm")
fecha2 = Replace(fecha1, "/", ".")

UListValue = UListValue

Range("C22").Value = UListValue
'Range("C22").Value = "CITIBANK"
Range("c23").Value = año
Range("C24").Value = mes
Range("C25").Value = fecha2


fecha1f = Format(fecha1, "dddd")

Pregunta = MsgBox("Hay un dia festivo en COLOMBIA de hoy hasta el cumplimiento?", vbYesNoCancel + vbQuestion, "EXCELeINFO")

Select Case Pregunta
 Case Is = 6
 If fecha1f = "viernes" Or fecha1f = "jueves" Or fecha1f = "miércoles" Then
    TD = Format(Date + 6, "dd-mm-yyyy")
 Else
    TD = Format(Date + 4, "dd-mm-yyyy")
 End If
 Case Is = 7
  If fecha1f = "viernes" Or fecha1f = "jueves" Or fecha1f = "miércoles" Then
    TD = Format((Date - 1) + 5, "dd-mm-yyyy")
 Else
    TD = Format(Date + 3, "dd-mm-yyyy")
 End If
 Case Else
  'InputBox
 End Select
 
 
 If Format(TD, "dddd") = "sabado" Then
    TD = Format(TD + 2, "dd-mm-yyyy")
 ElseIf Format(TD, "dddd") = "domingo" Then
    TD = Format(TD + 1, "dd-mm-yyyy")
 End If

 Range("A24").Value = TD
 
 Ini = Range("F22").Row
 Fin = Range("F22", Range("F22").End(xlDown)).Count
 

For i = Ini To Fin
    If Cells(i, 1) = UListValue Then
        BCCref = Cells(i, 1).Offset(0, 1).Value
    End If
i = i + 1
Next
  


Set OutlookApp = New Outlook.Application
Set Correo = OutlookApp.CreateItem(o)
'Crear el correo y mostrarlo
destino = Range("D21")
BCC = BCCref
mensaje = Range("A21")
cuerpo = Range("B21")
archivo = Range("C21")

With Correo
.Display
    .To = destino
    .BCC = BCC
    .Subject = mensaje
    '.Body = cuerpo & .Body
    .HTMLBody = cuerpo & .HTMLBody
    .Attachments.Add archivo
    '.Send

End With
With Application
.EnableEvents = True
.ScreenUpdating = True
End With

ActiveWorkbook.Close SaveChanges = True

End Sub

