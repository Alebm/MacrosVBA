Attribute VB_Name = "OyDrecomplementar"
  Sub OyDrecomplementar()
Attribute OyDrecomplementar.VB_Description = "Lista que sle de OyD para recomplementar organizacion y extraccion de Ordenes"
Attribute OyDrecomplementar.VB_ProcData.VB_Invoke_Func = "a\n14"

    Rows("2:2").Select
    Selection.Delete Shift:=xlUp
    Columns("F:F").Select
    Selection.Delete Shift:=xlToLeft
    Columns("I:I").Select
    Selection.Delete Shift:=xlToLeft
    Columns("R:U").Select
    Selection.Delete Shift:=xlToLeft
    Columns("S:S").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlToLeft
    Range("S2").Select
    Selection = 1
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight
    Range("C2").Select
    ActiveCell.FormulaR1C1 = "=+MID(RC[1],6,10)"
    Range("C2").Select
    'Selection.AutoFill Destination:=Range("C2:C")
    'Range("C2:CxlEnd").Select
    Selection.AutoFill Destination:=Range("C2:C" & Range("D" & Rows.Count).End(xlUp).Row)
    Range(Selection, Selection.End(xlDown)).Select
    
    'Selection.AutoFill Destination:=Range("C2").End(xlDown)s
    'Range("C2").End(xlDown).Select
    Selection.Copy
    Range("C2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
     Range("T2").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("C2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlMultiply, _
        SkipBlanks:=False, Transpose:=False
    Range("D1").Select
    Selection.Copy
    Range("C1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("D:D").Select
    Selection.Delete Shift:=xlToLeft
    


  End Sub
  
  
