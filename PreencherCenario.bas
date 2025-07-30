Option Explicit

Sub PreencherCenario()

    Dim PlanilhaCenario As Worksheet
    Dim PlanilhaConversor As Worksheet
    Dim UltLinhaConversor As Long
    Dim ConteudoConversor As Range
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set PlanilhaCenario = ActiveWorkbook.Sheets("ORGANICO")
    Set PlanilhaConversor = ThisWorkbook.Sheets("CONVERSOR DE X PARA")
    
    UltLinhaConversor = PlanilhaConversor.Cells(PlanilhaConversor.Rows.Count, "F").End(xlUp).Row
    Set ConteudoConversor = PlanilhaConversor.Range("F2:M" & UltLinhaConversor)
    
    
    PlanilhaCenario.Range("A5").Resize(ConteudoConversor.Rows.Count, ConteudoConversor.Columns.Count).value = ConteudoConversor.value
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
End Sub
