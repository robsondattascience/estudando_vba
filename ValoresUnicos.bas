Option Explicit

Sub ValoresUnicos()
    
    Dim wo As Worksheet
    Dim wd As Worksheet
    Dim dict As Object
    Dim celula As Range
    Dim chave As Variant
    Dim IntervaloOrigem As Integer
    Dim EncontraProg As String
    Dim PegarProg As String
    Dim LinhaDestino As Long
    Dim LinhaProg As Long
    Dim ValorEncontrado As String
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    Set wo = ThisWorkbook.Sheets("BASE NEO")
    Set wd = ThisWorkbook.Sheets("CONVERSOR DE X PARA")
    EncontraProg = wd.Range("C6").value
        
    IntervaloOrigem = wo.Cells(wo.Rows.Count, "L").End(xlUp).Row
    
  
    'Cria um Objeto(Dicionário) com chave e valor
    Set dict = CreateObject("Scripting.Dictionary")
    
    'Percorre a coluna de Programa compara com a C6 e pega somente os pertencentes a mesma
    For LinhaProg = 2 To IntervaloOrigem
        If wo.Cells(LinhaProg, "J").value = EncontraProg Then
        
            PegarProg = Trim(wo.Cells(LinhaProg, "L").value)
            
            If PegarProg <> "" Then
                If Not dict.exists(PegarProg) Then
                    dict.Add PegarProg, Nothing
                End If
            End If
        End If
    Next LinhaProg
        
    LinhaDestino = 2
    For Each chave In dict.Keys
        wd.Cells(LinhaDestino, "A").value = chave
        LinhaDestino = LinhaDestino + 1
    Next chave
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True

End Sub
