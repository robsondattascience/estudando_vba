Option Explicit

Sub ApagarDados()

    ' Desativa a atualização da tela e eventos automáticos
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    Dim ws As Worksheet
    Dim UltimaLinha As Long
    
    ' Defina a planilha onde você deseja apagar os dados
    Set ws = ActiveWorkbook.Sheets("ORGANICO") ' Altere o nome da sheet para o nome da sua planilha, se necessário
    
    ' Encontrar a última linha preenchida na planilha
    UltimaLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Verificar se há dados no intervalo
    If UltimaLinha >= 5 Then
        ' Apagar todas as linhas abaixo da linha 5
        ws.Range("A5:H" & UltimaLinha).ClearContents
    End If
    
    ' Ativa novamente a atualização da tela e eventos automáticos
    Application.ScreenUpdating = True
    Application.EnableEvents = True

End Sub

Sub ApagarDados()

    ' Desativa a atualização da tela e eventos automáticos
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    Dim ws As Worksheet
    Dim LastRow As Long
    
    ' Defina a planilha onde você deseja apagar os dados
    Set ws = ThisWorkbook.Sheets("Banco") ' Altere o nome da sheet para o nome da sua planilha, se necessário
    
    ' Encontrar a última linha preenchida na planilha
    LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Verificar se há dados abaixo da linha 3
    If LastRow >= 3 Then
        ' Apagar todas as linhas abaixo da linha 3
        ws.Rows("3:" & LastRow).Delete
    End If
    
    ' Ativa novamente a atualização da tela e eventos automáticos
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    MsgBox "Dados a partir da linha 3 foram apagados com sucesso!"

End Sub