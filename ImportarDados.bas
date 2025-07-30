Sub ImportarDados()

    Dim wbOrigin As Workbook
    Dim wsOrigin As Worksheet
    Dim wsDestin As Worksheet
    Dim filePath As String
    Dim countLastLine As Long
    Dim rangeCopy As Range
    Dim columnsRequired As String


    ' Desativa a atualização da tela e eventos automáticos
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    ' Defina aqui as colunas que deseja copiar (exemplo: A:B ou A,C,D)
    columnsRequired = "A2:L"

    
    ' Solicita ao usuário para escolher o arquivo origem
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Selecione o arquivo de origem"
        .Filters.Clear
        .Filters.Add "Arquivos Excel", "*.xls*"
        If .Show <> -1 Then Exit Sub
        filePath = .SelectedItems(1)
    End With

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Abre o arquivo origem
    Set wbOrigin = Workbooks.Open(filePath)

    ' Percorre cada planilha do arquivo origem
    For Each wsOrigin In wbOrigin.Worksheets
        ' Verifica última linha preenchida na planilha destino
        countLastLine = wsDestin.Cells(wsDestin.Rows.Count, 1).End(xlUp).Row + 1
        
        ' Ajusta se a primeira célula estiver vazia
        If countLastLine = 2 And wsDestin.Cells(1, 1).Value = "" Then countLastLine = 1
        
        ' Define o intervalo a copiar das colunas desejadas
        On Error Resume Next
        Set rangeCopy = wsOrigin.Range(columnsRequired).SpecialCells(xlCellTypeConstants)
        On Error GoTo 0

        If Not rangeCopy Is Nothing Then
            ' Copia os dados e cola na próxima linha disponível da planilha destino
            rangeCopy.Copy
            wsDestino.Cells(countLastLine, 1).PasteSpecial xlPasteValues
            Set rangeCopy = Nothing
        End If
    Next wsOrigem

    ' Fecha o arquivo de origem sem salvar alterações
    wbOrigem.Close SaveChanges:=False

    ' Ativa novamente a atualização da tela e eventos automáticos
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True


    MsgBox "Importação concluída com sucesso!", vbInformation

End Sub



Sub ApagarDados()

    ' Desativa a atualização da tela e eventos automáticos
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    Dim ws As Worksheet
    Dim lastRow As Long
    
    ' Defina a planilha onde você deseja apagar os dados
    Set ws = ThisWorkbook.Sheets("DADOS") ' Altere "Banco" para o nome da sua planilha, se necess�rio
    
    ' Encontrar a última linha preenchida na planilha
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Verificar se há dados abaixo da linha 3
    If lastRow >= 2 Then
        ' Apagar todas as linhas abaixo da linha 3
        ws.Rows("2:" & lastRow).Delete
    End If
    
    ' Ativa novamente a atualização da tela e eventos automáticos
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    MsgBox "Dados a partir da linha 2 foram apagados com sucesso!"

End Sub