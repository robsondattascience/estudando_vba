Option Explicit

Sub ColetarDadosDeArquivos()

    Dim FolderPath As String
    Dim Filename As String
    Dim wbSource As Workbook
    Dim wsSource As Worksheet
    Dim wbDestino As Workbook
    Dim wsDestino As Worksheet
    Dim LastRowU As Long
    Dim LastRowV As Long
    Dim LastRowA As Long
    Dim SourceLastRow As Long
    Dim DataRange1 As Range
    Dim DataRange2 As Range
    Dim SheetName As String
    Dim FileExt As String
    Dim FileNameWithoutExt As String ' Para armazenar o nome do arquivo sem extens�o
    
    ' Desativa a atualização da tela e eventos automáticos
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' Defina o caminho da pasta e a planilha que deseja buscar
    FolderPath = "" ' Altere para o caminho dos arquivos
    SheetName = "HIRING PLAN" ' Altere para o nome da sheet que você quer acessar
    FileExt = "*.xlsm*" ' Nome\Extensão dos arquivos (pode ser .xls, .xlsx, .xlsb, .xlsm  e etc.)

    ' Configura a planilha de destino
    Set wbDestino = ThisWorkbook ' O workbook onde a macro está sendo executada
    Set wsDestino = wbDestino.Sheets("Banco") ' Planilha de banco de dados

    ' Loop por todos os arquivos na pasta
    Filename = Dir(FolderPath & FileExt)
    
    Do While Filename <> ""
        ' Abrir o arquivo
        Set wbSource = Workbooks.Open(FolderPath & Filename)
        
        ' Verificar se a planilha existe
        On Error Resume Next
        Set wsSource = wbSource.Sheets(SheetName)
        On Error GoTo 0
        
        If Not wsSource Is Nothing Then
            ' -----------------------------------------------------
            ' Nova etapa: Copiar B5:B36 e colar na coluna U da planilha de "Banco"
            ' -----------------------------------------------------
            ' Encontrar a última linha preenchida na coluna U para colar abaixo
            LastRowU = wsDestino.Cells(wsDestino.Rows.Count, "U").End(xlUp).Row + 1
            
            ' Copiar os dados de B5:B36 e colar na coluna U
            Set DataRange1 = wsSource.Range("B6:B36")
            wsDestino.Range("U" & LastRowU & ":U" & LastRowU + DataRange1.Rows.Count - 1).Value = DataRange1.Value
            
            ' -----------------------------------------------------
            ' Primeira etapa: Copiar CR5:CR36 e CS5:CS36 para colunas N e O
            ' -----------------------------------------------------
            ' Encontrar a última linha preenchida nas colunas N e O para colar abaixo
            LastRowV = wsDestino.Cells(wsDestino.Rows.Count, "V").End(xlUp).Row + 1
            
            ' Copiar os dados de DS5:DS36 e colar na coluna N
            Set DataRange1 = wsSource.Range("DS6:DS36")
            wsDestino.Range("V" & LastRowV & ":V" & LastRowV + DataRange1.Rows.Count - 1).Value = DataRange1.Value
            
            ' Copiar os dados de DT5:DT36 e colar na coluna W
            Set DataRange1 = wsSource.Range("DT6:DT36")
            wsDestino.Range("W" & LastRowV & ":W" & LastRowV + DataRange1.Rows.Count - 1).Value = DataRange1.Value
            
            ' Colocar o nome do arquivo na coluna X ao lado dos dados copiados
            FileNameWithoutExt = Left(Filename, InStrRev(Filename, ".") - 1) ' Pegar nome do arquivo sem a extens�o
            wsDestino.Range("X" & LastRowV & ":X" & LastRowV + DataRange1.Rows.Count - 1).Value = FileNameWithoutExt
            
            ' -----------------------------------------------------
            ' Segunda etapa: Copiar EW5:FM e colar em A2:Q2
            ' -----------------------------------------------------
            ' Encontrar a última linha preenchida nas colunas A até Q para colar abaixo
            LastRowA = wsDestino.Cells(wsDestino.Rows.Count, "A").End(xlUp).Row + 1
            
            ' Encontre a última linha preenchida no arquivo de origem para a coluna DR:DZ
            SourceLastRow = wsSource.Cells(wsSource.Rows.Count, "EW").End(xlUp).Row
            
            ' Definir o intervalo de DR5:DZ até a última linha preenchida
            Set DataRange2 = wsSource.Range("EW6:FM" & SourceLastRow)
            
            ' Colar os dados na planilha de banco nas colunas A:Q
            wsDestino.Range("A" & LastRowA & ":Q" & LastRowA + DataRange2.Rows.Count - 1).Value = DataRange2.Value
            
            ' Adicionar o nome do arquivo na coluna R (ao lado dos dados colados nas colunas A:Q)
            wsDestino.Range("R" & LastRowA & ":R" & LastRowA + DataRange2.Rows.Count - 1).Value = FileNameWithoutExt
        End If
        
        ' Fechar o arquivo fonte sem salvar
        wbSource.Close SaveChanges:=False
        
        ' Próximo arquivo
        Filename = Dir
    Loop
    
        ' Ativa novamente a atualização da tela e eventos automáticos
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    MsgBox "Dados coletados com sucesso!"
    
End Sub