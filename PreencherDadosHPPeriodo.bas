Sub PreencherDadosHPNominalFerias()

    Dim FolderPath As String
    Dim FileExt As String
    Dim Filename As String
    Dim wbOrigem As Workbook
    Dim wbDestino As Workbook
    Dim wsOrigNominal As Worksheet
    Dim wsOrigFerias As Worksheet
    Dim wsDestNominal As Worksheet
    Dim wsDestFerias As Worksheet
    Dim ultimaLinha As Long

    ' Caminho da pasta dos arquivos de destino (alterar conforme necessário)
    FolderPath = "C:\SeuCaminho\Arquivos\"  ' << ALTERE AQUI
    FileExt = "*.xlsm"                      ' << OU "*.xlsx", "*.xls", etc.

    ' O arquivo de dados deve estar aberto antes de rodar o código
    Set wbOrigem = ActiveWorkbook ' Ou use: Workbooks("NomeDoArquivo.xlsx")

    ' Planilhas de origem
    Set wsOrigNominal = wbOrigem.Sheets("NOMINAL")
    Set wsOrigFerias = wbOrigem.Sheets("FERIAS")

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Loop pelos arquivos da pasta
    Filename = Dir(FolderPath & FileExt)
    Do While Filename <> ""

        ' Evita acidentalmente usar o mesmo arquivo da origem como destino
        If Filename <> wbOrigem.Name Then

            ' Abre o arquivo de destino
            Set wbDestino = Workbooks.Open(FolderPath & Filename)

            On Error Resume Next
            Set wsDestNominal = wbDestino.Sheets("NOMINAL OP")
            Set wsDestFerias = wbDestino.Sheets("FÉRIAS")
            On Error GoTo 0

            If Not wsDestNominal Is Nothing And Not wsDestFerias Is Nothing Then

                ' Limpa dados antigos (mantendo cabeçalhos)
                ultimaLinha = wsDestNominal.Cells(wsDestNominal.Rows.Count, "E").End(xlUp).Row
                If ultimaLinha >= 2 Then
                    wsDestNominal.Range("A2:E" & ultimaLinha).ClearContents
                End If

                ultimaLinha = wsDestFerias.Cells(wsDestFerias.Rows.Count, "D").End(xlUp).Row
                If ultimaLinha >= 2 Then
                    wsDestFerias.Range("A2:D" & ultimaLinha).ClearContents
                End If

                ' Copia dados da planilha "NOMINAL"
                ultimaLinha = wsOrigNominal.Cells(wsOrigNominal.Rows.Count, "E").End(xlUp).Row
                If ultimaLinha >= 2 Then
                    wsOrigNominal.Range("A2:E" & ultimaLinha).Copy
                    wsDestNominal.Range("A2").PasteSpecial xlPasteValues
                End If

                ' Copia dados da planilha "FERIAS"
                ultimaLinha = wsOrigFerias.Cells(wsOrigFerias.Rows.Count, "D").End(xlUp).Row
                If ultimaLinha >= 2 Then
                    wsOrigFerias.Range("A2:D" & ultimaLinha).Copy
                    wsDestFerias.Range("A2").PasteSpecial xlPasteValues
                End If

                ' Salva e fecha o destino
                wbDestino.Close SaveChanges:=True

            Else
                MsgBox "Erro: Sheets 'NOMINAL OP' ou 'FÉRIAS' não encontradas no arquivo " & Filename, vbExclamation
                wbDestino.Close SaveChanges:=False
            End If

        End If

        Filename = Dir() ' Próximo arquivo

    Loop

    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    MsgBox "Importação finalizada com sucesso!"

End Sub