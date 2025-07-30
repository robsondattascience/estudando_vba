Sub PreencherDadosHPNominalFerias()

    Dim FolderPath As String
    Dim FileExt As String
    Dim Filename As String
    Dim wbSource As Workbook
    Dim wbDestino As Workbook
    Dim wsOrigNominal As Worksheet
    Dim wsOrigFerias As Worksheet
    Dim wsDestNominal As Worksheet
    Dim wsDestFerias As Worksheet
    Dim ultimaLinha As Long
    Dim linhaDestinoNominal As Long
    Dim linhaDestinoFerias As Long

    ' Defina o caminho da pasta e a extensão dos arquivos
    FolderPath = "C:\SeuCaminho\Arquivos\"  ' << ALTERE AQUI
    FileExt = "*.xlsm"                      ' << OU "*.xlsx", "*.xls", etc.

    ' Configura a planilha de destino
    Set wbDestino = ThisWorkbook
    Set wsDestNominal = wbDestino.Sheets("NOMINAL OP")
    Set wsDestFerias = wbDestino.Sheets("FÉRIAS")

    ' Limpar apenas os dados, preservando cabeçalhos (linha 1)
    ultimaLinha = wsDestNominal.Cells(wsDestNominal.Rows.Count, "E").End(xlUp).Row
    If ultimaLinha >= 2 Then
        wsDestNominal.Range("A2:E" & ultimaLinha).ClearContents
    End If

    ultimaLinha = wsDestFerias.Cells(wsDestFerias.Rows.Count, "D").End(xlUp).Row
    If ultimaLinha >= 2 Then
        wsDestFerias.Range("A2:D" & ultimaLinha).ClearContents
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Inicia busca pelos arquivos
    Filename = Dir(FolderPath & FileExt)

    Do While Filename <> ""

        If Filename <> wbDestino.Name Then

            Set wbSource = Workbooks.Open(FolderPath & Filename)

            On Error Resume Next
            Set wsOrigNominal = wbSource.Sheets("NOMINAL")
            Set wsOrigFerias = wbSource.Sheets("FERIAS")
            On Error GoTo 0

            ' Copiar dados de NOMINAL
            If Not wsOrigNominal Is Nothing Then
                ultimaLinha = wsOrigNominal.Cells(wsOrigNominal.Rows.Count, "E").End(xlUp).Row
                If ultimaLinha >= 2 Then
                    linhaDestinoNominal = wsDestNominal.Cells(wsDestNominal.Rows.Count, "A").End(xlUp).Row + 1
                    If linhaDestinoNominal < 2 Then linhaDestinoNominal = 2
                    wsOrigNominal.Range("A2:E" & ultimaLinha).Copy
                    wsDestNominal.Range("A" & linhaDestinoNominal).PasteSpecial xlPasteValues
                End If
            End If

            ' Copiar dados de FERIAS
            If Not wsOrigFerias Is Nothing Then
                ultimaLinha = wsOrigFerias.Cells(wsOrigFerias.Rows.Count, "D").End(xlUp).Row
                If ultimaLinha >= 2 Then
                    linhaDestinoFerias = wsDestFerias.Cells(wsDestFerias.Rows.Count, "A").End(xlUp).Row + 1
                    If linhaDestinoFerias < 2 Then linhaDestinoFerias = 2
                    wsOrigFerias.Range("A2:D" & ultimaLinha).Copy
                    wsDestFerias.Range("A" & linhaDestinoFerias).PasteSpecial xlPasteValues
                End If
            End If

            wbSource.Close SaveChanges:=False

        End If

        Filename = Dir() ' Próximo arquivo

    Loop

    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    MsgBox "Importação finalizada com sucesso!"

End Sub