Sub SepararPorCaracteres()
    Dim filePath As String
    Dim fileContent As String
    Dim fileLine As String
    Dim fileNum As Integer
    Dim linha As Long
    Dim ws As Worksheet, wsTemplate As Worksheet
    Dim rowOffset As Long
    Dim totalLines As Long
    Dim pdfFilePath As String
    Dim numero As Integer

    ' Definir a planilha onde os dados serão exibidos e a planilha Template
    Set ws = ThisWorkbook.Sheets("Planilha1")
    Set wsTemplate = ThisWorkbook.Sheets("Template")
    
    ' Inicializa o número aleatório
    Randomize
    numero = Int((100 * Rnd)) ' Gera um número aleatório entre 0 e 99

    ' Exibir a caixa de diálogo para o usuário selecionar o arquivo .txt
    filePath = Application.GetOpenFilename("Text Files (*.txt), *.txt", , "Selecione o arquivo .txt")
    
    ' Verificar se o usuário selecionou um arquivo ou cancelou a seleção
    If filePath = "False" Then
        MsgBox "Nenhum arquivo foi selecionado. A macro será encerrada."
        Exit Sub
    End If
    
    ' Limpar dados existentes na planilha
    ws.Cells.Clear
    
    ' Exibir os nomes das variáveis na primeira linha
    ws.Cells(1, 1).Value = "ID"
    ws.Cells(1, 2).Value = "Tipo de Inscrição"
    ws.Cells(1, 3).Value = "CNPJ/CPF"
    ws.Cells(1, 4).Value = "Nome do Fornecedor"
    ws.Cells(1, 5).Value = "Endereço do Fornecedor"
    ws.Cells(1, 6).Value = "CEP do Fornecedor"
    ws.Cells(1, 7).Value = "Código do Banco"
    ws.Cells(1, 8).Value = "Código da Agência"
    ws.Cells(1, 9).Value = "Dígito da Agência"
    ws.Cells(1, 10).Value = "Conta Corrente"
    ws.Cells(1, 11).Value = "Dígito da Conta Corrente"
    ws.Cells(1, 12).Value = "Número do Pagamento"
    ws.Cells(1, 13).Value = "Carteira"
    ws.Cells(1, 14).Value = "Nosso Número"
    ws.Cells(1, 15).Value = "Seu Número"
    ws.Cells(1, 16).Value = "Data de Vencimento"
    ws.Cells(1, 17).Value = "Data de Emissão"
    ws.Cells(1, 18).Value = "Data Limite para Desconto"
    ws.Cells(1, 19).Value = "Fator de Vencimento"
    ws.Cells(1, 20).Value = "Valor do Documento"
    ws.Cells(1, 21).Value = "Valor do Pagamento"
    ws.Cells(1, 22).Value = "Valor do Desconto"
    ws.Cells(1, 23).Value = "Valor do Acréscimo"
    ws.Cells(1, 24).Value = "Tipo do DOC COMPE/TED"
    ws.Cells(1, 24).Value = "Tipo de Documento"
    ws.Cells(1, 24).Value = "Tipo de Documento"
    ws.Cells(1, 24).Value = "Tipo de Documento"
    ws.Cells(1, 24).Value = "Tipo de Documento"
    ws.Cells(1, 24).Value = "Tipo de Documento"
    
    ' Iniciar na segunda linha para os dados
    rowOffset = 2
    linha = 1

    ' Contar o total de linhas no arquivo
    fileNum = FreeFile
    Open filePath For Input As fileNum
    Do Until EOF(fileNum)
        Line Input #fileNum, fileLine
        totalLines = totalLines + 1
    Loop
    Close fileNum

    ' Reabrir o arquivo para processar as linhas, ignorando a primeira e a penúltima
    fileNum = FreeFile
    Open filePath For Input As fileNum
    Do Until EOF(fileNum)
        Line Input #fileNum, fileLine
        fileContent = Trim(fileLine)
        
        ' Ignorar a primeira e a penúltima linha
        If linha <> 1 And linha <> totalLines - 1 Then
            ' Verifica se a linha tem o tamanho esperado
            If Len(fileContent) < 251 Then
                MsgBox "Linha com formato inválido: " & fileContent
                GoTo NextLine
            End If
            
            ' Extraindo dados de acordo com os caracteres especificados
            ws.Cells(rowOffset, 1).Value = Mid(fileContent, 1, 1) ' ID
            ws.Cells(rowOffset, 2).Value = GetTipoInscricao(Mid(fileContent, 2, 1)) ' Tipo de Inscrição
            ws.Cells(rowOffset, 3).Value = "'" & Mid(fileContent, 3, 15) ' CNPJ/CPF como texto
            ws.Cells(rowOffset, 4).Value = Mid(fileContent, 18, 30) ' Nome do Fornecedor
            ws.Cells(rowOffset, 5).Value = Mid(fileContent, 48, 40) ' Endereço do Fornecedor
            ws.Cells(rowOffset, 6).Value = Mid(fileContent, 88, 5) & "-" & Mid(fileContent, 93, 3) ' CEP do Fornecedor
            ws.Cells(rowOffset, 7).Value = Mid(fileContent, 96, 3) ' Código do Banco
            ws.Cells(rowOffset, 8).Value = Mid(fileContent, 99, 5) ' Código da Agência
            ws.Cells(rowOffset, 9).Value = Mid(fileContent, 104, 1) ' Dígito da Agência
            ws.Cells(rowOffset, 10).Value = Mid(fileContent, 105, 13) ' Conta Corrente
            ws.Cells(rowOffset, 11).Value = Mid(fileContent, 118, 2) ' Dígito da Conta Corrente
            ws.Cells(rowOffset, 12).Value = Mid(fileContent, 120, 16) ' Número do Pagamento
            ws.Cells(rowOffset, 13).Value = Mid(fileContent, 136, 3) ' Carteira
            ws.Cells(rowOffset, 14).Value = Mid(fileContent, 139, 12) ' Nosso Número
            ws.Cells(rowOffset, 15).Value = Mid(fileContent, 151, 15) ' Seu Número
            ws.Cells(rowOffset, 16).Value = FormatDate(Mid(fileContent, 166, 8)) ' Data de Vencimento
            ws.Cells(rowOffset, 17).Value = FormatDate(Mid(fileContent, 174, 8)) ' Data de Emissão
            ws.Cells(rowOffset, 18).Value = FormatDate(Mid(fileContent, 182, 8)) ' Data Limite para Desconto
            ws.Cells(rowOffset, 19).Value = Mid(fileContent, 191, 4) ' Fator de Vencimento
            ws.Cells(rowOffset, 20).Value = Mid(fileContent, 195, 10) ' Valor do Documento
            ws.Cells(rowOffset, 21).Value = Mid(fileContent, 205, 15) ' Valor do Pagamento
            ws.Cells(rowOffset, 22).Value = Mid(fileContent, 220, 15) ' Valor do Desconto
            ws.Cells(rowOffset, 23).Value = Mid(fileContent, 235, 15) ' Valor do Acréscimo
            ws.Cells(rowOffset, 24).Value = GetTipoDocumento(Mid(fileContent, 250, 2)) ' Tipo de Documento
            ws.Cells(rowOffset, 25).Value = Mid(fileContent, 252, 10) ' Número Nota Fiscal/Fatura/Duplicata
            ws.Cells(rowOffset, 26).Value = GetModalidadePagamento(Mid(fileContent, 264, 2))
            ws.Cells(rowOffset, 27).Value = FormatDate(IIf(Mid(fileContent, 266, 8) = "", Mid(fileContent, 166, 8), Mid(fileContent, 266, 8)))
            ws.Cells(rowOffset, 28).Value = Mid(fileContent, 274, 3)
            ws.Cells(rowOffset, 29).Value = Mid(fileContent, 277, 2)
            ws.Cells(rowOffset, 30).Value = Mid(fileContent, 289, 1)
            ws.Cells(rowOffset, 31).Value = Mid(fileContent, 290, 2)
            ws.Cells(rowOffset, 32).Value = Mid(fileContent, 292, 4)
            ws.Cells(rowOffset, 33).Value = Mid(fileContent, 296, 15) ' Saldo Disponível
            ws.Cells(rowOffset, 34).Value = Mid(fileContent, 311, 15) ' Valor da taxa pré-funding
            ws.Cells(rowOffset, 35).Value = Mid(fileContent, 374, 1) ' Tipo do DOC COMPE/TED
            ws.Cells(rowOffset, 36).Value = Mid(fileContent, 375, 6) ' Número do DOC COMPE/TED

            ' Preencher o Template com os dados da linha processada
            wsTemplate.Range("A4").Value = ws.Cells(rowOffset, 12).Value ' Número do Pagamento
            wsTemplate.Range("E4").Value = ws.Cells(rowOffset, 24).Value ' Tipo de Documento
            wsTemplate.Range("L7").Value = ws.Cells(rowOffset, 20).Value ' Valor do Documento
            wsTemplate.Range("E7").Value = ws.Cells(rowOffset, 10).Value ' Conta Corrente
            wsTemplate.Range("B7").Value = ws.Cells(rowOffset, 7).Value ' Código do Banco
            wsTemplate.Range("A15").Value = ws.Cells(rowOffset, 4).Value ' Nome do Fornecedor
            wsTemplate.Range("H14").Value = ws.Cells(rowOffset, 8).Value ' Código da Agência
            wsTemplate.Range("K14").Value = ws.Cells(rowOffset, 10).Value ' Conta Corrente
            wsTemplate.Range("L7").Value = ValorPorExtenso(ws.Cells(rowOffset, 21).Value) ' Valor do Pagamento por extenso

            ' Atualiza a célula K7 da planilha Template com o número
            wsTemplate.Range("K7").Value = numero + linha - 1 ' Incrementa o número a cada linha processada
            
            ' Salvar o Template como PDF para cada linha
            pdfFilePath = ThisWorkbook.Path & "\Documento_" & linha & ".pdf"
            wsTemplate.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfFilePath

            ' Incrementar a linha para Planilha1
            rowOffset = rowOffset + 1
        End If
        
NextLine:
        linha = linha + 1
    Loop
    
    ' Fechar o arquivo
    Close fileNum
    
    MsgBox "Processamento e geração de PDFs concluídos!"
End Sub

Function GetTipoInscricao(tipo As String) As String
    Select Case tipo
        Case "1"
            GetTipoInscricao = "CPF"
        Case "2"
            GetTipoInscricao = "CNPJ"
        Case "3"
            GetTipoInscricao = "Outros"
        Case Else
            GetTipoInscricao = "Inválido"
    End Select
End Function

Function FormatDate(data As String) As String
    If Len(data) = 8 Then
        FormatDate = Mid(data, 1, 2) & "/" & Mid(data, 3, 2) & "/" & Mid(data, 5, 4)
    Else
        FormatDate = "Data Inválida"
    End If
End Function

Function GetTipoDocumento(tipo As String) As String
    Select Case tipo
        Case "01"
            GetTipoDocumento = "Nota Fiscal/Fatura"
        Case "02"
            GetTipoDocumento = "Fatura"
        Case "03"
            GetTipoDocumento = "Nota Fiscal"
        Case "04"
            GetTipoDocumento = "Duplicata"
        Case Else
            GetTipoDocumento = "Outro"
    End Select
End Function

Function ValorPorExtenso(Valor As Double) As String
    Dim TextoValor As String
    ' Aqui você pode adicionar uma função para converter números para texto por extenso
    TextoValor = Application.WorksheetFunction.Text(Valor, "#.00") & " reais"
    ValorPorExtenso = TextoValor
End Function
Function GetModalidadePagamento(tipo As String) As String
        Select Case tipo
            Case "01"
                GetModalidadePagamento = "Crédito c/c"
            Case "02"
                GetModalidadePagamento = "Cheque OP"
            Case "05"
                GetModalidadePagamento = "Crédito em c/c real time"
            Case "03"
                GetModalidadePagamento = "DOC COMPE"
            Case "08"
                GetModalidadePagamento = "TED"
            Case "30"
                GetModalidadePagamento = "Rastreamento de Títulos"
            Case "31"
                GetModalidadePagamento = "Títulos Terceiros"
            Case Else
                GetModalidadePagamento = "Modalidade Inválida"
        End Select
End Function
