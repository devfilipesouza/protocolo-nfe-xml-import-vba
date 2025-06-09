Sub ImportarDadosXML()
    ' Declaração das variáveis necessárias
    Dim ws As Worksheet
    Dim ultimaLinha As Long
    Dim caminhoPasta As String
    Dim chaveAcesso As String
    Dim caminhoArquivo As String
    Dim xmlDoc As Object
    Dim xmlNode As Object
    Dim dataEmissao As String ' Variável para armazenar a data de emissão como string temporariamente
    Dim cnpjEmitente As String
    Dim numeroNota As String
    Dim natOperacao As String ' Variável para armazenar a natureza da operação
    Dim formaPagamento As String ' Variável para armazenar a forma de pagamento
    Dim i As Long

    ' Define a planilha de trabalho onde os dados serão inseridos
    Set ws = ThisWorkbook.Sheets("PRNF")

    ' Caminho fixo onde os arquivos XML das NF-es estão armazenados
    caminhoPasta = "\\SRV-RELUZ\Users\ACESSO INTERNO\DOCUMENTOS FISCAIS\XML ENTRADA\"

    ' Descobre a última linha com dados na coluna J (onde estão as chaves de acesso)
    ultimaLinha = ws.Cells(ws.Rows.Count, 10).End(xlUp).Row

    ' Laço para percorrer cada linha com chave de acesso a partir da linha 13
    For i = 13 To ultimaLinha
        chaveAcesso = ws.Cells(i, 10).Value  ' Coluna J

        ' Reinicia as variáveis de texto para cada iteração
        natOperacao = ""
        formaPagamento = ""

        ' Só processa se a célula não estiver vazia
        If Trim(chaveAcesso) <> "" Then

            ' Concatena o caminho completo do arquivo XML com a chave de acesso
            caminhoArquivo = caminhoPasta & chaveAcesso & ".xml"

            ' Verifica se o arquivo existe fisicamente
            If Dir(caminhoArquivo) <> "" Then

                ' Instancia o objeto para leitura do XML
                Set xmlDoc = CreateObject("MSXML2.DOMDocument")
                xmlDoc.Load (caminhoArquivo)

                ' --- DADOS EXTRAÍDOS DO XML ---

                ' Data de emissão (tag: dhEmi)
                Set xmlNode = xmlDoc.SelectSingleNode("//ide/dhEmi")
                If Not xmlNode Is Nothing Then
                    dataEmissao = Mid(xmlNode.Text, 1, 10) ' Pega apenas a parte YYYY-MM-DD
                    ' Converte a string da data para um valor de data do Excel e atribui diretamente
                    ' O Excel irá armazenar como número e você pode formatar a célula na planilha.
                    ws.Cells(i, 3).Value = CDate(dataEmissao) ' Salva na Coluna C
                End If

                ' CNPJ do emitente
                Set xmlNode = xmlDoc.SelectSingleNode("//emit/CNPJ")
                If Not xmlNode Is Nothing Then
                    cnpjEmitente = xmlNode.Text
                    ws.Cells(i, 4).Value = cnpjEmitente ' Coluna D
                End If

                ' Número da nota fiscal
                Set xmlNode = xmlDoc.SelectSingleNode("//ide/nNF")
                If Not xmlNode Is Nothing Then
                    numeroNota = xmlNode.Text
                    ws.Cells(i, 7).Value = numeroNota ' Coluna G
                End If

                ' --- LÓGICA PARA FORMA DE PAGAMENTO (Coluna I) ---

                ' Primeiramente, verifica se existe a tag <dup> (duplicata)
                Set xmlNode = xmlDoc.SelectSingleNode("//cobr/dup")
                If Not xmlNode Is Nothing Then
                    formaPagamento = "FATURAMENTO"
                Else
                    ' Se não houver duplicata, verifica a Natureza da Operação
                    Set xmlNode = xmlDoc.SelectSingleNode("//ide/natOp")
                    If Not xmlNode Is Nothing Then
                        natOperacao = UCase(Trim(xmlNode.Text)) ' Converte para maiúsculas para comparação
                    End If

                    If InStr(natOperacao, "BONIFICACAO") > 0 Then
                        formaPagamento = "BONIFICAÇÃO"
                    ElseIf InStr(natOperacao, "DEVOLUCAO") > 0 Or InStr(natOperacao, "REMESSA") > 0 Then
                        formaPagamento = "REMESSA"
                    Else
                        ' Se nenhuma das condições anteriores for atendida, busca a forma de pagamento na tag infAdFisco
                        Set xmlNode = xmlDoc.SelectSingleNode("//infAdic/infAdFisco")
                        If Not xmlNode Is Nothing Then
                            ' Extrai a parte da forma de pagamento, se existir.
                            ' Procura por "FORMA PAGAMENTO: " e pega o que vem depois
                            Dim pos As Long
                            pos = InStr(xmlNode.Text, "FORMA PAGAMENTO:")
                            If pos > 0 Then
                                formaPagamento = Trim(Mid(xmlNode.Text, pos + Len("FORMA PAGAMENTO:")))
                            Else
                                formaPagamento = "Não Especificado" ' Caso não encontre o padrão "FORMA PAGAMENTO:"
                            End If
                        Else
                            ' Caso a tag infAdFisco não exista, atribui "À vista" conforme a nova regra
                            formaPagamento = "À VISTA"
                        End If
                    End If
                End If

                ' Salva a forma de pagamento na Coluna I
                ws.Cells(i, 9).Value = formaPagamento

                ' Indica que o XML foi processado corretamente
                ws.Cells(i, 11).Value = "XML VÁLIDO" ' Coluna K

            Else
                ' Se o arquivo XML não for localizado, limpa as células e marca como "XML NÃO ENCONTRADO"
                ws.Cells(i, 3).Value = ""  ' Coluna C (Data Emissão)
                ws.Cells(i, 4).Value = ""  ' Coluna D (CNPJ Emitente)
                ws.Cells(i, 7).Value = ""  ' Coluna G (Número da Nota)
                ws.Cells(i, 9).Value = ""  ' Coluna I (Forma de Pagamento)
                ws.Cells(i, 11).Value = "XML NÃO ENCONTRADO" ' Coluna K
            End If
        End If
    Next i

    ' --- LÓGICA DE ORDENAÇÃO DOS DADOS ---
    ' Verifica se há dados para ordenar (a partir da linha 13 até a última linha com dados na coluna J)
    If ultimaLinha >= 13 Then
        With ws.Sort
            .SortFields.Clear ' Limpa quaisquer campos de classificação anteriores
            ' Adiciona a coluna C (Data de Emissão) como chave de classificação
            .SortFields.Add Key:=ws.Columns("C"), _
                            SortOn:=xlSortOnValues, _
                            Order:=xlAscending, _
                            DataOption:=xlSortNormal
            .SetRange ws.Range("A13:K" & ultimaLinha) ' Define o intervalo a ser classificado (da linha 13 até a última linha com dados na coluna J, e as colunas até K)
            .Header = xlNo ' Indica que a primeira linha do intervalo não é um cabeçalho (pois os cabeçalhos estão antes da linha 13)
            .MatchCase = False ' Não diferencia maiúsculas de minúsculas
            .Orientation = xlTopToBottom ' Classifica por linhas
            .SortMethod = xlPinYin ' Método de classificação (útil para texto)
            .Apply ' Aplica a classificação
        End With
    End If

    ' Feedback ao usuário no final do processo
    MsgBox "Importação e ordenação concluídas! Verifique o status de cada arquivo.", vbInformation, "Processo Finalizado"

End Sub
