Sub ImportarDadosXML()
    ' Declaração das variáveis necessárias
    Dim ws As Worksheet
    Dim ultimaLinha As Long
    Dim caminhoPasta As String
    Dim chaveAcesso As String
    Dim caminhoArquivo As String
    Dim xmlDoc As Object
    Dim xmlNode As Object
    Dim dataEmissao As String
    Dim cnpjEmitente As String
    Dim numeroNota As String
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
                    ws.Cells(i, 3).Value = Format(CDate(dataEmissao), "dd/mm/yyyy") ' Salva na Coluna C
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
                
                ' Indica que o XML foi processado corretamente
                ws.Cells(i, 11).Value = "XML VÁLIDO" ' Coluna K

            Else
                ' Se o arquivo XML não for localizado
                ws.Cells(i, 11).Value = "XML NÃO ENCONTRADO" ' Coluna K
            End If
        End If
    Next i

    ' Feedback ao usuário no final do processo
    MsgBox "Importação concluída! Verifique o status de cada arquivo.", vbInformation, "Processo Finalizado"

End Sub
