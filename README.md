# 📥 Protocolo NFe - Importação de Dados XML via VBA

Este projeto em VBA automatiza a leitura de arquivos XML de Notas Fiscais Eletrônicas (NF-e) e importa dados essenciais diretamente para uma planilha do Excel. Ele foi desenvolvido para facilitar a conferência e registro de informações fiscais, otimizando o processo manual de lançamento.

## 🚀 Funcionalidades

- 📂 Leitura de arquivos XML localizados em uma pasta de rede.
- 🔍 Extração automática de:
  - Data de emissão
  - CNPJ do emitente
  - Número da nota fiscal
- ✅ Validação do XML e retorno de status direto na planilha
- 💡 Aviso ao final da importação com status individual por linha

## 📄 Estrutura da Planilha

A macro atua sobre a planilha chamada **`PRNF`**, utilizando as seguintes colunas:

| Coluna | Informação                         |
|--------|------------------------------------|
| C      | Data de Emissão da NF-e            |
| D      | CNPJ do Emitente                   |
| G      | Número da Nota Fiscal              |
| J      | Chave de Acesso (nome do XML)      |
| K      | Status de importação do XML        |

> ⚠️ A importação inicia a partir da linha 13, onde devem estar inseridas as chaves de acesso.

## 🛠️ Como Usar

1. Abra o arquivo Excel com a macro habilitada.
2. Certifique-se de que a planilha **PRNF** contenha as chaves de acesso na coluna J a partir da linha 13.
3. Execute a macro `ImportarDadosXML`.
4. Verifique o preenchimento das colunas e o status em "XML VÁLIDO" ou "XML NÃO ENCONTRADO".

## 📁 Local dos Arquivos XML

O código atualmente busca os arquivos XML em um caminho fixo de rede
