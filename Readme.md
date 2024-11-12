# Macro de Separação de Dados em Arquivo .txt e Geração de PDFs

Este projeto consiste em uma macro VBA que importa dados de um arquivo `.txt`, separa as informações com base na posição dos caracteres, e gera arquivos PDF com base em um template específico. O objetivo é simplificar o processamento e a geração de documentos padronizados a partir de dados brutos.

## Funcionalidades

1. **Importação de Arquivo `.txt`**: O usuário pode selecionar um arquivo `.txt` que será processado.
2. **Separação de Dados**: Cada linha do arquivo é lida, e os dados são extraídos de acordo com posições específicas de caracteres.
3. **Geração de PDF**: Para cada linha processada, um PDF é gerado com os dados no formato especificado pelo template.
4. **Campos Configuráveis**: Os dados extraídos incluem informações como ID, CNPJ/CPF, Nome do Fornecedor, Endereço, Código do Banco, entre outros.
5. **Tratamento de Exceções**: Ignora a primeira e a penúltima linha do arquivo, e verifica o tamanho das linhas para garantir o formato correto.
6. **Funções Auxiliares**: Funções auxiliares para formatação de datas, tipo de documento, tipo de inscrição, modalidade de pagamento, e valor por extenso.

## Estrutura do Código

- **Sub SepararPorCaracteres**: Procedimento principal que realiza a importação do arquivo, separação dos dados, preenchimento da planilha de dados e geração de PDFs.
- **Função GetTipoInscricao**: Retorna o tipo de inscrição (CPF, CNPJ, etc.) com base no caractere de identificação.
- **Função FormatDate**: Converte uma string de data no formato `DDMMAAAA` para `DD/MM/AAAA`.
- **Função GetTipoDocumento**: Identifica o tipo de documento com base no código.
- **Função ValorPorExtenso**: Converte o valor do pagamento para texto por extenso.
- **Função GetModalidadePagamento**: Identifica a modalidade de pagamento com base no código.

## Pré-requisitos

Para utilizar essa macro, é necessário:

- Ter o Excel instalado com suporte a VBA (Visual Basic for Applications).
- Salvar este código em um módulo VBA dentro de uma planilha Excel.

## Como Utilizar

1. Abra o Excel e a planilha onde deseja rodar a macro.
2. Pressione `ALT + F11` para abrir o editor VBA.
3. Insira o código fornecido em um novo módulo.
4. Salve a planilha com a extensão `.xlsm` para habilitar macros.
5. Execute a macro `SepararPorCaracteres`:
   - Será exibida uma caixa de diálogo para selecionar o arquivo `.txt`.
   - O código importará os dados, processará e gerará os PDFs automaticamente.

## Layout da Planilha

- **Planilha1**: Usada para armazenar os dados importados e processados.
- **Template**: Contém o layout do documento para a geração de PDFs.

## Exemplo de Uso

O arquivo `.txt` precisa ter as linhas de dados formatadas corretamente com um comprimento mínimo de 251 caracteres. A macro ignora a primeira e a penúltima linha e processa as demais, gerando um PDF para cada linha processada.

## Personalizações

Você pode ajustar o código VBA para:
- Alterar a posição de cada campo extraído, caso o layout do arquivo de entrada seja diferente.
- Adicionar novas funções de formatação ou processamento.
- Modificar o layout do template para refletir suas necessidades específicas de saída em PDF.

## Problemas Conhecidos

- **Formato Inválido**: Caso uma linha no arquivo `.txt` tenha menos de 251 caracteres, uma mensagem de erro será exibida.
- **Formato de Data**: A função `FormatDate` espera que as datas estejam no formato `DDMMAAAA`.

## Contribuição

Sinta-se à vontade para contribuir com melhorias. Envie um pull request ou reporte problemas para continuarmos aprimorando este projeto.

## Licença

Este projeto está licenciado sob a [MIT License](LICENSE).
