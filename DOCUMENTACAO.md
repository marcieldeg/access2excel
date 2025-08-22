# Documentação do Projeto: Access2Excel

## 1. Visão Geral

O Access2Excel é uma ferramenta de linha de comando desenvolvida em Java para converter bancos de dados Microsoft Access (.MDB ou .ACCDB) em planilhas do Microsoft Excel (.XLS ou .XLSX).

O projeto utiliza as seguintes bibliotecas principais:
-   **Apache POI**: Para a criação e manipulação de arquivos Excel.
-   **Jackcess**: Para a leitura de bancos de dados MS Access.
-   **Apache Commons CLI**: Para uma análise robusta dos argumentos da linha de comando.

## 2. Como Usar

A ferramenta é executada via linha de comando com as seguintes opções:

```bash
java -jar Access2Excel.jar -i <arquivo_de_entrada> [-o <arquivo_de_saida>] [-f <formato>]
```

**Argumentos:**

*   `-i, --inputFile <caminho>`: (Obrigatório) O caminho para o banco de dados Access (.MDB ou .ACCDB).
*   `-o, --outputFile <caminho>`: (Opcional) O nome do arquivo Excel a ser gerado.
*   `-f, --format <formato>`: (Opcional) O formato da planilha de saída. Pode ser "XLS" ou "XLSX".
*   `-h, --help`: Exibe a mensagem de ajuda.

## 3. Arquitetura e Fluxo de Dados

O fluxo de processamento principal é orquestrado pela classe `Access2Excel.java`:

1.  **Abertura do Banco de Dados**: A aplicação abre o arquivo de banco de dados do Access em modo somente leitura usando a biblioteca Jackcess.
2.  **Criação da Planilha**: Uma nova planilha Excel é criada em memória (usando a API de streaming `SXSSFWorkbook` para XLSX para otimizar o uso de memória).
3.  **Iteração de Tabelas**: O código itera sobre cada tabela encontrada no banco de dados do Access.
4.  **Criação de Abas**: Para cada tabela do Access, uma nova aba (Sheet) é criada na planilha Excel com o mesmo nome da tabela.
5.  **Escrita de Cabeçalhos**: Os nomes das colunas da tabela do Access são escritos na primeira linha da aba correspondente.
6.  **Escrita de Dados**: O código itera sobre cada linha da tabela do Access e escreve os dados nas células correspondentes da planilha, realizando a conversão de tipos de dados quando necessário.
7.  **Salvamento do Arquivo**: Após processar todas as tabelas, a planilha é salva no arquivo de saída especificado.

## 4. Análise de Segurança e Correções

Uma análise de segurança foi realizada no projeto, e as seguintes vulnerabilidades foram identificadas e corrigidas:

### 4.1. Vulnerabilidade de Dependência Transitiva (CVE-2025-48924)

-   **Descrição**: A dependência `jackcess:4.0.8` trazia consigo uma dependência transitiva para `commons-lang3:3.10`, que é vulnerável a um ataque de Negação de Serviço (Denial of Service) através de recursão não controlada na classe `ClassUtils`.
-   **Correção**: A versão da dependência `commons-lang3` foi forçada para `3.18.0` (uma versão segura) no arquivo `pom.xml`, sobrepondo a versão transitiva vulnerável.

### 4.2. Path Traversal

-   **Descrição**: A aplicação aceitava caminhos de arquivos de entrada e saída diretamente da linha de comando sem validação, permitindo que um usuário mal-intencionado lesse ou escrevesse arquivos em diretórios arbitrários no sistema de arquivos (ex: `../../etc/passwd`).
-   **Correção**: Foi adicionada uma verificação de segurança nos métodos `open()` e `convert()` da classe `Access2Excel`. A verificação garante que o caminho canônico do arquivo de entrada e de saída esteja contido dentro do diretório de trabalho da aplicação, lançando uma `SecurityException` caso contrário.

### 4.3. Negação de Serviço (Denial of Service) no Parser de OLE

-   **Descrição**: A classe `OleHeaderParser` lia metadados de objetos OLE de um array de bytes sem validar os tamanhos e deslocamentos. Um objeto OLE malformado em um arquivo de banco de dados poderia causar uma `ArrayIndexOutOfBoundsException`, resultando na interrupção abrupta da aplicação.
-   **Correção**: O construtor da classe `OleHeaderParser` foi robustecido para validar o cabeçalho OLE. Ele agora verifica o tamanho do array e a validade dos deslocamentos e tamanhos lidos. Em caso de dados inválidos, ele define um nome de objeto padrão ("Invalid OLE Data") em vez de lançar uma exceção, permitindo que a conversão continue sem interrupções.
