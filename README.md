# Extração de Dados e Geração de Planilha Excel

Este projeto consiste em um script Python que extrai informações estruturadas de um arquivo de texto e as salva em uma planilha Excel.

## Requisitos

Antes de executar o script, certifique-se de ter instalado as dependências necessárias. Você pode instalá-las usando:

```bash
pip install openpyxl pytz
```

## Como funciona

1. O script lê um arquivo de texto chamado `consulta.txt`.
2. Ele extrai informações estruturadas como nome, código, quantidade, unidade, valor unitário e valor total usando expressões regulares.
3. Os dados extraídos são salvos em um arquivo Excel (`.xlsx`) com um timestamp no nome para diferenciar versões.

## Como usar

1. Certifique-se de que o arquivo `consulta.txt` está no mesmo diretório do script.
2. Execute o script Python:

```bash
python script.py
```

3. O arquivo gerado será salvo no formato:

```
dados_extraidos_DD-MM-YYYY_HH-MM-SS.xlsx
```

onde `DD-MM-YYYY_HH-MM-SS` representa a data e hora da geração.

## Estrutura do Arquivo de Entrada (`consulta.txt`)

O arquivo de entrada deve conter dados formatados da seguinte forma:

```
Produto X (Código: 12345)
Qtde.: 10 UN: KG Vl. Unit.: 25,50
255,00

Produto Y (Código: 67890)
Qtde.: 5 UN: UN Vl. Unit.: 12,00
60,00
```

Cada item deve conter 3 linhas seguidas, separadas por uma linha em branco antes do próximo item.

## Estrutura do Arquivo de Saída (`.xlsx`)

O arquivo Excel conterá uma aba chamada "Dados Extraídos" com as colunas:

- **Nome**: Nome do produto
- **Código**: Código do produto
- **Quantidade**: Quantidade do produto
- **Unidade**: Unidade de medida
- **Valor Unitário**: Preço por unidade
- **Valor Total**: Valor total calculado

## Possíveis Erros

- Se o arquivo `consulta.txt` não for encontrado, uma mensagem de erro será exibida.
- Se os dados não estiverem no formato esperado, eles serão ignorados.

## Possivel melhoria
- Consultar direto pelo numero da nota fiscal e gerar o xlsx
