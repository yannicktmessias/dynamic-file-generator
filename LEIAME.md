
# Gerador Dinâmico de Arquivos PPTX

Possibilita gerar arquivos do PowerPoint (.pptx) com dados dinâmicos.


## Como utilizar

Para utilizar o gerador, são necessários dois arquivos:
 - Arquivo modelo em formato .pptx
 - Arquivo de dados em formato .csv ou .txt

Durante a execução do programa, basta selecionar os arquivos indicados acima e a pasta de destino, onde serão salvos os arquivos gerados dinamicamente.

O programa também permite exportar automaticamente os arquivos para o formato PDF, basta selecionar a opção "Gerar arquivos PDF".

### Modelo
O modelo é um arquivo em formato PowerPoint (.pptx) com espaços reservados (_placeholders_) para os dados variáveis. O único requisito para o modelo é que os _placeholders_ possuam um nome em formato de marcação (_tag_), ou seja, entre os símbolos '<' e '>'.

#### Exemplo:

```
<nome-completo>
```
No modelo, utilize para os _placeholders_ a mesma formatação de texto desejada para os dados variáveis nos arquivos gerados.

### Arquivo de dados
Os dados variáveis são fornecidos a partir de um arquivo de planilha em formato CSV (.csv). A primeira linha do arquivo deve conter os nomes dos _placeholders_ separados por ',' (vírgula) ou ';' (ponto e vírgula). As demais linhas devem conter os dados que serão inseridos nos _placeholders_, na mesma ordem em que as respectivas variáveis aparecem na primeira linha.

Os arquivos gerados são nomeados com os valores da primeira variável de cada linha.

#### Exemplo:

```csv
nome-completo; idade
João Mendes; 32
Maria Santos; 25
```
Arquivos em formato CSV também podem ser gerados a partir de planilhas do Excel.

## Autor

- @yannicktmessias

