---
title: Conceitos básicos de script para scripts do Office no Excel na Web
description: Informações de modelo de objeto e outros conceitos básicos para aprender antes de escrever scripts do Office.
ms.date: 01/27/2020
localization_priority: Priority
ms.openlocfilehash: 5a709c16e23c00ffc7ee7949a3cb11459dc2d530
ms.sourcegitcommit: b075eed5a6f275274fbbf6d62633219eac416f26
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/10/2020
ms.locfileid: "42700050"
---
# <a name="scripting-fundamentals-for-office-scripts-in-excel-on-the-web-preview"></a>Conceitos básicos de script para scripts do Office no Excel na Web (visualização)

Este artigo apresentará os aspectos técnicos dos scripts do Office. Você aprenderá como os objetos do Excel trabalham juntos e como o editor de código sincroniza com uma pasta de trabalho.

[!INCLUDE [Preview note](../includes/preview-note.md)]

## <a name="object-model"></a>Modelo de objetos

Para entender as APIs do Excel, você deve entender como os componentes de uma pasta de trabalho estão relacionados uns com os outros.

- Uma **pasta de trabalho** contém uma ou mais **planilhas**.
- Uma **planilha** dá acesso a células por meio de objetos **Range** .
- Um **intervalo** representa um grupo de células contíguas.
- Os **intervalos** são usados para criar e colocar **tabelas**, **gráficos**, **formas**e outros objetos de organização ou de visualização de dados.
- Uma **planilha** contém coleções desses objetos de dados que estão presentes na planilha individual.
- As **pastas de trabalho** contêm coleções de alguns desses objetos de dados (como **tabelas**) para a **pasta de trabalho**inteira.

### <a name="ranges"></a>Intervalos

Um intervalo é um grupo de células contíguas na pasta de trabalho. Os scripts normalmente usam notação de estilo a1 (por exemplo, **B3** para a célula única na linha **B** e coluna **3** ou **C2: F4** para as células das linhas de **C** a **F** e colunas **2** a **4**) para definir intervalos.

Os intervalos têm três propriedades principais `values`: `formulas`, e `format`. Essas propriedades obtêm ou definem os valores de célula, as fórmulas a serem avaliadas e a formatação visual das células.

#### <a name="range-sample"></a>Amostra de intervalo

O exemplo a seguir mostra como criar registros de vendas. Esse script usa `Range` objetos para definir os valores, fórmulas e formatos.

```TypeScript
async function main(context: Excel.RequestContext) {
  // Get the active worksheet.
  let sheet = context.workbook.worksheets.getActiveWorksheet();

  // Create the headers and format them to stand out.
  let headers = [
    ["Product", "Quantity", "Unit Price", "Totals"]
  ];
  let headerRange = sheet.getRange("B2:E2");
  headerRange.values = headers;
  headerRange.format.fill.color = "#4472C4";
  headerRange.format.font.color = "white";

  // Create the product data rows.
  let productData = [
    ["Almonds", 6, 7.5],
    ["Coffee", 20, 34.5],
    ["Chocolate", 10, 9.56],
  ];
  let dataRange = sheet.getRange("B3:D5");
  dataRange.values = productData;

  // Create the formulas to total the amounts sold.
  let totalFormulas = [
    ["=C3 * D3"],
    ["=C4 * D4"],
    ["=C5 * D5"],
    ["=SUM(E3:E5)"]
  ];
  let totalRange = sheet.getRange("E3:E6");
  totalRange.formulas = totalFormulas;
  totalRange.format.font.bold = true;

  // Display the totals as US dollar amounts.
  totalRange.numberFormat = [["$0.00"]];
}
```

A execução desse script cria os seguintes dados na planilha atual:

![Um registro de vendas mostrando linhas de valor, uma coluna de fórmula e cabeçalhos formatados.](../images/range-sample.png)

### <a name="charts-tables-and-other-data-objects"></a>Gráficos, tabelas e outros objetos de dados

Scripts podem criar e manipular as estruturas de dados e as visualizações no Excel. Tabelas e gráficos são dois dos objetos usados com mais frequência, mas as APIs dão suporte a tabelas dinâmicas, formas, imagens e muito mais.

#### <a name="creating-a-table"></a>Criar uma tabela

Criar tabelas usando intervalos preenchidos por dados. Controles de formatação e tabela (como filtros) são automaticamente aplicados ao intervalo.

O script a seguir cria uma tabela usando os intervalos do exemplo anterior.

```TypeScript
async function main(context: Excel.RequestContext) {
   let sheet = context.workbook.worksheets.getActiveWorksheet();
   sheet.tables.add("B2:E5", true);
}
```

A execução desse script na planilha com os dados anteriores cria a seguinte tabela:

![Uma tabela criada a partir do registro de vendas anterior.](../images/table-sample.png)

#### <a name="creating-a-chart"></a>Criar um gráfico

Criar gráficos para visualizar os dados em um intervalo. Os scripts permitem dezenas de variedades de gráficos, que podem ser personalizados para atender às suas necessidades.

O script a seguir cria um gráfico de colunas simples para três itens e o coloca 100 pixels abaixo da parte superior da planilha.

```TypeScript
async function main(context: Excel.RequestContext) {
  let sheet = context.workbook.worksheets.getActiveWorksheet();
  let chart = sheet.charts.add(Excel.ChartType.columnStacked, sheet.getRange("B3:C5"));
  chart.top = 100;
}
```

A execução desse script na planilha com a tabela anterior cria o seguinte gráfico:

![Um gráfico de colunas mostrando as quantidades de três itens do registro de vendas anterior.](../images/chart-sample.png)

### <a name="further-reading-on-the-object-model"></a>Leitura adicional no modelo de objeto

A [documentação de referência da API de scripts do Office](/javascript/api/office-scripts/overview) é uma lista abrangente dos objetos usados em scripts do Office. Lá, você pode usar o Sumário para navegar para qualquer classe sobre a qual você gostaria de saber mais. A seguir estão várias páginas exibidas com frequência.

- [Chart](/javascript/api/office-scripts/excel/excel.chart)
- [Comment](/javascript/api/office-scripts/excel/excel.comment)
- [PivotTable](/javascript/api/office-scripts/excel/excel.pivottable)
- [Range](/javascript/api/office-scripts/excel/excel.range)
- [RangeFormat](/javascript/api/office-scripts/excel/excel.rangeformat)
- [Shape](/javascript/api/office-scripts/excel/excel.shape)
- [Table](/javascript/api/office-scripts/excel/excel.table)
- [Workbook](/javascript/api/office-scripts/excel/excel.workbook)
- [Worksheet](/javascript/api/office-scripts/excel/excel.worksheet)

## <a name="main-function"></a>`main`AllFunctions

Cada script do Office deve conter `main` uma função com a assinatura a seguir, `Excel.RequestContext` incluindo a definição de tipo:

```TypeScript
async function main(context: Excel.RequestContext) {
    // Your Excel Script
}
```

O código dentro da `main` função é executado quando o script é executado. `main`pode chamar outras funções no seu script, mas o código que não está contido em uma função não será executado.

## <a name="context"></a>Contexto

A `main` função aceita um `Excel.RequestContext` parâmetro, chamado `context`. Considere `context` como ponte entre o script e a pasta de trabalho. O script acessa a pasta de trabalho com `context` o objeto e usa `context` -o para enviar dados de volta e para trás.

O `context` objeto é necessário porque o script e o Excel estão sendo executados em diferentes processos e locais. O script precisará fazer alterações ou consultar dados da pasta de trabalho na nuvem. O `context` objeto gerencia essas transações.

## <a name="sync-and-load"></a>Sincronizar e carregar

Como o script e a pasta de trabalho são executados em locais diferentes, qualquer transferência de dados entre os dois leva tempo. Para melhorar o desempenho do script, os comandos são enfileirados até que o `sync` script chame explicitamente a operação para sincronizar o script e a pasta de trabalho. O script pode funcionar independentemente até que seja necessário fazer um dos seguintes:

- Ler dados da pasta de trabalho (seguindo `load` uma operação).
- Gravar dados na pasta de trabalho (geralmente porque o script foi concluído).

A imagem a seguir mostra um fluxo de controle de exemplo entre o script e a pasta de trabalho:

![Um diagrama mostrando as operações de leitura e gravação indo para a pasta de trabalho do script.](../images/load-sync.png)

### <a name="sync"></a>Sincronizar

Sempre que o script precisar ler dados de ou gravar dados na pasta de trabalho, chame `RequestContext.sync` o método conforme mostrado aqui:

```TypeScript
await context.sync();
```

> [!NOTE]
> `context.sync()`é chamado implicitamente quando um script é encerrado.

Após a `sync` conclusão da operação, a pasta de trabalho é atualizada para refletir as operações de gravação especificadas pelo script. Uma operação de gravação está definindo qualquer propriedade em um objeto do Excel ( `range.format.fill.color = "red"`por exemplo,) ou chamando um método que altera uma propriedade ( `range.format.autoFitColumns()`por exemplo,). A `sync` operação também lê qualquer valor da pasta de trabalho que o script solicitou usando `load` uma operação (conforme discutido na próxima seção).

Sincronizar seu script com a pasta de trabalho pode levar tempo, dependendo da sua rede. Você deve minimizar o número de `sync` chamadas para ajudar seu script a ser executado rapidamente.  

### <a name="load"></a>Carregar

Um script deve carregar dados da pasta de trabalho antes de lê-lo. No entanto, freqüentemente carregar dados de toda a pasta de trabalho reduziria imensamente a velocidade do script. Em vez disso `load` , o método permite que o seu estado de script especificamente que os dados devem ser recuperados da pasta de trabalho.

O `load` método está disponível em cada objeto do Excel. O script deve carregar as propriedades de um objeto antes de poder lê-las. Se isso não for feito, ocorrerá um erro.

Os exemplos a seguir usam `Range` um objeto para mostrar as três maneiras `load` como o método pode ser usado para carregar dados.

|Intent |Comando de exemplo | Efeito |
|:--|:--|:--|
|Carregar uma propriedade |`myRange.load("values");` | Carrega uma única propriedade, neste caso, a matriz bidimensional de valores neste intervalo. |
|Carregar várias propriedades |`myRange.load("values, rowCount, columnCount");`| Carrega todas as propriedades de uma lista delimitada por vírgulas, neste exemplo, os valores, a contagem de linhas e a contagem de colunas. |
|Carregar tudo | `myRange.load();`|Carrega todas as propriedades no intervalo. Essa não é uma solução recomendada, já que ela tornará mais lento o script, obtendo dados desnecessários. Você só deve usar isso ao testar o script ou se precisar de todas as propriedades do objeto. |

O script deve chamar `context.sync()` antes de ler qualquer valor carregado.

```TypeScript
let range = selectedSheet.getRange("A1:B3");
range.load ("rowCount"); // Load the property.
await context.sync(); // Synchronize with the workbook to get the property.
console.log(range.rowCount); // Read and log the property value (3).
```

Você também pode carregar Propriedades em uma coleção inteira. Cada objeto de coleção tem `items` uma propriedade que é uma matriz que contém os objetos dessa coleção. Usando `items` como o início de uma chamada hierárquica (`items\myProperty`) para `load` carregar as propriedades especificadas em cada um desses itens. O exemplo a seguir carrega `resolved` a propriedade em `Comment` cada objeto no `CommentCollection` objeto de uma planilha.

```TypeScript
let comments = selectedSheet.comments;
comments.load("items/resolved"); // Load the `resolved` property from every comment in this collection.
await context.sync(); // Synchronize with the workbook to get the properties.
```

> [!TIP]
> Para saber mais sobre como trabalhar com coleções em scripts do Office, consulte a [seção matriz do artigo usando objetos JavaScript incorporados no Office scripts](javascript-objects.md#array) .

## <a name="see-also"></a>Confira também

- [Gravar, editar e criar scripts do Office no Excel na Web](../tutorials/excel-tutorial.md)
- [Ler dados de pasta de trabalho com scripts do Office no Excel na Web](../tutorials/excel-read-tutorial.md)
- [Referência da API de scripts do Office](/javascript/api/office-scripts/overview)
- [Usando objetos JavaScript internos em scripts do Office](javascript-objects.md)
