---
title: Scripts de exemplo para scripts do Office no Excel na Web
description: Uma coleção de exemplos de código para usar com scripts do Office no Excel na Web.
ms.date: 02/19/2020
localization_priority: Normal
ms.openlocfilehash: abb4064dfde8b644035e725832e481e6463e979e
ms.sourcegitcommit: b075eed5a6f275274fbbf6d62633219eac416f26
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/10/2020
ms.locfileid: "42700072"
---
# <a name="sample-scripts-for-office-scripts-in-excel-on-the-web-preview"></a>Scripts de exemplo para scripts do Office no Excel na Web (visualização)

Os exemplos a seguir são scripts simples para você experimentar em suas próprias pastas de trabalho. Para usá-los no Excel na Web:

1. Abra a guia **automatizar** .
2. Pressione **Editor de código**.
3. Pressione **novo script** no painel de tarefas do editor de código.
4. Substitua todo o script pelo exemplo de sua escolha.
5. Pressione **executar** no painel de tarefas do editor de código.

[!INCLUDE [Preview note](../includes/preview-note.md)]

## <a name="scripting-basics"></a>Noções básicas sobre scripts

Estes exemplos demonstram blocos de construção fundamentais para scripts do Office. Adicione-os aos seus scripts para estender sua solução e resolver problemas comuns.

### <a name="read-and-log-one-cell"></a>Ler e registrar uma célula

Este exemplo lê o valor de **a1** e o imprime no console.

``` TypeScript
async function main(context: Excel.RequestContext) {
  // Get the current worksheet.
  let selectedSheet = context.workbook.worksheets.getActiveWorksheet();

  // Get the value of cell A1.
  let range = selectedSheet.getRange("A1");
  range.load("values");
  await context.sync();

  // Print the value of A1.
  console.log(range.values);
}
```

### <a name="work-with-dates"></a>Trabalhar com datas

Este exemplo usa o objeto de [Data](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/date) JavaScript para obter a data e hora atuais e, em seguida, grava esses valores em duas células da planilha ativa.

```TypeScript
async function main(context: Excel.RequestContext) {
  // Get the cells at A1 and B1.
  let dateRange = context.workbook.worksheets.getActiveWorksheet().getRange("A1");
  let timeRange = context.workbook.worksheets.getActiveWorksheet().getRange("B1");

  // Get the current date and time with the JavaScript Date object.
  let date = new Date(Date.now());

  // Add the date string to A1.
  dateRange.values = [[date.toLocaleDateString()]];
  
  // Add the time string to B1.
  timeRange.values = [[date.toLocaleTimeString()]];
}
```

## <a name="display-data"></a>Exibir dados

Estes exemplos demonstram como trabalhar com dados de planilha e fornecer aos usuários uma melhor visualização ou organização.

### <a name="apply-conditional-formatting"></a>Aplicar formatação condicional

Este exemplo aplica formatação condicional ao intervalo atualmente usado na planilha. A formatação condicional é um preenchimento verde para os primeiros 10% dos valores.

```TypeScript
async function main(context: Excel.RequestContext) {
  // Get the current worksheet.
  let selectedSheet = context.workbook.worksheets.getActiveWorksheet();

  // Get the used range in the worksheet.
  let range = selectedSheet.getUsedRange();

  // Set the fill color to green for the top 10% of values in the range.
  let conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.topBottom);
  conditionalFormat.topBottom.format.fill.color = "green";
  conditionalFormat.topBottom.rule = {
    rank: 10, // The percentage threshold.
    type: Excel.ConditionalTopBottomCriterionType.topPercent // The type of the top/bottom condition.
  };
}
```

### <a name="create-a-sorted-table"></a>Criar uma tabela classificada

Este exemplo cria uma tabela a partir do intervalo usado da planilha atual e a classifica com base na primeira coluna.

```TypeScript
async function main(context: Excel.RequestContext) {
  // Get the current worksheet.
  let selectedSheet = context.workbook.worksheets.getActiveWorksheet();

  // Create a table with the used cells.
  let usedRange = selectedSheet.getUsedRange();
  let newTable = selectedSheet.tables.add(usedRange, true);

  // Sort the table using the first column.
  newTable.sort.apply([{ key: 0, ascending: true }]);
}
```

## <a name="collaboration"></a>Colaboração

Estes exemplos demonstram como trabalhar com recursos relacionados à colaboração do Excel, como comentários.

### <a name="delete-resolved-comments"></a>Excluir comentários resolvidos

Este exemplo exclui todos os comentários resolvidos da planilha atual.

```TypeScript
async function main(context: Excel.RequestContext) {
  // Get the current worksheet.
  let selectedSheet = context.workbook.worksheets.getActiveWorksheet();

  // Get the comments on this worksheet.
  let comments = selectedSheet.comments;
  comments.load("items/resolved");
  await context.sync();

  // Delete the resolved comments.
  comments.items.forEach((comment) => {
      if (comment.resolved) {
          comment.delete();
      }
  });
}
```

## <a name="scenario-samples"></a>Exemplos de cenário

Para obter exemplos de soluções maiores e reais, visite [exemplos de cenários de scripts do Office](scenarios/sample-scenario-overview.md).

## <a name="suggest-new-samples"></a>Sugerir novos exemplos

Boas-vindas de sugestões para novos exemplos. Se houver um cenário comum que ajudaria outros desenvolvedores de scripts, diga-nos na seção de comentários abaixo.
