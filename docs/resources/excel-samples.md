---
title: Scripts de exemplo para scripts do Office no Excel na Web
description: Uma coleção de exemplos de código para usar com scripts do Office no Excel na Web.
ms.date: 06/18/2020
localization_priority: Normal
ms.openlocfilehash: bfa6679595e6e28cc5d2ae3e3e487fd3e77738aa
ms.sourcegitcommit: aec3c971c6640429f89b6bb99d2c95ea06725599
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/25/2020
ms.locfileid: "44878672"
---
# <a name="sample-scripts-for-office-scripts-in-excel-on-the-web-preview"></a><span data-ttu-id="d1779-103">Scripts de exemplo para scripts do Office no Excel na Web (visualização)</span><span class="sxs-lookup"><span data-stu-id="d1779-103">Sample scripts for Office Scripts in Excel on the web (preview)</span></span>

<span data-ttu-id="d1779-104">Os exemplos a seguir são scripts simples para você experimentar em suas próprias pastas de trabalho.</span><span class="sxs-lookup"><span data-stu-id="d1779-104">The following samples are simple scripts for you to try on your own workbooks.</span></span> <span data-ttu-id="d1779-105">Para usá-los no Excel na Web:</span><span class="sxs-lookup"><span data-stu-id="d1779-105">To use them in Excel on the web:</span></span>

1. <span data-ttu-id="d1779-106">Abra a guia **Automação**.</span><span class="sxs-lookup"><span data-stu-id="d1779-106">Open the **Automate** tab.</span></span>
2. <span data-ttu-id="d1779-107">Pressione **Editor de código**.</span><span class="sxs-lookup"><span data-stu-id="d1779-107">Press **Code Editor**.</span></span>
3. <span data-ttu-id="d1779-108">Pressione **novo script** no painel de tarefas do editor de código.</span><span class="sxs-lookup"><span data-stu-id="d1779-108">Press **New Script** in the Code Editor's task pane.</span></span>
4. <span data-ttu-id="d1779-109">Substitua todo o script pelo exemplo de sua escolha.</span><span class="sxs-lookup"><span data-stu-id="d1779-109">Replace the entire script with the sample of your choice.</span></span>
5. <span data-ttu-id="d1779-110">Pressione **executar** no painel de tarefas do editor de código.</span><span class="sxs-lookup"><span data-stu-id="d1779-110">Press **Run** in the Code Editor's task pane.</span></span>

[!INCLUDE [Preview note](../includes/preview-note.md)]

## <a name="scripting-basics"></a><span data-ttu-id="d1779-111">Noções básicas sobre scripts</span><span class="sxs-lookup"><span data-stu-id="d1779-111">Scripting basics</span></span>

<span data-ttu-id="d1779-112">Estes exemplos demonstram blocos de construção fundamentais para scripts do Office.</span><span class="sxs-lookup"><span data-stu-id="d1779-112">These samples demonstrate fundamental building blocks for Office Scripts.</span></span> <span data-ttu-id="d1779-113">Adicione-os aos seus scripts para estender sua solução e resolver problemas comuns.</span><span class="sxs-lookup"><span data-stu-id="d1779-113">Add these to your scripts to extend your solution and solve common problems.</span></span>

### <a name="read-and-log-one-cell"></a><span data-ttu-id="d1779-114">Ler e registrar uma célula</span><span class="sxs-lookup"><span data-stu-id="d1779-114">Read and log one cell</span></span>

<span data-ttu-id="d1779-115">Este exemplo lê o valor de **a1** e o imprime no console.</span><span class="sxs-lookup"><span data-stu-id="d1779-115">This sample reads the value of **A1** and prints it to the console.</span></span>

```typescript
function main(workbook: ExcelScript.Workbook) {
  // Get the current worksheet.
  let selectedSheet = workbook.getActiveWorksheet();

  // Get the value of cell A1.
  let range = selectedSheet.getRange("A1");
  
  // Print the value of A1.
  console.log(range.getValue());
}
```

### <a name="read-the-active-cell"></a><span data-ttu-id="d1779-116">Ler a célula ativa</span><span class="sxs-lookup"><span data-stu-id="d1779-116">Read the active cell</span></span>

<span data-ttu-id="d1779-117">Este script registra o valor da célula ativa atual.</span><span class="sxs-lookup"><span data-stu-id="d1779-117">This script logs the value of the current active cell.</span></span> <span data-ttu-id="d1779-118">Se várias células estiverem selecionadas, a célula superior à extrema esquerda será registrada.</span><span class="sxs-lookup"><span data-stu-id="d1779-118">If multiple cells are selected, the top-leftmost cell will be logged.</span></span>

```typescript
function main(workbook: ExcelScript.Workbook) {
  // Get the current active cell in the workbook.
  let cell = workbook.getActiveCell();

  // Log that cell's value.
  console.log(`The current cell's value is ${cell.getValue()}`);
}
```

### <a name="change-an-adjacent-cell"></a><span data-ttu-id="d1779-119">Alterar uma célula adjacente</span><span class="sxs-lookup"><span data-stu-id="d1779-119">Change an adjacent cell</span></span>

<span data-ttu-id="d1779-120">Esse script Obtém células adjacentes usando referências relativas.</span><span class="sxs-lookup"><span data-stu-id="d1779-120">This script gets adjacent cells using relative references.</span></span> <span data-ttu-id="d1779-121">Observe que, se a célula ativa estiver na linha superior, parte do script falhará, pois ela faz referência à célula acima da selecionada atualmente.</span><span class="sxs-lookup"><span data-stu-id="d1779-121">Note that if the active cell is on the top row, part of the script fails, because it references the cell above the currently selected one.</span></span>

```typescript
function main(workbook: ExcelScript.Workbook) {
  // Get the currently active cell in the workbook.
  let activeCell = workbook.getActiveCell();
  console.log(`The active cell's address is: ${activeCell.getAddress()}`);

  // Get the cell to the right of the active cell and set its value and color.
  let rightCell = activeCell.getOffsetRange(0,1);
  rightCell.setValue("Right cell");
  console.log(`The right cell's address is: ${rightCell.getAddress()}`);
  rightCell.getFormat().getFont().setColor("Magenta");
  rightCell.getFormat().getFill().setColor("Cyan");

  // Get the cell to the above of the active cell and set its value and color.
  // Note that this operation will fail if the active cell is in the top row.
  let aboveCell = activeCell.getOffsetRange(-1, 0);
  aboveCell.setValue("Above cell");
  console.log(`The above cell's address is: ${aboveCell.getAddress()}`);
  aboveCell.getFormat().getFont().setColor("White");
  aboveCell.getFormat().getFill().setColor("Black");
}
```

### <a name="change-all-adjacent-cells"></a><span data-ttu-id="d1779-122">Alterar todas as células adjacentes</span><span class="sxs-lookup"><span data-stu-id="d1779-122">Change all adjacent cells</span></span>

<span data-ttu-id="d1779-123">Esse script copia a formatação na célula ativa para as células vizinhas.</span><span class="sxs-lookup"><span data-stu-id="d1779-123">This script copies the formatting in the active cell to the neighboring cells.</span></span> <span data-ttu-id="d1779-124">Observe que esse script só funciona quando a célula ativa não está em uma borda da planilha.</span><span class="sxs-lookup"><span data-stu-id="d1779-124">Note that this script only works when the active cell isn't on an edge of the worksheet.</span></span>

```typescript
function main(workbook: ExcelScript.Workbook) {
  // Get the active cell.
  let activeCell = workbook.getActiveCell();

  // Get the cell that's one row above and one column to the left of the active cell.
  let cornerCell = activeCell.getOffsetRange(-1,-1);

  // Get a range that includes all the cells surrounding the active cell.
  let surroundingRange = cornerCell.getResizedRange(2, 2)

  // Copy the formatting from the active cell to the new range.
  surroundingRange.copyFrom(
    activeCell, /* The source range. */
    ExcelScript.RangeCopyType.formats /* What to copy. */
    );
}
```

### <a name="work-with-dates"></a><span data-ttu-id="d1779-125">Trabalhar com datas</span><span class="sxs-lookup"><span data-stu-id="d1779-125">Work with dates</span></span>

<span data-ttu-id="d1779-126">Os exemplos nesta seção mostram como usar o objeto JavaScript [Date](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/date) .</span><span class="sxs-lookup"><span data-stu-id="d1779-126">The samples in this section show how to use the JavaScript [Date](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/date) object.</span></span>

<span data-ttu-id="d1779-127">O exemplo a seguir obtém a data e hora atuais e, em seguida, grava esses valores em duas células da planilha ativa.</span><span class="sxs-lookup"><span data-stu-id="d1779-127">The following sample gets the current date and time and then writes those values to two cells in the active worksheet.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the cells at A1 and B1.
  let dateRange = workbook.getActiveWorksheet().getRange("A1");
  let timeRange = workbook.getActiveWorksheet().getRange("B1");

  // Get the current date and time with the JavaScript Date object.
  let date = new Date(Date.now());

  // Add the date string to A1.
  dateRange.setValue(date.toLocaleDateString());

  // Add the time string to B1.
  timeRange.setValue(date.toLocaleTimeString());
}
```

<span data-ttu-id="d1779-128">A próxima amostra lê uma data que é armazenada no Excel e a converte para um objeto de data JavaScript.</span><span class="sxs-lookup"><span data-stu-id="d1779-128">The next sample reads a date that's stored in Excel and translates it to a JavaScript Date object.</span></span> <span data-ttu-id="d1779-129">Ele usa o [número de série numérico da data](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) como entrada para a data JavaScript.</span><span class="sxs-lookup"><span data-stu-id="d1779-129">It uses the [date's numeric serial number](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) as input for the JavaScript Date.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Read a date at cell A1 from Excel.
  let dateRange = workbook.getActiveWorksheet().getRange("A1");

  // Convert the Excel date to a JavaScript Date object.
  let excelDateValue = dateRange.getValue();
  let javaScriptDate = new Date(Math.round((excelDateValue - 25569) * 86400 * 1000));
  console.log(javaScriptDate);
}
```

## <a name="display-data"></a><span data-ttu-id="d1779-130">Exibir dados</span><span class="sxs-lookup"><span data-stu-id="d1779-130">Display data</span></span>

<span data-ttu-id="d1779-131">Estes exemplos demonstram como trabalhar com dados de planilha e fornecer aos usuários uma melhor visualização ou organização.</span><span class="sxs-lookup"><span data-stu-id="d1779-131">These samples demonstrate how to work with worksheet data and provide users with a better view or organization.</span></span>

### <a name="apply-conditional-formatting"></a><span data-ttu-id="d1779-132">Aplicar formatação condicional</span><span class="sxs-lookup"><span data-stu-id="d1779-132">Apply conditional formatting</span></span>

<span data-ttu-id="d1779-133">Este exemplo aplica formatação condicional ao intervalo atualmente usado na planilha.</span><span class="sxs-lookup"><span data-stu-id="d1779-133">This sample applies conditional formatting to the currently used range in the worksheet.</span></span> <span data-ttu-id="d1779-134">A formatação condicional é um preenchimento verde para os primeiros 10% dos valores.</span><span class="sxs-lookup"><span data-stu-id="d1779-134">The conditional formatting is a green fill for the top 10% of values.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the current worksheet.
  let selectedSheet = workbook.getActiveWorksheet();

  // Get the used range in the worksheet.
  let range = selectedSheet.getUsedRange();

  // Set the fill color to green for the top 10% of values in the range.
  let conditionalFormat = range.addConditionalFormat(ExcelScript.ConditionalFormatType.topBottom)
  conditionalFormat.getTopBottom().getFormat().getFill().setColor("green");
  conditionalFormat.getTopBottom().setRule({
    rank: 10, // The percentage threshold.
    type: ExcelScript.ConditionalTopBottomCriterionType.topPercent // The type of the top/bottom condition.
  });
}
```

### <a name="create-a-sorted-table"></a><span data-ttu-id="d1779-135">Criar uma tabela classificada</span><span class="sxs-lookup"><span data-stu-id="d1779-135">Create a sorted table</span></span>

<span data-ttu-id="d1779-136">Este exemplo cria uma tabela a partir do intervalo usado da planilha atual e a classifica com base na primeira coluna.</span><span class="sxs-lookup"><span data-stu-id="d1779-136">This sample creates a table from the current worksheet's used range, then sorts it based on the first column.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the current worksheet.
  let selectedSheet = workbook.getActiveWorksheet();

  // Create a table with the used cells.
  let usedRange = selectedSheet.getUsedRange();
  let newTable = selectedSheet.addTable(usedRange, true);

  // Sort the table using the first column.
  newTable.getSort().apply([{ key: 0, ascending: true }]);
}
```

### <a name="log-the-grand-total-values-from-a-pivottable"></a><span data-ttu-id="d1779-137">Registrar os valores de "total geral" de uma tabela dinâmica</span><span class="sxs-lookup"><span data-stu-id="d1779-137">Log the "Grand Total" values from a PivotTable</span></span>

<span data-ttu-id="d1779-138">Este exemplo localiza a primeira tabela dinâmica na pasta de trabalho e registra os valores nas células "total geral" (conforme realçado em verde na imagem abaixo).</span><span class="sxs-lookup"><span data-stu-id="d1779-138">This sample finds the first PivotTable in the workbook and logs the values in the "Grand Total" cells (as highlighted in green in the image below).</span></span>

![Uma tabela de vendas de frutas com a linha de total geral realçada verde.](../images/sample-pivottable-grand-total-row.png)

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the first PivotTable in the workbook.
  let pivotTable = workbook.getPivotTables()[0];

  // Get the names of each data column in the PivotTable.
  let pivotColumnLabelRange = pivotTable.getLayout().getColumnLabelRange();

  // Get the range displaying the pivoted data.
  let pivotDataRange = pivotTable.getLayout().getRangeBetweenHeaderAndTotal();

  // Get the range with the "grand totals" for the PivotTable columns.
  let grandTotalRange = pivotDataRange.getLastRow();

  // Print each of the "Grand Totals" to the console.
  grandTotalRange.getValues()[0].forEach((column, columnIndex) => {
    console.log(`Grand total of ${pivotColumnLabelRange.getValues()[0][columnIndex]}: ${grandTotalRange.getValues()[0][columnIndex]}`);
    // Example log: "Grand total of Sum of Crates Sold Wholesale: 11000"
  });
}
```

## <a name="scenario-samples"></a><span data-ttu-id="d1779-140">Exemplos de cenário</span><span class="sxs-lookup"><span data-stu-id="d1779-140">Scenario samples</span></span>

<span data-ttu-id="d1779-141">Para obter exemplos de soluções maiores e reais, visite [exemplos de cenários de scripts do Office](scenarios/sample-scenario-overview.md).</span><span class="sxs-lookup"><span data-stu-id="d1779-141">For samples showcasing larger, real-world solutions, visit [Sample scenarios for Office Scripts](scenarios/sample-scenario-overview.md).</span></span>

## <a name="suggest-new-samples"></a><span data-ttu-id="d1779-142">Sugerir novos exemplos</span><span class="sxs-lookup"><span data-stu-id="d1779-142">Suggest new samples</span></span>

<span data-ttu-id="d1779-143">Boas-vindas de sugestões para novos exemplos.</span><span class="sxs-lookup"><span data-stu-id="d1779-143">We welcome suggestions for new samples.</span></span> <span data-ttu-id="d1779-144">Se houver um cenário comum que ajudaria outros desenvolvedores de scripts, diga-nos na seção de comentários abaixo.</span><span class="sxs-lookup"><span data-stu-id="d1779-144">If there is a common scenario that would help other script developers, please tell us in the feedback section below.</span></span>
