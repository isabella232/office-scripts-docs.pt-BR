---
title: Scripts de exemplo para scripts do Office no Excel na Web
description: Uma coleção de exemplos de código para usar com scripts do Office no Excel na Web.
ms.date: 04/06/2020
localization_priority: Normal
ms.openlocfilehash: abf6b87b63ad027cca8ee5c947b687f54815409c
ms.sourcegitcommit: 0b2232c4c228b14d501edb8bb489fe0e84748b42
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/08/2020
ms.locfileid: "43191004"
---
# <a name="sample-scripts-for-office-scripts-in-excel-on-the-web-preview"></a><span data-ttu-id="2dab5-103">Scripts de exemplo para scripts do Office no Excel na Web (visualização)</span><span class="sxs-lookup"><span data-stu-id="2dab5-103">Sample scripts for Office Scripts in Excel on the web (preview)</span></span>

<span data-ttu-id="2dab5-104">Os exemplos a seguir são scripts simples para você experimentar em suas próprias pastas de trabalho.</span><span class="sxs-lookup"><span data-stu-id="2dab5-104">The following samples are simple scripts for you to try on your own workbooks.</span></span> <span data-ttu-id="2dab5-105">Para usá-los no Excel na Web:</span><span class="sxs-lookup"><span data-stu-id="2dab5-105">To use them in Excel on the web:</span></span>

1. <span data-ttu-id="2dab5-106">Abra a guia **Automação**.</span><span class="sxs-lookup"><span data-stu-id="2dab5-106">Open the **Automate** tab.</span></span>
2. <span data-ttu-id="2dab5-107">Pressione **Editor de código**.</span><span class="sxs-lookup"><span data-stu-id="2dab5-107">Press **Code Editor**.</span></span>
3. <span data-ttu-id="2dab5-108">Pressione **novo script** no painel de tarefas do editor de código.</span><span class="sxs-lookup"><span data-stu-id="2dab5-108">Press **New Script** in the Code Editor's task pane.</span></span>
4. <span data-ttu-id="2dab5-109">Substitua todo o script pelo exemplo de sua escolha.</span><span class="sxs-lookup"><span data-stu-id="2dab5-109">Replace the entire script with the sample of your choice.</span></span>
5. <span data-ttu-id="2dab5-110">Pressione **executar** no painel de tarefas do editor de código.</span><span class="sxs-lookup"><span data-stu-id="2dab5-110">Press **Run** in the Code Editor's task pane.</span></span>

[!INCLUDE [Preview note](../includes/preview-note.md)]

## <a name="scripting-basics"></a><span data-ttu-id="2dab5-111">Noções básicas sobre scripts</span><span class="sxs-lookup"><span data-stu-id="2dab5-111">Scripting basics</span></span>

<span data-ttu-id="2dab5-112">Estes exemplos demonstram blocos de construção fundamentais para scripts do Office.</span><span class="sxs-lookup"><span data-stu-id="2dab5-112">These samples demonstrate fundamental building blocks for Office Scripts.</span></span> <span data-ttu-id="2dab5-113">Adicione-os aos seus scripts para estender sua solução e resolver problemas comuns.</span><span class="sxs-lookup"><span data-stu-id="2dab5-113">Add these to your scripts to extend your solution and solve common problems.</span></span>

### <a name="read-and-log-one-cell"></a><span data-ttu-id="2dab5-114">Ler e registrar uma célula</span><span class="sxs-lookup"><span data-stu-id="2dab5-114">Read and log one cell</span></span>

<span data-ttu-id="2dab5-115">Este exemplo lê o valor de **a1** e o imprime no console.</span><span class="sxs-lookup"><span data-stu-id="2dab5-115">This sample reads the value of **A1** and prints it to the console.</span></span>

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

### <a name="work-with-dates"></a><span data-ttu-id="2dab5-116">Trabalhar com datas</span><span class="sxs-lookup"><span data-stu-id="2dab5-116">Work with dates</span></span>

<span data-ttu-id="2dab5-117">Os exemplos nesta seção mostram como usar o objeto JavaScript [Date](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/date) .</span><span class="sxs-lookup"><span data-stu-id="2dab5-117">The samples in this section show how to use the JavaScript [Date](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/date) object.</span></span>

<span data-ttu-id="2dab5-118">O exemplo a seguir obtém a data e hora atuais e, em seguida, grava esses valores em duas células da planilha ativa.</span><span class="sxs-lookup"><span data-stu-id="2dab5-118">The following sample gets the current date and time and then writes those values to two cells in the active worksheet.</span></span>

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

<span data-ttu-id="2dab5-119">A próxima amostra lê uma data que é armazenada no Excel e a converte para um objeto de data JavaScript.</span><span class="sxs-lookup"><span data-stu-id="2dab5-119">The next sample reads a date that's stored in Excel and translates it to a JavaScript Date object.</span></span> <span data-ttu-id="2dab5-120">Ele usa o [número de série numérico da data](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) como entrada para a data JavaScript.</span><span class="sxs-lookup"><span data-stu-id="2dab5-120">It uses the [date's numeric serial number](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) as input for the JavaScript Date.</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
  // Read a date at cell A1 from Excel.
  let dateRange = context.workbook.worksheets.getActiveWorksheet().getRange("A1");
  dateRange.load("values");
  await context.sync();

  // Convert the Excel date to a JavaScript Date object.
  let excelDateValue = dateRange.values[0][0];
  let javaScriptDate = new Date(Math.round((excelDateValue - 25569) * 86400 * 1000));
  console.log(javaScriptDate);
}
```

## <a name="display-data"></a><span data-ttu-id="2dab5-121">Exibir dados</span><span class="sxs-lookup"><span data-stu-id="2dab5-121">Display data</span></span>

<span data-ttu-id="2dab5-122">Estes exemplos demonstram como trabalhar com dados de planilha e fornecer aos usuários uma melhor visualização ou organização.</span><span class="sxs-lookup"><span data-stu-id="2dab5-122">These samples demonstrate how to work with worksheet data and provide users with a better view or organization.</span></span>

### <a name="apply-conditional-formatting"></a><span data-ttu-id="2dab5-123">Aplicar formatação condicional</span><span class="sxs-lookup"><span data-stu-id="2dab5-123">Apply conditional formatting</span></span>

<span data-ttu-id="2dab5-124">Este exemplo aplica formatação condicional ao intervalo atualmente usado na planilha.</span><span class="sxs-lookup"><span data-stu-id="2dab5-124">This sample applies conditional formatting to the currently used range in the worksheet.</span></span> <span data-ttu-id="2dab5-125">A formatação condicional é um preenchimento verde para os primeiros 10% dos valores.</span><span class="sxs-lookup"><span data-stu-id="2dab5-125">The conditional formatting is a green fill for the top 10% of values.</span></span>

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

### <a name="create-a-sorted-table"></a><span data-ttu-id="2dab5-126">Criar uma tabela classificada</span><span class="sxs-lookup"><span data-stu-id="2dab5-126">Create a sorted table</span></span>

<span data-ttu-id="2dab5-127">Este exemplo cria uma tabela a partir do intervalo usado da planilha atual e a classifica com base na primeira coluna.</span><span class="sxs-lookup"><span data-stu-id="2dab5-127">This sample creates a table from the current worksheet's used range, then sorts it based on the first column.</span></span>

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

## <a name="collaboration"></a><span data-ttu-id="2dab5-128">Colaboração</span><span class="sxs-lookup"><span data-stu-id="2dab5-128">Collaboration</span></span>

<span data-ttu-id="2dab5-129">Estes exemplos demonstram como trabalhar com recursos relacionados à colaboração do Excel, como comentários.</span><span class="sxs-lookup"><span data-stu-id="2dab5-129">These samples demonstrate how to work with collaboration-related features of Excel, such as comments.</span></span>

### <a name="delete-resolved-comments"></a><span data-ttu-id="2dab5-130">Excluir comentários resolvidos</span><span class="sxs-lookup"><span data-stu-id="2dab5-130">Delete resolved comments</span></span>

<span data-ttu-id="2dab5-131">Este exemplo exclui todos os comentários resolvidos da planilha atual.</span><span class="sxs-lookup"><span data-stu-id="2dab5-131">This sample deletes all resolved comments from the current worksheet.</span></span>

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

## <a name="scenario-samples"></a><span data-ttu-id="2dab5-132">Exemplos de cenário</span><span class="sxs-lookup"><span data-stu-id="2dab5-132">Scenario samples</span></span>

<span data-ttu-id="2dab5-133">Para obter exemplos de soluções maiores e reais, visite [exemplos de cenários de scripts do Office](scenarios/sample-scenario-overview.md).</span><span class="sxs-lookup"><span data-stu-id="2dab5-133">For samples showcasing larger, real-world solutions, visit [Sample scenarios for Office Scripts](scenarios/sample-scenario-overview.md).</span></span>

## <a name="suggest-new-samples"></a><span data-ttu-id="2dab5-134">Sugerir novos exemplos</span><span class="sxs-lookup"><span data-stu-id="2dab5-134">Suggest new samples</span></span>

<span data-ttu-id="2dab5-135">Boas-vindas de sugestões para novos exemplos.</span><span class="sxs-lookup"><span data-stu-id="2dab5-135">We welcome suggestions for new samples.</span></span> <span data-ttu-id="2dab5-136">Se houver um cenário comum que ajudaria outros desenvolvedores de scripts, diga-nos na seção de comentários abaixo.</span><span class="sxs-lookup"><span data-stu-id="2dab5-136">If there is a common scenario that would help other script developers, please tell us in the feedback section below.</span></span>
