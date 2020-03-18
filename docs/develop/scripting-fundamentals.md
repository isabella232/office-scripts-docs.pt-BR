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
# <a name="scripting-fundamentals-for-office-scripts-in-excel-on-the-web-preview"></a><span data-ttu-id="37ced-103">Conceitos básicos de script para scripts do Office no Excel na Web (visualização)</span><span class="sxs-lookup"><span data-stu-id="37ced-103">Scripting fundamentals for Office Scripts in Excel on the web (preview)</span></span>

<span data-ttu-id="37ced-104">Este artigo apresentará os aspectos técnicos dos scripts do Office.</span><span class="sxs-lookup"><span data-stu-id="37ced-104">This article will introduce you to the technical aspects of Office Scripts.</span></span> <span data-ttu-id="37ced-105">Você aprenderá como os objetos do Excel trabalham juntos e como o editor de código sincroniza com uma pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="37ced-105">You'll learn how the Excel objects work together and how the Code Editor synchronizes with a workbook.</span></span>

[!INCLUDE [Preview note](../includes/preview-note.md)]

## <a name="object-model"></a><span data-ttu-id="37ced-106">Modelo de objetos</span><span class="sxs-lookup"><span data-stu-id="37ced-106">Object model</span></span>

<span data-ttu-id="37ced-107">Para entender as APIs do Excel, você deve entender como os componentes de uma pasta de trabalho estão relacionados uns com os outros.</span><span class="sxs-lookup"><span data-stu-id="37ced-107">To understand the Excel APIs, you must understand how the components of a workbook are related to one another.</span></span>

- <span data-ttu-id="37ced-108">Uma **pasta de trabalho** contém uma ou mais **planilhas**.</span><span class="sxs-lookup"><span data-stu-id="37ced-108">A **Workbook** contains one or more **Worksheets**.</span></span>
- <span data-ttu-id="37ced-109">Uma **planilha** dá acesso a células por meio de objetos **Range** .</span><span class="sxs-lookup"><span data-stu-id="37ced-109">A **Worksheet** gives access to cells through **Range** objects.</span></span>
- <span data-ttu-id="37ced-110">Um **intervalo** representa um grupo de células contíguas.</span><span class="sxs-lookup"><span data-stu-id="37ced-110">A **Range** represents a group of contiguous cells.</span></span>
- <span data-ttu-id="37ced-111">Os **intervalos** são usados para criar e colocar **tabelas**, **gráficos**, **formas**e outros objetos de organização ou de visualização de dados.</span><span class="sxs-lookup"><span data-stu-id="37ced-111">**Ranges** are used to create and place **Tables**, **Charts**, **Shapes**, and other data visualization or organization objects.</span></span>
- <span data-ttu-id="37ced-112">Uma **planilha** contém coleções desses objetos de dados que estão presentes na planilha individual.</span><span class="sxs-lookup"><span data-stu-id="37ced-112">A **Worksheet** contains collections of those data objects that are present in the individual sheet.</span></span>
- <span data-ttu-id="37ced-113">As **pastas de trabalho** contêm coleções de alguns desses objetos de dados (como **tabelas**) para a **pasta de trabalho**inteira.</span><span class="sxs-lookup"><span data-stu-id="37ced-113">**Workbooks** contain collections of some of those data objects (such as **Tables**) for the entire **Workbook**.</span></span>

### <a name="ranges"></a><span data-ttu-id="37ced-114">Intervalos</span><span class="sxs-lookup"><span data-stu-id="37ced-114">Ranges</span></span>

<span data-ttu-id="37ced-115">Um intervalo é um grupo de células contíguas na pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="37ced-115">A range is a group of contiguous cells in the workbook.</span></span> <span data-ttu-id="37ced-116">Os scripts normalmente usam notação de estilo a1 (por exemplo, **B3** para a célula única na linha **B** e coluna **3** ou **C2: F4** para as células das linhas de **C** a **F** e colunas **2** a **4**) para definir intervalos.</span><span class="sxs-lookup"><span data-stu-id="37ced-116">Scripts typically use A1-style notation (e.g. **B3** for the single cell in row **B** and column **3** or **C2:F4** for the cells from rows **C** through **F** and columns **2** through **4**) to define ranges.</span></span>

<span data-ttu-id="37ced-117">Os intervalos têm três propriedades principais `values`: `formulas`, e `format`.</span><span class="sxs-lookup"><span data-stu-id="37ced-117">Ranges have three core properties: `values`, `formulas`, and `format`.</span></span> <span data-ttu-id="37ced-118">Essas propriedades obtêm ou definem os valores de célula, as fórmulas a serem avaliadas e a formatação visual das células.</span><span class="sxs-lookup"><span data-stu-id="37ced-118">These properties get or set the cell values, formulas to be evaluated, and the visual formatting of the cells.</span></span>

#### <a name="range-sample"></a><span data-ttu-id="37ced-119">Amostra de intervalo</span><span class="sxs-lookup"><span data-stu-id="37ced-119">Range sample</span></span>

<span data-ttu-id="37ced-120">O exemplo a seguir mostra como criar registros de vendas.</span><span class="sxs-lookup"><span data-stu-id="37ced-120">The following sample shows how to create sales records.</span></span> <span data-ttu-id="37ced-121">Esse script usa `Range` objetos para definir os valores, fórmulas e formatos.</span><span class="sxs-lookup"><span data-stu-id="37ced-121">This script uses `Range` objects to set the values, formulas, and formats.</span></span>

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

<span data-ttu-id="37ced-122">A execução desse script cria os seguintes dados na planilha atual:</span><span class="sxs-lookup"><span data-stu-id="37ced-122">Running this script creates the following data in the current worksheet:</span></span>

![Um registro de vendas mostrando linhas de valor, uma coluna de fórmula e cabeçalhos formatados.](../images/range-sample.png)

### <a name="charts-tables-and-other-data-objects"></a><span data-ttu-id="37ced-124">Gráficos, tabelas e outros objetos de dados</span><span class="sxs-lookup"><span data-stu-id="37ced-124">Charts, tables, and other data objects</span></span>

<span data-ttu-id="37ced-125">Scripts podem criar e manipular as estruturas de dados e as visualizações no Excel.</span><span class="sxs-lookup"><span data-stu-id="37ced-125">Scripts can create and manipulate the data structures and visualizations within Excel.</span></span> <span data-ttu-id="37ced-126">Tabelas e gráficos são dois dos objetos usados com mais frequência, mas as APIs dão suporte a tabelas dinâmicas, formas, imagens e muito mais.</span><span class="sxs-lookup"><span data-stu-id="37ced-126">Tables and charts are two of the more commonly used objects, but the APIs support PivotTables, shapes, images, and more.</span></span>

#### <a name="creating-a-table"></a><span data-ttu-id="37ced-127">Criar uma tabela</span><span class="sxs-lookup"><span data-stu-id="37ced-127">Creating a table</span></span>

<span data-ttu-id="37ced-128">Criar tabelas usando intervalos preenchidos por dados.</span><span class="sxs-lookup"><span data-stu-id="37ced-128">Create tables by using data-filled ranges.</span></span> <span data-ttu-id="37ced-129">Controles de formatação e tabela (como filtros) são automaticamente aplicados ao intervalo.</span><span class="sxs-lookup"><span data-stu-id="37ced-129">Formatting and table controls (such as filters) are automatically applied to the range.</span></span>

<span data-ttu-id="37ced-130">O script a seguir cria uma tabela usando os intervalos do exemplo anterior.</span><span class="sxs-lookup"><span data-stu-id="37ced-130">The following script creates a table using the ranges from the previous sample.</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
   let sheet = context.workbook.worksheets.getActiveWorksheet();
   sheet.tables.add("B2:E5", true);
}
```

<span data-ttu-id="37ced-131">A execução desse script na planilha com os dados anteriores cria a seguinte tabela:</span><span class="sxs-lookup"><span data-stu-id="37ced-131">Running this script on the worksheet with the previous data creates the following table:</span></span>

![Uma tabela criada a partir do registro de vendas anterior.](../images/table-sample.png)

#### <a name="creating-a-chart"></a><span data-ttu-id="37ced-133">Criar um gráfico</span><span class="sxs-lookup"><span data-stu-id="37ced-133">Creating a chart</span></span>

<span data-ttu-id="37ced-134">Criar gráficos para visualizar os dados em um intervalo.</span><span class="sxs-lookup"><span data-stu-id="37ced-134">Create charts to visualize the data in a range.</span></span> <span data-ttu-id="37ced-135">Os scripts permitem dezenas de variedades de gráficos, que podem ser personalizados para atender às suas necessidades.</span><span class="sxs-lookup"><span data-stu-id="37ced-135">Scripts allow for dozens of chart varieties, each of which can be customized to suit your needs.</span></span>

<span data-ttu-id="37ced-136">O script a seguir cria um gráfico de colunas simples para três itens e o coloca 100 pixels abaixo da parte superior da planilha.</span><span class="sxs-lookup"><span data-stu-id="37ced-136">The following script creates a simple column chart for three items and places it 100 pixels below the top of the worksheet.</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
  let sheet = context.workbook.worksheets.getActiveWorksheet();
  let chart = sheet.charts.add(Excel.ChartType.columnStacked, sheet.getRange("B3:C5"));
  chart.top = 100;
}
```

<span data-ttu-id="37ced-137">A execução desse script na planilha com a tabela anterior cria o seguinte gráfico:</span><span class="sxs-lookup"><span data-stu-id="37ced-137">Running this script on the worksheet with the previous table creates the following chart:</span></span>

![Um gráfico de colunas mostrando as quantidades de três itens do registro de vendas anterior.](../images/chart-sample.png)

### <a name="further-reading-on-the-object-model"></a><span data-ttu-id="37ced-139">Leitura adicional no modelo de objeto</span><span class="sxs-lookup"><span data-stu-id="37ced-139">Further reading on the object model</span></span>

<span data-ttu-id="37ced-140">A [documentação de referência da API de scripts do Office](/javascript/api/office-scripts/overview) é uma lista abrangente dos objetos usados em scripts do Office.</span><span class="sxs-lookup"><span data-stu-id="37ced-140">The [Office Scripts API reference documentation](/javascript/api/office-scripts/overview) is a comprehensive listing of the objects used in Office Scripts.</span></span> <span data-ttu-id="37ced-141">Lá, você pode usar o Sumário para navegar para qualquer classe sobre a qual você gostaria de saber mais.</span><span class="sxs-lookup"><span data-stu-id="37ced-141">There, you can use the table of contents to navigate to any class you'd like to learn more about.</span></span> <span data-ttu-id="37ced-142">A seguir estão várias páginas exibidas com frequência.</span><span class="sxs-lookup"><span data-stu-id="37ced-142">The following are several commonly viewed pages.</span></span>

- [<span data-ttu-id="37ced-143">Chart</span><span class="sxs-lookup"><span data-stu-id="37ced-143">Chart</span></span>](/javascript/api/office-scripts/excel/excel.chart)
- [<span data-ttu-id="37ced-144">Comment</span><span class="sxs-lookup"><span data-stu-id="37ced-144">Comment</span></span>](/javascript/api/office-scripts/excel/excel.comment)
- [<span data-ttu-id="37ced-145">PivotTable</span><span class="sxs-lookup"><span data-stu-id="37ced-145">PivotTable</span></span>](/javascript/api/office-scripts/excel/excel.pivottable)
- [<span data-ttu-id="37ced-146">Range</span><span class="sxs-lookup"><span data-stu-id="37ced-146">Range</span></span>](/javascript/api/office-scripts/excel/excel.range)
- [<span data-ttu-id="37ced-147">RangeFormat</span><span class="sxs-lookup"><span data-stu-id="37ced-147">RangeFormat</span></span>](/javascript/api/office-scripts/excel/excel.rangeformat)
- [<span data-ttu-id="37ced-148">Shape</span><span class="sxs-lookup"><span data-stu-id="37ced-148">Shape</span></span>](/javascript/api/office-scripts/excel/excel.shape)
- [<span data-ttu-id="37ced-149">Table</span><span class="sxs-lookup"><span data-stu-id="37ced-149">Table</span></span>](/javascript/api/office-scripts/excel/excel.table)
- [<span data-ttu-id="37ced-150">Workbook</span><span class="sxs-lookup"><span data-stu-id="37ced-150">Workbook</span></span>](/javascript/api/office-scripts/excel/excel.workbook)
- [<span data-ttu-id="37ced-151">Worksheet</span><span class="sxs-lookup"><span data-stu-id="37ced-151">Worksheet</span></span>](/javascript/api/office-scripts/excel/excel.worksheet)

## <a name="main-function"></a><span data-ttu-id="37ced-152">`main`AllFunctions</span><span class="sxs-lookup"><span data-stu-id="37ced-152">`main` function</span></span>

<span data-ttu-id="37ced-153">Cada script do Office deve conter `main` uma função com a assinatura a seguir, `Excel.RequestContext` incluindo a definição de tipo:</span><span class="sxs-lookup"><span data-stu-id="37ced-153">Every Office Script must contain a `main` function with the following signature, including the `Excel.RequestContext` type definition:</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
    // Your Excel Script
}
```

<span data-ttu-id="37ced-154">O código dentro da `main` função é executado quando o script é executado.</span><span class="sxs-lookup"><span data-stu-id="37ced-154">The code inside the `main` function runs when the script is run.</span></span> <span data-ttu-id="37ced-155">`main`pode chamar outras funções no seu script, mas o código que não está contido em uma função não será executado.</span><span class="sxs-lookup"><span data-stu-id="37ced-155">`main` can call other functions in your script, but code that's not contained in a function will not run.</span></span>

## <a name="context"></a><span data-ttu-id="37ced-156">Contexto</span><span class="sxs-lookup"><span data-stu-id="37ced-156">Context</span></span>

<span data-ttu-id="37ced-157">A `main` função aceita um `Excel.RequestContext` parâmetro, chamado `context`.</span><span class="sxs-lookup"><span data-stu-id="37ced-157">The `main` function accepts an `Excel.RequestContext` parameter, named `context`.</span></span> <span data-ttu-id="37ced-158">Considere `context` como ponte entre o script e a pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="37ced-158">Think of `context` as the bridge between your script and the workbook.</span></span> <span data-ttu-id="37ced-159">O script acessa a pasta de trabalho com `context` o objeto e usa `context` -o para enviar dados de volta e para trás.</span><span class="sxs-lookup"><span data-stu-id="37ced-159">Your script accesses the workbook with the `context` object and uses that `context` to send data back and forth.</span></span>

<span data-ttu-id="37ced-160">O `context` objeto é necessário porque o script e o Excel estão sendo executados em diferentes processos e locais.</span><span class="sxs-lookup"><span data-stu-id="37ced-160">The `context` object is necessary because the script and Excel are running in different processes and locations.</span></span> <span data-ttu-id="37ced-161">O script precisará fazer alterações ou consultar dados da pasta de trabalho na nuvem.</span><span class="sxs-lookup"><span data-stu-id="37ced-161">The script will need to make changes to or query data from the workbook in the cloud.</span></span> <span data-ttu-id="37ced-162">O `context` objeto gerencia essas transações.</span><span class="sxs-lookup"><span data-stu-id="37ced-162">The `context` object manages those transactions.</span></span>

## <a name="sync-and-load"></a><span data-ttu-id="37ced-163">Sincronizar e carregar</span><span class="sxs-lookup"><span data-stu-id="37ced-163">Sync and Load</span></span>

<span data-ttu-id="37ced-164">Como o script e a pasta de trabalho são executados em locais diferentes, qualquer transferência de dados entre os dois leva tempo.</span><span class="sxs-lookup"><span data-stu-id="37ced-164">Because your script and workbook run in different locations, any data transfer between the two takes time.</span></span> <span data-ttu-id="37ced-165">Para melhorar o desempenho do script, os comandos são enfileirados até que o `sync` script chame explicitamente a operação para sincronizar o script e a pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="37ced-165">To improve script performance, commands are queued up until the script explicitly calls the `sync` operation to synchronize the script and workbook.</span></span> <span data-ttu-id="37ced-166">O script pode funcionar independentemente até que seja necessário fazer um dos seguintes:</span><span class="sxs-lookup"><span data-stu-id="37ced-166">Your script can work independently until it needs to do either of the following:</span></span>

- <span data-ttu-id="37ced-167">Ler dados da pasta de trabalho (seguindo `load` uma operação).</span><span class="sxs-lookup"><span data-stu-id="37ced-167">Read data from the workbook (following a `load` operation).</span></span>
- <span data-ttu-id="37ced-168">Gravar dados na pasta de trabalho (geralmente porque o script foi concluído).</span><span class="sxs-lookup"><span data-stu-id="37ced-168">Write data to the workbook (usually because the script has finished).</span></span>

<span data-ttu-id="37ced-169">A imagem a seguir mostra um fluxo de controle de exemplo entre o script e a pasta de trabalho:</span><span class="sxs-lookup"><span data-stu-id="37ced-169">The following image shows an example control flow between the script and workbook:</span></span>

![Um diagrama mostrando as operações de leitura e gravação indo para a pasta de trabalho do script.](../images/load-sync.png)

### <a name="sync"></a><span data-ttu-id="37ced-171">Sincronizar</span><span class="sxs-lookup"><span data-stu-id="37ced-171">Sync</span></span>

<span data-ttu-id="37ced-172">Sempre que o script precisar ler dados de ou gravar dados na pasta de trabalho, chame `RequestContext.sync` o método conforme mostrado aqui:</span><span class="sxs-lookup"><span data-stu-id="37ced-172">Whenever your script needs to read data from or write data to the workbook, call the `RequestContext.sync` method as shown here:</span></span>

```TypeScript
await context.sync();
```

> [!NOTE]
> <span data-ttu-id="37ced-173">`context.sync()`é chamado implicitamente quando um script é encerrado.</span><span class="sxs-lookup"><span data-stu-id="37ced-173">`context.sync()` is implicitly called when a script ends.</span></span>

<span data-ttu-id="37ced-174">Após a `sync` conclusão da operação, a pasta de trabalho é atualizada para refletir as operações de gravação especificadas pelo script.</span><span class="sxs-lookup"><span data-stu-id="37ced-174">After the `sync` operation completes, the workbook updates to reflect any write operations that script has specified.</span></span> <span data-ttu-id="37ced-175">Uma operação de gravação está definindo qualquer propriedade em um objeto do Excel ( `range.format.fill.color = "red"`por exemplo,) ou chamando um método que altera uma propriedade ( `range.format.autoFitColumns()`por exemplo,).</span><span class="sxs-lookup"><span data-stu-id="37ced-175">A write operation is setting any property on a Excel object (e.g. `range.format.fill.color = "red"`) or calling a method that changes a property (e.g., `range.format.autoFitColumns()`).</span></span> <span data-ttu-id="37ced-176">A `sync` operação também lê qualquer valor da pasta de trabalho que o script solicitou usando `load` uma operação (conforme discutido na próxima seção).</span><span class="sxs-lookup"><span data-stu-id="37ced-176">The `sync` operation also reads any values from the workbook that the script requested by using a `load` operation (as discussed in the next section).</span></span>

<span data-ttu-id="37ced-177">Sincronizar seu script com a pasta de trabalho pode levar tempo, dependendo da sua rede.</span><span class="sxs-lookup"><span data-stu-id="37ced-177">Synchronizing your script with the workbook can take time, depending on your network.</span></span> <span data-ttu-id="37ced-178">Você deve minimizar o número de `sync` chamadas para ajudar seu script a ser executado rapidamente.</span><span class="sxs-lookup"><span data-stu-id="37ced-178">You should minimize the number of `sync` calls to help your script run fast.</span></span>  

### <a name="load"></a><span data-ttu-id="37ced-179">Carregar</span><span class="sxs-lookup"><span data-stu-id="37ced-179">Load</span></span>

<span data-ttu-id="37ced-180">Um script deve carregar dados da pasta de trabalho antes de lê-lo.</span><span class="sxs-lookup"><span data-stu-id="37ced-180">A script must load data from the workbook before reading it.</span></span> <span data-ttu-id="37ced-181">No entanto, freqüentemente carregar dados de toda a pasta de trabalho reduziria imensamente a velocidade do script.</span><span class="sxs-lookup"><span data-stu-id="37ced-181">However, frequently loading data from the entire workbook would greatly reduce the script's speed.</span></span> <span data-ttu-id="37ced-182">Em vez disso `load` , o método permite que o seu estado de script especificamente que os dados devem ser recuperados da pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="37ced-182">Instead, the `load` method lets your script state specifically which data should be retrieved from the workbook.</span></span>

<span data-ttu-id="37ced-183">O `load` método está disponível em cada objeto do Excel.</span><span class="sxs-lookup"><span data-stu-id="37ced-183">The `load` method is available on every Excel object.</span></span> <span data-ttu-id="37ced-184">O script deve carregar as propriedades de um objeto antes de poder lê-las.</span><span class="sxs-lookup"><span data-stu-id="37ced-184">Your script must load an object's properties before it can read them.</span></span> <span data-ttu-id="37ced-185">Se isso não for feito, ocorrerá um erro.</span><span class="sxs-lookup"><span data-stu-id="37ced-185">Not doing so will result in an error.</span></span>

<span data-ttu-id="37ced-186">Os exemplos a seguir usam `Range` um objeto para mostrar as três maneiras `load` como o método pode ser usado para carregar dados.</span><span class="sxs-lookup"><span data-stu-id="37ced-186">The following examples use a `Range` object to show the three ways the `load` method can be used to load data.</span></span>

|<span data-ttu-id="37ced-187">Intent</span><span class="sxs-lookup"><span data-stu-id="37ced-187">Intent</span></span> |<span data-ttu-id="37ced-188">Comando de exemplo</span><span class="sxs-lookup"><span data-stu-id="37ced-188">Example Command</span></span> | <span data-ttu-id="37ced-189">Efeito</span><span class="sxs-lookup"><span data-stu-id="37ced-189">Effect</span></span> |
|:--|:--|:--|
|<span data-ttu-id="37ced-190">Carregar uma propriedade</span><span class="sxs-lookup"><span data-stu-id="37ced-190">Load one property</span></span> |`myRange.load("values");` | <span data-ttu-id="37ced-191">Carrega uma única propriedade, neste caso, a matriz bidimensional de valores neste intervalo.</span><span class="sxs-lookup"><span data-stu-id="37ced-191">Loads a single property, in this case the two-dimensional array of values in this range.</span></span> |
|<span data-ttu-id="37ced-192">Carregar várias propriedades</span><span class="sxs-lookup"><span data-stu-id="37ced-192">Load multiple properties</span></span> |`myRange.load("values, rowCount, columnCount");`| <span data-ttu-id="37ced-193">Carrega todas as propriedades de uma lista delimitada por vírgulas, neste exemplo, os valores, a contagem de linhas e a contagem de colunas.</span><span class="sxs-lookup"><span data-stu-id="37ced-193">Loads all the properties from a comma-delimited list, in this example the values, row count, and column count.</span></span> |
|<span data-ttu-id="37ced-194">Carregar tudo</span><span class="sxs-lookup"><span data-stu-id="37ced-194">Load everything</span></span> | `myRange.load();`|<span data-ttu-id="37ced-195">Carrega todas as propriedades no intervalo.</span><span class="sxs-lookup"><span data-stu-id="37ced-195">Loads all the properties on the range.</span></span> <span data-ttu-id="37ced-196">Essa não é uma solução recomendada, já que ela tornará mais lento o script, obtendo dados desnecessários.</span><span class="sxs-lookup"><span data-stu-id="37ced-196">This is not a recommended solution, since it will slow down your script by getting unnecessary data.</span></span> <span data-ttu-id="37ced-197">Você só deve usar isso ao testar o script ou se precisar de todas as propriedades do objeto.</span><span class="sxs-lookup"><span data-stu-id="37ced-197">You should only use this while testing your script or if you need every property from the object.</span></span> |

<span data-ttu-id="37ced-198">O script deve chamar `context.sync()` antes de ler qualquer valor carregado.</span><span class="sxs-lookup"><span data-stu-id="37ced-198">Your script must call `context.sync()` before reading any loaded values.</span></span>

```TypeScript
let range = selectedSheet.getRange("A1:B3");
range.load ("rowCount"); // Load the property.
await context.sync(); // Synchronize with the workbook to get the property.
console.log(range.rowCount); // Read and log the property value (3).
```

<span data-ttu-id="37ced-199">Você também pode carregar Propriedades em uma coleção inteira.</span><span class="sxs-lookup"><span data-stu-id="37ced-199">You can also load properties across an entire collection.</span></span> <span data-ttu-id="37ced-200">Cada objeto de coleção tem `items` uma propriedade que é uma matriz que contém os objetos dessa coleção.</span><span class="sxs-lookup"><span data-stu-id="37ced-200">Every collection object has an `items` property that is an array containing the objects in that collection.</span></span> <span data-ttu-id="37ced-201">Usando `items` como o início de uma chamada hierárquica (`items\myProperty`) para `load` carregar as propriedades especificadas em cada um desses itens.</span><span class="sxs-lookup"><span data-stu-id="37ced-201">Using `items` as the start of a hierarchical call (`items\myProperty`) to `load` loads the specified properties on each of those items.</span></span> <span data-ttu-id="37ced-202">O exemplo a seguir carrega `resolved` a propriedade em `Comment` cada objeto no `CommentCollection` objeto de uma planilha.</span><span class="sxs-lookup"><span data-stu-id="37ced-202">The following example loads the `resolved` property on every `Comment` object in the `CommentCollection` object of a worksheet.</span></span>

```TypeScript
let comments = selectedSheet.comments;
comments.load("items/resolved"); // Load the `resolved` property from every comment in this collection.
await context.sync(); // Synchronize with the workbook to get the properties.
```

> [!TIP]
> <span data-ttu-id="37ced-203">Para saber mais sobre como trabalhar com coleções em scripts do Office, consulte a [seção matriz do artigo usando objetos JavaScript incorporados no Office scripts](javascript-objects.md#array) .</span><span class="sxs-lookup"><span data-stu-id="37ced-203">To learn more about working with collections in Office Scripts, see the [Array section of the Using built-in JavaScript objects in Office Scripts](javascript-objects.md#array) article.</span></span>

## <a name="see-also"></a><span data-ttu-id="37ced-204">Confira também</span><span class="sxs-lookup"><span data-stu-id="37ced-204">See also</span></span>

- [<span data-ttu-id="37ced-205">Gravar, editar e criar scripts do Office no Excel na Web</span><span class="sxs-lookup"><span data-stu-id="37ced-205">Record, edit, and create Office Scripts in Excel on the web</span></span>](../tutorials/excel-tutorial.md)
- [<span data-ttu-id="37ced-206">Ler dados de pasta de trabalho com scripts do Office no Excel na Web</span><span class="sxs-lookup"><span data-stu-id="37ced-206">Read workbook data with Office Scripts in Excel on the web</span></span>](../tutorials/excel-read-tutorial.md)
- [<span data-ttu-id="37ced-207">Referência da API de scripts do Office</span><span class="sxs-lookup"><span data-stu-id="37ced-207">Office Scripts API reference</span></span>](/javascript/api/office-scripts/overview)
- [<span data-ttu-id="37ced-208">Usando objetos JavaScript internos em scripts do Office</span><span class="sxs-lookup"><span data-stu-id="37ced-208">Using built-in JavaScript objects in Office Scripts</span></span>](javascript-objects.md)
