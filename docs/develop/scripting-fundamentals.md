---
title: Fundamentos de script para scripts do Office no Excel na Web
description: Informações sobre o modelo de objeto e outros fundamentos para saber mais antes de escrever scripts do Office.
ms.date: 01/27/2020
localization_priority: Priority
ms.openlocfilehash: 5a709c16e23c00ffc7ee7949a3cb11459dc2d530
ms.sourcegitcommit: d556aaefac80e55f53ac56b7f6ecbc657ebd426f
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/26/2020
ms.locfileid: "42978698"
---
# <a name="scripting-fundamentals-for-office-scripts-in-excel-on-the-web-preview"></a><span data-ttu-id="98847-103">Fundamentos de script para scripts do Office no Excel na Web (visualização)</span><span class="sxs-lookup"><span data-stu-id="98847-103">Scripting fundamentals for Office Scripts in Excel on the web (preview)</span></span>

<span data-ttu-id="98847-104">Este artigo apresentará os aspectos técnicos dos scripts do Office.</span><span class="sxs-lookup"><span data-stu-id="98847-104">This article will introduce you to the technical aspects of Office Scripts.</span></span> <span data-ttu-id="98847-105">Você saberá como os objetos do Excel funcionam em conjunto e como o editor de código se sincroniza com uma pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="98847-105">You'll learn how the Excel objects work together and how the Code Editor synchronizes with a workbook.</span></span>

[!INCLUDE [Preview note](../includes/preview-note.md)]

## <a name="object-model"></a><span data-ttu-id="98847-106">Modelo de objetos</span><span class="sxs-lookup"><span data-stu-id="98847-106">Object model</span></span>

<span data-ttu-id="98847-107">Para entender as APIs do Excel, você deve entender como os componentes de uma pasta de trabalho estão relacionados entre si.</span><span class="sxs-lookup"><span data-stu-id="98847-107">To understand the Excel APIs, you must understand how the components of a workbook are related to one another.</span></span>

- <span data-ttu-id="98847-108">Uma **Pasta de trabalho** contém uma ou mais **Planilhas**.</span><span class="sxs-lookup"><span data-stu-id="98847-108">A **Workbook** contains one or more **Worksheets**.</span></span>
- <span data-ttu-id="98847-109">Uma **Planilha** concede acesso a células por meio de objetos de **Intervalo**.</span><span class="sxs-lookup"><span data-stu-id="98847-109">A **Worksheet** gives access to cells through **Range** objects.</span></span>
- <span data-ttu-id="98847-110">Um **Intervalo** representa um grupo de células contíguas.</span><span class="sxs-lookup"><span data-stu-id="98847-110">A **Range** represents a group of contiguous cells.</span></span>
- <span data-ttu-id="98847-111">Os **Intervalos** são usados para criar e colocar **Tabelas**, **Gráficos**, **Formas** e outras visualizações de dados ou objetos da organização.</span><span class="sxs-lookup"><span data-stu-id="98847-111">**Ranges** are used to create and place **Tables**, **Charts**, **Shapes**, and other data visualization or organization objects.</span></span>
- <span data-ttu-id="98847-112">Uma **Planilha** contém coleções desses objetos de dados que estão presentes na planilha individual.</span><span class="sxs-lookup"><span data-stu-id="98847-112">A **Worksheet** contains collections of those data objects that are present in the individual sheet.</span></span>
- <span data-ttu-id="98847-113">As **Pastas de trabalho** contêm coleções de alguns desses objetos de dados (por exemplo, **Tabelas**) para toda a **Pasta de trabalho**.</span><span class="sxs-lookup"><span data-stu-id="98847-113">**Workbooks** contain collections of some of those data objects (such as **Tables**) for the entire **Workbook**.</span></span>

### <a name="ranges"></a><span data-ttu-id="98847-114">Intervalos</span><span class="sxs-lookup"><span data-stu-id="98847-114">Ranges</span></span>

<span data-ttu-id="98847-115">Um intervalo é um grupo de células contíguas na pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="98847-115">A range is a group of contiguous cells in the workbook.</span></span> <span data-ttu-id="98847-116">Os scripts geralmente usam a notação de estilo A1 (por exemplo, **B3** para a única célula na linha **B** e coluna **3** ou **C2:F4** para as células das linhas **C** a **F** e colunas **2** a **4**) para definir intervalos.</span><span class="sxs-lookup"><span data-stu-id="98847-116">Scripts typically use A1-style notation (e.g. **B3** for the single cell in row **B** and column **3** or **C2:F4** for the cells from rows **C** through **F** and columns **2** through **4**) to define ranges.</span></span>

<span data-ttu-id="98847-117">Os intervalos têm três propriedades principais: `values`, `formulas` e `format`.</span><span class="sxs-lookup"><span data-stu-id="98847-117">Ranges have three core properties: `values`, `formulas`, and `format`.</span></span> <span data-ttu-id="98847-118">Essas propriedades recebem ou definem os valores da célula, as fórmulas a serem avaliadas e a formatação visual das células.</span><span class="sxs-lookup"><span data-stu-id="98847-118">These properties get or set the cell values, formulas to be evaluated, and the visual formatting of the cells.</span></span>

#### <a name="range-sample"></a><span data-ttu-id="98847-119">Exemplo de intervalo</span><span class="sxs-lookup"><span data-stu-id="98847-119">Range sample</span></span>

<span data-ttu-id="98847-120">O exemplo a seguir mostra como criar registros de vendas.</span><span class="sxs-lookup"><span data-stu-id="98847-120">The following sample shows how to create sales records.</span></span> <span data-ttu-id="98847-121">Esse script usa objetos `Range` para definir os valores, fórmulas e formatos.</span><span class="sxs-lookup"><span data-stu-id="98847-121">This script uses `Range` objects to set the values, formulas, and formats.</span></span>

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

<span data-ttu-id="98847-122">Executar este script cria os seguintes dados na planilha atual:</span><span class="sxs-lookup"><span data-stu-id="98847-122">Running this script creates the following data in the current worksheet:</span></span>

![Um registro de vendas mostrando as linhas de valores, uma coluna de fórmulas e cabeçalhos formatados.](../images/range-sample.png)

### <a name="charts-tables-and-other-data-objects"></a><span data-ttu-id="98847-124">Gráficos, tabelas e outros objetos de dados</span><span class="sxs-lookup"><span data-stu-id="98847-124">Charts, tables, and other data objects</span></span>

<span data-ttu-id="98847-125">Os scripts podem criar e manipular estruturas de dados e visualizações no Excel.</span><span class="sxs-lookup"><span data-stu-id="98847-125">Scripts can create and manipulate the data structures and visualizations within Excel.</span></span> <span data-ttu-id="98847-126">As tabelas e gráficos são dois dos objetos mais usados, mas as APIs oferecem suporte a tabelas dinâmicas, formas, imagens e muito mais.</span><span class="sxs-lookup"><span data-stu-id="98847-126">Tables and charts are two of the more commonly used objects, but the APIs support PivotTables, shapes, images, and more.</span></span>

#### <a name="creating-a-table"></a><span data-ttu-id="98847-127">Criar uma tabela</span><span class="sxs-lookup"><span data-stu-id="98847-127">Creating a table</span></span>

<span data-ttu-id="98847-128">Criar tabelas usando intervalos de dados preenchidos.</span><span class="sxs-lookup"><span data-stu-id="98847-128">Create tables by using data-filled ranges.</span></span> <span data-ttu-id="98847-129">Controles de formatação e tabela (por exemplo, filtros) são aplicados automaticamente ao intervalo.</span><span class="sxs-lookup"><span data-stu-id="98847-129">Formatting and table controls (such as filters) are automatically applied to the range.</span></span>

<span data-ttu-id="98847-130">O script a seguir cria uma tabela usando os intervalos do exemplo anterior.</span><span class="sxs-lookup"><span data-stu-id="98847-130">The following script creates a table using the ranges from the previous sample.</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
   let sheet = context.workbook.worksheets.getActiveWorksheet();
   sheet.tables.add("B2:E5", true);
}
```

<span data-ttu-id="98847-131">Executar esse script na planilha com os dados anteriores cria a tabela a seguir:</span><span class="sxs-lookup"><span data-stu-id="98847-131">Running this script on the worksheet with the previous data creates the following table:</span></span>

![Uma tabela criada a partir do registro de vendas anterior.](../images/table-sample.png)

#### <a name="creating-a-chart"></a><span data-ttu-id="98847-133">Criar um gráfico</span><span class="sxs-lookup"><span data-stu-id="98847-133">Creating a chart</span></span>

<span data-ttu-id="98847-134">Crie gráficos para visualizar os dados em um intervalo.</span><span class="sxs-lookup"><span data-stu-id="98847-134">Create charts to visualize the data in a range.</span></span> <span data-ttu-id="98847-135">Os scripts permitem inúmeras variedades de gráficos que podem ser personalizadas de acordo com suas necessidades.</span><span class="sxs-lookup"><span data-stu-id="98847-135">Scripts allow for dozens of chart varieties, each of which can be customized to suit your needs.</span></span>

<span data-ttu-id="98847-136">O script a seguir cria um gráfico de colunas simples para três itens e o coloca 100 pixels abaixo da parte superior da planilha.</span><span class="sxs-lookup"><span data-stu-id="98847-136">The following script creates a simple column chart for three items and places it 100 pixels below the top of the worksheet.</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
  let sheet = context.workbook.worksheets.getActiveWorksheet();
  let chart = sheet.charts.add(Excel.ChartType.columnStacked, sheet.getRange("B3:C5"));
  chart.top = 100;
}
```

<span data-ttu-id="98847-137">Executar este script na planilha com a tabela anterior cria o seguinte gráfico:</span><span class="sxs-lookup"><span data-stu-id="98847-137">Running this script on the worksheet with the previous table creates the following chart:</span></span>

![Um gráfico de colunas mostrando as quantidades de três itens do registro de vendas anterior.](../images/chart-sample.png)

### <a name="further-reading-on-the-object-model"></a><span data-ttu-id="98847-139">Leituras adicionais sobre o modelo de objeto</span><span class="sxs-lookup"><span data-stu-id="98847-139">Further reading on the object model</span></span>

<span data-ttu-id="98847-140">A [documentação de referência de API dos scripts do Office](/javascript/api/office-scripts/overview) é uma lista completa dos objetos usados nos scripts do Office.</span><span class="sxs-lookup"><span data-stu-id="98847-140">The [Office Scripts API reference documentation](/javascript/api/office-scripts/overview) is a comprehensive listing of the objects used in Office Scripts.</span></span> <span data-ttu-id="98847-141">Lá, você pode usar o sumário para navegar para qualquer classe da qual quiser saber mais.</span><span class="sxs-lookup"><span data-stu-id="98847-141">There, you can use the table of contents to navigate to any class you'd like to learn more about.</span></span> <span data-ttu-id="98847-142">Estas são várias páginas exibidas com frequência.</span><span class="sxs-lookup"><span data-stu-id="98847-142">The following are several commonly viewed pages.</span></span>

- [<span data-ttu-id="98847-143">Gráfico</span><span class="sxs-lookup"><span data-stu-id="98847-143">Chart</span></span>](/javascript/api/office-scripts/excel/excel.chart)
- [<span data-ttu-id="98847-144">Comentário</span><span class="sxs-lookup"><span data-stu-id="98847-144">Comment</span></span>](/javascript/api/office-scripts/excel/excel.comment)
- [<span data-ttu-id="98847-145">PivotTable</span><span class="sxs-lookup"><span data-stu-id="98847-145">PivotTable</span></span>](/javascript/api/office-scripts/excel/excel.pivottable)
- [<span data-ttu-id="98847-146">Range</span><span class="sxs-lookup"><span data-stu-id="98847-146">Range</span></span>](/javascript/api/office-scripts/excel/excel.range)
- [<span data-ttu-id="98847-147">RangeFormat</span><span class="sxs-lookup"><span data-stu-id="98847-147">RangeFormat</span></span>](/javascript/api/office-scripts/excel/excel.rangeformat)
- [<span data-ttu-id="98847-148">Formato</span><span class="sxs-lookup"><span data-stu-id="98847-148">Shape</span></span>](/javascript/api/office-scripts/excel/excel.shape)
- [<span data-ttu-id="98847-149">Table</span><span class="sxs-lookup"><span data-stu-id="98847-149">Table</span></span>](/javascript/api/office-scripts/excel/excel.table)
- [<span data-ttu-id="98847-150">Pasta de trabalho</span><span class="sxs-lookup"><span data-stu-id="98847-150">Workbook</span></span>](/javascript/api/office-scripts/excel/excel.workbook)
- [<span data-ttu-id="98847-151">Planilha</span><span class="sxs-lookup"><span data-stu-id="98847-151">Worksheet</span></span>](/javascript/api/office-scripts/excel/excel.worksheet)

## <a name="main-function"></a><span data-ttu-id="98847-152">função `main`</span><span class="sxs-lookup"><span data-stu-id="98847-152">`main` function</span></span>

<span data-ttu-id="98847-153">Todos os scripts do Office devem conter uma função `main` com a seguinte assinatura, incluindo a definição de tipo `Excel.RequestContext`:</span><span class="sxs-lookup"><span data-stu-id="98847-153">Every Office Script must contain a `main` function with the following signature, including the `Excel.RequestContext` type definition:</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
    // Your Excel Script
}
```

<span data-ttu-id="98847-154">O código dentro da função `main` é executado quando o script é executado.</span><span class="sxs-lookup"><span data-stu-id="98847-154">The code inside the `main` function runs when the script is run.</span></span> <span data-ttu-id="98847-155">`main` pode chamar outras funções em seu script, mas o código que não estiver contido em uma função não será executado.</span><span class="sxs-lookup"><span data-stu-id="98847-155">`main` can call other functions in your script, but code that's not contained in a function will not run.</span></span>

## <a name="context"></a><span data-ttu-id="98847-156">Contexto</span><span class="sxs-lookup"><span data-stu-id="98847-156">Context</span></span>

<span data-ttu-id="98847-157">A função `main` aceita um parâmetro `Excel.RequestContext`, chamado `context`.</span><span class="sxs-lookup"><span data-stu-id="98847-157">The `main` function accepts an `Excel.RequestContext` parameter, named `context`.</span></span> <span data-ttu-id="98847-158">Imagine `context` como a ponte entre o seu script e a pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="98847-158">Think of `context` as the bridge between your script and the workbook.</span></span> <span data-ttu-id="98847-159">Seu script acessa a pasta de trabalho com o objeto `context` e usa esse `context` para troca de dados.</span><span class="sxs-lookup"><span data-stu-id="98847-159">Your script accesses the workbook with the `context` object and uses that `context` to send data back and forth.</span></span>

<span data-ttu-id="98847-160">O objeto `context` é necessário porque o script e o Excel estão sendo executados em processos e locais diferentes.</span><span class="sxs-lookup"><span data-stu-id="98847-160">The `context` object is necessary because the script and Excel are running in different processes and locations.</span></span> <span data-ttu-id="98847-161">O script precisará fazer alterações ou consultar dados da pasta de trabalho na nuvem.</span><span class="sxs-lookup"><span data-stu-id="98847-161">The script will need to make changes to or query data from the workbook in the cloud.</span></span> <span data-ttu-id="98847-162">O objeto `context` gerencia essas transações.</span><span class="sxs-lookup"><span data-stu-id="98847-162">The `context` object manages those transactions.</span></span>

## <a name="sync-and-load"></a><span data-ttu-id="98847-163">Sincronizar e carregar</span><span class="sxs-lookup"><span data-stu-id="98847-163">Sync and Load</span></span>

<span data-ttu-id="98847-164">Como o seu script e a pasta de trabalho são executados em locais diferentes, qualquer transferência de dados entre os dois levará algum tempo.</span><span class="sxs-lookup"><span data-stu-id="98847-164">Because your script and workbook run in different locations, any data transfer between the two takes time.</span></span> <span data-ttu-id="98847-165">Para melhorar o desempenho do script, os comandos são enfileirados até que o script chame explicitamente a operação `sync` para sincronizar o script e a pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="98847-165">To improve script performance, commands are queued up until the script explicitly calls the `sync` operation to synchronize the script and workbook.</span></span> <span data-ttu-id="98847-166">Seu script pode trabalhar de forma independente até que precise executar uma das seguintes ações:</span><span class="sxs-lookup"><span data-stu-id="98847-166">Your script can work independently until it needs to do either of the following:</span></span>

- <span data-ttu-id="98847-167">Ler dados da pasta de trabalho (após uma operação `load`).</span><span class="sxs-lookup"><span data-stu-id="98847-167">Read data from the workbook (following a `load` operation).</span></span>
- <span data-ttu-id="98847-168">Gravar dados na pasta de trabalho (geralmente porque o script terminou).</span><span class="sxs-lookup"><span data-stu-id="98847-168">Write data to the workbook (usually because the script has finished).</span></span>

<span data-ttu-id="98847-169">A imagem a seguir mostra um exemplo de fluxo de controle entre o script e a pasta de trabalho:</span><span class="sxs-lookup"><span data-stu-id="98847-169">The following image shows an example control flow between the script and workbook:</span></span>

![Um diagrama mostrando operações de leitura e gravação saindo do script e indo para a pasta de trabalho.](../images/load-sync.png)

### <a name="sync"></a><span data-ttu-id="98847-171">Sincronizar</span><span class="sxs-lookup"><span data-stu-id="98847-171">Sync</span></span>

<span data-ttu-id="98847-172">Sempre que o seu script precisa ler ou gravar dados na pasta de trabalho, chame o método `RequestContext.sync`, conforme mostrado aqui:</span><span class="sxs-lookup"><span data-stu-id="98847-172">Whenever your script needs to read data from or write data to the workbook, call the `RequestContext.sync` method as shown here:</span></span>

```TypeScript
await context.sync();
```

> [!NOTE]
> <span data-ttu-id="98847-173">`context.sync()` é chamado implicitamente quando um script termina.</span><span class="sxs-lookup"><span data-stu-id="98847-173">`context.sync()` is implicitly called when a script ends.</span></span>

<span data-ttu-id="98847-174">Após a conclusão da operação `sync`, a pasta de trabalho será atualizada para refletir as operações de gravação especificados por esse script.</span><span class="sxs-lookup"><span data-stu-id="98847-174">After the `sync` operation completes, the workbook updates to reflect any write operations that script has specified.</span></span> <span data-ttu-id="98847-175">Uma operação de gravação está definindo uma propriedade em um objeto do Excel (por exemplo, `range.format.fill.color = "red"`) ou chamando um método que altera uma propriedade (por exemplo, `range.format.autoFitColumns()`).</span><span class="sxs-lookup"><span data-stu-id="98847-175">A write operation is setting any property on a Excel object (e.g. `range.format.fill.color = "red"`) or calling a method that changes a property (e.g., `range.format.autoFitColumns()`).</span></span> <span data-ttu-id="98847-176">A operação `sync` também lê os valores da pasta de trabalho que o script solicitou usando uma operação `load` (conforme discutido na próxima seção).</span><span class="sxs-lookup"><span data-stu-id="98847-176">The `sync` operation also reads any values from the workbook that the script requested by using a `load` operation (as discussed in the next section).</span></span>

<span data-ttu-id="98847-177">A sincronização do seu script com a pasta de trabalho pode demorar, dependendo da sua rede.</span><span class="sxs-lookup"><span data-stu-id="98847-177">Synchronizing your script with the workbook can take time, depending on your network.</span></span> <span data-ttu-id="98847-178">Você deve minimizar o número de chamadas `sync` para ajudar seu script a ser executado rapidamente.</span><span class="sxs-lookup"><span data-stu-id="98847-178">You should minimize the number of `sync` calls to help your script run fast.</span></span>  

### <a name="load"></a><span data-ttu-id="98847-179">Carregar</span><span class="sxs-lookup"><span data-stu-id="98847-179">Load</span></span>

<span data-ttu-id="98847-180">Um script deve carregar os dados da pasta de trabalho antes de lê-los.</span><span class="sxs-lookup"><span data-stu-id="98847-180">A script must load data from the workbook before reading it.</span></span> <span data-ttu-id="98847-181">No entanto, o carregamento frequente de dados de uma pasta de trabalho reduz significativamente a velocidade do script.</span><span class="sxs-lookup"><span data-stu-id="98847-181">However, frequently loading data from the entire workbook would greatly reduce the script's speed.</span></span> <span data-ttu-id="98847-182">Em vez disso, o método `load` permite que o seu script indique especificamente quais dados devem ser recuperados da pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="98847-182">Instead, the `load` method lets your script state specifically which data should be retrieved from the workbook.</span></span>

<span data-ttu-id="98847-183">O método `load` está disponível em todos os objetos do Excel.</span><span class="sxs-lookup"><span data-stu-id="98847-183">The `load` method is available on every Excel object.</span></span> <span data-ttu-id="98847-184">Seu script deve carregar as propriedades de um objeto para poder lê-lo.</span><span class="sxs-lookup"><span data-stu-id="98847-184">Your script must load an object's properties before it can read them.</span></span> <span data-ttu-id="98847-185">Se isso não for feito, ocorrerá um erro.</span><span class="sxs-lookup"><span data-stu-id="98847-185">Not doing so will result in an error.</span></span>

<span data-ttu-id="98847-186">Os exemplos a seguir usam um objeto `Range` para mostrar as três maneiras de usar o método `load` para carregar dados.</span><span class="sxs-lookup"><span data-stu-id="98847-186">The following examples use a `Range` object to show the three ways the `load` method can be used to load data.</span></span>

|<span data-ttu-id="98847-187">Finalidade</span><span class="sxs-lookup"><span data-stu-id="98847-187">Intent</span></span> |<span data-ttu-id="98847-188">Comando de exemplo</span><span class="sxs-lookup"><span data-stu-id="98847-188">Example Command</span></span> | <span data-ttu-id="98847-189">Efeito</span><span class="sxs-lookup"><span data-stu-id="98847-189">Effect</span></span> |
|:--|:--|:--|
|<span data-ttu-id="98847-190">Carregar uma propriedade</span><span class="sxs-lookup"><span data-stu-id="98847-190">Load one property</span></span> |`myRange.load("values");` | <span data-ttu-id="98847-191">Carrega uma única propriedade, neste caso, a matriz bidimensional de valores nesse intervalo.</span><span class="sxs-lookup"><span data-stu-id="98847-191">Loads a single property, in this case the two-dimensional array of values in this range.</span></span> |
|<span data-ttu-id="98847-192">Carregar várias propriedades</span><span class="sxs-lookup"><span data-stu-id="98847-192">Load multiple properties</span></span> |`myRange.load("values, rowCount, columnCount");`| <span data-ttu-id="98847-193">Carrega todas as propriedades de uma lista delimitada por vírgulas, neste exemplo, os valores, a contagem de linhas e de colunas.</span><span class="sxs-lookup"><span data-stu-id="98847-193">Loads all the properties from a comma-delimited list, in this example the values, row count, and column count.</span></span> |
|<span data-ttu-id="98847-194">Carregar tudo</span><span class="sxs-lookup"><span data-stu-id="98847-194">Load everything</span></span> | `myRange.load();`|<span data-ttu-id="98847-195">Carrega todas as propriedades no intervalo.</span><span class="sxs-lookup"><span data-stu-id="98847-195">Loads all the properties on the range.</span></span> <span data-ttu-id="98847-196">Essa não é uma solução recomendada, uma vez que diminuirá a velocidade do seu script ao obter dados desnecessários.</span><span class="sxs-lookup"><span data-stu-id="98847-196">This is not a recommended solution, since it will slow down your script by getting unnecessary data.</span></span> <span data-ttu-id="98847-197">Você só deve usar isso enquanto testa seu script ou se precisar de todas as propriedades do objeto.</span><span class="sxs-lookup"><span data-stu-id="98847-197">You should only use this while testing your script or if you need every property from the object.</span></span> |

<span data-ttu-id="98847-198">Seu script deve chamar `context.sync()` antes de ler os valores carregados.</span><span class="sxs-lookup"><span data-stu-id="98847-198">Your script must call `context.sync()` before reading any loaded values.</span></span>

```TypeScript
let range = selectedSheet.getRange("A1:B3");
range.load ("rowCount"); // Load the property.
await context.sync(); // Synchronize with the workbook to get the property.
console.log(range.rowCount); // Read and log the property value (3).
```

<span data-ttu-id="98847-199">Você também pode carregar as propriedades em uma coleção.</span><span class="sxs-lookup"><span data-stu-id="98847-199">You can also load properties across an entire collection.</span></span> <span data-ttu-id="98847-200">Todos os objetos da coleção têm uma propriedade `items` que é uma matriz contendo os objetos dessa coleção.</span><span class="sxs-lookup"><span data-stu-id="98847-200">Every collection object has an `items` property that is an array containing the objects in that collection.</span></span> <span data-ttu-id="98847-201">Usar `items` como o início de uma chamada hierárquica (`items\myProperty`) para `load` carrega as propriedades especificadas em cada um desses itens.</span><span class="sxs-lookup"><span data-stu-id="98847-201">Using `items` as the start of a hierarchical call (`items\myProperty`) to `load` loads the specified properties on each of those items.</span></span> <span data-ttu-id="98847-202">O exemplo a seguir carrega a propriedade `resolved` em cada objeto `Comment` no objeto `CommentCollection` de uma planilha.</span><span class="sxs-lookup"><span data-stu-id="98847-202">The following example loads the `resolved` property on every `Comment` object in the `CommentCollection` object of a worksheet.</span></span>

```TypeScript
let comments = selectedSheet.comments;
comments.load("items/resolved"); // Load the `resolved` property from every comment in this collection.
await context.sync(); // Synchronize with the workbook to get the properties.
```

> [!TIP]
> <span data-ttu-id="98847-203">Para saber mais sobre como trabalhar com coleções nos scripts do Office, confira a seção Matriz do artigo [Usar objetos internos do JavaScript nos scripts do Office](javascript-objects.md#array).</span><span class="sxs-lookup"><span data-stu-id="98847-203">To learn more about working with collections in Office Scripts, see the [Array section of the Using built-in JavaScript objects in Office Scripts](javascript-objects.md#array) article.</span></span>

## <a name="see-also"></a><span data-ttu-id="98847-204">Confira também</span><span class="sxs-lookup"><span data-stu-id="98847-204">See also</span></span>

- [<span data-ttu-id="98847-205">Gravar, editar e criar scripts do Office no Excel na Web</span><span class="sxs-lookup"><span data-stu-id="98847-205">Record, edit, and create Office Scripts in Excel on the web</span></span>](../tutorials/excel-tutorial.md)
- [<span data-ttu-id="98847-206">Ler os dados da pasta de trabalho com scripts do Office no Excel na Web</span><span class="sxs-lookup"><span data-stu-id="98847-206">Read workbook data with Office Scripts in Excel on the web</span></span>](../tutorials/excel-read-tutorial.md)
- [<span data-ttu-id="98847-207">Referência da API de scripts do Office</span><span class="sxs-lookup"><span data-stu-id="98847-207">Office Scripts API reference</span></span>](/javascript/api/office-scripts/overview)
- [<span data-ttu-id="98847-208">Usar objetos internos do JavaScript nos scripts do Office</span><span class="sxs-lookup"><span data-stu-id="98847-208">Using built-in JavaScript objects in Office Scripts</span></span>](javascript-objects.md)
