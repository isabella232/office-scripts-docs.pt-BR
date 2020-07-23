---
title: Usar objetos internos do JavaScript nos scripts do Office
description: Como chamar APIs JavaScript internas de um script do Office no Excel na Web.
ms.date: 07/16/2020
localization_priority: Normal
ms.openlocfilehash: 4bb5fb5444887005ececbbfdf0130cba3784e0c4
ms.sourcegitcommit: 8d549884e68170f808d3d417104a4451a37da83c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/22/2020
ms.locfileid: "45229593"
---
# <a name="using-built-in-javascript-objects-in-office-scripts"></a><span data-ttu-id="5fd67-103">Usar objetos internos do JavaScript nos scripts do Office</span><span class="sxs-lookup"><span data-stu-id="5fd67-103">Using built-in JavaScript objects in Office Scripts</span></span>

<span data-ttu-id="5fd67-104">O JavaScript fornece vários objetos internos que você pode usar em seus scripts do Office, independentemente de você estar criando scripts em JavaScript ou [TypeScript](../overview/code-editor-environment.md) (um superconjunto de JavaScript).</span><span class="sxs-lookup"><span data-stu-id="5fd67-104">JavaScript provides several built-in objects that you can use in your Office Scripts, regardless of whether you're scripting in JavaScript or [TypeScript](../overview/code-editor-environment.md) (a superset of JavaScript).</span></span> <span data-ttu-id="5fd67-105">Este artigo descreve como você pode usar alguns dos objetos JavaScript internos em scripts do Office para Excel na Web.</span><span class="sxs-lookup"><span data-stu-id="5fd67-105">This article describes how you can use some of the built-in JavaScript objects in Office Scripts for Excel on the web.</span></span>

> [!NOTE]
> <span data-ttu-id="5fd67-106">Para obter uma lista completa de todos os objetos JavaScript internos, consulte o artigo sobre [objetos internos padrão](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects) do Mozilla.</span><span class="sxs-lookup"><span data-stu-id="5fd67-106">For a complete list of all built-in JavaScript objects, see Mozilla's [Standard built-in objects](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects) article.</span></span>

## <a name="array"></a><span data-ttu-id="5fd67-107">Matriz</span><span class="sxs-lookup"><span data-stu-id="5fd67-107">Array</span></span>

<span data-ttu-id="5fd67-108">O objeto [array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) oferece uma maneira padronizada de trabalhar com matrizes no seu script.</span><span class="sxs-lookup"><span data-stu-id="5fd67-108">The [Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) object provides a standardized way to work with arrays in your script.</span></span> <span data-ttu-id="5fd67-109">Embora matrizes sejam construções JavaScript padrão, elas se relacionam aos scripts do Office de duas maneiras principais: intervalos e coleções.</span><span class="sxs-lookup"><span data-stu-id="5fd67-109">While arrays are standard JavaScript constructs, they relate to Office Scripts in two major ways: ranges and collections.</span></span>

### <a name="working-with-ranges"></a><span data-ttu-id="5fd67-110">Trabalhar com intervalos</span><span class="sxs-lookup"><span data-stu-id="5fd67-110">Working with ranges</span></span>

<span data-ttu-id="5fd67-111">Intervalos contêm várias matrizes bidimensionais que mapeiam diretamente para as células naquele intervalo.</span><span class="sxs-lookup"><span data-stu-id="5fd67-111">Ranges contain several two-dimensional arrays that directly map to the cells in that range.</span></span> <span data-ttu-id="5fd67-112">Essas matrizes contêm informações específicas sobre cada célula desse intervalo.</span><span class="sxs-lookup"><span data-stu-id="5fd67-112">These arrays contain specific information about each cell in that range.</span></span> <span data-ttu-id="5fd67-113">Por exemplo, `Range.getValues` retorna todos os valores dessas células (com as linhas e colunas do mapeamento de duas dimensões bidimensionais para as linhas e colunas dessa subseção de planilha).</span><span class="sxs-lookup"><span data-stu-id="5fd67-113">For example, `Range.getValues` returns all the values in those cells (with the rows and columns of the two-dimensional array mapping to the rows and columns of that worksheet subsection).</span></span> <span data-ttu-id="5fd67-114">`Range.getFormulas`e `Range.getNumberFormats` são outros métodos usados com frequência que retornam matrizes, como `Range.getValues` .</span><span class="sxs-lookup"><span data-stu-id="5fd67-114">`Range.getFormulas` and `Range.getNumberFormats` are other frequently used methods that return arrays like `Range.getValues`.</span></span>

<span data-ttu-id="5fd67-115">O script a seguir pesquisa o intervalo **a1: D4** para qualquer formato de número que contenha um "$".</span><span class="sxs-lookup"><span data-stu-id="5fd67-115">The following script searches the **A1:D4** range for any number format containing a "$".</span></span> <span data-ttu-id="5fd67-116">O script define a cor de preenchimento dessas células como "amarelo".</span><span class="sxs-lookup"><span data-stu-id="5fd67-116">The script sets the fill color in those cells to "yellow".</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the range From A1 to D4.
  let range = workbook.getActiveWorksheet().getRange("A1:D4");

  // Get the number formats for each cell in the range.
  let rangeNumberFormats = range.getNumberFormats();
  // Iterate through the arrays of rows and columns corresponding to those in the range.
  rangeNumberFormats.forEach((rowItem, rowIndex) => {
    rangeNumberFormats[rowIndex].forEach((columnItem, columnIndex) => {
      // Treat the numberFormat as a string so we can do text comparisons.
      let columnItemText = columnItem as string;
      if (columnItemText.indexOf("$") >= 0) {
        // Set the cell's fill to yellow.
        range.getCell(rowIndex, columnIndex).getFormat().getFill().setColor("yellow");
      }
    });
  });
}
```

### <a name="working-with-collections"></a><span data-ttu-id="5fd67-117">Trabalhar com coleções</span><span class="sxs-lookup"><span data-stu-id="5fd67-117">Working with collections</span></span>

<span data-ttu-id="5fd67-118">Muitos objetos do Excel estão contidos em uma coleção.</span><span class="sxs-lookup"><span data-stu-id="5fd67-118">Many Excel objects are contained in a collection.</span></span> <span data-ttu-id="5fd67-119">A coleção é gerenciada pela API de scripts do Office e exposta como uma matriz.</span><span class="sxs-lookup"><span data-stu-id="5fd67-119">The collection is managed by the Office Scripts API and exposed as an array.</span></span> <span data-ttu-id="5fd67-120">Por exemplo, todas as [formas](/javascript/api/office-scripts/excelscript/excelscript.shape) em uma planilha estão contidas em um `Shape[]` que é retornado pelo `Worksheet.getShapes` método.</span><span class="sxs-lookup"><span data-stu-id="5fd67-120">For example, all [Shapes](/javascript/api/office-scripts/excelscript/excelscript.shape) in a worksheet are contained in a `Shape[]` that is returned by the `Worksheet.getShapes` method.</span></span> <span data-ttu-id="5fd67-121">Você pode usar essa matriz para ler valores da coleção ou pode acessar objetos específicos dos métodos do objeto pai `get*` .</span><span class="sxs-lookup"><span data-stu-id="5fd67-121">You can use this array to read values from the collection, or you can access specific objects from the parent object's `get*` methods.</span></span>

> [!NOTE]
> <span data-ttu-id="5fd67-122">Não adicione nem remova manualmente objetos dessas matrizes de coleção.</span><span class="sxs-lookup"><span data-stu-id="5fd67-122">Do not manually add or remove objects from these collection arrays.</span></span> <span data-ttu-id="5fd67-123">Use os `add` métodos nos objetos pai e os `delete` métodos nos objetos do tipo coleção.</span><span class="sxs-lookup"><span data-stu-id="5fd67-123">Use the `add` methods on the parent objects and the `delete` methods on the collection-type objects.</span></span> <span data-ttu-id="5fd67-124">Por exemplo, adicione uma [tabela](/javascript/api/office-scripts/excelscript/excelscript.table) a uma [planilha](/javascript/api/office-scripts/excelscript/excelscript.worksheet) com o `Worksheet.addTable` método e remova o `Table` usando `Table.delete` .</span><span class="sxs-lookup"><span data-stu-id="5fd67-124">For example, add a [Table](/javascript/api/office-scripts/excelscript/excelscript.table) to a [Worksheet](/javascript/api/office-scripts/excelscript/excelscript.worksheet) with the `Worksheet.addTable` method and remove the `Table` using `Table.delete`.</span></span>

<span data-ttu-id="5fd67-125">O script a seguir registra o tipo de todas as formas na planilha atual.</span><span class="sxs-lookup"><span data-stu-id="5fd67-125">The following script logs the type of every shape in the current worksheet.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the current worksheet.
  let selectedSheet = workbook.getActiveWorksheet();

  // Get the shapes in this worksheet.
  let shapes = selectedSheet.getShapes();

  // Log the type of every shape in the collection.
  shapes.forEach((shape) => {
    console.log(shape.getType());
  });
}
```

<span data-ttu-id="5fd67-126">O script a seguir exclui a forma mais antiga na planilha atual.</span><span class="sxs-lookup"><span data-stu-id="5fd67-126">The following script deletes the oldest shape in the current worksheet.</span></span>

```Typescript
function main(workbook: ExcelScript.Workbook) {
  // Get the current worksheet.
  let selectedSheet = workbook.getActiveWorksheet();

  // Get the first (oldest) shape in the worksheet.
  // Note that this script will thrown an error if there are no shapes.
  let shape = selectedSheet.getShapes()[0];

  // Remove the shape from the worksheet.
  shape.delete();
}
```

## <a name="date"></a><span data-ttu-id="5fd67-127">Data</span><span class="sxs-lookup"><span data-stu-id="5fd67-127">Date</span></span>

<span data-ttu-id="5fd67-128">O objeto [Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date) fornece uma maneira padronizada para trabalhar com datas no seu script.</span><span class="sxs-lookup"><span data-stu-id="5fd67-128">The [Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date) object provides a standardized way to work with dates in your script.</span></span> <span data-ttu-id="5fd67-129">`Date.now()`gera um objeto com data e hora atuais, o que é útil ao adicionar carimbos de data/hora à entrada de dados do script.</span><span class="sxs-lookup"><span data-stu-id="5fd67-129">`Date.now()` generates an object with the current date and time, which is useful when adding timestamps to your script's data entry.</span></span>

<span data-ttu-id="5fd67-130">O script a seguir adiciona a data atual à planilha.</span><span class="sxs-lookup"><span data-stu-id="5fd67-130">The following script adds the current date to the worksheet.</span></span> <span data-ttu-id="5fd67-131">Observe que, usando o `toLocaleDateString` método, o Excel reconhece o valor como uma data e altera automaticamente o formato de número da célula.</span><span class="sxs-lookup"><span data-stu-id="5fd67-131">Note that by using the `toLocaleDateString` method, Excel recognizes the value as a date and changes the number format of the cell automatically.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the range for cell A1.
  let range = workbook.getActiveWorksheet().getRange("A1");

  // Get the current date and time.
  let date = new Date(Date.now());

  // Set the value at A1 to the current date, using a localized string.
  range.setValue(date.toLocaleDateString());
}
```

<span data-ttu-id="5fd67-132">A seção [trabalhar com datas](../resources/excel-samples.md#dates) dos exemplos tem mais scripts relacionados a datas.</span><span class="sxs-lookup"><span data-stu-id="5fd67-132">The [Work with dates](../resources/excel-samples.md#dates) section of the samples has more date-related scripts.</span></span>

## <a name="math"></a><span data-ttu-id="5fd67-133">Matemática</span><span class="sxs-lookup"><span data-stu-id="5fd67-133">Math</span></span>

<span data-ttu-id="5fd67-134">O objeto [Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math) fornece métodos e constantes para operações matemáticas comuns.</span><span class="sxs-lookup"><span data-stu-id="5fd67-134">The [Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math) object provides methods and constants for common mathematical operations.</span></span> <span data-ttu-id="5fd67-135">Eles fornecem muitas funções também disponíveis no Excel, sem a necessidade de usar o mecanismo de cálculo da pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="5fd67-135">These provide many functions also available in Excel, without the need to use the workbook's calculation engine.</span></span> <span data-ttu-id="5fd67-136">Isso salva o script de ter que consultar a pasta de trabalho, o que melhora o desempenho.</span><span class="sxs-lookup"><span data-stu-id="5fd67-136">This saves your script from having to query the workbook, which improves performance.</span></span>

<span data-ttu-id="5fd67-137">O script a seguir usa `Math.min` para localizar e registrar o menor número no intervalo **a1: D4** .</span><span class="sxs-lookup"><span data-stu-id="5fd67-137">The following script uses `Math.min` to find and log the smallest number in the **A1:D4** range.</span></span> <span data-ttu-id="5fd67-138">Observe que este exemplo pressupõe que o intervalo inteiro contenha apenas números, e não cadeias de caracteres.</span><span class="sxs-lookup"><span data-stu-id="5fd67-138">Note that this sample assumes the entire range contains only numbers, not strings.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the range from A1 to D4.
  let comparisonRange = workbook.getActiveWorksheet().getRange("A1:D4");

  // Load the range's values.
  let comparisonRangeValues = comparisonRange.getValues();

  // Set the minimum values as the first value.
  let minimum = comparisonRangeValues[0][0];

  // Iterate over each row looking for the smallest value.
  comparisonRangeValues.forEach((rowItem, rowIndex) => {
    // Iterate over each column looking for the smallest value.
    comparisonRangeValues[rowIndex].forEach((columnItem) => {
      // Use `Math.min` to set the smallest value as either the current cell's value or the previous minimum.
      minimum = Math.min(minimum, columnItem);
    });
  });

  console.log(minimum);
}

```

## <a name="use-of-external-javascript-libraries-is-not-supported"></a><span data-ttu-id="5fd67-139">Não há suporte para o uso de bibliotecas JavaScript externas</span><span class="sxs-lookup"><span data-stu-id="5fd67-139">Use of external JavaScript libraries is not supported</span></span>

<span data-ttu-id="5fd67-140">Os scripts do Office não oferecem suporte ao uso de bibliotecas externas de terceiros.</span><span class="sxs-lookup"><span data-stu-id="5fd67-140">Office Scripts don't support the use of external, third-party libraries.</span></span> <span data-ttu-id="5fd67-141">O script só pode usar os objetos JavaScript internos e as APIs de scripts do Office.</span><span class="sxs-lookup"><span data-stu-id="5fd67-141">Your script can only use the built-in JavaScript objects and the Office Scripts APIs.</span></span>

## <a name="see-also"></a><span data-ttu-id="5fd67-142">Confira também</span><span class="sxs-lookup"><span data-stu-id="5fd67-142">See also</span></span>

- [<span data-ttu-id="5fd67-143">Objetos internos padrão</span><span class="sxs-lookup"><span data-stu-id="5fd67-143">Standard built-in objects</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects)
- [<span data-ttu-id="5fd67-144">Ambiente do editor de código de scripts do Office</span><span class="sxs-lookup"><span data-stu-id="5fd67-144">Office Scripts Code Editor environment</span></span>](../overview/code-editor-environment.md)
