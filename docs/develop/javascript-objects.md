---
title: Usando objetos JavaScript internos em scripts do Office
description: Como chamar APIs JavaScript internas de um script do Office no Excel na Web.
ms.date: 01/21/2020
localization_priority: Normal
ms.openlocfilehash: e0fcd98117125ead18e55675e195415ff59c0c5d
ms.sourcegitcommit: b075eed5a6f275274fbbf6d62633219eac416f26
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/10/2020
ms.locfileid: "42700051"
---
# <a name="using-built-in-javascript-objects-in-office-scripts"></a><span data-ttu-id="2f3a0-103">Usando objetos JavaScript internos em scripts do Office</span><span class="sxs-lookup"><span data-stu-id="2f3a0-103">Using built-in JavaScript objects in Office Scripts</span></span>

<span data-ttu-id="2f3a0-104">O JavaScript fornece vários objetos internos que você pode usar em seus scripts do Office, independentemente de você estar criando scripts em JavaScript ou [TypeScript](../overview/code-editor-environment.md) (um superconjunto de JavaScript).</span><span class="sxs-lookup"><span data-stu-id="2f3a0-104">JavaScript provides several built-in objects that you can use in your Office Scripts, regardless of whether you're scripting in JavaScript or [TypeScript](../overview/code-editor-environment.md) (a superset of JavaScript).</span></span> <span data-ttu-id="2f3a0-105">Este artigo descreve como você pode usar alguns dos objetos JavaScript internos em scripts do Office para Excel na Web.</span><span class="sxs-lookup"><span data-stu-id="2f3a0-105">This article describes how you can use some of the built-in JavaScript objects in Office Scripts for Excel on the web.</span></span>

> [!NOTE]
> <span data-ttu-id="2f3a0-106">Para obter uma lista completa de todos os objetos JavaScript internos, consulte o artigo sobre [objetos internos padrão](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects) do Mozilla.</span><span class="sxs-lookup"><span data-stu-id="2f3a0-106">For a complete list of all built-in JavaScript objects, see Mozilla's [Standard built-in objects](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects) article.</span></span>

## <a name="array"></a><span data-ttu-id="2f3a0-107">Matriz</span><span class="sxs-lookup"><span data-stu-id="2f3a0-107">Array</span></span>

<span data-ttu-id="2f3a0-108">O objeto [array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) oferece uma maneira padronizada de trabalhar com matrizes no seu script.</span><span class="sxs-lookup"><span data-stu-id="2f3a0-108">The [Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) object provides a standardized way to work with arrays in your script.</span></span> <span data-ttu-id="2f3a0-109">Embora matrizes sejam construções JavaScript padrão, elas se relacionam aos scripts do Office de duas maneiras principais: intervalos e coleções.</span><span class="sxs-lookup"><span data-stu-id="2f3a0-109">While arrays are standard JavaScript constructs, they relate to Office Scripts in two major ways: ranges and collections.</span></span>

### <a name="working-with-ranges"></a><span data-ttu-id="2f3a0-110">Trabalhar com intervalos</span><span class="sxs-lookup"><span data-stu-id="2f3a0-110">Working with ranges</span></span>

<span data-ttu-id="2f3a0-111">Intervalos contêm várias matrizes bidimensionais que mapeiam diretamente para as células naquele intervalo.</span><span class="sxs-lookup"><span data-stu-id="2f3a0-111">Ranges contain several two-dimensional arrays that directly map to the cells in that range.</span></span> <span data-ttu-id="2f3a0-112">Elas incluem propriedades como `values`, `formulas`e. `numberFormat`</span><span class="sxs-lookup"><span data-stu-id="2f3a0-112">These include properties such as `values`, `formulas`, and `numberFormat`.</span></span> <span data-ttu-id="2f3a0-113">As propriedades do tipo matriz devem ser [carregadas](scripting-fundamentals.md#sync-and-load) como qualquer outra propriedade.</span><span class="sxs-lookup"><span data-stu-id="2f3a0-113">Array-type properties must be [loaded](scripting-fundamentals.md#sync-and-load) like any other properties.</span></span>

<span data-ttu-id="2f3a0-114">O script a seguir pesquisa o intervalo **a1: D4** para qualquer formato de número que contenha um "$".</span><span class="sxs-lookup"><span data-stu-id="2f3a0-114">The following script searches the **A1:D4** range for any number format containing a "$".</span></span> <span data-ttu-id="2f3a0-115">O script define a cor de preenchimento dessas células como "amarelo".</span><span class="sxs-lookup"><span data-stu-id="2f3a0-115">The script sets the fill color in those cells to "yellow".</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
  // Get the range From A1 to D4.
  let range = context.workbook.worksheets.getActiveWorksheet().getRange("A1:D4");

  // Load the numberFormat property on the range.
  range.load("numberFormat");
  await context.sync();

  // Iterate through the arrays of rows and columns corresponding to those in the range.
  range.numberFormat.forEach((rowItem, rowIndex) => {
    range.numberFormat[rowIndex].forEach((columnItem, columnIndex) => {
      // Treat the numberFormat as a string so we can do text comparisons.
      let columnItemText = columnItem as string;
      if (columnItemText.indexOf("$") >= 0) {
        // Set the cell's fill to yellow.
        range.getCell(rowIndex, columnIndex).format.fill.color = "yellow";
      }
    });
  });
}
```

### <a name="working-with-collections"></a><span data-ttu-id="2f3a0-116">Trabalhar com coleções</span><span class="sxs-lookup"><span data-stu-id="2f3a0-116">Working with collections</span></span>

<span data-ttu-id="2f3a0-117">Muitos objetos do Excel estão contidos em uma coleção.</span><span class="sxs-lookup"><span data-stu-id="2f3a0-117">Many Excel objects are contained in a collection.</span></span> <span data-ttu-id="2f3a0-118">Por exemplo, todas as [formas](/javascript/api/office-scripts/excel/excel.shape) em uma planilha estão contidas em uma [ShapeCollection](/javascript/api/office-scripts/excel/excel.shapecollection) ( `Worksheet.shapes` como a propriedade).</span><span class="sxs-lookup"><span data-stu-id="2f3a0-118">For example, all [Shapes](/javascript/api/office-scripts/excel/excel.shape) in a worksheet are contained in a [ShapeCollection](/javascript/api/office-scripts/excel/excel.shapecollection) (as the `Worksheet.shapes` property).</span></span> <span data-ttu-id="2f3a0-119">Cada `*Collection` objeto contém uma `items` Propriedade, que é uma matriz que armazena os objetos dentro dessa coleção.</span><span class="sxs-lookup"><span data-stu-id="2f3a0-119">Each `*Collection` object contains an `items` property, which is an array that stores the objects inside that collection.</span></span> <span data-ttu-id="2f3a0-120">Isso pode ser tratado como uma matriz JavaScript normal, mas os itens da coleção precisam ser carregados primeiro.</span><span class="sxs-lookup"><span data-stu-id="2f3a0-120">This can be treated like a normal JavaScript array, but the items in the collection have to first be loaded.</span></span> <span data-ttu-id="2f3a0-121">Se você precisar trabalhar com uma propriedade em cada objeto da coleção, use uma instrução de carga hierárquica (`items/propertyName`).</span><span class="sxs-lookup"><span data-stu-id="2f3a0-121">If you need to work with a property on every object in the collection, use a hierarchal load statement (`items/propertyName`).</span></span>

<span data-ttu-id="2f3a0-122">O script a seguir registra o tipo de todas as formas na planilha atual.</span><span class="sxs-lookup"><span data-stu-id="2f3a0-122">The following script logs the type of every shape in the current worksheet.</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
  // Get the current worksheet.
  let selectedSheet = context.workbook.worksheets.getActiveWorksheet();

  // Get the shapes in this worksheet.
  let shapes = selectedSheet.shapes;
  shapes.load("items/type");
  await context.sync();

  // Log the type of every shape in the collection.
  shapes.items.forEach((shape) => {
    console.log(shape.type);
  });
}
```

<span data-ttu-id="2f3a0-123">Você pode carregar objetos individuais de uma coleção usando os `getItem` métodos `getItemAt` ou.</span><span class="sxs-lookup"><span data-stu-id="2f3a0-123">You can load individual objects from a collection using the `getItem` or `getItemAt` methods.</span></span> <span data-ttu-id="2f3a0-124">`getItem`Obtém um objeto usando um identificador exclusivo como um nome (esses nomes geralmente são especificados pelo script).</span><span class="sxs-lookup"><span data-stu-id="2f3a0-124">`getItem` gets an object by using a unique identifier like a name (such names are often specified by your script).</span></span> <span data-ttu-id="2f3a0-125">`getItemAt`Obtém um objeto usando seu índice na coleção.</span><span class="sxs-lookup"><span data-stu-id="2f3a0-125">`getItemAt` gets an object by using its index in the collection.</span></span> <span data-ttu-id="2f3a0-126">Qualquer uma das chamadas deve ser seguida `await context.sync();` por um comando para que o objeto possa ser usado.</span><span class="sxs-lookup"><span data-stu-id="2f3a0-126">Either call must be followed by a `await context.sync();` command before the object can be used.</span></span>

<span data-ttu-id="2f3a0-127">O script a seguir exclui a forma mais antiga na planilha atual.</span><span class="sxs-lookup"><span data-stu-id="2f3a0-127">The following script deletes the oldest shape in the current worksheet.</span></span>

```Typescript
async function main(context: Excel.RequestContext) {
  // Get the current worksheet.
  let selectedSheet = context.workbook.worksheets.getActiveWorksheet();

  // Get the first (oldest) shape in the worksheet.
  // Note that this script will thrown an error if there are no shapes.
  let shape = selectedSheet.shapes.getItemAt(0);

  // Sync to load `shape` from the collection.
  await context.sync();

  // Remove the shape from the worksheet.
  shape.delete();
}
```

## <a name="date"></a><span data-ttu-id="2f3a0-128">Data</span><span class="sxs-lookup"><span data-stu-id="2f3a0-128">Date</span></span>

<span data-ttu-id="2f3a0-129">O objeto [Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date) fornece uma maneira padronizada para trabalhar com datas no seu script.</span><span class="sxs-lookup"><span data-stu-id="2f3a0-129">The [Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date) object provides a standardized way to work with dates in your script.</span></span> <span data-ttu-id="2f3a0-130">`Date.now()`gera um objeto com data e hora atuais, o que é útil ao adicionar carimbos de data/hora à entrada de dados do script.</span><span class="sxs-lookup"><span data-stu-id="2f3a0-130">`Date.now()` generates an object with the current date and time, which is useful when adding timestamps to your script's data entry.</span></span>

<span data-ttu-id="2f3a0-131">O script a seguir adiciona a data atual à planilha.</span><span class="sxs-lookup"><span data-stu-id="2f3a0-131">The following script adds the current date to the worksheet.</span></span> <span data-ttu-id="2f3a0-132">Observe que, usando o `toLocaleDateString` método, o Excel reconhece o valor como uma data e altera automaticamente o formato de número da célula.</span><span class="sxs-lookup"><span data-stu-id="2f3a0-132">Note that by using the `toLocaleDateString` method, Excel recognizes the value as a date and changes the number format of the cell automatically.</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
  // Get the range for cell A1.
  let range = context.workbook.worksheets.getActiveWorksheet().getRange("A1");

  // Get the current date and time.
  let date = new Date(Date.now());

  // Set the value at A1 to the current date, using a localized string.
  range.values = [[date.toLocaleDateString()]];
}
```

## <a name="math"></a><span data-ttu-id="2f3a0-133">Matemática</span><span class="sxs-lookup"><span data-stu-id="2f3a0-133">Math</span></span>

<span data-ttu-id="2f3a0-134">O objeto [Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math) fornece métodos e constantes para operações matemáticas comuns.</span><span class="sxs-lookup"><span data-stu-id="2f3a0-134">The [Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math) object provides methods and constants for common mathematical operations.</span></span> <span data-ttu-id="2f3a0-135">Eles fornecem muitas funções também disponíveis no Excel, sem a necessidade de usar o mecanismo de cálculo da pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="2f3a0-135">These provide many functions also available in Excel, without the need to use the workbook's calculation engine.</span></span> <span data-ttu-id="2f3a0-136">Isso salva o script de ter que consultar a pasta de trabalho, o que melhora o desempenho.</span><span class="sxs-lookup"><span data-stu-id="2f3a0-136">This saves your script from having to query the workbook, which improves performance.</span></span>

<span data-ttu-id="2f3a0-137">O script a seguir `Math.min` usa para localizar e registrar o menor número no intervalo **a1: D4** .</span><span class="sxs-lookup"><span data-stu-id="2f3a0-137">The following script uses `Math.min` to find and log the smallest number in the **A1:D4** range.</span></span> <span data-ttu-id="2f3a0-138">Observe que este exemplo pressupõe que o intervalo inteiro contenha apenas números, e não cadeias de caracteres.</span><span class="sxs-lookup"><span data-stu-id="2f3a0-138">Note that this sample assumes the entire range contains only numbers, not strings.</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
  // Get the range from A1 to D4.
  let comparisonRange = context.workbook.worksheets.getActiveWorksheet().getRange("A1:D4");
  
  // Load the range's values.
  comparisonRange.load("values");
  await context.sync();

  // Set the minimum values as the first value.
  let minimum = comparisonRange.values[0][0];

  // Iterate over each row looking for the smallest value.
  comparisonRange.values.forEach((rowItem, rowIndex) => {
    // Iterate over each column looking for the smallest value.
    comparisonRange.values[rowIndex].forEach((columnItem) => {
      // Use `Math.min` to set the smallest value as either the current cell's value or the previous minimum.
      minimum = Math.min(minimum, columnItem);
    });
  });
  
  console.log(minimum);
}

```

## <a name="see-also"></a><span data-ttu-id="2f3a0-139">Confira também</span><span class="sxs-lookup"><span data-stu-id="2f3a0-139">See also</span></span>

- [<span data-ttu-id="2f3a0-140">Objetos internos padrão</span><span class="sxs-lookup"><span data-stu-id="2f3a0-140">Standard built-in objects</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects)
- [<span data-ttu-id="2f3a0-141">Ambiente do editor de código de scripts do Office</span><span class="sxs-lookup"><span data-stu-id="2f3a0-141">Office Scripts Code Editor environment</span></span>](../overview/code-editor-environment.md)
