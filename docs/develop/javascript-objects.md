---
title: Usar objetos internos do JavaScript nos scripts do Office
description: Como chamar APIs JavaScript internas de um script do Office no Excel na Web.
ms.date: 04/06/2020
localization_priority: Normal
ms.openlocfilehash: a4b698215edea5f266e159fee0e08690904dd379
ms.sourcegitcommit: 0b2232c4c228b14d501edb8bb489fe0e84748b42
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/08/2020
ms.locfileid: "43191011"
---
# <a name="using-built-in-javascript-objects-in-office-scripts"></a>Usar objetos internos do JavaScript nos scripts do Office

O JavaScript fornece vários objetos internos que você pode usar em seus scripts do Office, independentemente de você estar criando scripts em JavaScript ou [TypeScript](../overview/code-editor-environment.md) (um superconjunto de JavaScript). Este artigo descreve como você pode usar alguns dos objetos JavaScript internos em scripts do Office para Excel na Web.

> [!NOTE]
> Para obter uma lista completa de todos os objetos JavaScript internos, consulte o artigo sobre [objetos internos padrão](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects) do Mozilla.

## <a name="array"></a>Matriz

O objeto [array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) oferece uma maneira padronizada de trabalhar com matrizes no seu script. Embora matrizes sejam construções JavaScript padrão, elas se relacionam aos scripts do Office de duas maneiras principais: intervalos e coleções.

### <a name="working-with-ranges"></a>Trabalhar com intervalos

Intervalos contêm várias matrizes bidimensionais que mapeiam diretamente para as células naquele intervalo. Elas incluem propriedades como `values`, `formulas`e. `numberFormat` As propriedades do tipo matriz devem ser [carregadas](scripting-fundamentals.md#sync-and-load) como qualquer outra propriedade.

O script a seguir pesquisa o intervalo **a1: D4** para qualquer formato de número que contenha um "$". O script define a cor de preenchimento dessas células como "amarelo".

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

### <a name="working-with-collections"></a>Trabalhar com coleções

Muitos objetos do Excel estão contidos em uma coleção. Por exemplo, todas as [formas](/javascript/api/office-scripts/excel/excel.shape) em uma planilha estão contidas em uma [ShapeCollection](/javascript/api/office-scripts/excel/excel.shapecollection) ( `Worksheet.shapes` como a propriedade). Cada `*Collection` objeto contém uma `items` Propriedade, que é uma matriz que armazena os objetos dentro dessa coleção. Isso pode ser tratado como uma matriz JavaScript normal, mas os itens da coleção precisam ser carregados primeiro. Se você precisar trabalhar com uma propriedade em cada objeto da coleção, use uma instrução de carga hierárquica (`items/propertyName`).

O script a seguir registra o tipo de todas as formas na planilha atual.

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

Você pode carregar objetos individuais de uma coleção usando os `getItem` métodos `getItemAt` ou. `getItem`Obtém um objeto usando um identificador exclusivo como um nome (esses nomes geralmente são especificados pelo script). `getItemAt`Obtém um objeto usando seu índice na coleção. Qualquer uma das chamadas deve ser seguida `await context.sync();` por um comando para que o objeto possa ser usado.

O script a seguir exclui a forma mais antiga na planilha atual.

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

## <a name="date"></a>Data

O objeto [Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date) fornece uma maneira padronizada para trabalhar com datas no seu script. `Date.now()`gera um objeto com data e hora atuais, o que é útil ao adicionar carimbos de data/hora à entrada de dados do script.

O script a seguir adiciona a data atual à planilha. Observe que, usando o `toLocaleDateString` método, o Excel reconhece o valor como uma data e altera automaticamente o formato de número da célula.

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

A seção [trabalhar com datas](../resources/excel-samples.md#work-with-dates) dos exemplos tem mais scripts relacionados a datas.

## <a name="math"></a>Matemática

O objeto [Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math) fornece métodos e constantes para operações matemáticas comuns. Eles fornecem muitas funções também disponíveis no Excel, sem a necessidade de usar o mecanismo de cálculo da pasta de trabalho. Isso salva o script de ter que consultar a pasta de trabalho, o que melhora o desempenho.

O script a seguir `Math.min` usa para localizar e registrar o menor número no intervalo **a1: D4** . Observe que este exemplo pressupõe que o intervalo inteiro contenha apenas números, e não cadeias de caracteres.

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

## <a name="see-also"></a>Confira também

- [Objetos internos padrão](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects)
- [Ambiente do editor de código de scripts do Office](../overview/code-editor-environment.md)
