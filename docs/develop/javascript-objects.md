---
title: Usar objetos internos do JavaScript nos scripts do Office
description: Como chamar APIs JavaScript internas de um script do Office no Excel na Web.
ms.date: 04/24/2020
localization_priority: Normal
ms.openlocfilehash: b5d70e77aef79c38a8cfd680c9d03bb126c402b2
ms.sourcegitcommit: aec3c971c6640429f89b6bb99d2c95ea06725599
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/25/2020
ms.locfileid: "44878532"
---
# <a name="using-built-in-javascript-objects-in-office-scripts"></a>Usar objetos internos do JavaScript nos scripts do Office

O JavaScript fornece vários objetos internos que você pode usar em seus scripts do Office, independentemente de você estar criando scripts em JavaScript ou [TypeScript](../overview/code-editor-environment.md) (um superconjunto de JavaScript). Este artigo descreve como você pode usar alguns dos objetos JavaScript internos em scripts do Office para Excel na Web.

> [!NOTE]
> Para obter uma lista completa de todos os objetos JavaScript internos, consulte o artigo sobre [objetos internos padrão](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects) do Mozilla.

## <a name="array"></a>Matriz

O objeto [array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) oferece uma maneira padronizada de trabalhar com matrizes no seu script. Embora matrizes sejam construções JavaScript padrão, elas se relacionam aos scripts do Office de duas maneiras principais: intervalos e coleções.

### <a name="working-with-ranges"></a>Trabalhar com intervalos

Intervalos contêm várias matrizes bidimensionais que mapeiam diretamente para as células naquele intervalo. Essas matrizes contêm informações específicas sobre cada célula desse intervalo. Por exemplo, `Range.getValues` retorna todos os valores dessas células (com as linhas e colunas do mapeamento de duas dimensões bidimensionais para as linhas e colunas dessa subseção de planilha). `Range.getFormulas`e `Range.getNumberFormats` são outros métodos usados com frequência que retornam matrizes, como `Range.getValues` .

O script a seguir pesquisa o intervalo **a1: D4** para qualquer formato de número que contenha um "$". O script define a cor de preenchimento dessas células como "amarelo".

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

### <a name="working-with-collections"></a>Trabalhar com coleções

Muitos objetos do Excel estão contidos em uma coleção. A coleção é gerenciada pela API de scripts do Office e exposta como uma matriz. Por exemplo, todas as [formas](/javascript/api/office-scripts/excel/excelscript.shape) em uma planilha estão contidas em um `Shape[]` que é retornado pelo `Worksheet.getShapes` método. Você pode usar essa matriz para ler valores da coleção ou pode acessar objetos específicos dos métodos do objeto pai `get*` .

> [!NOTE]
> Não adicione nem remova manualmente objetos dessas matrizes de coleção. Use os `add` métodos nos objetos pai e os `delete` métodos nos objetos do tipo coleção. Por exemplo, adicione uma [tabela](/javascript/api/office-scripts/excel/excelscript.table) a uma [planilha](/javascript/api/office-scripts/excel/excelscript.worksheet) com o `Worksheet.addTable` método e remova o `Table` usando `Table.delete` .

O script a seguir registra o tipo de todas as formas na planilha atual.

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

O script a seguir exclui a forma mais antiga na planilha atual.

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

## <a name="date"></a>Data

O objeto [Date](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date) fornece uma maneira padronizada para trabalhar com datas no seu script. `Date.now()`gera um objeto com data e hora atuais, o que é útil ao adicionar carimbos de data/hora à entrada de dados do script.

O script a seguir adiciona a data atual à planilha. Observe que, usando o `toLocaleDateString` método, o Excel reconhece o valor como uma data e altera automaticamente o formato de número da célula.

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

A seção [trabalhar com datas](../resources/excel-samples.md#work-with-dates) dos exemplos tem mais scripts relacionados a datas.

## <a name="math"></a>Matemática

O objeto [Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math) fornece métodos e constantes para operações matemáticas comuns. Eles fornecem muitas funções também disponíveis no Excel, sem a necessidade de usar o mecanismo de cálculo da pasta de trabalho. Isso salva o script de ter que consultar a pasta de trabalho, o que melhora o desempenho.

O script a seguir usa `Math.min` para localizar e registrar o menor número no intervalo **a1: D4** . Observe que este exemplo pressupõe que o intervalo inteiro contenha apenas números, e não cadeias de caracteres.

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

## <a name="use-of-external-javascript-libraries-is-not-supported"></a>Não há suporte para o uso de bibliotecas JavaScript externas

Os scripts do Office não oferecem suporte ao uso de bibliotecas externas de terceiros. O script só pode usar os objetos JavaScript internos e as APIs de scripts do Office.

## <a name="see-also"></a>Confira também

- [Objetos internos padrão](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects)
- [Ambiente do editor de código de scripts do Office](../overview/code-editor-environment.md)
