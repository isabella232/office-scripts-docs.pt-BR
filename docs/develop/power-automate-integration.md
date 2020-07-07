---
title: Executar scripts do Office com automatização de energia
description: Como obter scripts do Office para Excel na Web trabalhando com um fluxo de trabalho automatizado de energia.
ms.date: 06/29/2020
localization_priority: Normal
ms.openlocfilehash: 0ea58324998d23020e04cb37dfeea065791757f5
ms.sourcegitcommit: bf9f33c37c6f7805d6b408aa648bb9785a7cd133
ms.contentlocale: pt-BR
ms.lasthandoff: 07/06/2020
ms.locfileid: "45043381"
---
# <a name="run-office-scripts-with-power-automate"></a>Executar scripts do Office com automatização de energia

A [automatização de energia](https://flow.microsoft.com) permite que você adicione scripts do Office a um fluxo de trabalho maior e automatizado. Você pode usar a automatização de energia, como adicionar o conteúdo de um email à tabela de uma planilha ou criar ações em suas ferramentas de gerenciamento de projeto com base nos comentários da pasta de trabalho. Se você for novo para a automatização de energia, recomendamos [a visita de introdução à automatização de energia](/power-automate/getting-started). Lá, você pode saber mais sobre como automatizar seus fluxos de trabalho em vários serviços.

> [!IMPORTANT]
> No momento, não é possível executar scripts do Office a partir de um [fluxo compartilhado](/power-automate/share-buttons). Somente o usuário que criou um script pode executá-lo, mesmo através da automatização de energia.

## <a name="getting-started"></a>Introdução

Para começar a combinar os scripts do Office e automatizados de energia, siga o tutorial [começar a usar scripts com a automatização de energia](../tutorials/excel-power-automate-manual.md). Isso ensina como criar um fluxo que chama um script simples. Depois de concluir o tutorial e [executar automaticamente os scripts com o tutorial automatizar de energia](../tutorials/excel-power-automate-trigger.md) , retorne aqui para obter informações detalhadas sobre como conectar scripts do Office para automatizar fluxos de energia.

## <a name="excel-online-business-connector"></a>Conector do Excel online (comercial)

Os [conectores](/connectors/connectors) são as pontes entre automatização e aplicativos. O [conector do Excel online (Business)](/connectors/excelonlinebusiness) fornece aos seus fluxos acesso às pastas de trabalho do Excel. A ação "executar script" permite chamar qualquer script do Office acessível por meio da pasta de trabalho selecionada. Não só é possível executar scripts por meio de um fluxo, você pode passar dados de e para a pasta de trabalho com o fluxo pelos scripts.

> [!IMPORTANT]
> A ação "executar script" fornece às pessoas que usam o conector Excel acesso significativo à sua pasta de trabalho e seus dados. Além disso, há riscos de segurança com scripts que fazem chamadas à API externa, conforme explicado em [chamadas externas da automatização de energia](external-calls.md). Se seu administrador estiver preocupado com a exposição de dados altamente confidenciais, eles poderão desativar o conector do Excel online ou restringir o acesso a scripts do Office por meio dos [controles de administrador de scripts do Office](https://support.microsoft.com/office/19d3c51a-6ca2-40ab-978d-60fa49554dcf).

## <a name="passing-data-from-power-automate-into-a-script"></a>Passar dados da energia automatizar para um script

Todas as entradas de script são especificadas como parâmetros adicionais para a `main` função. Por exemplo, se você quisesse que um script aceita um `string` que representa um nome como entrada, você alteraria a `main` assinatura para `function main(workbook: ExcelScript.Workbook, name: string)` .

Quando você estiver configurando um fluxo em automatização de energia, poderá especificar a entrada de script como valores estáticos, [expressões](/power-automate/use-expressions-in-conditions)ou conteúdo dinâmico. Os detalhes sobre o conector de um serviço individual podem ser encontrados na [documentação do conector automatizado de energia](/connectors/).

Ao adicionar parâmetros de entrada para a função de um script `main` , considere as seguintes permissões e restrições.

1. O primeiro parâmetro deve ser do tipo `ExcelScript.Workbook` . O nome do parâmetro não importa.

2. Todo parâmetro deve ter um tipo.

3. Os tipos básicos,,,,, `string` `number` `boolean` `any` `unknown` `object` e `undefined` são suportados.

4. Há suporte para matrizes dos tipos básicos listados anteriormente.

5. Há suporte para matrizes aninhadas como parâmetros (mas não como tipos de retorno).

6. Os tipos de União são permitidos se eles forem uma União de literais pertencentes a um único tipo ( `string` , `number` , ou `boolean` ). Também há suporte para Undefined de um tipo com suporte.

7. Os tipos de objeto são permitidos se contiverem Propriedades de tipo `string` , `number` , `boolean` matrizes com suporte ou outros objetos com suporte. O exemplo a seguir mostra objetos aninhados suportados como tipos de parâmetros:

    ```TypeScript
    // Office Scripts can return an Employee object because Position only contains strings and numbers.
    interface Employee {
        name: string;
        job: Position;
    }

    interface Position {
        id: number;
        title: string;
    }
    ```

8. Os objetos devem ter sua definição de interface ou de classe definida no script. Um objeto também pode ser definido de forma anônima, como no exemplo a seguir:

    ```TypeScript
    function main(workbook: ExcelScript.Workbook): {name: string, email: string}
    ```

9. Parâmetros opcionais são permitidos e podem ser indicados por meio do modificador opcional `?` (por exemplo, `function main(workbook: ExcelScript.Workbook, Name?: string)` ).

10. Os valores de parâmetro padrão são permitidos (por exemplo `async function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')` .

## <a name="returning-data-from-a-script-back-to-power-automate"></a>Retornando dados de um script de volta para automatizar a energia

Os scripts podem retornar dados da pasta de trabalho para serem usados como conteúdo dinâmico em um fluxo automatizado de energia. Como nos parâmetros de entrada, a automatização de energia coloca algumas restrições no tipo de retorno.

1. Os tipos básicos `string` , `number` , `boolean` , `void` e `undefined` são suportados.

2. Tipos de União usados como tipos de retorno seguem as mesmas restrições que eles fazem quando usados como parâmetros de script.

3. Tipos de matriz são permitidos se forem do tipo `string` , `number` ou `boolean` . Eles também são permitidos se o tipo for um tipo de União ou tipo literal suportado.

4. Tipos de objeto usados como tipos de retorno seguem as mesmas restrições que eles fazem quando usados como parâmetros de script.

5. Há suporte para digitação implícita, embora seja necessário seguir as mesmas regras que um tipo definido.

## <a name="avoid-using-relative-references"></a>Evitar o uso de referências relativas

A automatização de energia executa o script na pasta de trabalho do Excel escolhida em seu nome. A pasta de trabalho pode ser fechada quando isso acontecer. Qualquer API que se baseia no estado atual do usuário, como `Workbook.getActiveWorksheet` , falhará quando for executada através da automatização de energia. Ao projetar seus scripts, certifique-se de usar referências absolutas para planilhas e intervalos.

As funções a seguir apresentarão um erro e falharão quando chamadas de um script em um fluxo automatizado de energia.

- `Chart.activate`
- `Range.select`
- `Workbook.getActiveCell`
- `Workbook.getActiveChart`
- `Workbook.getActiveChartOrNullObject`
- `Workbook.getActiveSlicer`
- `Workbook.getActiveSlicerOrNullObject`
- `Workbook.getActiveWorksheet`
- `Workbook.getSelectedRange`
- `Workbook.getSelectedRanges`
- `Worksheet.activate`

## <a name="example"></a>Exemplo

A captura de tela a seguir mostra um fluxo automatizado de energia que é disparado sempre que um problema do [GitHub](https://github.com/) é atribuído a você. O fluxo executa um script que adiciona o problema a uma tabela em uma pasta de trabalho do Excel. Se houver cinco ou mais problemas nessa tabela, o fluxo enviará um lembrete por email.

![O fluxo de exemplo mostrado no editor de fluxo automatizar energia.](../images/power-automate-parameter-return-sample.png)

A `main` função do script especifica a ID do problema e o título do problema como parâmetros de entrada, e o script retorna o número de linhas na tabela de saída.

```TypeScript
function main(
  workbook: ExcelScript.Workbook,
  issueId: string,
  issueTitle: string): number {
  // Get the "GitHub" worksheet.
  let worksheet = workbook.getWorksheet("GitHub");

  // Get the first table in this worksheet, which contains the table of GitHub issues.
  let issueTable = worksheet.getTables()[0];

  // Add the issue ID and issue title as a row.
  issueTable.addRow(-1, [issueId, issueTitle]);

  // Return the number of rows in the table, which represents how many issues are assigned to this user.
  return issueTable.getRangeBetweenHeaderAndTotal().getRowCount();
}
```

## <a name="see-also"></a>Confira também

- [Executar scripts do Office no Excel na Web com a automatização de energia](../tutorials/excel-power-automate-manual.md)
- [Executar automaticamente scripts com Power Automate](../tutorials/excel-power-automate-trigger.md)
- [Fundamentos de script para scripts do Office no Excel na Web](scripting-fundamentals.md)
- [Começar a usar o Power Automate](/power-automate/getting-started)
- [Documentação de referência do conector do Excel online (Business)](/connectors/excelonlinebusiness/)
