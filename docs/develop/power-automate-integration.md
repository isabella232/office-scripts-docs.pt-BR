---
title: Executar scripts do Office com automatização de energia
description: Como obter scripts do Office para Excel na Web trabalhando com um fluxo de trabalho automatizado de energia.
ms.date: 07/10/2020
localization_priority: Normal
ms.openlocfilehash: bd8fea08b7a9303ad2ceace787de6457a33fb979
ms.sourcegitcommit: ebd1079c7e2695ac0e7e4c616f2439975e196875
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/17/2020
ms.locfileid: "45160443"
---
# <a name="run-office-scripts-with-power-automate"></a>Executar scripts do Office com automatização de energia

A [automatização de energia](https://flow.microsoft.com) permite que você adicione scripts do Office a um fluxo de trabalho maior e automatizado. Você pode usar a automatização de energia, como adicionar o conteúdo de um email à tabela de uma planilha ou criar ações em suas ferramentas de gerenciamento de projeto com base nos comentários da pasta de trabalho.

## <a name="getting-started"></a>Introdução

Se você for novo para a automatização de energia, recomendamos [a visita de introdução à automatização de energia](/power-automate/getting-started). Lá, você pode saber mais sobre todas as possibilidades de automação disponíveis para você. Os documentos aqui se concentram em como os scripts do Office trabalham com a automatização de energia e como isso pode ajudar a melhorar a experiência do Excel.

Para começar a combinar os scripts do Office e automatizados de energia, siga o tutorial [começar a usar scripts com a automatização de energia](../tutorials/excel-power-automate-manual.md). Isso ensina como criar um fluxo que chama um script simples. Após concluir o tutorial e a passagem dos [dados para scripts em um tutorial de fluxo automático automatizado de energia automatizada](../tutorials/excel-power-automate-trigger.md) , retorne aqui para obter informações detalhadas sobre como conectar scripts do Office para automatizar fluxos de energia.

## <a name="excel-online-business-connector"></a>Conector do Excel online (comercial)

Os [conectores](/connectors/connectors) são as pontes entre automatização e aplicativos. O [conector do Excel online (Business)](/connectors/excelonlinebusiness) fornece aos seus fluxos acesso às pastas de trabalho do Excel. A ação "executar script" permite chamar qualquer script do Office acessível por meio da pasta de trabalho selecionada. Não só é possível executar scripts por meio de um fluxo, você pode passar dados de e para a pasta de trabalho com o fluxo pelos scripts.

> [!IMPORTANT]
> A ação "executar script" fornece às pessoas que usam o conector Excel acesso significativo à sua pasta de trabalho e seus dados. Além disso, há riscos de segurança com scripts que fazem chamadas à API externa, conforme explicado em [chamadas externas da automatização de energia](external-calls.md). Se seu administrador estiver preocupado com a exposição de dados altamente confidenciais, eles poderão desativar o conector do Excel online ou restringir o acesso a scripts do Office por meio dos [controles de administrador de scripts do Office](https://support.microsoft.com/office/19d3c51a-6ca2-40ab-978d-60fa49554dcf).

## <a name="data-transfer-in-flows-for-scripts"></a>Transferência de dados em fluxos para scripts

A automatização de energia permite que você passe dados entre as etapas do seu fluxo. Os scripts podem ser configurados para aceitar qualquer tipo de informação que você precisa e retornar qualquer coisa da sua pasta de trabalho que você deseja em seu fluxo. A entrada para o seu script é especificada adicionando parâmetros à `main` função (além de `workbook: ExcelScript.Workbook` ). A saída do script é declarada pela adição de um tipo de retorno a `main` .

> [!NOTE]
> Quando você cria um bloco de "script de execução" em seu fluxo, os parâmetros aceitos e os tipos retornados são preenchidos. Se você alterar os parâmetros ou retornar tipos de seu script, será necessário refazer o bloco "executar script" do seu fluxo. Isso garante que os dados estão sendo analisados corretamente.

As seções a seguir abrangem os detalhes de entrada e saída para scripts usados na automatização de energia. Se você gostaria de obter uma abordagem prática para aprender este tópico, experimente os dados de [passagem para scripts em um tutorial de fluxo automático automatizado de fluxo](../tutorials/excel-power-automate-trigger.md) automático ou explore o cenário de exemplo de [lembretes automáticos de tarefas](../resources/scenarios/task-reminders.md) .

### <a name="main-parameters-passing-data-to-a-script"></a>`main`Parâmetros: passagem de dados para um script

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

### <a name="returning-data-from-a-script"></a>Retornar dados de um script

Os scripts podem retornar dados da pasta de trabalho para serem usados como conteúdo dinâmico em um fluxo automatizado de energia. Como nos parâmetros de entrada, a automatização de energia coloca algumas restrições no tipo de retorno.

1. Os tipos básicos `string` , `number` , `boolean` , `void` e `undefined` são suportados.

2. Tipos de União usados como tipos de retorno seguem as mesmas restrições que eles fazem quando usados como parâmetros de script.

3. Tipos de matriz são permitidos se forem do tipo `string` , `number` ou `boolean` . Eles também são permitidos se o tipo for um tipo de União ou tipo literal suportado.

4. Tipos de objeto usados como tipos de retorno seguem as mesmas restrições que eles fazem quando usados como parâmetros de script.

5. Há suporte para digitação implícita, embora seja necessário seguir as mesmas regras que um tipo definido.

## <a name="avoid-using-relative-references"></a>Evitar o uso de referências relativas

A automatização de energia executa o script na pasta de trabalho do Excel escolhida em seu nome. A pasta de trabalho pode ser fechada quando isso acontecer. Qualquer API que se baseia no estado atual do usuário, como `Workbook.getActiveWorksheet` , falhará quando for executada através da automatização de energia. Ao projetar seus scripts, certifique-se de usar referências absolutas para planilhas e intervalos.

Os métodos a seguir gerarão um erro e falharão quando chamados de um script em um fluxo automatizado de energia.

| Classe | Método |
|--|--|
| [Gráfico](/javascript/api/office-scripts/excelscript/excelscript.chart) | `activate` |
| [Range](/javascript/api/office-scripts/excelscript/excelscript.range) | `select` |
| [Pasta de trabalho](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveCell` |
| [Pasta de trabalho](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveChart` |
| [Pasta de trabalho](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveSlicer` |
| [Pasta de trabalho](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveWorksheet` |
| [Pasta de trabalho](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRange` |
| [Pasta de trabalho](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRanges` |
| [Planilha](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `activate` |

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
- [Transmitir dados para scripts em um fluxo automático de energia de execução automatizada](../tutorials/excel-power-automate-trigger.md)
- [Fundamentos de script para scripts do Office no Excel na Web](scripting-fundamentals.md)
- [Começar a usar o Power Automate](/power-automate/getting-started)
- [Documentação de referência do conector do Excel online (Business)](/connectors/excelonlinebusiness/)
