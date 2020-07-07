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
# <a name="run-office-scripts-with-power-automate"></a><span data-ttu-id="50e4f-103">Executar scripts do Office com automatização de energia</span><span class="sxs-lookup"><span data-stu-id="50e4f-103">Run Office Scripts with Power Automate</span></span>

<span data-ttu-id="50e4f-104">A [automatização de energia](https://flow.microsoft.com) permite que você adicione scripts do Office a um fluxo de trabalho maior e automatizado.</span><span class="sxs-lookup"><span data-stu-id="50e4f-104">[Power Automate](https://flow.microsoft.com) lets you add Office Scripts to a larger, automated workflow.</span></span> <span data-ttu-id="50e4f-105">Você pode usar a automatização de energia, como adicionar o conteúdo de um email à tabela de uma planilha ou criar ações em suas ferramentas de gerenciamento de projeto com base nos comentários da pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="50e4f-105">You can use Power Automate do things like add the contents of an email to a worksheet's table or create actions in your project management tools based on workbook comments.</span></span> <span data-ttu-id="50e4f-106">Se você for novo para a automatização de energia, recomendamos [a visita de introdução à automatização de energia](/power-automate/getting-started).</span><span class="sxs-lookup"><span data-stu-id="50e4f-106">If you are new to Power Automate, we recommend visiting [Get started with Power Automate](/power-automate/getting-started).</span></span> <span data-ttu-id="50e4f-107">Lá, você pode saber mais sobre como automatizar seus fluxos de trabalho em vários serviços.</span><span class="sxs-lookup"><span data-stu-id="50e4f-107">There, you can learn more about automating your workflows across multiple services.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="50e4f-108">No momento, não é possível executar scripts do Office a partir de um [fluxo compartilhado](/power-automate/share-buttons).</span><span class="sxs-lookup"><span data-stu-id="50e4f-108">Currently, you can't run Office Scripts from a [shared flow](/power-automate/share-buttons).</span></span> <span data-ttu-id="50e4f-109">Somente o usuário que criou um script pode executá-lo, mesmo através da automatização de energia.</span><span class="sxs-lookup"><span data-stu-id="50e4f-109">Only the user who created a script can run it, even through Power Automate.</span></span>

## <a name="getting-started"></a><span data-ttu-id="50e4f-110">Introdução</span><span class="sxs-lookup"><span data-stu-id="50e4f-110">Getting started</span></span>

<span data-ttu-id="50e4f-111">Para começar a combinar os scripts do Office e automatizados de energia, siga o tutorial [começar a usar scripts com a automatização de energia](../tutorials/excel-power-automate-manual.md).</span><span class="sxs-lookup"><span data-stu-id="50e4f-111">To begin combining Power Automate and Office Scripts, follow the tutorial [Start using scripts with Power Automate](../tutorials/excel-power-automate-manual.md).</span></span> <span data-ttu-id="50e4f-112">Isso ensina como criar um fluxo que chama um script simples.</span><span class="sxs-lookup"><span data-stu-id="50e4f-112">This will teach you how to create a flow that calls a simple script.</span></span> <span data-ttu-id="50e4f-113">Depois de concluir o tutorial e [executar automaticamente os scripts com o tutorial automatizar de energia](../tutorials/excel-power-automate-trigger.md) , retorne aqui para obter informações detalhadas sobre como conectar scripts do Office para automatizar fluxos de energia.</span><span class="sxs-lookup"><span data-stu-id="50e4f-113">After you've completed that tutorial and the [Automatically run scripts with Power Automate](../tutorials/excel-power-automate-trigger.md) tutorial, return here for detailed information about connecting Office Scripts to Power Automate flows.</span></span>

## <a name="excel-online-business-connector"></a><span data-ttu-id="50e4f-114">Conector do Excel online (comercial)</span><span class="sxs-lookup"><span data-stu-id="50e4f-114">Excel Online (Business) connector</span></span>

<span data-ttu-id="50e4f-115">Os [conectores](/connectors/connectors) são as pontes entre automatização e aplicativos.</span><span class="sxs-lookup"><span data-stu-id="50e4f-115">[Connectors](/connectors/connectors) are the bridges between Power Automate and applications.</span></span> <span data-ttu-id="50e4f-116">O [conector do Excel online (Business)](/connectors/excelonlinebusiness) fornece aos seus fluxos acesso às pastas de trabalho do Excel.</span><span class="sxs-lookup"><span data-stu-id="50e4f-116">The [Excel Online (Business) connector](/connectors/excelonlinebusiness) gives your flows access to Excel workbooks.</span></span> <span data-ttu-id="50e4f-117">A ação "executar script" permite chamar qualquer script do Office acessível por meio da pasta de trabalho selecionada.</span><span class="sxs-lookup"><span data-stu-id="50e4f-117">The "Run script" action lets you call any Office Script accessible through the selected workbook.</span></span> <span data-ttu-id="50e4f-118">Não só é possível executar scripts por meio de um fluxo, você pode passar dados de e para a pasta de trabalho com o fluxo pelos scripts.</span><span class="sxs-lookup"><span data-stu-id="50e4f-118">Not only can you run scripts through a flow, you can pass data to and from the workbook with the flow through the scripts.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="50e4f-119">A ação "executar script" fornece às pessoas que usam o conector Excel acesso significativo à sua pasta de trabalho e seus dados.</span><span class="sxs-lookup"><span data-stu-id="50e4f-119">The "Run script" action gives people who use the Excel connector significant access to your workbook and its data.</span></span> <span data-ttu-id="50e4f-120">Além disso, há riscos de segurança com scripts que fazem chamadas à API externa, conforme explicado em [chamadas externas da automatização de energia](external-calls.md).</span><span class="sxs-lookup"><span data-stu-id="50e4f-120">Additionally, there are security risks with scripts that make external API calls, as explained in [External calls from Power Automate](external-calls.md).</span></span> <span data-ttu-id="50e4f-121">Se seu administrador estiver preocupado com a exposição de dados altamente confidenciais, eles poderão desativar o conector do Excel online ou restringir o acesso a scripts do Office por meio dos [controles de administrador de scripts do Office](https://support.microsoft.com/office/19d3c51a-6ca2-40ab-978d-60fa49554dcf).</span><span class="sxs-lookup"><span data-stu-id="50e4f-121">If your admin is concerned with the exposure of highly sensitive data, they can either turn off the Excel Online connector or restrict access to Office Scripts through the [Office Scripts administrator controls](https://support.microsoft.com/office/19d3c51a-6ca2-40ab-978d-60fa49554dcf).</span></span>

## <a name="passing-data-from-power-automate-into-a-script"></a><span data-ttu-id="50e4f-122">Passar dados da energia automatizar para um script</span><span class="sxs-lookup"><span data-stu-id="50e4f-122">Passing data from Power Automate into a script</span></span>

<span data-ttu-id="50e4f-123">Todas as entradas de script são especificadas como parâmetros adicionais para a `main` função.</span><span class="sxs-lookup"><span data-stu-id="50e4f-123">All script input is specified as additional parameters for the `main` function.</span></span> <span data-ttu-id="50e4f-124">Por exemplo, se você quisesse que um script aceita um `string` que representa um nome como entrada, você alteraria a `main` assinatura para `function main(workbook: ExcelScript.Workbook, name: string)` .</span><span class="sxs-lookup"><span data-stu-id="50e4f-124">For example, if you wanted a script to accept a `string` that represents a name as input, you would change the `main` signature to `function main(workbook: ExcelScript.Workbook, name: string)`.</span></span>

<span data-ttu-id="50e4f-125">Quando você estiver configurando um fluxo em automatização de energia, poderá especificar a entrada de script como valores estáticos, [expressões](/power-automate/use-expressions-in-conditions)ou conteúdo dinâmico.</span><span class="sxs-lookup"><span data-stu-id="50e4f-125">When you're configuring a flow in Power Automate, you can specify script input as static values, [expressions](/power-automate/use-expressions-in-conditions), or dynamic content.</span></span> <span data-ttu-id="50e4f-126">Os detalhes sobre o conector de um serviço individual podem ser encontrados na [documentação do conector automatizado de energia](/connectors/).</span><span class="sxs-lookup"><span data-stu-id="50e4f-126">Details on an individual service's connector can be found in the [Power Automate Connector documentation](/connectors/).</span></span>

<span data-ttu-id="50e4f-127">Ao adicionar parâmetros de entrada para a função de um script `main` , considere as seguintes permissões e restrições.</span><span class="sxs-lookup"><span data-stu-id="50e4f-127">When adding input parameters to a script's `main` function, consider the following allowances and restrictions.</span></span>

1. <span data-ttu-id="50e4f-128">O primeiro parâmetro deve ser do tipo `ExcelScript.Workbook` .</span><span class="sxs-lookup"><span data-stu-id="50e4f-128">The first parameter must be of type `ExcelScript.Workbook`.</span></span> <span data-ttu-id="50e4f-129">O nome do parâmetro não importa.</span><span class="sxs-lookup"><span data-stu-id="50e4f-129">Its parameter name does not matter.</span></span>

2. <span data-ttu-id="50e4f-130">Todo parâmetro deve ter um tipo.</span><span class="sxs-lookup"><span data-stu-id="50e4f-130">Every parameter must have a type.</span></span>

3. <span data-ttu-id="50e4f-131">Os tipos básicos,,,,, `string` `number` `boolean` `any` `unknown` `object` e `undefined` são suportados.</span><span class="sxs-lookup"><span data-stu-id="50e4f-131">The basic types `string`, `number`, `boolean`, `any`, `unknown`, `object`, and `undefined` are supported.</span></span>

4. <span data-ttu-id="50e4f-132">Há suporte para matrizes dos tipos básicos listados anteriormente.</span><span class="sxs-lookup"><span data-stu-id="50e4f-132">Arrays of the previously listed basic types are supported.</span></span>

5. <span data-ttu-id="50e4f-133">Há suporte para matrizes aninhadas como parâmetros (mas não como tipos de retorno).</span><span class="sxs-lookup"><span data-stu-id="50e4f-133">Nested arrays are supported as parameters (but not as return types).</span></span>

6. <span data-ttu-id="50e4f-134">Os tipos de União são permitidos se eles forem uma União de literais pertencentes a um único tipo ( `string` , `number` , ou `boolean` ).</span><span class="sxs-lookup"><span data-stu-id="50e4f-134">Union types are allowed if they are a union of literals belonging to a single type (`string`, `number`, or `boolean`).</span></span> <span data-ttu-id="50e4f-135">Também há suporte para Undefined de um tipo com suporte.</span><span class="sxs-lookup"><span data-stu-id="50e4f-135">Unions of a supported type with undefined are also supported.</span></span>

7. <span data-ttu-id="50e4f-136">Os tipos de objeto são permitidos se contiverem Propriedades de tipo `string` , `number` , `boolean` matrizes com suporte ou outros objetos com suporte.</span><span class="sxs-lookup"><span data-stu-id="50e4f-136">Object types are allowed if they contain properties of type `string`, `number`, `boolean`, supported arrays, or other supported objects.</span></span> <span data-ttu-id="50e4f-137">O exemplo a seguir mostra objetos aninhados suportados como tipos de parâmetros:</span><span class="sxs-lookup"><span data-stu-id="50e4f-137">The following example shows nested objects that are supported as parameter types:</span></span>

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

8. <span data-ttu-id="50e4f-138">Os objetos devem ter sua definição de interface ou de classe definida no script.</span><span class="sxs-lookup"><span data-stu-id="50e4f-138">Objects must have their interface or class definition defined in the script.</span></span> <span data-ttu-id="50e4f-139">Um objeto também pode ser definido de forma anônima, como no exemplo a seguir:</span><span class="sxs-lookup"><span data-stu-id="50e4f-139">An object can also be defined anonymously inline, as in the following example:</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook): {name: string, email: string}
    ```

9. <span data-ttu-id="50e4f-140">Parâmetros opcionais são permitidos e podem ser indicados por meio do modificador opcional `?` (por exemplo, `function main(workbook: ExcelScript.Workbook, Name?: string)` ).</span><span class="sxs-lookup"><span data-stu-id="50e4f-140">Optional parameters are allowed and can be denoted as such by using the optional modifier `?` (for example, `function main(workbook: ExcelScript.Workbook, Name?: string)`).</span></span>

10. <span data-ttu-id="50e4f-141">Os valores de parâmetro padrão são permitidos (por exemplo `async function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')` .</span><span class="sxs-lookup"><span data-stu-id="50e4f-141">Default parameter values are allowed (for example `async function main(workbook: ExcelScript.Workbook, Name: string = 'Jane Doe')`.</span></span>

## <a name="returning-data-from-a-script-back-to-power-automate"></a><span data-ttu-id="50e4f-142">Retornando dados de um script de volta para automatizar a energia</span><span class="sxs-lookup"><span data-stu-id="50e4f-142">Returning data from a script back to Power Automate</span></span>

<span data-ttu-id="50e4f-143">Os scripts podem retornar dados da pasta de trabalho para serem usados como conteúdo dinâmico em um fluxo automatizado de energia.</span><span class="sxs-lookup"><span data-stu-id="50e4f-143">Scripts can return data from the workbook to be used as dynamic content in a Power Automate flow.</span></span> <span data-ttu-id="50e4f-144">Como nos parâmetros de entrada, a automatização de energia coloca algumas restrições no tipo de retorno.</span><span class="sxs-lookup"><span data-stu-id="50e4f-144">As with input parameters, Power Automate places some restrictions on the return type.</span></span>

1. <span data-ttu-id="50e4f-145">Os tipos básicos `string` , `number` , `boolean` , `void` e `undefined` são suportados.</span><span class="sxs-lookup"><span data-stu-id="50e4f-145">The basic types `string`, `number`, `boolean`, `void`, and `undefined` are supported.</span></span>

2. <span data-ttu-id="50e4f-146">Tipos de União usados como tipos de retorno seguem as mesmas restrições que eles fazem quando usados como parâmetros de script.</span><span class="sxs-lookup"><span data-stu-id="50e4f-146">Union types used as return types follow the same restrictions as they do when used as script parameters.</span></span>

3. <span data-ttu-id="50e4f-147">Tipos de matriz são permitidos se forem do tipo `string` , `number` ou `boolean` .</span><span class="sxs-lookup"><span data-stu-id="50e4f-147">Array types are allowed if they are of type `string`, `number`, or `boolean`.</span></span> <span data-ttu-id="50e4f-148">Eles também são permitidos se o tipo for um tipo de União ou tipo literal suportado.</span><span class="sxs-lookup"><span data-stu-id="50e4f-148">They are also allowed if the type is a supported union or supported literal type.</span></span>

4. <span data-ttu-id="50e4f-149">Tipos de objeto usados como tipos de retorno seguem as mesmas restrições que eles fazem quando usados como parâmetros de script.</span><span class="sxs-lookup"><span data-stu-id="50e4f-149">Object types used as return types follow the same restrictions as they do when used as script parameters.</span></span>

5. <span data-ttu-id="50e4f-150">Há suporte para digitação implícita, embora seja necessário seguir as mesmas regras que um tipo definido.</span><span class="sxs-lookup"><span data-stu-id="50e4f-150">Implicit typing is supported, though it must follow the same rules as a defined type.</span></span>

## <a name="avoid-using-relative-references"></a><span data-ttu-id="50e4f-151">Evitar o uso de referências relativas</span><span class="sxs-lookup"><span data-stu-id="50e4f-151">Avoid using relative references</span></span>

<span data-ttu-id="50e4f-152">A automatização de energia executa o script na pasta de trabalho do Excel escolhida em seu nome.</span><span class="sxs-lookup"><span data-stu-id="50e4f-152">Power Automate runs your script in the chosen Excel workbook on your behalf.</span></span> <span data-ttu-id="50e4f-153">A pasta de trabalho pode ser fechada quando isso acontecer.</span><span class="sxs-lookup"><span data-stu-id="50e4f-153">The workbook might be closed when this happens.</span></span> <span data-ttu-id="50e4f-154">Qualquer API que se baseia no estado atual do usuário, como `Workbook.getActiveWorksheet` , falhará quando for executada através da automatização de energia.</span><span class="sxs-lookup"><span data-stu-id="50e4f-154">Any API that relies on the user's current state, such as `Workbook.getActiveWorksheet`, will fail when run through Power Automate.</span></span> <span data-ttu-id="50e4f-155">Ao projetar seus scripts, certifique-se de usar referências absolutas para planilhas e intervalos.</span><span class="sxs-lookup"><span data-stu-id="50e4f-155">When designing your scripts, be sure to use absolute references for worksheets and ranges.</span></span>

<span data-ttu-id="50e4f-156">As funções a seguir apresentarão um erro e falharão quando chamadas de um script em um fluxo automatizado de energia.</span><span class="sxs-lookup"><span data-stu-id="50e4f-156">The following functions will throw an error and fail when called from a script in a Power Automate flow.</span></span>

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

## <a name="example"></a><span data-ttu-id="50e4f-157">Exemplo</span><span class="sxs-lookup"><span data-stu-id="50e4f-157">Example</span></span>

<span data-ttu-id="50e4f-158">A captura de tela a seguir mostra um fluxo automatizado de energia que é disparado sempre que um problema do [GitHub](https://github.com/) é atribuído a você.</span><span class="sxs-lookup"><span data-stu-id="50e4f-158">The following screenshot shows a Power Automate flow that's triggered whenever a [GitHub](https://github.com/) issue is assigned to you.</span></span> <span data-ttu-id="50e4f-159">O fluxo executa um script que adiciona o problema a uma tabela em uma pasta de trabalho do Excel.</span><span class="sxs-lookup"><span data-stu-id="50e4f-159">The flow runs a script that adds the issue to a table in an Excel workbook.</span></span> <span data-ttu-id="50e4f-160">Se houver cinco ou mais problemas nessa tabela, o fluxo enviará um lembrete por email.</span><span class="sxs-lookup"><span data-stu-id="50e4f-160">If there are five or more issues in that table, the flow sends an email reminder.</span></span>

![O fluxo de exemplo mostrado no editor de fluxo automatizar energia.](../images/power-automate-parameter-return-sample.png)

<span data-ttu-id="50e4f-162">A `main` função do script especifica a ID do problema e o título do problema como parâmetros de entrada, e o script retorna o número de linhas na tabela de saída.</span><span class="sxs-lookup"><span data-stu-id="50e4f-162">The `main` function of the script specifies the issue ID and issue title as input parameters, and the script returns the number of rows in the issue table.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="50e4f-163">Confira também</span><span class="sxs-lookup"><span data-stu-id="50e4f-163">See also</span></span>

- [<span data-ttu-id="50e4f-164">Executar scripts do Office no Excel na Web com a automatização de energia</span><span class="sxs-lookup"><span data-stu-id="50e4f-164">Run Office Scripts in Excel on the web with Power Automate</span></span>](../tutorials/excel-power-automate-manual.md)
- [<span data-ttu-id="50e4f-165">Executar automaticamente scripts com Power Automate</span><span class="sxs-lookup"><span data-stu-id="50e4f-165">Automatically run scripts with Power Automate</span></span>](../tutorials/excel-power-automate-trigger.md)
- [<span data-ttu-id="50e4f-166">Fundamentos de script para scripts do Office no Excel na Web</span><span class="sxs-lookup"><span data-stu-id="50e4f-166">Scripting fundamentals for Office Scripts in Excel on the web</span></span>](scripting-fundamentals.md)
- [<span data-ttu-id="50e4f-167">Começar a usar o Power Automate</span><span class="sxs-lookup"><span data-stu-id="50e4f-167">Get started with Power Automate</span></span>](/power-automate/getting-started)
- [<span data-ttu-id="50e4f-168">Documentação de referência do conector do Excel online (Business)</span><span class="sxs-lookup"><span data-stu-id="50e4f-168">Excel Online (Business) connector reference documentation</span></span>](/connectors/excelonlinebusiness/)
