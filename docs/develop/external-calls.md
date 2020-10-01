---
title: Chamada de API externa nos scripts do Office
description: Suporte e orientação para fazer chamadas de API externa em um script do Office.
ms.date: 09/24/2020
localization_priority: Normal
ms.openlocfilehash: fa77e606e2b3ab90144507660d71561b278e82e5
ms.sourcegitcommit: ce72354381561dc167ea0092efd915642a9161b3
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/30/2020
ms.locfileid: "48319627"
---
# <a name="external-api-call-support-in-office-scripts"></a><span data-ttu-id="d3eac-103">Chamada de API externa nos scripts do Office</span><span class="sxs-lookup"><span data-stu-id="d3eac-103">External API call support in Office Scripts</span></span>

<span data-ttu-id="d3eac-104">A plataforma de scripts do Office não dá suporte a chamadas para [APIs externas](https://developer.mozilla.org/docs/Web/API).</span><span class="sxs-lookup"><span data-stu-id="d3eac-104">The Office Scripts platform doesn't support calls to [external APIs](https://developer.mozilla.org/docs/Web/API).</span></span> <span data-ttu-id="d3eac-105">No entanto, essas chamadas podem ser executadas sob as circunstâncias certas.</span><span class="sxs-lookup"><span data-stu-id="d3eac-105">However, these calls can be run under the right circumstances.</span></span> <span data-ttu-id="d3eac-106">Chamadas externas só podem ser feitas por meio do cliente Excel, não através da automatização de energia [sob circunstâncias normais](#external-calls-from-power-automate).</span><span class="sxs-lookup"><span data-stu-id="d3eac-106">External calls can be only be made through the Excel client, not through Power Automate [under normal circumstances](#external-calls-from-power-automate).</span></span>

<span data-ttu-id="d3eac-107">Os autores de script devem esperar um comportamento consistente ao usar APIs externas durante a fase de visualização da plataforma.</span><span class="sxs-lookup"><span data-stu-id="d3eac-107">Script authors shouldn't expect consistent behavior when using external APIs during the platform's preview phase.</span></span> <span data-ttu-id="d3eac-108">Isso se deve ao modo como o tempo de execução do JavaScript gerencia a interação com a pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="d3eac-108">This is due how the JavaScript runtime manages interacting with the workbook.</span></span> <span data-ttu-id="d3eac-109">O script pode terminar antes que a chamada da API seja concluída (ou sua `Promise` está totalmente resolvida).</span><span class="sxs-lookup"><span data-stu-id="d3eac-109">The script may end before the API call completes (or its `Promise` is fully resolved).</span></span> <span data-ttu-id="d3eac-110">Assim, não confie em APIs externas para cenários de script críticos.</span><span class="sxs-lookup"><span data-stu-id="d3eac-110">As such, do not rely on external APIs for critical script scenarios.</span></span>

> [!CAUTION]
> <span data-ttu-id="d3eac-111">As chamadas externas podem resultar em dados confidenciais expostos a pontos de extremidade indesejados.</span><span class="sxs-lookup"><span data-stu-id="d3eac-111">External calls may result in sensitive data being exposed to undesirable endpoints.</span></span> <span data-ttu-id="d3eac-112">Seu administrador pode estabelecer proteção de firewall contra essas chamadas.</span><span class="sxs-lookup"><span data-stu-id="d3eac-112">Your admin can establish firewall protection against such calls.</span></span>

## <a name="definition-files-for-external-apis"></a><span data-ttu-id="d3eac-113">Arquivos de definição para APIs externas</span><span class="sxs-lookup"><span data-stu-id="d3eac-113">Definition files for external APIs</span></span>

<span data-ttu-id="d3eac-114">Os arquivos de definição para APIs externas não estão incluídos em scripts do Office.</span><span class="sxs-lookup"><span data-stu-id="d3eac-114">The definition files for external APIs aren't included with Office Scripts.</span></span> <span data-ttu-id="d3eac-115">O uso dessas APIs gera erros de tempo de compilação para definições ausentes.</span><span class="sxs-lookup"><span data-stu-id="d3eac-115">The use of such APIs generates compile-time errors for missing definitions.</span></span> <span data-ttu-id="d3eac-116">As APIs ainda são executadas (embora somente quando executadas pelo cliente do Excel), conforme mostrado no seguinte script:</span><span class="sxs-lookup"><span data-stu-id="d3eac-116">The APIs still run (though only when run through the Excel client), as shown in the following script:</span></span>

```typescript
async function main(workbook: ExcelScript.Workbook): Promise <void> {
  /* The following line of code generates the error:
   * "Cannot find name 'fetch'".
   * It will still run and return the JSON from the testing service.
   */
  let fetchResult = await fetch('https://jsonplaceholder.typicode.com/todos/1');
  let json = await fetchResult.json();

  // Displays the content from https://jsonplaceholder.typicode.com/todos/1
  console.log(JSON.stringify(json));
}
```

## <a name="external-calls-from-power-automate"></a><span data-ttu-id="d3eac-117">Chamadas externas da automatização de energia</span><span class="sxs-lookup"><span data-stu-id="d3eac-117">External calls from Power Automate</span></span>

<span data-ttu-id="d3eac-118">Qualquer chamada de API externa falha quando um script é executado com automatização de energia.</span><span class="sxs-lookup"><span data-stu-id="d3eac-118">Any external API calls fail when a script is run with Power Automate.</span></span> <span data-ttu-id="d3eac-119">Essa é uma diferença comportamental entre a execução de um script por meio do cliente do Excel e através da automatização de energia.</span><span class="sxs-lookup"><span data-stu-id="d3eac-119">This is a behavioral difference between running a script through the Excel client and through Power Automate.</span></span> <span data-ttu-id="d3eac-120">Certifique-se de verificar os scripts para essas referências antes de criá-las em um fluxo.</span><span class="sxs-lookup"><span data-stu-id="d3eac-120">Be sure to check your scripts for such references before building them into a flow.</span></span>

> [!WARNING]
> <span data-ttu-id="d3eac-121">A falha de chamadas externas do [Excel online Connector](/connectors/excelonlinebusiness) em energia automatizada está lá para ajudar a sustentar as políticas de prevenção de perda de dados existentes.</span><span class="sxs-lookup"><span data-stu-id="d3eac-121">The failure of external calls [Excel Online connector](/connectors/excelonlinebusiness) in Power Automate is there to help uphold existing data loss prevention policies.</span></span> <span data-ttu-id="d3eac-122">No entanto, os scripts executados através da automatização de energia estão prontos para fora da sua organização e fora dos firewalls da sua organização.</span><span class="sxs-lookup"><span data-stu-id="d3eac-122">However, the scripts run through Power Automate are done so outside of your organization, and outside of your organization's firewalls.</span></span> <span data-ttu-id="d3eac-123">Para obter proteção adicional de usuários mal-intencionados nesse ambiente externo, seu administrador pode controlar o uso de scripts do Office.</span><span class="sxs-lookup"><span data-stu-id="d3eac-123">For additional protection from malicious users in this external environment, your admin can control the use of Office Scripts.</span></span> <span data-ttu-id="d3eac-124">O administrador pode desabilitar o conector do Excel online para automatizar ou desativar scripts do Office para Excel na Web por meio dos [controles de administrador de scripts do Office](/microsoft-365/admin/manage/manage-office-scripts-settings).</span><span class="sxs-lookup"><span data-stu-id="d3eac-124">Your admin can either disable the Excel Online connector in Power Automate or turn off Office Scripts for Excel on the web through the [Office Scripts administrator controls](/microsoft-365/admin/manage/manage-office-scripts-settings).</span></span>

## <a name="see-also"></a><span data-ttu-id="d3eac-125">Confira também</span><span class="sxs-lookup"><span data-stu-id="d3eac-125">See also</span></span>

- [<span data-ttu-id="d3eac-126">Usar objetos internos do JavaScript nos scripts do Office</span><span class="sxs-lookup"><span data-stu-id="d3eac-126">Using built-in JavaScript objects in Office Scripts</span></span>](javascript-objects.md)