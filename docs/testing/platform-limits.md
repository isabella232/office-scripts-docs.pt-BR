---
title: Limites e requisitos de plataforma com scripts do Office
description: Limites de recurso e suporte de navegador para scripts do Office quando usados com o Excel na Web
ms.date: 07/23/2020
localization_priority: Normal
ms.openlocfilehash: 6e297cba0b9f984f2d541cc3c441a666f9ebfcef
ms.sourcegitcommit: ff7fde04ce5a66d8df06ed505951c8111e2e9833
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/11/2020
ms.locfileid: "46618153"
---
# <a name="platform-limits-and-requirements-with-office-scripts"></a><span data-ttu-id="10827-103">Limites e requisitos de plataforma com scripts do Office</span><span class="sxs-lookup"><span data-stu-id="10827-103">Platform limits and requirements with Office Scripts</span></span>

<span data-ttu-id="10827-104">Há algumas limitações de plataforma das quais você deve estar ciente ao desenvolver scripts do Office.</span><span class="sxs-lookup"><span data-stu-id="10827-104">There are some platform limitations of which you should be aware when developing Office Scripts.</span></span> <span data-ttu-id="10827-105">Este artigo detalha o suporte do navegador e os limites de dados para scripts do Office para Excel na Web.</span><span class="sxs-lookup"><span data-stu-id="10827-105">This article details the browser support and data limits for Office Scripts for Excel on the web.</span></span>

## <a name="browser-support"></a><span data-ttu-id="10827-106">Suporte do navegador</span><span class="sxs-lookup"><span data-stu-id="10827-106">Browser support</span></span>

<span data-ttu-id="10827-107">Os scripts do Office funcionam em qualquer navegador que [ofereça suporte para o Office para a Web](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452).</span><span class="sxs-lookup"><span data-stu-id="10827-107">Office Scripts work in any browser that [supports Office for the web](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452).</span></span> <span data-ttu-id="10827-108">No entanto, alguns recursos JavaScript não são compatíveis com o Internet Explorer 11 (IE 11).</span><span class="sxs-lookup"><span data-stu-id="10827-108">However, some JavaScript features aren't supported in Internet Explorer 11 (IE 11).</span></span> <span data-ttu-id="10827-109">Quaisquer recursos introduzidos no [ES6 ou posterior](https://www.w3schools.com/Js/js_es6.asp) não funcionarão com o IE 11.</span><span class="sxs-lookup"><span data-stu-id="10827-109">Any features introduced in [ES6 or later](https://www.w3schools.com/Js/js_es6.asp) won't work with IE 11.</span></span> <span data-ttu-id="10827-110">Se as pessoas na sua organização ainda usarem esse navegador, certifique-se de testar seus scripts nesse ambiente ao compartilhá-los.</span><span class="sxs-lookup"><span data-stu-id="10827-110">If people in your organization still use that browser, be sure to test your scripts in that environment when sharing them.</span></span>

### <a name="third-party-cookies"></a><span data-ttu-id="10827-111">Cookies de terceiros</span><span class="sxs-lookup"><span data-stu-id="10827-111">Third-party cookies</span></span>

<span data-ttu-id="10827-112">Seu navegador precisa de cookies de terceiros habilitados para mostrar a guia **automatizada** no Excel na Web.</span><span class="sxs-lookup"><span data-stu-id="10827-112">Your browser needs third-party cookies enabled to show the **Automate** tab in Excel on the web.</span></span> <span data-ttu-id="10827-113">Verifique as configurações do navegador se a guia não estiver sendo exibida.</span><span class="sxs-lookup"><span data-stu-id="10827-113">Check your browser settings if the tab isn't being displayed.</span></span> <span data-ttu-id="10827-114">Se você estiver usando uma sessão privada do navegador, talvez seja necessário habilitar novamente essa configuração a cada vez.</span><span class="sxs-lookup"><span data-stu-id="10827-114">If you're using a private browser session, you may need to re-enable this setting each time.</span></span>

> [!NOTE]
> <span data-ttu-id="10827-115">Alguns navegadores se referem a essa configuração como "todos os cookies", em vez de "cookies terceirizados".</span><span class="sxs-lookup"><span data-stu-id="10827-115">Some browsers refer to this setting as "all cookies", instead of "third-party cookies".</span></span>

## <a name="data-limits"></a><span data-ttu-id="10827-116">Limites de dados</span><span class="sxs-lookup"><span data-stu-id="10827-116">Data limits</span></span>

<span data-ttu-id="10827-117">Há limites para a quantidade de dados do Excel que podem ser transferidos ao mesmo tempo e quantas transações de automatização de energia individuais podem ser conduzidas.</span><span class="sxs-lookup"><span data-stu-id="10827-117">There are limits on how much Excel data can be transferred at once and how many individual Power Automate transactions can be conducted.</span></span>

### <a name="excel"></a><span data-ttu-id="10827-118">Excel</span><span class="sxs-lookup"><span data-stu-id="10827-118">Excel</span></span>

<span data-ttu-id="10827-119">O Excel para a Web tem as seguintes limitações ao fazer chamadas para a pasta de trabalho por meio de um script:</span><span class="sxs-lookup"><span data-stu-id="10827-119">Excel for the web has the following limitations when making calls to the workbook through a script:</span></span>

- <span data-ttu-id="10827-120">As solicitações e respostas são limitadas a **5 MB**.</span><span class="sxs-lookup"><span data-stu-id="10827-120">Requests and responses are limited to **5MB**.</span></span>
- <span data-ttu-id="10827-121">Um intervalo é limitado a **5 milhões células**.</span><span class="sxs-lookup"><span data-stu-id="10827-121">A range is limited to **five million cells**.</span></span>

<span data-ttu-id="10827-122">Se você estiver encontrando erros ao lidar com grandes conjuntos de grandes, tente usar vários intervalos menores em vez de intervalos maiores.</span><span class="sxs-lookup"><span data-stu-id="10827-122">If you're encountering errors when dealing with large datasets, try using multiple smaller ranges instead of larger ranges.</span></span> <span data-ttu-id="10827-123">Você também pode APIs como [Range. getSpecialCells](/javascript/api/office-scripts/excelscript/excelscript.range#getspecialcells-celltype--cellvaluetype-) para direcionar células específicas em vez de intervalos grandes.</span><span class="sxs-lookup"><span data-stu-id="10827-123">You can also APIs like [Range.getSpecialCells](/javascript/api/office-scripts/excelscript/excelscript.range#getspecialcells-celltype--cellvaluetype-) to target specific cells instead of large ranges.</span></span>

### <a name="power-automate"></a><span data-ttu-id="10827-124">Power Automate</span><span class="sxs-lookup"><span data-stu-id="10827-124">Power Automate</span></span>

<span data-ttu-id="10827-125">Ao usar scripts do Office com a automatização de energia, você está limitado a **200 chamadas por dia**.</span><span class="sxs-lookup"><span data-stu-id="10827-125">When using Office Scripts with Power Automate, you're limited to **200 calls per day**.</span></span> <span data-ttu-id="10827-126">Esse limite é redefinido às 12:00 AM UTC.</span><span class="sxs-lookup"><span data-stu-id="10827-126">This limit resets at 12:00 AM UTC.</span></span>

<span data-ttu-id="10827-127">A plataforma de automatização de energia também tem limitações de uso, que podem ser encontradas no artigo [limites e configuração da energia automatizada](/power-automate/limits-and-config).</span><span class="sxs-lookup"><span data-stu-id="10827-127">The Power Automate platform also has usage limitations, which can be found in the article [Limits and configuration in Power Automate](/power-automate/limits-and-config).</span></span>

## <a name="see-also"></a><span data-ttu-id="10827-128">Confira também</span><span class="sxs-lookup"><span data-stu-id="10827-128">See also</span></span>

- [<span data-ttu-id="10827-129">Solução de problemas dos scripts do Office</span><span class="sxs-lookup"><span data-stu-id="10827-129">Troubleshooting Office Scripts</span></span>](troubleshooting.md)
- [<span data-ttu-id="10827-130">Desfazer os efeitos de um script do Office</span><span class="sxs-lookup"><span data-stu-id="10827-130">Undo the effects of an Office Script</span></span>](undo.md)
- [<span data-ttu-id="10827-131">Melhorar o desempenho dos scripts do Office</span><span class="sxs-lookup"><span data-stu-id="10827-131">Improve the performance of your Office Scripts</span></span>](../develop/web-client-performance.md)
- [<span data-ttu-id="10827-132">Conceitos básicos de script para scripts do Office no Excel na Web</span><span class="sxs-lookup"><span data-stu-id="10827-132">Scripting Fundamentals for Office Scripts in Excel on the web</span></span>](../develop/scripting-fundamentals.md)
