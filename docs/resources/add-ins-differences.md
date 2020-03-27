---
title: Diferenças entre scripts do Office e suplementos do Office
description: As diferenças de comportamento e API entre scripts do Office e suplementos do Office.
ms.date: 03/23/2020
localization_priority: Normal
ms.openlocfilehash: 2290d4e34b7a7286d67443de9e9c64bad4fcd4b7
ms.sourcegitcommit: d556aaefac80e55f53ac56b7f6ecbc657ebd426f
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/26/2020
ms.locfileid: "42978696"
---
# <a name="differences-between-office-scripts-and-office-add-ins"></a><span data-ttu-id="d23ae-103">Diferenças entre scripts do Office e suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="d23ae-103">Differences between Office Scripts and Office Add-ins</span></span>

<span data-ttu-id="d23ae-104">Os suplementos do Office e os scripts do Office têm muito em comum.</span><span class="sxs-lookup"><span data-stu-id="d23ae-104">Office Add-ins and Office Scripts have a lot in common.</span></span> <span data-ttu-id="d23ae-105">Ambas oferecem controle automatizado de uma pasta de trabalho do `Excel` Excel por meio do namespace da API JavaScript do Office.</span><span class="sxs-lookup"><span data-stu-id="d23ae-105">They both offer automated control of an Excel workbook through the `Excel` namespace of the Office JavaScript API.</span></span> <span data-ttu-id="d23ae-106">No entanto, os scripts do Office são mais limitados em seu escopo.</span><span class="sxs-lookup"><span data-stu-id="d23ae-106">However, Office Scripts are more limited in their scope.</span></span>

![Um diagrama de quatro quadrantes mostrando as áreas de foco para diferentes soluções de extensibilidade do Office.](../images/office-programmability-diagram.png)

<span data-ttu-id="d23ae-109">Os scripts do Office são executados para conclusão com um botão manual ou como uma etapa da [automatização de energia](https://flow.microsoft.com/), enquanto os suplementos do Office são persistentes enquanto seus painéis de tarefas estão abertos.</span><span class="sxs-lookup"><span data-stu-id="d23ae-109">Office Scripts run to completion with a manual button press or as a step in [Power Automate](https://flow.microsoft.com/), whereas Office Add-ins persist while their task panes are open.</span></span> <span data-ttu-id="d23ae-110">Isso significa que os suplementos podem manter o estado durante uma sessão, enquanto os scripts do Office não mantêm um estado interno entre as execuções.</span><span class="sxs-lookup"><span data-stu-id="d23ae-110">This means the add-ins can maintain state during a session, whereas Office Scripts do not maintain an internal state between runs.</span></span> <span data-ttu-id="d23ae-111">Se você descobrir que sua extensão do Excel precisa exceder os recursos da plataforma de script, visite a [documentação de suplementos do Office](/office/dev/add-ins) para saber mais sobre os suplementos do Office.</span><span class="sxs-lookup"><span data-stu-id="d23ae-111">If you find that your Excel extension needs to exceed the scripting platform's capabilities, visit the [Office Add-ins documentation](/office/dev/add-ins) to learn more about Office Add-ins.</span></span>

<span data-ttu-id="d23ae-112">O restante deste artigo descreve as principais diferenças entre os suplementos do Office e os scripts do Office.</span><span class="sxs-lookup"><span data-stu-id="d23ae-112">The rest of this article describes on the main differences between Office Add-ins and Office Scripts.</span></span>

## <a name="platform-support"></a><span data-ttu-id="d23ae-113">Suporte à plataforma</span><span class="sxs-lookup"><span data-stu-id="d23ae-113">Platform Support</span></span>

<span data-ttu-id="d23ae-114">Os suplementos do Office são de plataforma cruzada.</span><span class="sxs-lookup"><span data-stu-id="d23ae-114">Office Add-ins are cross-platform.</span></span> <span data-ttu-id="d23ae-115">Eles funcionam nas plataformas de área de trabalho do Windows, Mac, iOS e Web e fornecem a mesma experiência em cada.</span><span class="sxs-lookup"><span data-stu-id="d23ae-115">They work across Windows desktop, Mac, iOS, and web platforms and provide the same experience on each.</span></span> <span data-ttu-id="d23ae-116">Qualquer exceção a isso é indicada na documentação da API individual.</span><span class="sxs-lookup"><span data-stu-id="d23ae-116">Any exception to this is noted in the documentation of the individual API.</span></span>

<span data-ttu-id="d23ae-117">Atualmente, os scripts do Office só têm suporte no Excel na Web.</span><span class="sxs-lookup"><span data-stu-id="d23ae-117">Office Scripts are currently only supported by for Excel on the web.</span></span> <span data-ttu-id="d23ae-118">Toda gravação, edição e execução é feita na plataforma da Web.</span><span class="sxs-lookup"><span data-stu-id="d23ae-118">All recording, editing, and running is done on the web platform.</span></span>

## <a name="apis"></a><span data-ttu-id="d23ae-119">APIs</span><span class="sxs-lookup"><span data-stu-id="d23ae-119">APIs</span></span>

<span data-ttu-id="d23ae-120">Os scripts do Office oferecem suporte à maioria das APIs JavaScript do Excel, o que significa que há muita sobreposição de funcionalidade entre as duas plataformas.</span><span class="sxs-lookup"><span data-stu-id="d23ae-120">Office Scripts support most of the Excel JavaScript APIs, which means there's  a lot of functionality overlap between the two platforms.</span></span> <span data-ttu-id="d23ae-121">Há duas exceções: eventos e APIs comuns.</span><span class="sxs-lookup"><span data-stu-id="d23ae-121">There are two exceptions: events and Common APIs.</span></span>

### <a name="events"></a><span data-ttu-id="d23ae-122">Eventos</span><span class="sxs-lookup"><span data-stu-id="d23ae-122">Events</span></span>

<span data-ttu-id="d23ae-123">Scripts do Office não dão suporte a [eventos](/office/dev/add-ins/excel/excel-add-ins-events).</span><span class="sxs-lookup"><span data-stu-id="d23ae-123">Office Scripts do not support [events](/office/dev/add-ins/excel/excel-add-ins-events).</span></span> <span data-ttu-id="d23ae-124">Cada script executa o código em um único `main` método e, em seguida, termina.</span><span class="sxs-lookup"><span data-stu-id="d23ae-124">Every script runs the code in a single `main` method, then ends.</span></span> <span data-ttu-id="d23ae-125">Ele não reativa quando os eventos são acionados e, portanto, não podem registrar eventos.</span><span class="sxs-lookup"><span data-stu-id="d23ae-125">It does not reactivate when events are triggered, and thus, cannot register events.</span></span>

### <a name="common-apis"></a><span data-ttu-id="d23ae-126">APIs comuns</span><span class="sxs-lookup"><span data-stu-id="d23ae-126">Common APIs</span></span>

<span data-ttu-id="d23ae-127">Scripts do Office não podem usar [APIs comuns](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="d23ae-127">Office Scripts cannot use [Common APIs](/javascript/api/office).</span></span> <span data-ttu-id="d23ae-128">Se você precisar de autenticação, de janelas de diálogo ou de outros recursos que são suportados apenas por APIs comuns, provavelmente precisará criar um suplemento do Office em vez de um script do Office.</span><span class="sxs-lookup"><span data-stu-id="d23ae-128">If you need authentication, dialog windows, or other features that are only supported by Common APIs, you'll likely need to create an Office Add-in instead of an Office Script.</span></span>

## <a name="see-also"></a><span data-ttu-id="d23ae-129">Confira também</span><span class="sxs-lookup"><span data-stu-id="d23ae-129">See also</span></span>

- [<span data-ttu-id="d23ae-130">Scripts do Office no Excel na Web</span><span class="sxs-lookup"><span data-stu-id="d23ae-130">Office Scripts in Excel on the web</span></span>](../overview/excel.md)
- [<span data-ttu-id="d23ae-131">Diferenças entre scripts do Office e macros VBA</span><span class="sxs-lookup"><span data-stu-id="d23ae-131">Differences between Office Scripts and VBA macros</span></span>](vba-differences.md)
- [<span data-ttu-id="d23ae-132">Solução de problemas dos scripts do Office</span><span class="sxs-lookup"><span data-stu-id="d23ae-132">Troubleshooting Office Scripts</span></span>](../testing/troubleshooting.md)
- [<span data-ttu-id="d23ae-133">Criar um suplemento do painel de tarefas do Excel</span><span class="sxs-lookup"><span data-stu-id="d23ae-133">Build an Excel task pane add-in</span></span>](/office/dev/add-ins/quickstarts/excel-quickstart-jquery)
