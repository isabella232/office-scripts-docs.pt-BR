---
title: Diferenças entre scripts do Office e suplementos do Office
description: As diferenças de comportamento e API entre scripts do Office e suplementos do Office.
ms.date: 12/12/2019
localization_priority: Normal
ms.openlocfilehash: 4626afb66b54c94a72f29b039c601435c089d64d
ms.sourcegitcommit: b075eed5a6f275274fbbf6d62633219eac416f26
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/10/2020
ms.locfileid: "42700065"
---
# <a name="differences-between-office-scripts-and-office-add-ins"></a><span data-ttu-id="56d9e-103">Diferenças entre scripts do Office e suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="56d9e-103">Differences between Office Scripts and Office Add-ins</span></span>

<span data-ttu-id="56d9e-104">Os suplementos do Office e os scripts do Office têm muito em comum.</span><span class="sxs-lookup"><span data-stu-id="56d9e-104">Office Add-ins and Office Scripts have a lot in common.</span></span> <span data-ttu-id="56d9e-105">Ambas oferecem controle automatizado de uma pasta de trabalho do `Excel` Excel por meio do namespace da API JavaScript do Office.</span><span class="sxs-lookup"><span data-stu-id="56d9e-105">They both offer automated control of an Excel workbook through the `Excel` namespace of the Office JavaScript API.</span></span> <span data-ttu-id="56d9e-106">No entanto, os scripts do Office são mais limitados em seu escopo.</span><span class="sxs-lookup"><span data-stu-id="56d9e-106">However, Office Scripts are more limited in their scope.</span></span>

<span data-ttu-id="56d9e-107">Os scripts do Office são executados para conclusão com um pressionamento de botão manual, enquanto os suplementos do Office dependem da interação do usuário e continuam enquanto a pasta de trabalho está em uso.</span><span class="sxs-lookup"><span data-stu-id="56d9e-107">Office Scripts run to completion with a manual button press, whereas Office Add-ins rely on user interaction and persist while the workbook is in use.</span></span> <span data-ttu-id="56d9e-108">Se você descobrir que sua extensão do Excel precisa exceder os recursos da plataforma de script, visite a [documentação de suplementos do Office](/office/dev/add-ins) para saber mais sobre os suplementos do Office.</span><span class="sxs-lookup"><span data-stu-id="56d9e-108">If you find that your Excel extension needs to exceed the scripting platform's capabilities, visit the [Office Add-ins documentation](/office/dev/add-ins) to learn more about Office Add-ins.</span></span>

<span data-ttu-id="56d9e-109">O restante deste artigo descreve as principais diferenças entre os suplementos do Office e os scripts do Office.</span><span class="sxs-lookup"><span data-stu-id="56d9e-109">The rest of this article describes on the main differences between Office Add-ins and Office Scripts.</span></span>

## <a name="platform-support"></a><span data-ttu-id="56d9e-110">Suporte à plataforma</span><span class="sxs-lookup"><span data-stu-id="56d9e-110">Platform Support</span></span>

<span data-ttu-id="56d9e-111">Os suplementos do Office são de plataforma cruzada.</span><span class="sxs-lookup"><span data-stu-id="56d9e-111">Office Add-ins are cross-platform.</span></span> <span data-ttu-id="56d9e-112">Eles funcionam nas plataformas de área de trabalho do Windows, Mac, iOS e Web e fornecem a mesma experiência em cada.</span><span class="sxs-lookup"><span data-stu-id="56d9e-112">They work across Windows desktop, Mac, iOS, and web platforms and provide the same experience on each.</span></span> <span data-ttu-id="56d9e-113">Qualquer exceção a isso é indicada na documentação da API individual.</span><span class="sxs-lookup"><span data-stu-id="56d9e-113">Any exception to this is noted in the documentation of the individual API.</span></span>

<span data-ttu-id="56d9e-114">Atualmente, os scripts do Office só têm suporte no Excel na Web.</span><span class="sxs-lookup"><span data-stu-id="56d9e-114">Office Scripts are currently only supported by for Excel on the web.</span></span> <span data-ttu-id="56d9e-115">Toda gravação, edição e execução é feita na plataforma da Web.</span><span class="sxs-lookup"><span data-stu-id="56d9e-115">All recording, editing, and running is done on the web platform.</span></span>

## <a name="apis"></a><span data-ttu-id="56d9e-116">APIs</span><span class="sxs-lookup"><span data-stu-id="56d9e-116">APIs</span></span>

<span data-ttu-id="56d9e-117">Os scripts do Office oferecem suporte à maioria das APIs JavaScript do Excel, o que significa que há muita sobreposição de funcionalidade entre as duas plataformas.</span><span class="sxs-lookup"><span data-stu-id="56d9e-117">Office Scripts support most of the Excel JavaScript APIs, which means there's  a lot of functionality overlap between the two platforms.</span></span> <span data-ttu-id="56d9e-118">Há duas exceções: eventos e APIs comuns.</span><span class="sxs-lookup"><span data-stu-id="56d9e-118">There are two exceptions: events and Common APIs.</span></span>

### <a name="events"></a><span data-ttu-id="56d9e-119">Eventos</span><span class="sxs-lookup"><span data-stu-id="56d9e-119">Events</span></span>

<span data-ttu-id="56d9e-120">Scripts do Office não dão suporte a [eventos](/office/dev/add-ins/excel/excel-add-ins-events).</span><span class="sxs-lookup"><span data-stu-id="56d9e-120">Office Scripts do not support [events](/office/dev/add-ins/excel/excel-add-ins-events).</span></span> <span data-ttu-id="56d9e-121">Cada script executa o código em um único `main` método e, em seguida, termina.</span><span class="sxs-lookup"><span data-stu-id="56d9e-121">Every script runs the code in a single `main` method, then ends.</span></span> <span data-ttu-id="56d9e-122">Ele não reativa quando os eventos são acionados e, portanto, não podem registrar eventos.</span><span class="sxs-lookup"><span data-stu-id="56d9e-122">It does not reactivate when events are triggered, and thus, cannot register events.</span></span>

### <a name="common-apis"></a><span data-ttu-id="56d9e-123">APIs comuns</span><span class="sxs-lookup"><span data-stu-id="56d9e-123">Common APIs</span></span>

<span data-ttu-id="56d9e-124">Scripts do Office não podem usar [APIs comuns](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="56d9e-124">Office Scripts cannot use [Common APIs](/javascript/api/office).</span></span> <span data-ttu-id="56d9e-125">Se você precisar de autenticação, de janelas de diálogo ou de outros recursos que são suportados apenas por APIs comuns, provavelmente precisará criar um suplemento do Office em vez de um script do Office.</span><span class="sxs-lookup"><span data-stu-id="56d9e-125">If you need authentication, dialog windows, or other features that are only supported by Common APIs, you'll likely need to create an Office Add-in instead of an Office Script.</span></span>

## <a name="see-also"></a><span data-ttu-id="56d9e-126">Confira também</span><span class="sxs-lookup"><span data-stu-id="56d9e-126">See also</span></span>

- [<span data-ttu-id="56d9e-127">Scripts do Office no Excel na Web</span><span class="sxs-lookup"><span data-stu-id="56d9e-127">Office Scripts in Excel on the web</span></span>](../overview/excel.md)
- [<span data-ttu-id="56d9e-128">Solucionando problemas de scripts do Office</span><span class="sxs-lookup"><span data-stu-id="56d9e-128">Troubleshooting Office Scripts</span></span>](../testing/troubleshooting.md)
- [<span data-ttu-id="56d9e-129">Criar um suplemento do painel de tarefas do Excel</span><span class="sxs-lookup"><span data-stu-id="56d9e-129">Build an Excel task pane add-in</span></span>](/office/dev/add-ins/quickstarts/excel-quickstart-jquery)