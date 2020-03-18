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
# <a name="differences-between-office-scripts-and-office-add-ins"></a>Diferenças entre scripts do Office e suplementos do Office

Os suplementos do Office e os scripts do Office têm muito em comum. Ambas oferecem controle automatizado de uma pasta de trabalho do `Excel` Excel por meio do namespace da API JavaScript do Office. No entanto, os scripts do Office são mais limitados em seu escopo.

Os scripts do Office são executados para conclusão com um pressionamento de botão manual, enquanto os suplementos do Office dependem da interação do usuário e continuam enquanto a pasta de trabalho está em uso. Se você descobrir que sua extensão do Excel precisa exceder os recursos da plataforma de script, visite a [documentação de suplementos do Office](/office/dev/add-ins) para saber mais sobre os suplementos do Office.

O restante deste artigo descreve as principais diferenças entre os suplementos do Office e os scripts do Office.

## <a name="platform-support"></a>Suporte à plataforma

Os suplementos do Office são de plataforma cruzada. Eles funcionam nas plataformas de área de trabalho do Windows, Mac, iOS e Web e fornecem a mesma experiência em cada. Qualquer exceção a isso é indicada na documentação da API individual.

Atualmente, os scripts do Office só têm suporte no Excel na Web. Toda gravação, edição e execução é feita na plataforma da Web.

## <a name="apis"></a>APIs

Os scripts do Office oferecem suporte à maioria das APIs JavaScript do Excel, o que significa que há muita sobreposição de funcionalidade entre as duas plataformas. Há duas exceções: eventos e APIs comuns.

### <a name="events"></a>Eventos

Scripts do Office não dão suporte a [eventos](/office/dev/add-ins/excel/excel-add-ins-events). Cada script executa o código em um único `main` método e, em seguida, termina. Ele não reativa quando os eventos são acionados e, portanto, não podem registrar eventos.

### <a name="common-apis"></a>APIs comuns

Scripts do Office não podem usar [APIs comuns](/javascript/api/office). Se você precisar de autenticação, de janelas de diálogo ou de outros recursos que são suportados apenas por APIs comuns, provavelmente precisará criar um suplemento do Office em vez de um script do Office.

## <a name="see-also"></a>Confira também

- [Scripts do Office no Excel na Web](../overview/excel.md)
- [Solucionando problemas de scripts do Office](../testing/troubleshooting.md)
- [Criar um suplemento do painel de tarefas do Excel](/office/dev/add-ins/quickstarts/excel-quickstart-jquery)