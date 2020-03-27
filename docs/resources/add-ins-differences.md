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
# <a name="differences-between-office-scripts-and-office-add-ins"></a>Diferenças entre scripts do Office e suplementos do Office

Os suplementos do Office e os scripts do Office têm muito em comum. Ambas oferecem controle automatizado de uma pasta de trabalho do `Excel` Excel por meio do namespace da API JavaScript do Office. No entanto, os scripts do Office são mais limitados em seu escopo.

![Um diagrama de quatro quadrantes mostrando as áreas de foco para diferentes soluções de extensibilidade do Office. Os scripts do Office e os suplementos Web do Office são focados na Web e na colaboração, mas os scripts do Office atendem aos usuários finais (enquanto os desenvolvedores profissionais de destino dos Web suplementos do Office).)](../images/office-programmability-diagram.png)

Os scripts do Office são executados para conclusão com um botão manual ou como uma etapa da [automatização de energia](https://flow.microsoft.com/), enquanto os suplementos do Office são persistentes enquanto seus painéis de tarefas estão abertos. Isso significa que os suplementos podem manter o estado durante uma sessão, enquanto os scripts do Office não mantêm um estado interno entre as execuções. Se você descobrir que sua extensão do Excel precisa exceder os recursos da plataforma de script, visite a [documentação de suplementos do Office](/office/dev/add-ins) para saber mais sobre os suplementos do Office.

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
- [Diferenças entre scripts do Office e macros VBA](vba-differences.md)
- [Solução de problemas dos scripts do Office](../testing/troubleshooting.md)
- [Criar um suplemento do painel de tarefas do Excel](/office/dev/add-ins/quickstarts/excel-quickstart-jquery)
