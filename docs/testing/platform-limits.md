---
title: Limites e requisitos de plataforma com scripts do Office
description: Limites de recurso e suporte de navegador para scripts do Office quando usados com o Excel na Web
ms.date: 10/23/2020
localization_priority: Normal
ms.openlocfilehash: 61f5c55be278ae056014d3b01e4176354d913f87
ms.sourcegitcommit: d3e7681e262bdccc281fcb7b3c719494202e846b
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/06/2020
ms.locfileid: "48930075"
---
# <a name="platform-limits-and-requirements-with-office-scripts"></a>Limites e requisitos de plataforma com scripts do Office

Há algumas limitações de plataforma das quais você deve estar ciente ao desenvolver scripts do Office. Este artigo detalha o suporte do navegador e os limites de dados para scripts do Office para Excel na Web.

## <a name="browser-support"></a>Suporte do navegador

Os scripts do Office funcionam em qualquer navegador que [ofereça suporte para o Office para a Web](https://support.microsoft.com/office/ad1303e0-a318-47aa-b409-d3a5eb44e452). No entanto, alguns recursos JavaScript não são compatíveis com o Internet Explorer 11 (IE 11). Quaisquer recursos introduzidos no [ES6 ou posterior](https://www.w3schools.com/Js/js_es6.asp) não funcionarão com o IE 11. Se as pessoas na sua organização ainda usarem esse navegador, certifique-se de testar seus scripts nesse ambiente ao compartilhá-los.

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

### <a name="third-party-cookies"></a>Cookies de terceiros

Seu navegador precisa de cookies de terceiros habilitados para mostrar a guia **automatizada** no Excel na Web. Verifique as configurações do navegador se a guia não estiver sendo exibida. Se você estiver usando uma sessão privada do navegador, talvez seja necessário habilitar novamente essa configuração a cada vez.

> [!NOTE]
> Alguns navegadores se referem a essa configuração como "todos os cookies", em vez de "cookies terceirizados".

#### <a name="instructions-for-adjusting-cookie-settings-in-popular-browsers"></a>Instruções para ajustar as configurações de cookies em navegadores populares

- [Chrome](https://support.google.com/chrome/answer/95647)
- [Borda](https://support.microsoft.com/microsoft-edge/temporarily-allow-cookies-and-site-data-in-microsoft-edge-597f04f2-c0ce-f08c-7c2b-541086362bd2)
- [Firefox](https://support.mozilla.org/kb/disable-third-party-cookies)
- [Safari](https://support.apple.com/guide/safari/manage-cookies-and-website-data-sfri11471/mac)

## <a name="data-limits"></a>Limites de dados

Há limites para a quantidade de dados do Excel que podem ser transferidos ao mesmo tempo e quantas transações de automatização de energia individuais podem ser conduzidas.

### <a name="excel"></a>Excel

O Excel para a Web tem as seguintes limitações ao fazer chamadas para a pasta de trabalho por meio de um script:

- As solicitações e respostas são limitadas a **5 MB**.
- Um intervalo é limitado a **5 milhões células**.

Se você estiver encontrando erros ao lidar com grandes conjuntos de grandes, tente usar vários intervalos menores em vez de intervalos maiores. Você também pode APIs como [Range. getSpecialCells](/javascript/api/office-scripts/excelscript/excelscript.range#getspecialcells-celltype--cellvaluetype-) para direcionar células específicas em vez de intervalos grandes.

### <a name="power-automate"></a>Power Automate

Ao usar scripts do Office com a automatização de energia, você está limitado a **200 chamadas por dia**. Esse limite é redefinido às 12:00 AM UTC.

A plataforma de automatização de energia também tem limitações de uso, que podem ser encontradas no artigo [limites e configuração da energia automatizada](/power-automate/limits-and-config).

## <a name="see-also"></a>Confira também

- [Solução de problemas dos scripts do Office](troubleshooting.md)
- [Desfazer os efeitos de um script do Office](undo.md)
- [Melhorar o desempenho dos scripts do Office](../develop/web-client-performance.md)
- [Conceitos básicos de script para scripts do Office no Excel na Web](../develop/scripting-fundamentals.md)
