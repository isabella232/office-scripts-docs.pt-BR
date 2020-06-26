---
title: Suporte à chamada de API externa em scripts do Office
description: Suporte e orientação para fazer chamadas de API externa em um script do Office.
ms.date: 06/25/2020
localization_priority: Normal
ms.openlocfilehash: ec8281551cbe7c500eee40ec86067e5efbfcfc31
ms.sourcegitcommit: aec3c971c6640429f89b6bb99d2c95ea06725599
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/25/2020
ms.locfileid: "44878693"
---
# <a name="external-api-call-support-in-office-scripts"></a>Suporte à chamada de API externa em scripts do Office

A plataforma de scripts do Office não dá suporte a chamadas para [APIs externas](https://developer.mozilla.org/docs/Web/API). No entanto, essas chamadas podem ser executadas sob as circunstâncias certas. Chamadas externas só podem ser feitas por meio do cliente Excel, não através da automatização de energia [sob circunstâncias normais](#external-calls-from-power-automate).

Os autores de script devem esperar um comportamento consistente ao usar APIs externas durante a fase de visualização da plataforma. Isso se deve ao modo como o tempo de execução do JavaScript gerencia a interação com a pasta de trabalho. O script pode terminar antes que a chamada da API seja concluída (ou sua `Promise` está totalmente resolvida). Assim, não confie em APIs externas para cenários de script críticos.

> [!CAUTION]
> As chamadas externas podem resultar em dados confidenciais expostos a pontos de extremidade indesejados. Seu administrador pode estabelecer proteção de firewall contra essas chamadas.

## <a name="definition-files-for-external-apis"></a>Arquivos de definição para APIs externas

Os arquivos de definição para APIs externas não estão incluídos em scripts do Office. O uso dessas APIs gera erros de tempo de compilação para definições ausentes. As APIs ainda são executadas (embora somente quando executadas pelo cliente do Excel), conforme mostrado no seguinte script:

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

## <a name="external-calls-from-power-automate"></a>Chamadas externas da automatização de energia

Qualquer chamada de API externa falha quando um script é executado com automatização de energia. Essa é uma diferença comportamental entre a execução de um script por meio do cliente do Excel e através da automatização de energia. Certifique-se de verificar os scripts para essas referências antes de criá-las em um fluxo.

> [!WARNING]
> A falha de chamadas externas do [Excel online Connector](/connectors/excelonlinebusiness) em energia automatizada está lá para ajudar a sustentar as políticas de prevenção de perda de dados existentes. No entanto, os scripts executados através da automatização de energia estão prontos para fora da sua organização e fora dos firewalls da sua organização. Para obter proteção adicional de usuários mal-intencionados nesse ambiente externo, seu administrador pode controlar o uso de scripts do Office. O administrador pode desabilitar o conector do Excel online para automatizar ou desativar scripts do Office para Excel na Web por meio dos [controles de administrador de scripts do Office](https://support.microsoft.com/office/19d3c51a-6ca2-40ab-978d-60fa49554dcf).

## <a name="see-also"></a>Confira também

- [Usar objetos internos do JavaScript nos scripts do Office](javascript-objects.md)