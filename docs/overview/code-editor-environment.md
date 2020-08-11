---
title: Ambiente do editor de código de scripts do Office
description: Os pré-requisitos e as informações de ambiente para scripts do Office no Excel na Web.
ms.date: 07/23/2020
localization_priority: Normal
ms.openlocfilehash: 6a496d6c245879eae60e60b9b0cd6fced9e9259a
ms.sourcegitcommit: ff7fde04ce5a66d8df06ed505951c8111e2e9833
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/11/2020
ms.locfileid: "46616693"
---
# <a name="office-scripts-code-editor-environment"></a>Ambiente do editor de código de scripts do Office

Os scripts do Office são escritos em [TypeScript ou JavaScript](#scripting-language-typescript-or-javascript) e usam as [APIs JavaScript de scripts do Office](#office-scripts-javascript-api) para interagir com uma pasta de trabalho do Excel.

## <a name="scripting-language-typescript-or-javascript"></a>Linguagem de script: TypeScript ou JavaScript

Os scripts do Office são gravados no [TypeScript](https://www.typescriptlang.org/docs/home.html) ou no [JavaScript](https://developer.mozilla.org/docs/Web/JavaScript). O gravador de ação gera código em TypeScript (que é um superconjunto de JavaScript). A documentação de scripts do Office usa TypeScript, mas se você estiver mais confortável com JavaScript, poderá usá-lo em vez disso.

Os scripts do Office são partes de código amplamente contidas. Apenas uma pequena parte da funcionalidade do TypeScript é usada. Portanto, você pode editar scripts sem ter que aprender as complexidades do TypeScript. O editor de código também trata a instalação, a compilação e a execução de código, de modo que você não precisa se preocupar em nada, exceto no próprio script. É possível aprender o idioma e criar scripts sem conhecimento de programação anterior. No entanto, se você é novo para programação, recomendamos aprender alguns conceitos básicos antes de prosseguir com os scripts do Office:

- Saiba mais sobre o JavaScript. Você deve se familiarizar com conceitos como variáveis, fluxo de controle, funções e tipos de dados. [O Mozilla oferece um tutorial bom e abrangente sobre o JavaScript](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).
- Saiba mais sobre tipos no TypeScript. O TypeScript é criado no JavaScript garantindo no momento da compilação os tipos corretos são usados para as chamadas de método e atribuições. A documentação do TypeScript em [interfaces](https://www.typescriptlang.org/docs/handbook/interfaces.html), [classes](https://www.typescriptlang.org/docs/handbook/classes.html), [inferência de tipo](https://www.typescriptlang.org/docs/handbook/type-inference.html)e compatibilidade de [tipo](https://www.typescriptlang.org/docs/handbook/type-compatibility.html) será a mais útil.

## <a name="office-scripts-javascript-api"></a>API JavaScript de scripts do Office

Os scripts do Office usam uma versão especializada das APIs JavaScript do Office para [suplementos do Office](/office/dev/add-ins/overview/index). Embora haja semelhanças nas duas APIs, você não deve presumir que o código pode ser portado entre as duas plataformas. As diferenças entre as duas plataformas são descritas no artigo [diferenças entre scripts do Office e suplementos do Office](../resources/add-ins-differences.md#apis) . Você pode exibir todas as APIs disponíveis para o seu script na [documentação de referência da API de scripts do Office](/javascript/api/office-scripts/overview).

## <a name="intellisense"></a>Eventual

O IntelliSense é um recurso do editor de código que ajuda a evitar erros ortográficos e de sintaxe à medida que você edita o script. Exibe os possíveis nomes de objeto e campo à medida que você digita, bem como a documentação embutida para cada API.

O editor de código do Excel usa o mesmo mecanismo IntelliSense que o Visual Studio Code. Para saber mais sobre o recurso, visite os [recursos do IntelliSense do Visual Studio Code](https://code.visualstudio.com/docs/editor/intellisense#_intellisense-features).

## <a name="external-library-support"></a>Suporte à biblioteca externa

Os scripts do Office não oferecem suporte ao uso de bibliotecas JavaScript externas de terceiros. No momento, você não pode chamar nenhuma biblioteca além das APIs de scripts do Office de um script. Você ainda tem acesso a qualquer [objeto JavaScript interno](../develop/javascript-objects.md), como [Math](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Math).

## <a name="see-also"></a>Confira também

- [Referência da API de scripts do Office](/javascript/api/office-scripts/overview)
- [Solução de problemas dos scripts do Office](../testing/troubleshooting.md)
- [Usar objetos internos do JavaScript nos scripts do Office](../develop/javascript-objects.md)
