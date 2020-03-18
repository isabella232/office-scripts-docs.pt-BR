---
title: Solucionando problemas de scripts do Office
description: Dicas e técnicas de depuração para scripts do Office, bem como recursos da ajuda.
ms.date: 12/13/2019
localization_priority: Normal
ms.openlocfilehash: 959faff875f342dc1b1ab158ad9ded24732b0894
ms.sourcegitcommit: b075eed5a6f275274fbbf6d62633219eac416f26
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/10/2020
ms.locfileid: "42700052"
---
# <a name="troubleshooting-office-scripts"></a>Solucionando problemas de scripts do Office

Ao desenvolver scripts do Office, você pode cometer erros. Não há problema. Temos ferramentas que ajudam a encontrar os problemas e a fazer com que seus scripts funcionem perfeitamente.

## <a name="console-logs"></a>Logs do console

Às vezes, durante a solução de problemas, convém imprimir mensagens na tela. Eles podem mostrar o valor atual de variáveis ou quais caminhos de código estão sendo disparados. Para fazer isso, faça o log do texto no console.

```TypeScript
console.log("Logging my range's address.");
myRange.load("address");
await context.sync();
console.log(myRange.address);
```

> [!IMPORTANT]
> Não se `load` esqueça dos dados da `sync` planilha e da pasta de trabalho antes de registrar as propriedades do objeto.

As cadeias`console.log` de caracteres passadas para serão exibidas no console de registro em log do editor de código. Para ativar o console, pressione o botão **reticências** e selecione **logs...**

Os logs não afetam a pasta de trabalho.

## <a name="error-messages"></a>Mensagens de erro

Quando o script do Excel encontra um problema em execução, ele produz um erro. Você verá um pop-up de solicitação perguntando se deseja **exibir os logs**. Pressione esse botão para abrir o console e exibir quaisquer erros.

## <a name="help-resources"></a>Recursos de ajuda

[Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts) é uma comunidade de desenvolvedores que desejam ajudar com problemas de codificação. Muitas vezes, você poderá encontrar a solução para o problema por meio de uma pesquisa rápida de estouro de pilha. Caso contrário, faça a pergunta e marque-a com a marca "Office-scripts". Não deixe de mencionar que você está criando um *script*do Office, não um *suplemento*do Office.

Se você encontrar um problema com a API JavaScript do Office, crie um problema no repositório do GitHub do [OfficeDev/Office-js](https://github.com/OfficeDev/office-js) . Os membros da equipe do produto responderão a problemas e fornecerão mais assistência. A criação de um problema no repositório **OfficeDev/Office-js** indica que você encontrou uma falha na biblioteca da API JavaScript do Office para a qual a equipe de produto deve tratar.

Se houver um problema com o gravador de ação ou editor, envie comentários através do botão **ajuda > comentários** no Excel.

## <a name="see-also"></a>Confira também

- [Scripts do Office no Excel na Web](../overview/excel.md)
- [Conceitos básicos de script para scripts do Office no Excel na Web](../develop/scripting-fundamentals.md)
- [Desfazer os efeitos de um script do Office](undo.md)
