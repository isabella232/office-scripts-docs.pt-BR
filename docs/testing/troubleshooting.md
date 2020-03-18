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
# <a name="troubleshooting-office-scripts"></a><span data-ttu-id="08089-103">Solucionando problemas de scripts do Office</span><span class="sxs-lookup"><span data-stu-id="08089-103">Troubleshooting Office Scripts</span></span>

<span data-ttu-id="08089-104">Ao desenvolver scripts do Office, você pode cometer erros.</span><span class="sxs-lookup"><span data-stu-id="08089-104">As you develop Office Scripts, you may make mistakes.</span></span> <span data-ttu-id="08089-105">Não há problema.</span><span class="sxs-lookup"><span data-stu-id="08089-105">It's okay.</span></span> <span data-ttu-id="08089-106">Temos ferramentas que ajudam a encontrar os problemas e a fazer com que seus scripts funcionem perfeitamente.</span><span class="sxs-lookup"><span data-stu-id="08089-106">We have tools that help find the problems and get your scripts working perfectly.</span></span>

## <a name="console-logs"></a><span data-ttu-id="08089-107">Logs do console</span><span class="sxs-lookup"><span data-stu-id="08089-107">Console logs</span></span>

<span data-ttu-id="08089-108">Às vezes, durante a solução de problemas, convém imprimir mensagens na tela.</span><span class="sxs-lookup"><span data-stu-id="08089-108">Sometimes while troubleshooting, you'll want to print messages to the screen.</span></span> <span data-ttu-id="08089-109">Eles podem mostrar o valor atual de variáveis ou quais caminhos de código estão sendo disparados.</span><span class="sxs-lookup"><span data-stu-id="08089-109">These can show you the current value of variables or which code paths are being triggered.</span></span> <span data-ttu-id="08089-110">Para fazer isso, faça o log do texto no console.</span><span class="sxs-lookup"><span data-stu-id="08089-110">To do this, log text to the console.</span></span>

```TypeScript
console.log("Logging my range's address.");
myRange.load("address");
await context.sync();
console.log(myRange.address);
```

> [!IMPORTANT]
> <span data-ttu-id="08089-111">Não se `load` esqueça dos dados da `sync` planilha e da pasta de trabalho antes de registrar as propriedades do objeto.</span><span class="sxs-lookup"><span data-stu-id="08089-111">Don't forget to `load` worksheet data and `sync` with the workbook before logging object properties.</span></span>

<span data-ttu-id="08089-112">As cadeias`console.log` de caracteres passadas para serão exibidas no console de registro em log do editor de código.</span><span class="sxs-lookup"><span data-stu-id="08089-112">Strings passed to`console.log` will be displayed in the Code Editor's logging console.</span></span> <span data-ttu-id="08089-113">Para ativar o console, pressione o botão **reticências** e selecione **logs...**</span><span class="sxs-lookup"><span data-stu-id="08089-113">To turn on the console, press the **Ellipses** button and select **Logs...**</span></span>

<span data-ttu-id="08089-114">Os logs não afetam a pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="08089-114">Logs do not affect the workbook.</span></span>

## <a name="error-messages"></a><span data-ttu-id="08089-115">Mensagens de erro</span><span class="sxs-lookup"><span data-stu-id="08089-115">Error messages</span></span>

<span data-ttu-id="08089-116">Quando o script do Excel encontra um problema em execução, ele produz um erro.</span><span class="sxs-lookup"><span data-stu-id="08089-116">When your Excel Script encounters a problem running, it produces an error.</span></span> <span data-ttu-id="08089-117">Você verá um pop-up de solicitação perguntando se deseja **exibir os logs**.</span><span class="sxs-lookup"><span data-stu-id="08089-117">You'll see a prompt pop-up asking if you want to **View Logs**.</span></span> <span data-ttu-id="08089-118">Pressione esse botão para abrir o console e exibir quaisquer erros.</span><span class="sxs-lookup"><span data-stu-id="08089-118">Press that button to open the console and display any errors.</span></span>

## <a name="help-resources"></a><span data-ttu-id="08089-119">Recursos de ajuda</span><span class="sxs-lookup"><span data-stu-id="08089-119">Help resources</span></span>

<span data-ttu-id="08089-120">[Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts) é uma comunidade de desenvolvedores que desejam ajudar com problemas de codificação.</span><span class="sxs-lookup"><span data-stu-id="08089-120">[Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts) is a community of developers willing to help with coding problems.</span></span> <span data-ttu-id="08089-121">Muitas vezes, você poderá encontrar a solução para o problema por meio de uma pesquisa rápida de estouro de pilha.</span><span class="sxs-lookup"><span data-stu-id="08089-121">Often, you'll be able to find the solution to your problem through a quick Stack Overflow search.</span></span> <span data-ttu-id="08089-122">Caso contrário, faça a pergunta e marque-a com a marca "Office-scripts".</span><span class="sxs-lookup"><span data-stu-id="08089-122">If not, ask your question and tag it with the "office-scripts" tag.</span></span> <span data-ttu-id="08089-123">Não deixe de mencionar que você está criando um *script*do Office, não um *suplemento*do Office.</span><span class="sxs-lookup"><span data-stu-id="08089-123">Be sure to mention you're creating an Office *Script*, not an Office *Add-in*.</span></span>

<span data-ttu-id="08089-124">Se você encontrar um problema com a API JavaScript do Office, crie um problema no repositório do GitHub do [OfficeDev/Office-js](https://github.com/OfficeDev/office-js) .</span><span class="sxs-lookup"><span data-stu-id="08089-124">If you encounter a problem with the Office JavaScript API, create an issue in the [OfficeDev/office-js](https://github.com/OfficeDev/office-js) GitHub repository.</span></span> <span data-ttu-id="08089-125">Os membros da equipe do produto responderão a problemas e fornecerão mais assistência.</span><span class="sxs-lookup"><span data-stu-id="08089-125">Members of the product team will respond to issues and provide further assistance.</span></span> <span data-ttu-id="08089-126">A criação de um problema no repositório **OfficeDev/Office-js** indica que você encontrou uma falha na biblioteca da API JavaScript do Office para a qual a equipe de produto deve tratar.</span><span class="sxs-lookup"><span data-stu-id="08089-126">Creating an issue in the **OfficeDev/office-js** repository indicates you have found a flaw in the Office JavaScript API library that the product team should address.</span></span>

<span data-ttu-id="08089-127">Se houver um problema com o gravador de ação ou editor, envie comentários através do botão **ajuda > comentários** no Excel.</span><span class="sxs-lookup"><span data-stu-id="08089-127">If there is a problem with the Action Recorder or Editor, send feedback through the **Help > Feedback** button in Excel.</span></span>

## <a name="see-also"></a><span data-ttu-id="08089-128">Confira também</span><span class="sxs-lookup"><span data-stu-id="08089-128">See also</span></span>

- [<span data-ttu-id="08089-129">Scripts do Office no Excel na Web</span><span class="sxs-lookup"><span data-stu-id="08089-129">Office Scripts in Excel on the web</span></span>](../overview/excel.md)
- [<span data-ttu-id="08089-130">Conceitos básicos de script para scripts do Office no Excel na Web</span><span class="sxs-lookup"><span data-stu-id="08089-130">Scripting Fundamentals for Office Scripts in Excel on the web</span></span>](../develop/scripting-fundamentals.md)
- [<span data-ttu-id="08089-131">Desfazer os efeitos de um script do Office</span><span class="sxs-lookup"><span data-stu-id="08089-131">Undo the effects of an Office Script</span></span>](undo.md)
