---
title: Scripts do Office no Excel na Web
description: Uma breve introdução ao Gravador de ação e ao Editor de códigos de scripts do Office.
ms.date: 02/24/2020
localization_priority: Priority
ms.openlocfilehash: fb1d32068f9a738bb99412c2892cf22b4119b9b1
ms.sourcegitcommit: d556aaefac80e55f53ac56b7f6ecbc657ebd426f
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/26/2020
ms.locfileid: "42978346"
---
# <a name="office-scripts-in-excel-on-the-web-preview"></a><span data-ttu-id="9f705-103">Scripts do Office no Excel na Web (visualização)</span><span class="sxs-lookup"><span data-stu-id="9f705-103">Office Scripts in Excel on the web (preview)</span></span>

<span data-ttu-id="9f705-104">Os scripts do Office no Excel na Web permitem automatizar suas tarefas diárias.</span><span class="sxs-lookup"><span data-stu-id="9f705-104">Office Scripts in Excel on the web let you automate your day-to-day tasks.</span></span> <span data-ttu-id="9f705-105">Você pode gravar suas ações do Excel com o Gravador de ações, o qual cria um script.</span><span class="sxs-lookup"><span data-stu-id="9f705-105">You can record your Excel actions with the Action Recorder, which creates a script.</span></span> <span data-ttu-id="9f705-106">Você também pode criar e editar scripts com o Editor de códigos.</span><span class="sxs-lookup"><span data-stu-id="9f705-106">You can also create and edit scripts with the Code Editor.</span></span> <span data-ttu-id="9f705-107">Esta série de documentos ensina como usar essas ferramentas.</span><span class="sxs-lookup"><span data-stu-id="9f705-107">This series of documents teaches you how to use these tools.</span></span> <span data-ttu-id="9f705-108">Você será apresentado ao Gravador de ações e verá como gravar suas ações frequentes do Excel.</span><span class="sxs-lookup"><span data-stu-id="9f705-108">You'll be introduced to the Action Recorder and see how to record your frequent Excel actions.</span></span> <span data-ttu-id="9f705-109">Você também aprenderá a criar ou atualizar seus próprios scripts com o Editor de códigos.</span><span class="sxs-lookup"><span data-stu-id="9f705-109">You'll also learn how to make or update your own scripts with the Code Editor.</span></span>

<br>

> [!VIDEO https://www.microsoft.com/videoplayer/embed/RE4qdFF]

[!INCLUDE [Preview note](../includes/preview-note.md)]

## <a name="when-to-use-office-scripts"></a><span data-ttu-id="9f705-110">Quando usar scripts do Office</span><span class="sxs-lookup"><span data-stu-id="9f705-110">When to use Office Scripts</span></span>

<span data-ttu-id="9f705-111">Os scripts permitem gravar e reproduzir suas ações do Excel em diferentes pastas de trabalho e planilhas.</span><span class="sxs-lookup"><span data-stu-id="9f705-111">Scripts allow you to record and replay your Excel actions on different workbooks and worksheets.</span></span> <span data-ttu-id="9f705-112">Se você se encontrar fazendo as mesmas coisas repetidamente, um Script do Office pode ajudá-lo, reduzindo todo o fluxo de trabalho a um único pressionar de botão.</span><span class="sxs-lookup"><span data-stu-id="9f705-112">If you find yourself doing the same things over and over again, an Office Script can help you by reducing your whole workflow to a single button press.</span></span>

<span data-ttu-id="9f705-113">Como exemplo, digamos que você comece seu dia de trabalho abrindo um arquivo .csv em um site de contabilidade no Excel.</span><span class="sxs-lookup"><span data-stu-id="9f705-113">As an example, say you start your work day by opening a .csv file from an accounting site in Excel.</span></span> <span data-ttu-id="9f705-114">Então você gasta alguns minutos excluindo colunas desnecessárias, formatando uma tabela, adicionando fórmulas e criando uma tabela dinâmica em uma nova planilha.</span><span class="sxs-lookup"><span data-stu-id="9f705-114">You then spend several minutes deleting unnecessary columns, formatting a table, adding formulas, and creating a PivotTable in a new worksheet.</span></span> <span data-ttu-id="9f705-115">As ações repetidas diariamente podem ser gravadas uma vez com o Gravador de ações.</span><span class="sxs-lookup"><span data-stu-id="9f705-115">Those actions you repeat daily can be recorded once with the Action Recorder.</span></span> <span data-ttu-id="9f705-116">A partir daí, a execução do script cuidará da sua conversão .csv.</span><span class="sxs-lookup"><span data-stu-id="9f705-116">From then on, running the script will take care of your entire .csv conversion.</span></span> <span data-ttu-id="9f705-117">Além de remover o risco de esquecer as etapas, você poderá compartilhar seu processo com outras pessoas sem precisar ensinar nada a elas.</span><span class="sxs-lookup"><span data-stu-id="9f705-117">You'll not only remove the risk of forgetting steps, but be able to share your process with others without having to teach them anything.</span></span> <span data-ttu-id="9f705-118">Os scripts do Office automatizam suas tarefas comuns para que você e seu local de trabalho possam ser mais eficientes e produtivos.</span><span class="sxs-lookup"><span data-stu-id="9f705-118">Office Scripts automate your common tasks so you and your workplace can be more efficient and productive.</span></span>

## <a name="action-recorder"></a><span data-ttu-id="9f705-119">Gravador de ações</span><span class="sxs-lookup"><span data-stu-id="9f705-119">Action Recorder</span></span>

![O Gravador de ações depois de gravar várias ações.](../images/action-recorder-intro.png)

<span data-ttu-id="9f705-121">O Gravador de ações registra as ações que você executa no Excel e as converte em um script.</span><span class="sxs-lookup"><span data-stu-id="9f705-121">The Action Recorder records actions you take in Excel and translates them into a script.</span></span> <span data-ttu-id="9f705-122">Com o Gravador de ações em execução, você pode capturar as ações do Excel enquanto edita células, altera a formatação e cria tabelas.</span><span class="sxs-lookup"><span data-stu-id="9f705-122">With the Action recorder running, you can capture the Excel actions as you edit cells, change formatting, and create tables.</span></span> <span data-ttu-id="9f705-123">O script resultante pode ser executado em outras planilhas e pastas de trabalho para recriar suas ações originais.</span><span class="sxs-lookup"><span data-stu-id="9f705-123">The resulting script can be run on other worksheets and workbooks to recreate your original actions.</span></span>

## <a name="code-editor"></a><span data-ttu-id="9f705-124">Editor de códigos</span><span class="sxs-lookup"><span data-stu-id="9f705-124">Code Editor</span></span>

![O Editor de códigos exibe o código do script acima.](../images/code-editor-intro.png)

<span data-ttu-id="9f705-126">Todos os scripts gravados com o Gravador de ações podem ser editados através do Editor de códigos.</span><span class="sxs-lookup"><span data-stu-id="9f705-126">All scripts recorded with the Action Recorder can be edited through the Code Editor.</span></span> <span data-ttu-id="9f705-127">Isso permite que você ajuste e personalize o script para melhor atender às suas necessidades.</span><span class="sxs-lookup"><span data-stu-id="9f705-127">This lets you tweak and customize the script to better suit your exact needs.</span></span> <span data-ttu-id="9f705-128">Você também pode adicionar lógica e funcionalidade que não são acessíveis de forma direta pela interface do usuário do Excel, como instruções condicionais (se/senão) e loops.</span><span class="sxs-lookup"><span data-stu-id="9f705-128">You can also add logic and functionality that is not directly accessible through the Excel UI, such as conditional statements (if/else) and loops.</span></span>

<span data-ttu-id="9f705-129">Uma maneira fácil de começar a aprender sobre os recursos dos scripts do Office é gravá-los no Excel na Web e exibir o código resultante.</span><span class="sxs-lookup"><span data-stu-id="9f705-129">One easy way to start learning the capabilities of Office Scripts is to record scripts in Excel on the web and view the resulting code.</span></span> <span data-ttu-id="9f705-130">Outra opção é seguir nossos [tutoriais](../tutorials/excel-tutorial.md) para aprender de uma maneira mais guiada e estruturada.</span><span class="sxs-lookup"><span data-stu-id="9f705-130">Another option is to follow our [tutorials](../tutorials/excel-tutorial.md) to learn in a more guided and structured way.</span></span>

## <a name="next-steps"></a><span data-ttu-id="9f705-131">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="9f705-131">Next steps</span></span>

<span data-ttu-id="9f705-132">Conclua o [tutorial Scripts do Office no Excel na Web](../tutorials/excel-tutorial.md) para aprender como criar seus primeiros scripts do Office.</span><span class="sxs-lookup"><span data-stu-id="9f705-132">Complete the [Office Scripts in Excel on the web tutorial](../tutorials/excel-tutorial.md) to learn how to create your first Office Scripts.</span></span>

## <a name="see-also"></a><span data-ttu-id="9f705-133">Também confira</span><span class="sxs-lookup"><span data-stu-id="9f705-133">See also</span></span>

- [<span data-ttu-id="9f705-134">Fundamentos de script para script do Office no Excel na Web</span><span class="sxs-lookup"><span data-stu-id="9f705-134">Scripting fundamentals for Office Script in Excel on the web</span></span>](../develop/scripting-fundamentals.md)
- [<span data-ttu-id="9f705-135">Referência da API de scripts do Office</span><span class="sxs-lookup"><span data-stu-id="9f705-135">Office Scripts API reference</span></span>](/javascript/api/office-scripts/overview)
- [<span data-ttu-id="9f705-136">Solução de problemas dos scripts do Office</span><span class="sxs-lookup"><span data-stu-id="9f705-136">Troubleshooting Office Scripts</span></span>](../testing/troubleshooting.md)
- [<span data-ttu-id="9f705-137">Configurações dos scripts do Office no M365</span><span class="sxs-lookup"><span data-stu-id="9f705-137">Office Scripts settings in M365</span></span>](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf)
- [<span data-ttu-id="9f705-138">Introdução aos scripts do Office no Excel (em support.office.com)</span><span class="sxs-lookup"><span data-stu-id="9f705-138">Introduction to Office Scripts in Excel (on support.office.com)</span></span>](https://support.office.com/article/introduction-to-office-scripts-in-excel-9fbe283d-adb8-4f13-a75b-a81c6baf163a)
