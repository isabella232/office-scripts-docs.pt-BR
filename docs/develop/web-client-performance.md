---
title: Melhorar o desempenho dos scripts do Office
description: Crie scripts mais rápidos compreendendo a comunicação entre a pasta de trabalho do Excel e seu script.
ms.date: 06/15/2020
localization_priority: Normal
ms.openlocfilehash: 4d5b7c70f14e3fc598b95a6226e3ef8caf89f651
ms.sourcegitcommit: aec3c971c6640429f89b6bb99d2c95ea06725599
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/25/2020
ms.locfileid: "44878719"
---
# <a name="improve-the-performance-of-your-office-scripts"></a><span data-ttu-id="2a441-103">Melhorar o desempenho dos scripts do Office</span><span class="sxs-lookup"><span data-stu-id="2a441-103">Improve the performance of your Office Scripts</span></span>

<span data-ttu-id="2a441-104">A finalidade dos scripts do Office é automatizar a série de tarefas realizadas com frequência para poupar tempo.</span><span class="sxs-lookup"><span data-stu-id="2a441-104">The purpose of Office Scripts is to automate commonly performed series of tasks to save you time.</span></span> <span data-ttu-id="2a441-105">Um script lento pode parecer que não acelera o fluxo de trabalho.</span><span class="sxs-lookup"><span data-stu-id="2a441-105">A slow script can feel like it doesn't speed up your workflow.</span></span> <span data-ttu-id="2a441-106">Na maioria das vezes, seu script será perfeitamente bom e será executado conforme o esperado.</span><span class="sxs-lookup"><span data-stu-id="2a441-106">Most of the time, your script will be perfectly fine and run as expected.</span></span> <span data-ttu-id="2a441-107">No entanto, há alguns cenários do avoidable que podem afetar o desempenho.</span><span class="sxs-lookup"><span data-stu-id="2a441-107">However, there are a few, avoidable scenarios that can affect performance.</span></span>

<span data-ttu-id="2a441-108">O motivo mais comum para um script lento é a comunicação excessiva com a pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="2a441-108">The most common reason for a slow script is excessive communication with the workbook.</span></span> <span data-ttu-id="2a441-109">O script é executado no computador local, enquanto a pasta de trabalho existe na nuvem.</span><span class="sxs-lookup"><span data-stu-id="2a441-109">Your script runs on your local machine, while the workbook exists in the cloud.</span></span> <span data-ttu-id="2a441-110">Em determinados momentos, o script sincroniza seus dados locais com o da pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="2a441-110">At certain times, your script synchronizes its local data with that of the workbook.</span></span> <span data-ttu-id="2a441-111">Isso significa que qualquer operação de gravação (como `workbook.addWorksheet()` ) só será aplicada à pasta de trabalho quando essa sincronização por trás da cena acontecer.</span><span class="sxs-lookup"><span data-stu-id="2a441-111">This means that any write operations (such as `workbook.addWorksheet()`) are only applied to the workbook when this behind-the-scenes synchronization happens.</span></span> <span data-ttu-id="2a441-112">Da mesma forma, qualquer operação de leitura (como `myRange.getValues()` ) só obterá dados da pasta de trabalho para o script nesses horários.</span><span class="sxs-lookup"><span data-stu-id="2a441-112">Likewise, any read operations (such as `myRange.getValues()`) only get data from the workbook for the script at those times.</span></span> <span data-ttu-id="2a441-113">Em ambos os casos, o script busca informações antes que ele atue nos dados.</span><span class="sxs-lookup"><span data-stu-id="2a441-113">In either case, the script fetches information before it acts on the data.</span></span> <span data-ttu-id="2a441-114">Por exemplo, o código a seguir registra com precisão o número de linhas no intervalo usado.</span><span class="sxs-lookup"><span data-stu-id="2a441-114">For example, the following code will accurately log the number of rows in the used range.</span></span>

```TypeScript
let usedRange = workbook.getActiveWorksheet().getUsedRange();
let rowCount = usedRange.getRowCount();
// The script will read the range and row count from
// the workbook before logging the information.
console.log(rowCount);
```

<span data-ttu-id="2a441-115">As APIs de scripts do Office garantem que todos os dados na pasta de trabalho ou script sejam precisos e atualizados quando necessário.</span><span class="sxs-lookup"><span data-stu-id="2a441-115">Office Scripts APIs ensure any data in the workbook or script is accurate and up-to-date when necessary.</span></span> <span data-ttu-id="2a441-116">Você não precisa se preocupar com essas sincronizações para que o script seja executado corretamente.</span><span class="sxs-lookup"><span data-stu-id="2a441-116">You don't need to worry about these synchronizations for your script to run correctly.</span></span> <span data-ttu-id="2a441-117">No entanto, um reconhecimento dessa comunicação de script para nuvem pode ajudar você a evitar chamadas de rede desnecessárias.</span><span class="sxs-lookup"><span data-stu-id="2a441-117">However, an awareness of this script-to-cloud communication can help you avoid unneeded network calls.</span></span>

## <a name="performance-optimizations"></a><span data-ttu-id="2a441-118">Otimizações de desempenho</span><span class="sxs-lookup"><span data-stu-id="2a441-118">Performance optimizations</span></span>

<span data-ttu-id="2a441-119">Você pode aplicar técnicas simples para ajudar a reduzir a comunicação com a nuvem.</span><span class="sxs-lookup"><span data-stu-id="2a441-119">You can apply simple techniques to help reduce the communication to the cloud.</span></span> <span data-ttu-id="2a441-120">Os padrões a seguir ajudam a acelerar seus scripts.</span><span class="sxs-lookup"><span data-stu-id="2a441-120">The following patterns help speed up your scripts.</span></span>

- <span data-ttu-id="2a441-121">Leia os dados da pasta de trabalho uma vez em vez de repetidamente em um loop.</span><span class="sxs-lookup"><span data-stu-id="2a441-121">Read workbook data once instead of repeatedly in a loop.</span></span>
- <span data-ttu-id="2a441-122">Remover instruções desnecessárias `console.log` .</span><span class="sxs-lookup"><span data-stu-id="2a441-122">Remove unnecessary `console.log` statements.</span></span>
- <span data-ttu-id="2a441-123">Evite usar blocos try/catch.</span><span class="sxs-lookup"><span data-stu-id="2a441-123">Avoid using try/catch blocks.</span></span>

### <a name="read-workbook-data-outside-of-a-loop"></a><span data-ttu-id="2a441-124">Ler dados de pasta de trabalho fora de um loop</span><span class="sxs-lookup"><span data-stu-id="2a441-124">Read workbook data outside of a loop</span></span>

<span data-ttu-id="2a441-125">Qualquer método que obtém dados da pasta de trabalho pode disparar uma chamada de rede.</span><span class="sxs-lookup"><span data-stu-id="2a441-125">Any method that gets data from the workbook can trigger a network call.</span></span> <span data-ttu-id="2a441-126">Em vez de fazer repetidamente a mesma chamada, você deve salvar dados localmente, sempre que possível.</span><span class="sxs-lookup"><span data-stu-id="2a441-126">Rather than repeatedly making the same call, you should save data locally whenever possible.</span></span> <span data-ttu-id="2a441-127">Isso se aplica especialmente ao lidar com loops.</span><span class="sxs-lookup"><span data-stu-id="2a441-127">This is especially true when dealing with loops.</span></span>

<span data-ttu-id="2a441-128">Considere um script para obter a contagem de números negativos no intervalo usado de uma planilha.</span><span class="sxs-lookup"><span data-stu-id="2a441-128">Consider a script to get the count of negative numbers in the used range of a worksheet.</span></span> <span data-ttu-id="2a441-129">O script precisa iterar em todas as células do intervalo usado.</span><span class="sxs-lookup"><span data-stu-id="2a441-129">The script needs to iterate over every cell in the used range.</span></span> <span data-ttu-id="2a441-130">Para fazer isso, é necessário o intervalo, o número de linhas e o número de colunas.</span><span class="sxs-lookup"><span data-stu-id="2a441-130">To do that, it needs the range, the number of rows, and the number of columns.</span></span> <span data-ttu-id="2a441-131">Você deve armazená-las como variáveis locais antes de iniciar o loop.</span><span class="sxs-lookup"><span data-stu-id="2a441-131">You should store those as local variables before starting the loop.</span></span> <span data-ttu-id="2a441-132">Caso contrário, cada iteração do loop forçará um retorno à pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="2a441-132">Otherwise, each iteration of the loop will force a return to the workbook.</span></span>

```TypeScript
/**
 * This script provides the count of negative numbers that are present
 * in the used range of the current worksheet.
 */
function main(workbook: ExcelScript.Workbook) {
  // Get the working range.
  let usedRange = workbook.getActiveWorksheet().getUsedRange();

  // Save the values locally to avoid repeatedly asking the workbook.
  let usedRangeValues = usedRange.getValues();

  // Start the negative number counter.
  let negativeCount = 0;

  // Iterate over the entire range looking for negative numbers.
  for (let i = 0; i < usedRangeValues.length; i++) {
    for (let j = 0; j < usedRangeValues[i].length; j++) {
      if (usedRangeValues[i][j] < 0) {
        negativeCount++;
      }
    }
  }

  // Log the negative number count to the console.
  console.log(negativeCount);
}
```

> [!NOTE]
> <span data-ttu-id="2a441-133">Como experimento, tente substituir `usedRangeValues` no loop por `usedRange.getValues()` .</span><span class="sxs-lookup"><span data-stu-id="2a441-133">As an experiment, try replacing `usedRangeValues` in the loop with `usedRange.getValues()`.</span></span> <span data-ttu-id="2a441-134">Você pode notar que o script demora consideravelmente mais para ser executado ao lidar com intervalos grandes.</span><span class="sxs-lookup"><span data-stu-id="2a441-134">You may notice the script takes considerably longer to run when dealing with large ranges.</span></span>

### <a name="remove-unnecessary-consolelog-statements"></a><span data-ttu-id="2a441-135">Remover instruções desnecessárias `console.log`</span><span class="sxs-lookup"><span data-stu-id="2a441-135">Remove unnecessary `console.log` statements</span></span>

<span data-ttu-id="2a441-136">O registro em log do console é uma ferramenta vital para [depurar seus scripts](../testing/troubleshooting.md).</span><span class="sxs-lookup"><span data-stu-id="2a441-136">Console logging is a vital tool for [debugging your scripts](../testing/troubleshooting.md).</span></span> <span data-ttu-id="2a441-137">No entanto, ele força o script a sincronizar com a pasta de trabalho para garantir que as informações registradas estejam atualizadas.</span><span class="sxs-lookup"><span data-stu-id="2a441-137">However, it does force the script to synchronize with the workbook to ensure the logged information is up-to-date.</span></span> <span data-ttu-id="2a441-138">Considere remover declarações desnecessárias de registro em log (como as usadas para teste) antes de compartilhar seu script.</span><span class="sxs-lookup"><span data-stu-id="2a441-138">Consider removing unnecessary logging statements (such as those used for testing) before sharing your script.</span></span> <span data-ttu-id="2a441-139">Isso normalmente não causará um problema de desempenho perceptível, a menos que a `console.log()` instrução esteja em um loop.</span><span class="sxs-lookup"><span data-stu-id="2a441-139">This typically won't cause a noticeable performance issue, unless the `console.log()` statement is in a loop.</span></span>

### <a name="avoid-using-trycatch-blocks"></a><span data-ttu-id="2a441-140">Evite usar blocos try/catch</span><span class="sxs-lookup"><span data-stu-id="2a441-140">Avoid using try/catch blocks</span></span>

<span data-ttu-id="2a441-141">Não recomendamos o uso de [ `try` / `catch` blocos](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) como parte do fluxo de controle esperado de um script.</span><span class="sxs-lookup"><span data-stu-id="2a441-141">We don't recommend using [`try`/`catch` blocks](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) as part of a script's expected control flow.</span></span> <span data-ttu-id="2a441-142">A maioria dos erros pode ser evitada verificando os objetos retornados da pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="2a441-142">Most errors can be avoided by checking objects returned from the workbook.</span></span> <span data-ttu-id="2a441-143">Por exemplo, o script a seguir verifica se a tabela retornada pela pasta de trabalho existe antes de tentar adicionar uma linha.</span><span class="sxs-lookup"><span data-stu-id="2a441-143">For example, the following script checks that the table returned by the workbook exists before trying to add a row.</span></span>

```TypeScript
/**
 * This script adds a row to "MyTable", if that table is present.
 */
function main(workbook: ExcelScript.Workbook) {
  let table = workbook.getTable("MyTable");

  // Check if the table exists.
  if (table) {
    // Add the row.
    table.addRow(-1, ["2012", "Yes", "Maybe"]);
  } else {
    // Report the missing table.
    console.log("MyTable not found.");
  }
}
```

## <a name="case-by-case-help"></a><span data-ttu-id="2a441-144">Ajuda caso a caso</span><span class="sxs-lookup"><span data-stu-id="2a441-144">Case-by-case help</span></span>

<span data-ttu-id="2a441-145">Como a plataforma de scripts do Office se expande para trabalhar com [automatização de energia](https://flow.microsoft.com/), [cartões adaptáveis](https://docs.microsoft.com/adaptive-cards)e outros recursos entre produtos, os detalhes da comunicação de pasta de trabalho de script ficam mais complexos.</span><span class="sxs-lookup"><span data-stu-id="2a441-145">As the Office Scripts platform expands to work with [Power Automate](https://flow.microsoft.com/), [Adaptive Cards](https://docs.microsoft.com/adaptive-cards), and other cross-product features, the details of the script-workbook communication become more intricate.</span></span> <span data-ttu-id="2a441-146">Se você precisar de ajuda para executar o script com mais rapidez, saia pelo [estouro de pilha](https://stackoverflow.com/questions/tagged/office-scripts).</span><span class="sxs-lookup"><span data-stu-id="2a441-146">If you need help making your script run faster, please reach out through [Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts).</span></span> <span data-ttu-id="2a441-147">Certifique-se de marcar sua pergunta com "Office-scripts" para que os especialistas possam encontrá-lo e ajudá-lo.</span><span class="sxs-lookup"><span data-stu-id="2a441-147">Be sure to tag your question with "office-scripts" so experts can find it and help.</span></span>

## <a name="see-also"></a><span data-ttu-id="2a441-148">Confira também</span><span class="sxs-lookup"><span data-stu-id="2a441-148">See also</span></span>

- [<span data-ttu-id="2a441-149">Fundamentos de script para scripts do Office no Excel na Web</span><span class="sxs-lookup"><span data-stu-id="2a441-149">Scripting fundamentals for Office Scripts in Excel on the web</span></span>](scripting-fundamentals.md)
- [<span data-ttu-id="2a441-150">MDN Web docs: loops e iteração</span><span class="sxs-lookup"><span data-stu-id="2a441-150">MDN web docs: Loops and iteration</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Loops_and_iteration)
