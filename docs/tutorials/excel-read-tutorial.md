---
title: Ler os dados da pasta de trabalho com scripts do Office no Excel na Web.
description: Um tutorial de scripts do Office sobre a leitura de dados de pastas de trabalho e avaliação desses dados no script.
ms.date: 07/10/2020
localization_priority: Priority
ms.openlocfilehash: fef1df7cab70ccef67a12ee466af5a89803d0992
ms.sourcegitcommit: ebd1079c7e2695ac0e7e4c616f2439975e196875
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/17/2020
ms.locfileid: "45160403"
---
# <a name="read-workbook-data-with-office-scripts-in-excel-on-the-web"></a><span data-ttu-id="e1e29-103">Ler os dados da pasta de trabalho com scripts do Office no Excel na Web.</span><span class="sxs-lookup"><span data-stu-id="e1e29-103">Read workbook data with Office Scripts in Excel on the web</span></span>

<span data-ttu-id="e1e29-104">Esse tutorial ensina a ler dados de uma pasta de trabalho com scripts do Office para o Excel na Web.</span><span class="sxs-lookup"><span data-stu-id="e1e29-104">This tutorial teaches you how to read data from a workbook with an Office Script for Excel on the web.</span></span> <span data-ttu-id="e1e29-105">Em seguida, edite os dados lidos e coloque-os de volta na pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="e1e29-105">You'll then edit the data you read and put it back in the workbook.</span></span>

> [!TIP]
> <span data-ttu-id="e1e29-106">Se você não tiver experiência com os scripts do Office, recomendamos começar com o tutorial [Grave, edite e crie scripts do Office no Excel na Web](excel-tutorial.md).</span><span class="sxs-lookup"><span data-stu-id="e1e29-106">If you are new to Office Scripts, we recommend starting with the [Record, edit, and create Office Scripts in Excel on the web](excel-tutorial.md) tutorial.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="e1e29-107">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="e1e29-107">Prerequisites</span></span>

[!INCLUDE [Tutorial prerequisites](../includes/tutorial-prerequisites.md)]

> [!IMPORTANT]
> <span data-ttu-id="e1e29-108">Este tutorial é destinado a pessoas com conhecimento básico ou de nível intermediário de JavaScript ou TypeScript.</span><span class="sxs-lookup"><span data-stu-id="e1e29-108">This tutorial is intended for people with beginner to intermediate-level knowledge of JavaScript or TypeScript.</span></span> <span data-ttu-id="e1e29-109">Se você não conhece o JavaScript, recomendamos que revise o [tutorial do Mozilla JavaScript](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).</span><span class="sxs-lookup"><span data-stu-id="e1e29-109">If you're new to JavaScript, we recommend reviewing the [Mozilla JavaScript tutorial](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).</span></span> <span data-ttu-id="e1e29-110">Visite [Scripts do Office no Excel na Web](../overview/excel.md) para saber mais sobre o ambiente de scripts.</span><span class="sxs-lookup"><span data-stu-id="e1e29-110">Visit [Office Scripts in Excel on the web](../overview/excel.md) to learn more about the script environment.</span></span>

## <a name="read-a-cell"></a><span data-ttu-id="e1e29-111">Ler uma célula</span><span class="sxs-lookup"><span data-stu-id="e1e29-111">Read a cell</span></span>

<span data-ttu-id="e1e29-112">Os scripts feitos com o Gravador de Ação só podem gravar informações na pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="e1e29-112">Scripts made with the Action Recorder can only write information to the workbook.</span></span> <span data-ttu-id="e1e29-113">Com o Editor de Códigos, é possível editar e criar scripts que também leem dados de uma pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="e1e29-113">With the Code Editor, you can edit and make scripts that also read data from a workbook.</span></span>

<span data-ttu-id="e1e29-114">Vamos criar um script que leia dados e atue com base no que foi lido.</span><span class="sxs-lookup"><span data-stu-id="e1e29-114">Let's make a script that reads data and acts based on what was read.</span></span> <span data-ttu-id="e1e29-115">Vamos usar um exemplo de um extrato bancário.</span><span class="sxs-lookup"><span data-stu-id="e1e29-115">We're going to work with a sample banking statement.</span></span> <span data-ttu-id="e1e29-116">Essa instrução é um relatório combinado de verificação de crédito.</span><span class="sxs-lookup"><span data-stu-id="e1e29-116">This statement is a combined checking and credit statement.</span></span> <span data-ttu-id="e1e29-117">Infelizmente, eles relatam alterações no balanço de forma diferente.</span><span class="sxs-lookup"><span data-stu-id="e1e29-117">Unfortunately, they report balance changes differently.</span></span> <span data-ttu-id="e1e29-118">A declaração de verificação exibe o rendimento como crédito positivo e custos como débito negativo.</span><span class="sxs-lookup"><span data-stu-id="e1e29-118">The checking statement gives income as positive credit and costs as negative debit.</span></span> <span data-ttu-id="e1e29-119">O demonstrativo de crédito faz o oposto.</span><span class="sxs-lookup"><span data-stu-id="e1e29-119">The credit statement does the opposite.</span></span>

<span data-ttu-id="e1e29-120">No resto do tutorial, normalizaremos os dados usando um script.</span><span class="sxs-lookup"><span data-stu-id="e1e29-120">Over the rest of the tutorial, we will normalize this data using a script.</span></span> <span data-ttu-id="e1e29-121">Primeiro, vamos aprender a ler os dados da pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="e1e29-121">First, let's learn how to read data from the workbook.</span></span>

1. <span data-ttu-id="e1e29-122">Crie uma nova planilha na pasta de trabalho usada para o resto do tutorial.</span><span class="sxs-lookup"><span data-stu-id="e1e29-122">Create a new worksheet in the workbook you've used for the rest of the tutorial.</span></span>
2. <span data-ttu-id="e1e29-123">Copie os seguintes dados e cole-os na nova planilha, começando na célula **A1**.</span><span class="sxs-lookup"><span data-stu-id="e1e29-123">Copy the following data and paste it into the new worksheet, starting at cell **A1**.</span></span>

    |<span data-ttu-id="e1e29-124">Data</span><span class="sxs-lookup"><span data-stu-id="e1e29-124">Date</span></span> |<span data-ttu-id="e1e29-125">Conta</span><span class="sxs-lookup"><span data-stu-id="e1e29-125">Account</span></span> |<span data-ttu-id="e1e29-126">Descrição</span><span class="sxs-lookup"><span data-stu-id="e1e29-126">Description</span></span> |<span data-ttu-id="e1e29-127">Débito</span><span class="sxs-lookup"><span data-stu-id="e1e29-127">Debit</span></span> |<span data-ttu-id="e1e29-128">Crédito</span><span class="sxs-lookup"><span data-stu-id="e1e29-128">Credit</span></span> |
    |:--|:--|:--|:--|:--|
    |<span data-ttu-id="e1e29-129">10/10/2019</span><span class="sxs-lookup"><span data-stu-id="e1e29-129">10/10/2019</span></span> |<span data-ttu-id="e1e29-130">Verificando</span><span class="sxs-lookup"><span data-stu-id="e1e29-130">Checking</span></span> |<span data-ttu-id="e1e29-131">Vinícola Coho</span><span class="sxs-lookup"><span data-stu-id="e1e29-131">Coho Vineyard</span></span> |<span data-ttu-id="e1e29-132">-20.05</span><span class="sxs-lookup"><span data-stu-id="e1e29-132">-20.05</span></span> | |
    |<span data-ttu-id="e1e29-133">11/10/2019</span><span class="sxs-lookup"><span data-stu-id="e1e29-133">10/11/2019</span></span> |<span data-ttu-id="e1e29-134">Crédito</span><span class="sxs-lookup"><span data-stu-id="e1e29-134">Credit</span></span> |<span data-ttu-id="e1e29-135">A Companhia Telefônica</span><span class="sxs-lookup"><span data-stu-id="e1e29-135">The Phone Company</span></span> |<span data-ttu-id="e1e29-136">99.95</span><span class="sxs-lookup"><span data-stu-id="e1e29-136">99.95</span></span> | |
    |<span data-ttu-id="e1e29-137">13/10/2019</span><span class="sxs-lookup"><span data-stu-id="e1e29-137">10/13/2019</span></span> |<span data-ttu-id="e1e29-138">Crédito</span><span class="sxs-lookup"><span data-stu-id="e1e29-138">Credit</span></span> |<span data-ttu-id="e1e29-139">Vinícola Coho</span><span class="sxs-lookup"><span data-stu-id="e1e29-139">Coho Vineyard</span></span> |<span data-ttu-id="e1e29-140">154.43</span><span class="sxs-lookup"><span data-stu-id="e1e29-140">154.43</span></span> | |
    |<span data-ttu-id="e1e29-141">15/10/2019</span><span class="sxs-lookup"><span data-stu-id="e1e29-141">10/15/2019</span></span> |<span data-ttu-id="e1e29-142">Verificando</span><span class="sxs-lookup"><span data-stu-id="e1e29-142">Checking</span></span> |<span data-ttu-id="e1e29-143">Depósito externo</span><span class="sxs-lookup"><span data-stu-id="e1e29-143">External Deposit</span></span> | |<span data-ttu-id="e1e29-144">1000</span><span class="sxs-lookup"><span data-stu-id="e1e29-144">1000</span></span> |
    |<span data-ttu-id="e1e29-145">20/10/2019</span><span class="sxs-lookup"><span data-stu-id="e1e29-145">10/20/2019</span></span> |<span data-ttu-id="e1e29-146">Crédito</span><span class="sxs-lookup"><span data-stu-id="e1e29-146">Credit</span></span> |<span data-ttu-id="e1e29-147">Vinícola Coho – Reembolso</span><span class="sxs-lookup"><span data-stu-id="e1e29-147">Coho Vineyard - Refund</span></span> | |<span data-ttu-id="e1e29-148">-35.45</span><span class="sxs-lookup"><span data-stu-id="e1e29-148">-35.45</span></span> |
    |<span data-ttu-id="e1e29-149">25/10/2019</span><span class="sxs-lookup"><span data-stu-id="e1e29-149">10/25/2019</span></span> |<span data-ttu-id="e1e29-150">Verificando</span><span class="sxs-lookup"><span data-stu-id="e1e29-150">Checking</span></span> |<span data-ttu-id="e1e29-151">Ideal para sua empresa de produtos orgânicos</span><span class="sxs-lookup"><span data-stu-id="e1e29-151">Best For You Organics Company</span></span> | <span data-ttu-id="e1e29-152">-85.64</span><span class="sxs-lookup"><span data-stu-id="e1e29-152">-85.64</span></span> | |
    |<span data-ttu-id="e1e29-153">01/11/2019</span><span class="sxs-lookup"><span data-stu-id="e1e29-153">11/01/2019</span></span> |<span data-ttu-id="e1e29-154">Verificando</span><span class="sxs-lookup"><span data-stu-id="e1e29-154">Checking</span></span> |<span data-ttu-id="e1e29-155">Depósito externo</span><span class="sxs-lookup"><span data-stu-id="e1e29-155">External Deposit</span></span> | |<span data-ttu-id="e1e29-156">1000</span><span class="sxs-lookup"><span data-stu-id="e1e29-156">1000</span></span> |

3. <span data-ttu-id="e1e29-157">Abra o **Editor de códigos** e escolha **Novo script**.</span><span class="sxs-lookup"><span data-stu-id="e1e29-157">Open the **Code Editor** and select **New Script**.</span></span>
4. <span data-ttu-id="e1e29-158">Vamos limpar a formatação.</span><span class="sxs-lookup"><span data-stu-id="e1e29-158">Let's clean up the formatting.</span></span> <span data-ttu-id="e1e29-159">Este é um documento financeiro, iremos alterar a formatação dos números nas colunas **Débito** e **Crédito** para mostrar os valores em dólares.</span><span class="sxs-lookup"><span data-stu-id="e1e29-159">This is a financial document, so let's change the number formatting in the **Debit** and **Credit** columns to show values as dollar amounts.</span></span> <span data-ttu-id="e1e29-160">Também iremos ajustar a largura da coluna para os dados.</span><span class="sxs-lookup"><span data-stu-id="e1e29-160">Let's also fit the column width to the data.</span></span>

    <span data-ttu-id="e1e29-161">Substitua o conteúdo do script pelo código a seguir:</span><span class="sxs-lookup"><span data-stu-id="e1e29-161">Replace the script contents with the following code:</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
        // Get the current worksheet.
        let selectedSheet = workbook.getActiveWorksheet();

        // Format the range to display numerical dollar amounts.
        selectedSheet.getRange("D2:E8").setNumberFormat("$#,##0.00");

        // Fit the width of all the used columns to the data.
        selectedSheet.getUsedRange().getFormat().autofitColumns();
    }
    ```

5. <span data-ttu-id="e1e29-162">Agora, leremos um valor de uma das colunas de número.</span><span class="sxs-lookup"><span data-stu-id="e1e29-162">Now let's read a value from one of the number columns.</span></span> <span data-ttu-id="e1e29-163">Adicione o seguinte código no final do script (antes do encerramento `}`):</span><span class="sxs-lookup"><span data-stu-id="e1e29-163">Add the following code to the end of the script (before the closing `}`):</span></span>

    ```TypeScript
    // Get the value of cell D2.
    let range = selectedSheet.getRange("D2");
    console.log(range.getValues());
    ```

6. <span data-ttu-id="e1e29-164">Execute o script.</span><span class="sxs-lookup"><span data-stu-id="e1e29-164">Run the script.</span></span>
7. <span data-ttu-id="e1e29-165">Você deverá ver `[Array[1]]` no console.</span><span class="sxs-lookup"><span data-stu-id="e1e29-165">You should see `[Array[1]]` in the console.</span></span> <span data-ttu-id="e1e29-166">Isso não é um número porque os intervalos são matrizes bidimensionais de dados.</span><span class="sxs-lookup"><span data-stu-id="e1e29-166">This is not a number because ranges are two-dimensional arrays of data.</span></span> <span data-ttu-id="e1e29-167">Esse intervalo bidimensional está sendo registrado diretamente no console.</span><span class="sxs-lookup"><span data-stu-id="e1e29-167">That two-dimensional range is being logged to the console directly.</span></span> <span data-ttu-id="e1e29-168">Felizmente, o Editor de Códigos permite visualizar o conteúdo da matriz.</span><span class="sxs-lookup"><span data-stu-id="e1e29-168">Luckily, the Code Editor lets you see the contents of the array.</span></span>
8. <span data-ttu-id="e1e29-169">Quando uma matriz bidimensional é registrada no console, ela agrupa os valores de coluna em cada linha.</span><span class="sxs-lookup"><span data-stu-id="e1e29-169">When a two-dimensional array is logged to the console, it groups column values under each row.</span></span> <span data-ttu-id="e1e29-170">Expanda o log de matriz pressionando o triângulo azul.</span><span class="sxs-lookup"><span data-stu-id="e1e29-170">Expand the array log by pressing the blue triangle.</span></span>
9. <span data-ttu-id="e1e29-171">Expanda o segundo nível da matriz, pressionando o triângulo azul exibido recentemente.</span><span class="sxs-lookup"><span data-stu-id="e1e29-171">Expand the second level of the array by pressing the newly revealed blue triangle.</span></span> <span data-ttu-id="e1e29-172">Agora, você deverá ver isto:</span><span class="sxs-lookup"><span data-stu-id="e1e29-172">You should now see this:</span></span>

    ![O log do console mostrando a saída "-20.05", aninhada sob duas matrizes.](../images/tutorial-4.png)

## <a name="modify-the-value-of-a-cell"></a><span data-ttu-id="e1e29-174">Modificar o valor de uma célula</span><span class="sxs-lookup"><span data-stu-id="e1e29-174">Modify the value of a cell</span></span>

<span data-ttu-id="e1e29-175">Agora que podemos ler os dados, usaremos eles para modificar a pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="e1e29-175">Now that we can read data, let's use that data to modify the workbook.</span></span> <span data-ttu-id="e1e29-176">Deixaremos o valor da célula **D2** positivo com a função `Math.abs`.</span><span class="sxs-lookup"><span data-stu-id="e1e29-176">We'll make the value of the cell **D2** positive with the `Math.abs` function.</span></span> <span data-ttu-id="e1e29-177">O objeto [Matemática](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/math) contém várias funções às quais seus scripts têm acesso.</span><span class="sxs-lookup"><span data-stu-id="e1e29-177">The [Math](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/math) object contains many functions to which your scripts have access.</span></span> <span data-ttu-id="e1e29-178">É possível encontrar mais informações sobre `Math` e outros objetos internos [Usando objetos JavaScript internos nos scripts do Office](../develop/javascript-objects.md).</span><span class="sxs-lookup"><span data-stu-id="e1e29-178">More information about `Math` and other built-in objects can be found at [Using built-in JavaScript objects in Office Scripts](../develop/javascript-objects.md).</span></span>

1. <span data-ttu-id="e1e29-179">Adicione o seguinte código ao final do script:</span><span class="sxs-lookup"><span data-stu-id="e1e29-179">Add the following code to the end of the script:</span></span>

    ```TypeScript
    // Run the `Math.abs` function with the value at D2 and apply that value back to D2.
    let positiveValue = Math.abs(range.getValue());
    range.setValue(positiveValue);
    ```

    <span data-ttu-id="e1e29-180">Observe que estamos usando `getValue` e `setValue`.</span><span class="sxs-lookup"><span data-stu-id="e1e29-180">Note that we're using `getValue` and `setValue`.</span></span> <span data-ttu-id="e1e29-181">Esses métodos funcionam em uma única célula.</span><span class="sxs-lookup"><span data-stu-id="e1e29-181">These methods work on a single cell.</span></span> <span data-ttu-id="e1e29-182">Ao lidar com intervalos de várias células, use `getValues` e `setValues`.</span><span class="sxs-lookup"><span data-stu-id="e1e29-182">When handling multi-cell ranges, you'll want to use `getValues` and `setValues`.</span></span>

2. <span data-ttu-id="e1e29-183">O valor da célula **D2** agora deverá ser positivo.</span><span class="sxs-lookup"><span data-stu-id="e1e29-183">The value of cell **D2** should now be positive.</span></span>

## <a name="modify-the-values-of-a-column"></a><span data-ttu-id="e1e29-184">Modificar os valores de uma coluna</span><span class="sxs-lookup"><span data-stu-id="e1e29-184">Modify the values of a column</span></span>

<span data-ttu-id="e1e29-185">Agora que sabemos ler e escrever em uma única célula, vamos generalizar o script para trabalhar em todas as colunas de **Débito** e **Crédito**.</span><span class="sxs-lookup"><span data-stu-id="e1e29-185">Now that we know how to read and write to a single cell, let's generalize the script to work on the entire **Debit** and **Credit** columns.</span></span>

1. <span data-ttu-id="e1e29-186">Remova o código que afeta apenas uma única célula (o código de valor absoluto anterior), de modo que o script agora se pareça com este:</span><span class="sxs-lookup"><span data-stu-id="e1e29-186">Remove the code that affects only a single cell (the previous absolute value code), such that your script now looks like this:</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
        // Get the current worksheet.
        let selectedSheet = workbook.getActiveWorksheet();

        // Format the range to display numerical dollar amounts.
        selectedSheet.getRange("D2:E8").setNumberFormat("$#,##0.00");

        // Fit the width of all the used columns to the data.
        selectedSheet.getUsedRange().getFormat().autofitColumns();
    }
    ```

2. <span data-ttu-id="e1e29-187">Adicione um loop que percorra as linhas nas duas últimas colunas.</span><span class="sxs-lookup"><span data-stu-id="e1e29-187">Add a loop to the end of the script that iterates through the rows in the last two columns.</span></span> <span data-ttu-id="e1e29-188">Para cada célula, o script define o valor para o valor absoluto do valor atual.</span><span class="sxs-lookup"><span data-stu-id="e1e29-188">For each cell, the script sets the value to the current value's absolute value.</span></span>

    <span data-ttu-id="e1e29-189">Observe que a matriz que define a localização das células é baseada em zero.</span><span class="sxs-lookup"><span data-stu-id="e1e29-189">Note that the array defining cell locations is zero-based.</span></span> <span data-ttu-id="e1e29-190">Isso significa que a célula **A1** é `range[0][0]`.</span><span class="sxs-lookup"><span data-stu-id="e1e29-190">That means cell **A1** is `range[0][0]`.</span></span>

    ```TypeScript
    // Get the values of the used range.
    let range = selectedSheet.getUsedRange();
    let rangeValues = range.getValues();

    // Iterate over the fourth and fifth columns and set their values to their absolute value.
    let rowCount = range.getRowCount();
    for (let i = 1; i < rowCount; i++) {
        // The column at index 3 is column "4" in the worksheet.
        if (rangeValues[i][3] != 0) {
            let positiveValue = Math.abs(rangeValues[i][3]);
            selectedSheet.getCell(i, 3).setValue(positiveValue);
        }

        // The column at index 4 is column "5" in the worksheet.
        if (rangeValues[i][4] != 0) {
            let positiveValue = Math.abs(rangeValues[i][4]);
            selectedSheet.getCell(i, 4).setValue(positiveValue);
        }
    }
    ```

    <span data-ttu-id="e1e29-191">Essa parte do script faz várias tarefas importantes.</span><span class="sxs-lookup"><span data-stu-id="e1e29-191">This portion of the script does several important tasks.</span></span> <span data-ttu-id="e1e29-192">Primeiro, ela obtém os valores e a contagem de linhas do intervalo usado.</span><span class="sxs-lookup"><span data-stu-id="e1e29-192">First, it gets the values and row count of the used range.</span></span> <span data-ttu-id="e1e29-193">Isso nos permite ver os valores e saber quando parar.</span><span class="sxs-lookup"><span data-stu-id="e1e29-193">This lets us look at values and know when to stop.</span></span> <span data-ttu-id="e1e29-194">Segundo, ela reitera através do intervalo usado, verificando cada célula nas colunas **Débito** ou **Crédito**.</span><span class="sxs-lookup"><span data-stu-id="e1e29-194">Second, it iterates through the used range, checking each cell in the **Debit** or **Credit** columns.</span></span> <span data-ttu-id="e1e29-195">Por fim, se o valor na célula não for 0, ele será substituído pelo valor absoluto.</span><span class="sxs-lookup"><span data-stu-id="e1e29-195">Finally, if the value in the cell is not 0, it is replaced by its absolute value.</span></span> <span data-ttu-id="e1e29-196">Estamos evitando zeros, para que possamos deixar as células em branco.</span><span class="sxs-lookup"><span data-stu-id="e1e29-196">We're avoiding zeroes so we can leave the blank cells as they were.</span></span>

3. <span data-ttu-id="e1e29-197">Execute o script.</span><span class="sxs-lookup"><span data-stu-id="e1e29-197">Run the script.</span></span>

    <span data-ttu-id="e1e29-198">Seu extrato bancário agora deverá ter a seguinte aparência:</span><span class="sxs-lookup"><span data-stu-id="e1e29-198">Your banking statement should now look like this:</span></span>

    ![O extrato bancário como uma tabela formatada apenas com valores positivos.](../images/tutorial-5.png)

## <a name="next-steps"></a><span data-ttu-id="e1e29-200">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="e1e29-200">Next steps</span></span>

<span data-ttu-id="e1e29-201">Abra o Editor de códigos e experimente alguns dos [Scripts de exemplo para scripts do Office no Excel na Web](../resources/excel-samples.md).</span><span class="sxs-lookup"><span data-stu-id="e1e29-201">Open the Code Editor and try out some of our [Sample scripts for Office Scripts in Excel on the web](../resources/excel-samples.md).</span></span> <span data-ttu-id="e1e29-202">Visite também [Fundamentos de Scripts do Office no Excel na Web](../develop/scripting-fundamentals.md) para saber mais sobre como criar scripts do Office.</span><span class="sxs-lookup"><span data-stu-id="e1e29-202">You can also visit [Scripting Fundamentals for Office Scripts in Excel on the web](../develop/scripting-fundamentals.md) to learn more about creating Office Scripts.</span></span>
