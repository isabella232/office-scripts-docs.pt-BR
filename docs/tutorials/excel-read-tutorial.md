---
title: Ler os dados da pasta de trabalho com scripts do Office no Excel na Web.
description: Um tutorial de scripts do Office sobre a leitura de dados de pastas de trabalho e avaliação desses dados no script.
ms.date: 04/23/2020
localization_priority: Priority
ms.openlocfilehash: 93204184d4b5947b2a67107b1fd73c178a73c32e
ms.sourcegitcommit: aec3c971c6640429f89b6bb99d2c95ea06725599
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/25/2020
ms.locfileid: "44878680"
---
# <a name="read-workbook-data-with-office-scripts-in-excel-on-the-web"></a><span data-ttu-id="459b6-103">Ler os dados da pasta de trabalho com scripts do Office no Excel na Web.</span><span class="sxs-lookup"><span data-stu-id="459b6-103">Read workbook data with Office Scripts in Excel on the web</span></span>

<span data-ttu-id="459b6-104">Esse tutorial ensina a ler dados de uma pasta de trabalho com scripts do Office para o Excel na Web.</span><span class="sxs-lookup"><span data-stu-id="459b6-104">This tutorial teaches you how to read data from a workbook with an Office Script for Excel on the web.</span></span> <span data-ttu-id="459b6-105">Em seguida, edite os dados lidos e coloque-os de volta na pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="459b6-105">You'll then edit the data you read and put it back in the workbook.</span></span>

> [!TIP]
> <span data-ttu-id="459b6-106">Se você não tiver experiência com os scripts do Office, recomendamos começar com o tutorial [Grave, edite e crie scripts do Office no Excel na Web](excel-tutorial.md).</span><span class="sxs-lookup"><span data-stu-id="459b6-106">If you are new to Office Scripts, we recommend starting with the [Record, edit, and create Office Scripts in Excel on the web](excel-tutorial.md) tutorial.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="459b6-107">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="459b6-107">Prerequisites</span></span>

[!INCLUDE [Tutorial prerequisites](../includes/tutorial-prerequisites.md)]

> [!IMPORTANT]
> <span data-ttu-id="459b6-108">Este tutorial é destinado a pessoas com conhecimento básico ou de nível intermediário de JavaScript ou TypeScript.</span><span class="sxs-lookup"><span data-stu-id="459b6-108">This tutorial is intended for people with beginner to intermediate-level knowledge of JavaScript or TypeScript.</span></span> <span data-ttu-id="459b6-109">Se você não conhece o JavaScript, recomendamos que revise o [tutorial do Mozilla JavaScript](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).</span><span class="sxs-lookup"><span data-stu-id="459b6-109">If you're new to JavaScript, we recommend reviewing the [Mozilla JavaScript tutorial](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).</span></span> <span data-ttu-id="459b6-110">Visite [Scripts do Office no Excel na Web](../overview/excel.md) para saber mais sobre o ambiente de scripts.</span><span class="sxs-lookup"><span data-stu-id="459b6-110">Visit [Office Scripts in Excel on the web](../overview/excel.md) to learn more about the script environment.</span></span>

## <a name="read-a-cell"></a><span data-ttu-id="459b6-111">Ler uma célula</span><span class="sxs-lookup"><span data-stu-id="459b6-111">Read a cell</span></span>

<span data-ttu-id="459b6-112">Os scripts feitos com o Gravador de Ação só podem gravar informações na pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="459b6-112">Scripts made with the Action Recorder can only write information to the workbook.</span></span> <span data-ttu-id="459b6-113">Com o Editor de Códigos, é possível editar e criar scripts que também leem dados de uma pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="459b6-113">With the Code Editor, you can edit and make scripts that also read data from a workbook.</span></span>

<span data-ttu-id="459b6-114">Vamos criar um script que leia dados e atue com base no que foi lido.</span><span class="sxs-lookup"><span data-stu-id="459b6-114">Let's make a script that reads data and acts based on what was read.</span></span> <span data-ttu-id="459b6-115">Vamos usar um exemplo de um extrato bancário.</span><span class="sxs-lookup"><span data-stu-id="459b6-115">We're going to work with a sample banking statement.</span></span> <span data-ttu-id="459b6-116">Essa instrução é um relatório combinado de verificação de crédito.</span><span class="sxs-lookup"><span data-stu-id="459b6-116">This statement is a combined checking and credit statement.</span></span> <span data-ttu-id="459b6-117">Infelizmente, eles relatam alterações no balanço de forma diferente.</span><span class="sxs-lookup"><span data-stu-id="459b6-117">Unfortunately, they report balance changes differently.</span></span> <span data-ttu-id="459b6-118">A declaração de verificação exibe o rendimento como crédito positivo e custos como débito negativo.</span><span class="sxs-lookup"><span data-stu-id="459b6-118">The checking statement gives income as positive credit and costs as negative debit.</span></span> <span data-ttu-id="459b6-119">O demonstrativo de crédito faz o oposto.</span><span class="sxs-lookup"><span data-stu-id="459b6-119">The credit statement does the opposite.</span></span>

<span data-ttu-id="459b6-120">No resto do tutorial, normalizaremos os dados usando um script.</span><span class="sxs-lookup"><span data-stu-id="459b6-120">Over the rest of the tutorial, we will normalize this data using a script.</span></span> <span data-ttu-id="459b6-121">Primeiro, vamos aprender a ler os dados da pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="459b6-121">First, let's learn how to read data from the workbook.</span></span>

1. <span data-ttu-id="459b6-122">Crie uma nova planilha na pasta de trabalho usada para o resto do tutorial.</span><span class="sxs-lookup"><span data-stu-id="459b6-122">Create a new worksheet in the workbook you've used for the rest of the tutorial.</span></span>
2. <span data-ttu-id="459b6-123">Copie os seguintes dados e cole-os na nova planilha, começando na célula **A1**.</span><span class="sxs-lookup"><span data-stu-id="459b6-123">Copy the following data and paste it into the new worksheet, starting at cell **A1**.</span></span>

    |<span data-ttu-id="459b6-124">Data</span><span class="sxs-lookup"><span data-stu-id="459b6-124">Date</span></span> |<span data-ttu-id="459b6-125">Conta</span><span class="sxs-lookup"><span data-stu-id="459b6-125">Account</span></span> |<span data-ttu-id="459b6-126">Descrição</span><span class="sxs-lookup"><span data-stu-id="459b6-126">Description</span></span> |<span data-ttu-id="459b6-127">Débito</span><span class="sxs-lookup"><span data-stu-id="459b6-127">Debit</span></span> |<span data-ttu-id="459b6-128">Crédito</span><span class="sxs-lookup"><span data-stu-id="459b6-128">Credit</span></span> |
    |:--|:--|:--|:--|:--|
    |<span data-ttu-id="459b6-129">10/10/2019</span><span class="sxs-lookup"><span data-stu-id="459b6-129">10/10/2019</span></span> |<span data-ttu-id="459b6-130">Verificando</span><span class="sxs-lookup"><span data-stu-id="459b6-130">Checking</span></span> |<span data-ttu-id="459b6-131">Vinícola Coho</span><span class="sxs-lookup"><span data-stu-id="459b6-131">Coho Vineyard</span></span> |<span data-ttu-id="459b6-132">-20.05</span><span class="sxs-lookup"><span data-stu-id="459b6-132">-20.05</span></span> | |
    |<span data-ttu-id="459b6-133">11/10/2019</span><span class="sxs-lookup"><span data-stu-id="459b6-133">10/11/2019</span></span> |<span data-ttu-id="459b6-134">Crédito</span><span class="sxs-lookup"><span data-stu-id="459b6-134">Credit</span></span> |<span data-ttu-id="459b6-135">A Companhia Telefônica</span><span class="sxs-lookup"><span data-stu-id="459b6-135">The Phone Company</span></span> |<span data-ttu-id="459b6-136">99.95</span><span class="sxs-lookup"><span data-stu-id="459b6-136">99.95</span></span> | |
    |<span data-ttu-id="459b6-137">13/10/2019</span><span class="sxs-lookup"><span data-stu-id="459b6-137">10/13/2019</span></span> |<span data-ttu-id="459b6-138">Crédito</span><span class="sxs-lookup"><span data-stu-id="459b6-138">Credit</span></span> |<span data-ttu-id="459b6-139">Vinícola Coho</span><span class="sxs-lookup"><span data-stu-id="459b6-139">Coho Vineyard</span></span> |<span data-ttu-id="459b6-140">154.43</span><span class="sxs-lookup"><span data-stu-id="459b6-140">154.43</span></span> | |
    |<span data-ttu-id="459b6-141">15/10/2019</span><span class="sxs-lookup"><span data-stu-id="459b6-141">10/15/2019</span></span> |<span data-ttu-id="459b6-142">Verificando</span><span class="sxs-lookup"><span data-stu-id="459b6-142">Checking</span></span> |<span data-ttu-id="459b6-143">Depósito externo</span><span class="sxs-lookup"><span data-stu-id="459b6-143">External Deposit</span></span> | |<span data-ttu-id="459b6-144">1000</span><span class="sxs-lookup"><span data-stu-id="459b6-144">1000</span></span> |
    |<span data-ttu-id="459b6-145">20/10/2019</span><span class="sxs-lookup"><span data-stu-id="459b6-145">10/20/2019</span></span> |<span data-ttu-id="459b6-146">Crédito</span><span class="sxs-lookup"><span data-stu-id="459b6-146">Credit</span></span> |<span data-ttu-id="459b6-147">Vinícola Coho – Reembolso</span><span class="sxs-lookup"><span data-stu-id="459b6-147">Coho Vineyard - Refund</span></span> | |<span data-ttu-id="459b6-148">-35.45</span><span class="sxs-lookup"><span data-stu-id="459b6-148">-35.45</span></span> |
    |<span data-ttu-id="459b6-149">25/10/2019</span><span class="sxs-lookup"><span data-stu-id="459b6-149">10/25/2019</span></span> |<span data-ttu-id="459b6-150">Verificando</span><span class="sxs-lookup"><span data-stu-id="459b6-150">Checking</span></span> |<span data-ttu-id="459b6-151">Ideal para sua empresa de produtos orgânicos</span><span class="sxs-lookup"><span data-stu-id="459b6-151">Best For You Organics Company</span></span> | <span data-ttu-id="459b6-152">-85.64</span><span class="sxs-lookup"><span data-stu-id="459b6-152">-85.64</span></span> | |
    |<span data-ttu-id="459b6-153">01/11/2019</span><span class="sxs-lookup"><span data-stu-id="459b6-153">11/01/2019</span></span> |<span data-ttu-id="459b6-154">Verificando</span><span class="sxs-lookup"><span data-stu-id="459b6-154">Checking</span></span> |<span data-ttu-id="459b6-155">Depósito externo</span><span class="sxs-lookup"><span data-stu-id="459b6-155">External Deposit</span></span> | |<span data-ttu-id="459b6-156">1000</span><span class="sxs-lookup"><span data-stu-id="459b6-156">1000</span></span> |

3. <span data-ttu-id="459b6-157">Abra o **Editor de códigos** e escolha **Novo script**.</span><span class="sxs-lookup"><span data-stu-id="459b6-157">Open the **Code Editor** and select **New Script**.</span></span>
4. <span data-ttu-id="459b6-158">Vamos limpar a formatação.</span><span class="sxs-lookup"><span data-stu-id="459b6-158">Let's clean up the formatting.</span></span> <span data-ttu-id="459b6-159">Este é um documento financeiro, iremos alterar a formatação dos números nas colunas **Débito** e **Crédito** para mostrar os valores em dólares.</span><span class="sxs-lookup"><span data-stu-id="459b6-159">This is a financial document, so let's change the number formatting in the **Debit** and **Credit** columns to show values as dollar amounts.</span></span> <span data-ttu-id="459b6-160">Também iremos ajustar a largura da coluna para os dados.</span><span class="sxs-lookup"><span data-stu-id="459b6-160">Let's also fit the column width to the data.</span></span>

    <span data-ttu-id="459b6-161">Substitua o conteúdo do script pelo código a seguir:</span><span class="sxs-lookup"><span data-stu-id="459b6-161">Replace the script contents with the following code:</span></span>

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

5. <span data-ttu-id="459b6-162">Agora, leremos um valor de uma das colunas de número.</span><span class="sxs-lookup"><span data-stu-id="459b6-162">Now let's read a value from one of the number columns.</span></span> <span data-ttu-id="459b6-163">Adicione o seguinte código no final do script (antes do encerramento `}`):</span><span class="sxs-lookup"><span data-stu-id="459b6-163">Add the following code to the end of the script (before the closing `}`):</span></span>

    ```TypeScript
    // Get the value of cell D2.
    let range = selectedSheet.getRange("D2");
    console.log(range.getValues());
    ```

6. <span data-ttu-id="459b6-164">Execute o script.</span><span class="sxs-lookup"><span data-stu-id="459b6-164">Run the script.</span></span>
7. <span data-ttu-id="459b6-165">Abra o console.</span><span class="sxs-lookup"><span data-stu-id="459b6-165">Open the console.</span></span> <span data-ttu-id="459b6-166">Vá para o menu **Reticências** e pressione **Logs...**.</span><span class="sxs-lookup"><span data-stu-id="459b6-166">Go to the **Ellipses** menu and press **Logs...**.</span></span>
8. <span data-ttu-id="459b6-167">Você deverá ver `[Array[1]]` no console.</span><span class="sxs-lookup"><span data-stu-id="459b6-167">You should see `[Array[1]]` in the console.</span></span> <span data-ttu-id="459b6-168">Isso não é um número porque os intervalos são matrizes bidimensionais de dados.</span><span class="sxs-lookup"><span data-stu-id="459b6-168">This is not a number because ranges are two-dimensional arrays of data.</span></span> <span data-ttu-id="459b6-169">Esse intervalo bidimensional está sendo registrado diretamente no console.</span><span class="sxs-lookup"><span data-stu-id="459b6-169">That two-dimensional range is being logged to the console directly.</span></span> <span data-ttu-id="459b6-170">Felizmente, o Editor de códigos permite visualizar o conteúdo da matriz.</span><span class="sxs-lookup"><span data-stu-id="459b6-170">Luckily, the Code Editor does let you see the contents of the array.</span></span>
9. <span data-ttu-id="459b6-171">Quando uma matriz bidimensional é registrada no console, ela agrupa os valores de coluna em cada linha.</span><span class="sxs-lookup"><span data-stu-id="459b6-171">When a two-dimensional array is logged to the console, it groups column values under each row.</span></span> <span data-ttu-id="459b6-172">Expanda o log de matriz pressionando o triângulo azul.</span><span class="sxs-lookup"><span data-stu-id="459b6-172">Expand the array log by pressing the blue triangle.</span></span>
10. <span data-ttu-id="459b6-173">Expanda o segundo nível da matriz, pressionando o triângulo azul exibido recentemente.</span><span class="sxs-lookup"><span data-stu-id="459b6-173">Expand the second level of the array by pressing the newly revealed blue triangle.</span></span> <span data-ttu-id="459b6-174">Agora, você deverá ver isto:</span><span class="sxs-lookup"><span data-stu-id="459b6-174">You should now see this:</span></span>

    ![O log do console mostrando a saída "-20.05", aninhada sob duas matrizes.](../images/tutorial-4.png)

## <a name="modify-the-value-of-a-cell"></a><span data-ttu-id="459b6-176">Modificar o valor de uma célula</span><span class="sxs-lookup"><span data-stu-id="459b6-176">Modify the value of a cell</span></span>

<span data-ttu-id="459b6-177">Agora que podemos ler os dados, usaremos eles para modificar a pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="459b6-177">Now that we can read data, let's use that data to modify the workbook.</span></span> <span data-ttu-id="459b6-178">Deixaremos o valor da célula **D2** positivo com a função `Math.abs`.</span><span class="sxs-lookup"><span data-stu-id="459b6-178">We'll make the value of the cell **D2** positive with the `Math.abs` function.</span></span> <span data-ttu-id="459b6-179">O objeto [Matemática](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/math) contém várias funções às quais seus scripts têm acesso.</span><span class="sxs-lookup"><span data-stu-id="459b6-179">The [Math](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/math) object contains many functions to which your scripts have access.</span></span> <span data-ttu-id="459b6-180">É possível encontrar mais informações sobre `Math` e outros objetos internos [Usando objetos JavaScript internos nos scripts do Office](../develop/javascript-objects.md).</span><span class="sxs-lookup"><span data-stu-id="459b6-180">More information about `Math` and other built-in objects can be found at [Using built-in JavaScript objects in Office Scripts](../develop/javascript-objects.md).</span></span>

1. <span data-ttu-id="459b6-181">Adicione o seguinte código ao final do script:</span><span class="sxs-lookup"><span data-stu-id="459b6-181">Add the following code to the end of the script:</span></span>

    ```TypeScript
        // Run the `Math.abs` function with the value at D2 and apply that value back to D2.
    let positiveValue = Math.abs(range.getValue());
    range.setValue(positiveValue);
    ```

    <span data-ttu-id="459b6-182">Observe que estamos usando `getValue` e `setValue`.</span><span class="sxs-lookup"><span data-stu-id="459b6-182">Note that we're using `getValue` and `setValue`.</span></span> <span data-ttu-id="459b6-183">Esses métodos funcionam em uma única célula.</span><span class="sxs-lookup"><span data-stu-id="459b6-183">These methods work on a single cell.</span></span> <span data-ttu-id="459b6-184">Ao lidar com intervalos de várias células, use `getValues` e `setValues`.</span><span class="sxs-lookup"><span data-stu-id="459b6-184">When handling multi-cell ranges, you'll want to use `getValues` and `setValues`.</span></span>

2. <span data-ttu-id="459b6-185">O valor da célula **D2** agora deverá ser positivo.</span><span class="sxs-lookup"><span data-stu-id="459b6-185">The value of cell **D2** should now be positive.</span></span>

## <a name="modify-the-values-of-a-column"></a><span data-ttu-id="459b6-186">Modificar os valores de uma coluna</span><span class="sxs-lookup"><span data-stu-id="459b6-186">Modify the values of a column</span></span>

<span data-ttu-id="459b6-187">Agora que sabemos ler e escrever em uma única célula, vamos generalizar o script para trabalhar em todas as colunas de **Débito** e **Crédito**.</span><span class="sxs-lookup"><span data-stu-id="459b6-187">Now that we know how to read and write to a single cell, let's generalize the script to work on the entire **Debit** and **Credit** columns.</span></span>

1. <span data-ttu-id="459b6-188">Remova o código que afeta apenas uma única célula (o código de valor absoluto anterior), de modo que o script agora se pareça com este:</span><span class="sxs-lookup"><span data-stu-id="459b6-188">Remove the code that affects only a single cell (the previous absolute value code), such that your script now looks like this:</span></span>

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

2. <span data-ttu-id="459b6-189">Adicione um loop que percorra as linhas nas duas últimas colunas.</span><span class="sxs-lookup"><span data-stu-id="459b6-189">Add a loop to the end of the script that iterates through the rows in the last two columns.</span></span> <span data-ttu-id="459b6-190">Para cada célula, o script define o valor para o valor absoluto do valor atual.</span><span class="sxs-lookup"><span data-stu-id="459b6-190">For each cell, the script sets the value to the current value's absolute value.</span></span>

    <span data-ttu-id="459b6-191">Observe que a matriz que define a localização das células é baseada em zero.</span><span class="sxs-lookup"><span data-stu-id="459b6-191">Note that the array defining cell locations is zero-based.</span></span> <span data-ttu-id="459b6-192">Isso significa que a célula **A1** é `range[0][0]`.</span><span class="sxs-lookup"><span data-stu-id="459b6-192">That means cell **A1** is `range[0][0]`.</span></span>

    ```TypeScript
    // Get the values of the used range.
    let range = selectedSheet.getUsedRange();
    let rangeValues = range.getValues();

    // Iterate over the fourth and fifth columns and set their values to their absolute value.
    for (let i = 1; i < range.getRowCount(); i++) {
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

    <span data-ttu-id="459b6-193">Essa parte do script faz várias tarefas importantes.</span><span class="sxs-lookup"><span data-stu-id="459b6-193">This portion of the script does several important tasks.</span></span> <span data-ttu-id="459b6-194">Primeiro, ela obtém os valores e a contagem de linhas do intervalo usado.</span><span class="sxs-lookup"><span data-stu-id="459b6-194">First, it gets the values and row count of the used range.</span></span> <span data-ttu-id="459b6-195">Isso nos permite ver os valores e saber quando parar.</span><span class="sxs-lookup"><span data-stu-id="459b6-195">This lets us look at values and know when to stop.</span></span> <span data-ttu-id="459b6-196">Segundo, ela reitera através do intervalo usado, verificando cada célula nas colunas **Débito** ou **Crédito**.</span><span class="sxs-lookup"><span data-stu-id="459b6-196">Second, it iterates through the used range, checking each cell in the **Debit** or **Credit** columns.</span></span> <span data-ttu-id="459b6-197">Por fim, se o valor na célula não for 0, ele será substituído pelo valor absoluto.</span><span class="sxs-lookup"><span data-stu-id="459b6-197">Finally, if the value in the cell is not 0, it is replaced by its absolute value.</span></span> <span data-ttu-id="459b6-198">Estamos evitando zeros, para que possamos deixar as células em branco.</span><span class="sxs-lookup"><span data-stu-id="459b6-198">We're avoiding zeroes so we can leave the blank cells as they were.</span></span>

3. <span data-ttu-id="459b6-199">Execute o script.</span><span class="sxs-lookup"><span data-stu-id="459b6-199">Run the script.</span></span>

    <span data-ttu-id="459b6-200">Seu extrato bancário agora deverá ter a seguinte aparência:</span><span class="sxs-lookup"><span data-stu-id="459b6-200">Your banking statement should now look like this:</span></span>

    ![O extrato bancário como uma tabela formatada apenas com valores positivos.](../images/tutorial-5.png)

## <a name="next-steps"></a><span data-ttu-id="459b6-202">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="459b6-202">Next steps</span></span>

<span data-ttu-id="459b6-203">Abra o Editor de códigos e experimente alguns dos [Scripts de exemplo para scripts do Office no Excel na Web](../resources/excel-samples.md).</span><span class="sxs-lookup"><span data-stu-id="459b6-203">Open the Code Editor and try out some of our [Sample scripts for Office Scripts in Excel on the web](../resources/excel-samples.md).</span></span> <span data-ttu-id="459b6-204">Visite também [Fundamentos de Scripts do Office no Excel na Web](../develop/scripting-fundamentals.md) para saber mais sobre como criar scripts do Office.</span><span class="sxs-lookup"><span data-stu-id="459b6-204">You can also visit [Scripting Fundamentals for Office Scripts in Excel on the web](../develop/scripting-fundamentals.md) to learn more about creating Office Scripts.</span></span>
