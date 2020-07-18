---
title: Comece a usar scripts de um fluxo manual do Power Automate
description: Um tutorial sobre o uso de Scripts do Office no Power Automate por meio de um acionamento manual.
ms.date: 07/14/2020
localization_priority: Priority
ms.openlocfilehash: 70fca2620973ecefe9eda40f02e28f064b713677
ms.sourcegitcommit: ebd1079c7e2695ac0e7e4c616f2439975e196875
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/17/2020
ms.locfileid: "45160430"
---
# <a name="call-scripts-from-a-manual-power-automate-flow-preview"></a><span data-ttu-id="22070-103">Comece a usar scripts de um fluxo manual do Power Automate (pré-visualização)</span><span class="sxs-lookup"><span data-stu-id="22070-103">Call scripts from a manual Power Automate flow (preview)</span></span>

<span data-ttu-id="22070-104">Este tutorial ensina como executar um Script do Office para o Excel na web por meio do [Power Automate](https://flow.microsoft.com).</span><span class="sxs-lookup"><span data-stu-id="22070-104">This tutorial teaches you how to run an Office Script for Excel on the web through [Power Automate](https://flow.microsoft.com).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="22070-105">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="22070-105">Prerequisites</span></span>

[!INCLUDE [Tutorial prerequisites](../includes/power-automate-tutorial-prerequisites.md)]

> [!IMPORTANT]
> <span data-ttu-id="22070-106">Este tutorial pressupõe que você tenha concluído o tutorial [Registrar, editar e criar Scripts do Office no Excel na Web](excel-tutorial.md).</span><span class="sxs-lookup"><span data-stu-id="22070-106">This tutorial assumes you have completed the [Record, edit, and create Office Scripts in Excel on the web](excel-tutorial.md) tutorial.</span></span>

## <a name="prepare-the-workbook"></a><span data-ttu-id="22070-107">Preparar a pasta de trabalho</span><span class="sxs-lookup"><span data-stu-id="22070-107">Prepare the workbook</span></span>

<span data-ttu-id="22070-108">O Power Automate não consegue usar referências relativas como `Workbook.getActiveWorksheet` para acessar os componentes da pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="22070-108">Power Automate can't use relative references like `Workbook.getActiveWorksheet` to access workbook components.</span></span> <span data-ttu-id="22070-109">Portanto, precisamos de uma pasta de trabalho e de uma planilha com nomes consistentes que o Power Automate consiga consultar.</span><span class="sxs-lookup"><span data-stu-id="22070-109">So, we need a workbook and worksheet with consistent names that Power Automate can reference.</span></span>

1. <span data-ttu-id="22070-110">Crie uma pasta de trabalho intitulada **MyWorkbook**.</span><span class="sxs-lookup"><span data-stu-id="22070-110">Create a new workbook named **MyWorkbook**.</span></span>

2. <span data-ttu-id="22070-111">Na pasta de trabalho **MyWorkbook**, crie uma planilha intitulada **TutorialWorksheet**.</span><span class="sxs-lookup"><span data-stu-id="22070-111">In the **MyWorkbook** workbook, create a worksheet called **TutorialWorksheet**.</span></span>

## <a name="create-an-office-script"></a><span data-ttu-id="22070-112">Criar um Script do Office</span><span class="sxs-lookup"><span data-stu-id="22070-112">Create an Office Script</span></span>

1. <span data-ttu-id="22070-113">Vá para a guia **Automatizar** e selecione **Editor de Códigos**.</span><span class="sxs-lookup"><span data-stu-id="22070-113">Go to the **Automate** tab and select **Code Editor**.</span></span>

2. <span data-ttu-id="22070-114">Selecione **Novo Script**.</span><span class="sxs-lookup"><span data-stu-id="22070-114">Select **New Script**.</span></span>

3. <span data-ttu-id="22070-115">Substitua o script padrão pelo script abaixo.</span><span class="sxs-lookup"><span data-stu-id="22070-115">Replace the default script with the following script.</span></span> <span data-ttu-id="22070-116">Esse script adiciona a data e hora atuais às duas primeiras células da planilha **TutorialWorksheet**.</span><span class="sxs-lookup"><span data-stu-id="22070-116">This script adds the current date and time to the first two cells of the **TutorialWorksheet** worksheet.</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
      // Get the "TutorialWorksheet" worksheet from the workbook.
      let worksheet = workbook.getWorksheet("TutorialWorksheet");

      // Get the cells at A1 and B1.
      let dateRange = worksheet.getRange("A1");
      let timeRange = worksheet.getRange("B1");

      // Get the current date and time using the JavaScript Date object.
      let date = new Date(Date.now());

      // Add the date string to A1.
      dateRange.setValue(date.toLocaleDateString());

      // Add the time string to B1.
      timeRange.setValue(date.toLocaleTimeString());
    }
    ```

4. <span data-ttu-id="22070-117">Renomeie o script como **Definir data e hora**.</span><span class="sxs-lookup"><span data-stu-id="22070-117">Rename the script to **Set date and time**.</span></span> <span data-ttu-id="22070-118">Pressione o nome do script para alterá-lo.</span><span class="sxs-lookup"><span data-stu-id="22070-118">Press the script name to change it.</span></span>

5. <span data-ttu-id="22070-119">Salve o script pressionando **Salvar Script**.</span><span class="sxs-lookup"><span data-stu-id="22070-119">Save the script by pressing **Save Script**.</span></span>

## <a name="create-an-automated-workflow-with-power-automate"></a><span data-ttu-id="22070-120">Criar um fluxo de trabalho automatizado com o Power Automate</span><span class="sxs-lookup"><span data-stu-id="22070-120">Create an automated workflow with Power Automate</span></span>

1. <span data-ttu-id="22070-121">Entre no [site do Power Automate](https://flow.microsoft.com).</span><span class="sxs-lookup"><span data-stu-id="22070-121">Sign in to the [Power Automate site](https://flow.microsoft.com).</span></span>

2. <span data-ttu-id="22070-122">No menu exibido do lado esquerdo da tela, pressione **Criar**.</span><span class="sxs-lookup"><span data-stu-id="22070-122">In the menu that's displayed on the left side of the screen, press **Create**.</span></span> <span data-ttu-id="22070-123">Isso o conduzirá a uma lista de maneiras de criar novos fluxos de trabalho.</span><span class="sxs-lookup"><span data-stu-id="22070-123">This brings you to list of ways to create new workflows.</span></span>

    ![Botão Criar no Power Automate.](../images/power-automate-tutorial-1.png)

3. <span data-ttu-id="22070-125">Na seção **Começar no espaço em branco**, selecione **Fluxo instantâneo**.</span><span class="sxs-lookup"><span data-stu-id="22070-125">In the **Start from blank** section, select **Instant flow**.</span></span> <span data-ttu-id="22070-126">Isso irá criar um fluxo de trabalho ativado manualmente.</span><span class="sxs-lookup"><span data-stu-id="22070-126">This creates a manually activated workflow.</span></span>

    ![Opção Fluxo instantâneo para criar um novo fluxo de trabalho.](../images/power-automate-tutorial-2.png)

4. <span data-ttu-id="22070-128">Na janela da caixa de diálogo que aparece, insira um nome para o seu fluxo na caixa de texto **Nome do fluxo**; selecione **Acionar um fluxo manualmente** na lista de opções em **Escolher como acionar o fluxo**, e pressione **Criar**.</span><span class="sxs-lookup"><span data-stu-id="22070-128">In the dialog window that appears, enter a name for your flow in the **Flow name** text box, select **Manually trigger a flow** from the list of options under **Choose how to trigger the flow**, and press **Create**.</span></span>

    ![Opção acionamento manual para a criação de um novo fluxo instantâneo.](../images/power-automate-tutorial-3.png)

    <span data-ttu-id="22070-130">Observe que o fluxo acionado manualmente é apenas um entre os diversos tipos de fluxo.</span><span class="sxs-lookup"><span data-stu-id="22070-130">Note that a manually triggered flow is just one of many types of flows.</span></span> <span data-ttu-id="22070-131">No tutorial a seguir, você criará um fluxo que é executado automaticamente quando você recebe um email.</span><span class="sxs-lookup"><span data-stu-id="22070-131">In the next tutorial, you'll make a flow that automatically runs when you receive an email.</span></span>

5. <span data-ttu-id="22070-132">Pressione **Nova etapa**.</span><span class="sxs-lookup"><span data-stu-id="22070-132">Press **New step**.</span></span>

6. <span data-ttu-id="22070-133">Selecione a guia **Padrão** e, em seguida, selecione **Excel Online (Business)**.</span><span class="sxs-lookup"><span data-stu-id="22070-133">Select the **Standard** tab, then select **Excel Online (Business)**.</span></span>

    ![Opção Power Automate para o Excel Online (Business).](../images/power-automate-tutorial-4.png)

7. <span data-ttu-id="22070-135">Em **Ações**, selecione **Executar script (pré-visualização)**.</span><span class="sxs-lookup"><span data-stu-id="22070-135">Under **Actions**, select **Run script (preview)**.</span></span>

    ![Opção de ação do Power Automate para Executar script (pré-visualização).](../images/power-automate-tutorial-5.png)

8. <span data-ttu-id="22070-137">Especifique as seguintes configurações para o conector **Executar script**:</span><span class="sxs-lookup"><span data-stu-id="22070-137">Specify the following settings for the **Run script** connector:</span></span>

    - <span data-ttu-id="22070-138">**Localização**: OneDrive for Business</span><span class="sxs-lookup"><span data-stu-id="22070-138">**Location**: OneDrive for Business</span></span>
    - <span data-ttu-id="22070-139">**Biblioteca de Documentos**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="22070-139">**Document Library**: OneDrive</span></span>
    - <span data-ttu-id="22070-140">**Arquivo**: MyWorkbook.xlsx</span><span class="sxs-lookup"><span data-stu-id="22070-140">**File**: MyWorkbook.xlsx</span></span>
    - <span data-ttu-id="22070-141">**Script**: Definir data e hora</span><span class="sxs-lookup"><span data-stu-id="22070-141">**Script**: Set date and time</span></span>

    ![Configurações do conector para executar um script no Power Automate.](../images/power-automate-tutorial-6.png)

9. <span data-ttu-id="22070-143">Pressione **Salvar**.</span><span class="sxs-lookup"><span data-stu-id="22070-143">Press **Save**.</span></span>

<span data-ttu-id="22070-144">Seu fluxo agora está pronto para ser executado por meio do Power Automate.</span><span class="sxs-lookup"><span data-stu-id="22070-144">Your flow is now ready to be run through Power Automate.</span></span> <span data-ttu-id="22070-145">Você pode testá-lo usando o botão **Testar** no editor de fluxo ou seguir as etapas restantes do tutorial para executar o fluxo a partir da sua coleção de fluxos.</span><span class="sxs-lookup"><span data-stu-id="22070-145">You can test it using the **Test** button in the flow editor or follow the remaining tutorial steps to run the flow from your flow collection.</span></span>

## <a name="run-the-script-through-power-automate"></a><span data-ttu-id="22070-146">Executar o script por meio da automação</span><span class="sxs-lookup"><span data-stu-id="22070-146">Run the script through Power Automate</span></span>

1. <span data-ttu-id="22070-147">Na página principal do Power Automate, selecione **Meus fluxos**.</span><span class="sxs-lookup"><span data-stu-id="22070-147">From the main Power Automate page, select **My flows**.</span></span>

    ![Botão Meus fluxos no Power Automate.](../images/power-automate-tutorial-7.png)

2. <span data-ttu-id="22070-149">Selecione **Fluxo do meu tutorial** na lista de fluxos exibida na guia **Meus fluxos**. Isso irá lhe mostrar os detalhes do fluxo que criamos anteriormente.</span><span class="sxs-lookup"><span data-stu-id="22070-149">Select **My tutorial flow** from the list of flows displayed in the **My flows** tab. This shows the details of the flow we previously created.</span></span>

3. <span data-ttu-id="22070-150">Pressione **Executar**.</span><span class="sxs-lookup"><span data-stu-id="22070-150">Press **Run**.</span></span>

    ![Botão Executar no Power Automate.](../images/power-automate-tutorial-8.png)

4. <span data-ttu-id="22070-152">Um painel de tarefas irá aparecer para executar o fluxo.</span><span class="sxs-lookup"><span data-stu-id="22070-152">A task pane will appear for running the flow.</span></span> <span data-ttu-id="22070-153">Se você for solicitado a **Entrar** no Excel Online, faça o login pressionando **Continuar**.</span><span class="sxs-lookup"><span data-stu-id="22070-153">If you are asked to **Sign in** to Excel Online, do so by pressing **Continue**.</span></span>

5. <span data-ttu-id="22070-154">Pressione **Executar o fluxo**.</span><span class="sxs-lookup"><span data-stu-id="22070-154">Press **Run flow**.</span></span> <span data-ttu-id="22070-155">Isso executará o fluxo, que, por sua vez, executará o Script do Office associado.</span><span class="sxs-lookup"><span data-stu-id="22070-155">This runs the flow, which runs the related Office Script.</span></span>

6. <span data-ttu-id="22070-156">Pressione **Concluído**.</span><span class="sxs-lookup"><span data-stu-id="22070-156">Press **Done**.</span></span> <span data-ttu-id="22070-157">Você deverá ver a seção **Executar** ser atualizada de acordo.</span><span class="sxs-lookup"><span data-stu-id="22070-157">You should see the **Runs** section update accordingly.</span></span>

7. <span data-ttu-id="22070-158">Atualize a página para ver os resultados do Power Automate.</span><span class="sxs-lookup"><span data-stu-id="22070-158">Refresh the page to see the results of the Power Automate.</span></span> <span data-ttu-id="22070-159">Se o script tiver sido bem-sucedido, vá para a pasta de trabalho para ver as células atualizadas.</span><span class="sxs-lookup"><span data-stu-id="22070-159">If it succeeded, go to the workbook to see the updated cells.</span></span> <span data-ttu-id="22070-160">Se tiver falhado, verifique as configurações do fluxo e execute-o novamente.</span><span class="sxs-lookup"><span data-stu-id="22070-160">If it failed, verify the flow's settings and run it a second time.</span></span>

    ![Resultado do Power Automate mostrando um fluxo executado com sucesso.](../images/power-automate-tutorial-9.png)

## <a name="next-steps"></a><span data-ttu-id="22070-162">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="22070-162">Next steps</span></span>

<span data-ttu-id="22070-163">Faça o tutorial [Transferir dados para scripts em um fluxo executado automaticamente pelo Power Automate](excel-power-automate-trigger.md).</span><span class="sxs-lookup"><span data-stu-id="22070-163">Complete the [Pass data to scripts in an automatically-run Power Automate flow](excel-power-automate-trigger.md) tutorial.</span></span> <span data-ttu-id="22070-164">O tutorial ensinará como transferir dados de um serviço de fluxo de trabalho para o seu Script do Office e executar o fluxo do Power Automate quando certos eventos ocorrerem.</span><span class="sxs-lookup"><span data-stu-id="22070-164">It teaches you how to pass data from a workflow service to your Office Script and run the Power Automate flow when certain events occur.</span></span>
