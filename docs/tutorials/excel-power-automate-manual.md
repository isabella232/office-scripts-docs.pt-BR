---
title: Começar a usar scripts com a automatização de energia
description: Tutorial sobre como integrar a automatização de energia com scripts do Office usando um gatilho manual.
ms.date: 06/09/2020
localization_priority: Priority
ms.openlocfilehash: 37c2d9ae4c5456a1355362c70695fc61c236a725
ms.sourcegitcommit: aec3c971c6640429f89b6bb99d2c95ea06725599
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/25/2020
ms.locfileid: "44878710"
---
# <a name="start-using-scripts-with-power-automate-preview"></a><span data-ttu-id="9a450-103">Começar a usar scripts com a automatização de energia (visualização)</span><span class="sxs-lookup"><span data-stu-id="9a450-103">Start using scripts with Power Automate (preview)</span></span>

<span data-ttu-id="9a450-104">Este tutorial ensina a executar um script do Office para Excel na Web através da [automatização de energia](https://flow.microsoft.com).</span><span class="sxs-lookup"><span data-stu-id="9a450-104">This tutorial teaches you how to run an Office Script for Excel on the web through [Power Automate](https://flow.microsoft.com).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="9a450-105">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="9a450-105">Prerequisites</span></span>

[!INCLUDE [Tutorial prerequisites](../includes/tutorial-prerequisites.md)]

> [!IMPORTANT]
> <span data-ttu-id="9a450-106">Este tutorial pressupõe que você tenha concluído o tutorial [gravar, editar e criar scripts do Office no Excel na Web](excel-tutorial.md) .</span><span class="sxs-lookup"><span data-stu-id="9a450-106">This tutorial assumes you have completed the [Record, edit, and create Office Scripts in Excel on the web](excel-tutorial.md) tutorial.</span></span>

## <a name="prepare-the-workbook"></a><span data-ttu-id="9a450-107">Preparar a pasta de trabalho</span><span class="sxs-lookup"><span data-stu-id="9a450-107">Prepare the workbook</span></span>

<span data-ttu-id="9a450-108">A automatização de energia não pode usar referências relativas como `Workbook.getActiveWorksheet` acessar componentes de pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="9a450-108">Power Automate can't use relative references like `Workbook.getActiveWorksheet` to access workbook components.</span></span> <span data-ttu-id="9a450-109">Portanto, precisamos de uma pasta de trabalho e planilha com nomes consistentes que os recursos de automatização podem fazer referência.</span><span class="sxs-lookup"><span data-stu-id="9a450-109">So, we need a workbook and worksheet with consistent names that Power Automate can reference.</span></span>

1. <span data-ttu-id="9a450-110">Crie uma nova pasta de trabalho chamada **myworkbook**.</span><span class="sxs-lookup"><span data-stu-id="9a450-110">Create a new workbook named **MyWorkbook**.</span></span>

2. <span data-ttu-id="9a450-111">Na pasta de trabalho **myworkbook** , crie uma planilha chamada **TutorialWorksheet**.</span><span class="sxs-lookup"><span data-stu-id="9a450-111">In the **MyWorkbook** workbook, create a worksheet called **TutorialWorksheet**.</span></span>

## <a name="create-an-office-script"></a><span data-ttu-id="9a450-112">Criar um script do Office</span><span class="sxs-lookup"><span data-stu-id="9a450-112">Create an Office Script</span></span>

1. <span data-ttu-id="9a450-113">Vá para a guia **automatizar** e selecione **Editor de código**.</span><span class="sxs-lookup"><span data-stu-id="9a450-113">Go to the **Automate** tab and select **Code Editor**.</span></span>

2. <span data-ttu-id="9a450-114">Selecione **novo script**.</span><span class="sxs-lookup"><span data-stu-id="9a450-114">Select **New Script**.</span></span>

3. <span data-ttu-id="9a450-115">Substitua o script padrão pelo script a seguir.</span><span class="sxs-lookup"><span data-stu-id="9a450-115">Replace the default script with the following script.</span></span> <span data-ttu-id="9a450-116">Esse script adiciona a data e hora atuais às duas primeiras células da planilha **TutorialWorksheet** .</span><span class="sxs-lookup"><span data-stu-id="9a450-116">This script adds the current date and time to the first two cells of the **TutorialWorksheet** worksheet.</span></span>

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

4. <span data-ttu-id="9a450-117">Renomeie o script para **Definir data e hora**.</span><span class="sxs-lookup"><span data-stu-id="9a450-117">Rename the script to **Set date and time**.</span></span> <span data-ttu-id="9a450-118">Pressione o nome do script para alterá-lo.</span><span class="sxs-lookup"><span data-stu-id="9a450-118">Press the script name to change it.</span></span>

5. <span data-ttu-id="9a450-119">Salve o script pressionando **Salvar script**.</span><span class="sxs-lookup"><span data-stu-id="9a450-119">Save the script by pressing **Save Script**.</span></span>

## <a name="create-an-automated-workflow-with-power-automate"></a><span data-ttu-id="9a450-120">Criar um fluxo de trabalho automatizado com a automatização de energia</span><span class="sxs-lookup"><span data-stu-id="9a450-120">Create an automated workflow with Power Automate</span></span>

1. <span data-ttu-id="9a450-121">Entre no site de [visualização de energia automatizada](https://flow.microsoft.com).</span><span class="sxs-lookup"><span data-stu-id="9a450-121">Sign in to the [Power Automate preview site](https://flow.microsoft.com).</span></span>

2. <span data-ttu-id="9a450-122">No menu que é exibido no lado esquerdo da tela, pressione **criar**.</span><span class="sxs-lookup"><span data-stu-id="9a450-122">In the menu that's displayed on the left side of the screen, press **Create**.</span></span> <span data-ttu-id="9a450-123">Isso lhe permite listar maneiras de criar novos fluxos de trabalho.</span><span class="sxs-lookup"><span data-stu-id="9a450-123">This brings you to list of ways to create new workflows.</span></span>

    ![O botão criar na automatização de energia.](../images/power-automate-tutorial-1.png)

3. <span data-ttu-id="9a450-125">Na seção **Iniciar com base em branco** , selecione **fluxo instantâneo**.</span><span class="sxs-lookup"><span data-stu-id="9a450-125">In the **Start from blank** section, select **Instant flow**.</span></span> <span data-ttu-id="9a450-126">Isso cria um fluxo de trabalho ativado manualmente.</span><span class="sxs-lookup"><span data-stu-id="9a450-126">This creates a manually activated workflow.</span></span>

    ![A opção de fluxo instantâneo para a criação de um novo fluxo de trabalho.](../images/power-automate-tutorial-2.png)

4. <span data-ttu-id="9a450-128">Na janela de diálogo exibida, insira um nome para o fluxo na caixa de **texto nome do fluxo** , selecione **acionar manualmente um fluxo** na lista de opções em **escolher como acionar o fluxo**e pressione **criar**.</span><span class="sxs-lookup"><span data-stu-id="9a450-128">In the dialog window that appears, enter a name for your flow in the **Flow name** text box, select **Manually trigger a flow** from the list of options under **Choose how to trigger the flow**, and press **Create**.</span></span>

    ![A opção de gatilho manual para criar um novo fluxo instantâneo.](../images/power-automate-tutorial-3.png)

5. <span data-ttu-id="9a450-130">Pressione **nova etapa**.</span><span class="sxs-lookup"><span data-stu-id="9a450-130">Press **New step**.</span></span>

6. <span data-ttu-id="9a450-131">Selecione a guia **padrão** e, em seguida, selecione **Excel online (comercial)**.</span><span class="sxs-lookup"><span data-stu-id="9a450-131">Select the **Standard** tab, then select **Excel Online (Business)**.</span></span>

    ![A opção de automatização de energia para o Excel online (Business).](../images/power-automate-tutorial-4.png)

7. <span data-ttu-id="9a450-133">Em **ações**, selecione **Executar script (versão prévia)**.</span><span class="sxs-lookup"><span data-stu-id="9a450-133">Under **Actions**, select **Run script (preview)**.</span></span>

    ![A opção de ação automatizar a energia para executar script (visualização).](../images/power-automate-tutorial-5.png)

8. <span data-ttu-id="9a450-135">Especifique as seguintes configurações para executar o conector de **script** :</span><span class="sxs-lookup"><span data-stu-id="9a450-135">Specify the following settings for the **Run script** connector:</span></span>

    - <span data-ttu-id="9a450-136">**Local**: onedrive for Business</span><span class="sxs-lookup"><span data-stu-id="9a450-136">**Location**: OneDrive for Business</span></span>
    - <span data-ttu-id="9a450-137">**Biblioteca de documentos**: onedrive</span><span class="sxs-lookup"><span data-stu-id="9a450-137">**Document Library**: OneDrive</span></span>
    - <span data-ttu-id="9a450-138">**Arquivo**: MyWorkbook.xlsx</span><span class="sxs-lookup"><span data-stu-id="9a450-138">**File**: MyWorkbook.xlsx</span></span>
    - <span data-ttu-id="9a450-139">**Script**: Definir data e hora</span><span class="sxs-lookup"><span data-stu-id="9a450-139">**Script**: Set date and time</span></span>

    ![As configurações de conector para executar um script em automatização de energia.](../images/power-automate-tutorial-6.png)

9. <span data-ttu-id="9a450-141">Pressione **salvar**.</span><span class="sxs-lookup"><span data-stu-id="9a450-141">Press **Save**.</span></span>

<span data-ttu-id="9a450-142">Agora, o fluxo está pronto para ser executado através da automatização de energia.</span><span class="sxs-lookup"><span data-stu-id="9a450-142">Your flow is now ready to be run through Power Automate.</span></span> <span data-ttu-id="9a450-143">Você pode testá-lo usando o botão **testar** no editor de fluxo ou siga as etapas restantes do tutorial para executar o fluxo de sua coleção de fluxo.</span><span class="sxs-lookup"><span data-stu-id="9a450-143">You can test it using the **Test** button in the flow editor or follow the remaining tutorial steps to run the flow from your flow collection.</span></span>

## <a name="run-the-script-through-power-automate"></a><span data-ttu-id="9a450-144">Executar o script através da automatização de energia</span><span class="sxs-lookup"><span data-stu-id="9a450-144">Run the script through Power Automate</span></span>

1. <span data-ttu-id="9a450-145">Na página automatizar alimentação principal, selecione **meus fluxos**.</span><span class="sxs-lookup"><span data-stu-id="9a450-145">From the main Power Automate page, select **My flows**.</span></span>

    ![O botão meus fluxos em automatização de energia.](../images/power-automate-tutorial-7.png)

2. <span data-ttu-id="9a450-147">Selecione **meu fluxo de tutorial** na lista de fluxos exibida na guia **meus fluxos** . Isso mostra os detalhes do fluxo que criamos anteriormente.</span><span class="sxs-lookup"><span data-stu-id="9a450-147">Select **My tutorial flow** from the list of flows displayed in the **My flows** tab. This shows the details of the flow we previously created.</span></span>

3. <span data-ttu-id="9a450-148">Pressione **executar**.</span><span class="sxs-lookup"><span data-stu-id="9a450-148">Press **Run**.</span></span>

    ![O botão Executar em automatização de energia.](../images/power-automate-tutorial-8.png)

4. <span data-ttu-id="9a450-150">Um painel de tarefas será exibido para executar o fluxo.</span><span class="sxs-lookup"><span data-stu-id="9a450-150">A task pane will appear for running the flow.</span></span> <span data-ttu-id="9a450-151">Se você for solicitado a **entrar no** Excel online, pressione **continuar**.</span><span class="sxs-lookup"><span data-stu-id="9a450-151">If you are asked to **Sign in** to Excel Online, do so by pressing **Continue**.</span></span>

5. <span data-ttu-id="9a450-152">Pressione o **fluxo de execução**.</span><span class="sxs-lookup"><span data-stu-id="9a450-152">Press **Run flow**.</span></span> <span data-ttu-id="9a450-153">Isso executa o fluxo, que executa o script relacionado do Office.</span><span class="sxs-lookup"><span data-stu-id="9a450-153">This runs the flow, which runs the related Office Script.</span></span>

6. <span data-ttu-id="9a450-154">Pressione **concluído**.</span><span class="sxs-lookup"><span data-stu-id="9a450-154">Press **Done**.</span></span> <span data-ttu-id="9a450-155">Você deve ver a seção **runs** Update de acordo.</span><span class="sxs-lookup"><span data-stu-id="9a450-155">You should see the **Runs** section update accordingly.</span></span>

7. <span data-ttu-id="9a450-156">Atualize a página para ver os resultados da automatização de energia.</span><span class="sxs-lookup"><span data-stu-id="9a450-156">Refresh the page to see the results of the Power Automate.</span></span> <span data-ttu-id="9a450-157">Se tiver êxito, vá para a pasta de trabalho para ver as células atualizadas.</span><span class="sxs-lookup"><span data-stu-id="9a450-157">If it succeeded, go to the workbook to see the updated cells.</span></span> <span data-ttu-id="9a450-158">Se ele falhar, verifique as configurações do fluxo e execute-o uma segunda vez.</span><span class="sxs-lookup"><span data-stu-id="9a450-158">If it failed, verify the flow's settings and run it a second time.</span></span>

    ![Saída automatizada de energia mostrando uma execução de fluxo bem-sucedida.](../images/power-automate-tutorial-9.png)

## <a name="next-steps"></a><span data-ttu-id="9a450-160">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="9a450-160">Next steps</span></span>

<span data-ttu-id="9a450-161">Preencha os [scripts executados automaticamente com o tutorial automatizar de energia](excel-power-automate-trigger.md) .</span><span class="sxs-lookup"><span data-stu-id="9a450-161">Complete the [Automatically run scripts with Power Automate](excel-power-automate-trigger.md) tutorial.</span></span> <span data-ttu-id="9a450-162">Ele ensina como transmitir dados de um serviço de fluxo de trabalho para o script do Office.</span><span class="sxs-lookup"><span data-stu-id="9a450-162">It teaches you how to pass data from a workflow service to your Office Script.</span></span>
