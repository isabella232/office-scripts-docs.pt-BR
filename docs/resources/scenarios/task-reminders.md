---
title: 'Cenário de exemplo de scripts do Office: lembretes de tarefas automatizadas'
description: Um exemplo que usa os cartões automatizados de energia e adaptável automatizar lembretes de tarefas em uma planilha de gerenciamento de projetos.
ms.date: 06/09/2020
localization_priority: Normal
ms.openlocfilehash: f764c37dafdd964e9435d504770d10b1608428b8
ms.sourcegitcommit: aec3c971c6640429f89b6bb99d2c95ea06725599
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/25/2020
ms.locfileid: "44878721"
---
# <a name="office-scripts-sample-scenario-automated-task-reminders"></a><span data-ttu-id="9f870-103">Cenário de exemplo de scripts do Office: lembretes de tarefas automatizadas</span><span class="sxs-lookup"><span data-stu-id="9f870-103">Office Scripts sample scenario: Automated task reminders</span></span>

<span data-ttu-id="9f870-104">Neste cenário, você está gerenciando um projeto.</span><span class="sxs-lookup"><span data-stu-id="9f870-104">In this scenario you're managing a project.</span></span> <span data-ttu-id="9f870-105">Você usa uma planilha do Excel para acompanhar o status de seus funcionários todos os meses.</span><span class="sxs-lookup"><span data-stu-id="9f870-105">You use an Excel worksheet to track your employees' status every month.</span></span> <span data-ttu-id="9f870-106">Você geralmente precisa lembrar as pessoas para preencher seu status, portanto, você optou por automatizar esse processo de lembrete.</span><span class="sxs-lookup"><span data-stu-id="9f870-106">You often need to remind people to fill out their status, so you've decided to automate that reminder process.</span></span>

<span data-ttu-id="9f870-107">Você criará um fluxo automatizado de energia para mensagens com campos de status ausentes e aplicará as respostas à planilha.</span><span class="sxs-lookup"><span data-stu-id="9f870-107">You'll create a Power Automate flow to message people with missing status fields and apply their responses to the spreadsheet.</span></span> <span data-ttu-id="9f870-108">Para fazer isso, você desenvolverá um par de scripts para lidar com a pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="9f870-108">To do this, you'll develop a pair of scripts to handle the working with the workbook.</span></span> <span data-ttu-id="9f870-109">O primeiro script Obtém uma lista de pessoas com status em branco e o segundo script adiciona uma cadeia de caracteres de status à linha à direita.</span><span class="sxs-lookup"><span data-stu-id="9f870-109">The first script gets a list of people with blank statuses and the second script adds a status string to the right row.</span></span> <span data-ttu-id="9f870-110">Você também fará uso de [cartões adaptáveis do teams](/microsoftteams/platform/task-modules-and-cards/what-are-cards) para que os funcionários insiram o status diretamente da notificação.</span><span class="sxs-lookup"><span data-stu-id="9f870-110">You'll also make use of [Teams Adaptive Cards](/microsoftteams/platform/task-modules-and-cards/what-are-cards) to have employees enter their status directly from the notification.</span></span>

## <a name="scripting-skills-covered"></a><span data-ttu-id="9f870-111">Habilidades de script abordadas</span><span class="sxs-lookup"><span data-stu-id="9f870-111">Scripting skills covered</span></span>

- <span data-ttu-id="9f870-112">Criar fluxos em automatização de energia</span><span class="sxs-lookup"><span data-stu-id="9f870-112">Create flows in Power Automate</span></span>
- <span data-ttu-id="9f870-113">Transmitir dados para scripts</span><span class="sxs-lookup"><span data-stu-id="9f870-113">Pass data to scripts</span></span>
- <span data-ttu-id="9f870-114">Retornar dados de scripts</span><span class="sxs-lookup"><span data-stu-id="9f870-114">Return data from scripts</span></span>
- <span data-ttu-id="9f870-115">Cartões adaptáveis do teams</span><span class="sxs-lookup"><span data-stu-id="9f870-115">Teams Adaptive Cards</span></span>
- <span data-ttu-id="9f870-116">Tabelas</span><span class="sxs-lookup"><span data-stu-id="9f870-116">Tables</span></span>

## <a name="prerequisites"></a><span data-ttu-id="9f870-117">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="9f870-117">Prerequisites</span></span>

<span data-ttu-id="9f870-118">Este cenário usa [automatização de energia](https://flow.microsoft.com) e [o Microsoft Teams](https://www.microsoft.com/microsoft-365/microsoft-teams/group-chat-software).</span><span class="sxs-lookup"><span data-stu-id="9f870-118">This scenario uses [Power Automate](https://flow.microsoft.com) and [Microsoft Teams](https://www.microsoft.com/microsoft-365/microsoft-teams/group-chat-software).</span></span> <span data-ttu-id="9f870-119">Você precisará associar-se à conta que você usa para desenvolver scripts do Office.</span><span class="sxs-lookup"><span data-stu-id="9f870-119">You will need both associated with the account that you use for developing Office Scripts.</span></span> <span data-ttu-id="9f870-120">Para obter acesso gratuito a uma assinatura de desenvolvedor da Microsoft para saber mais sobre o e trabalhar com esses aplicativos, considere participar do [programa de desenvolvedor do Microsoft 365](https://developer.microsoft.com/microsoft-365/dev-program).</span><span class="sxs-lookup"><span data-stu-id="9f870-120">For free access to a Microsoft Developer subscription to learn about and work with these applications, consider joining the [Microsoft 365 Developer Program](https://developer.microsoft.com/microsoft-365/dev-program).</span></span>

## <a name="setup-instructions"></a><span data-ttu-id="9f870-121">Instruções de configuração</span><span class="sxs-lookup"><span data-stu-id="9f870-121">Setup instructions</span></span>

1. <span data-ttu-id="9f870-122">Baixe <a href="task-reminders.xlsx">task-reminders.xlsx</a> para o onedrive.</span><span class="sxs-lookup"><span data-stu-id="9f870-122">Download <a href="task-reminders.xlsx">task-reminders.xlsx</a> to your OneDrive.</span></span>

2. <span data-ttu-id="9f870-123">Abra a pasta de trabalho no Excel na Web.</span><span class="sxs-lookup"><span data-stu-id="9f870-123">Open the workbook in Excel on the web.</span></span>

3. <span data-ttu-id="9f870-124">Na guia **automatizar** , abra o **Editor de código**.</span><span class="sxs-lookup"><span data-stu-id="9f870-124">Under the **Automate** tab, open the **Code Editor**.</span></span>

4. <span data-ttu-id="9f870-125">Primeiro, precisamos de um script para obter todos os funcionários com relatórios de status ausentes da planilha.</span><span class="sxs-lookup"><span data-stu-id="9f870-125">First, we need a script to get all the employees with status reports that are missing from the spreadsheet.</span></span> <span data-ttu-id="9f870-126">No painel de tarefas **Editor de código** , pressione **novo script** e cole o script a seguir no editor.</span><span class="sxs-lookup"><span data-stu-id="9f870-126">In the **Code Editor** task pane, press **New Script** and paste the following script into the editor.</span></span>

    ```typescript
    /**
     * This script looks for missing status reports in a project management table.
     *
     * @returns An array of Employee objects (containing their names and emails).
     */
    function main(workbook: ExcelScript.Workbook): Employee[] {
      // Get the first worksheet and the first table on that worksheet.
      let sheet = workbook.getFirstWorksheet()
      let table = sheet.getTables()[0];

      // Give the column indices names matching their expected content.
      const NAME_INDEX = 0;
      const EMAIL_INDEX = 1;
      const STATUS_REPORT_INDEX = 2;

      // Get the data for the whole table.
      let bodyRangeValues = table.getRangeBetweenHeaderAndTotal().getValues();

      // Create the array of Employee objects to return.
      let people: Employee[] = [];

      // Loop through the table and check each row for completion.
      for (let i = 0; i < bodyRangeValues.length; i++) {
        let row = bodyRangeValues[i];
        if (row[STATUS_REPORT_INDEX] === "") {
          // Save the email to return.
          people.push({ name: row[NAME_INDEX], email: row[EMAIL_INDEX] });
        }
      }

      // Log the array to verify we're getting the right rows.
      console.log(people);

      // Return the array of Employees.
      return people;
    }

    /**
     * An interface representing an employee.
     * An array of Employees will be returned from the script
     * for the Power Automate flow.
     */
    interface Employee {
      name: string;
      email: string;
    }
    ```

5. <span data-ttu-id="9f870-127">Salve o script com o nome **obter pessoas**.</span><span class="sxs-lookup"><span data-stu-id="9f870-127">Save the script with the name **Get People**.</span></span>

6. <span data-ttu-id="9f870-128">Em seguida, precisamos de um segundo script para processar os cartões de relatório de status e colocar as novas informações na planilha.</span><span class="sxs-lookup"><span data-stu-id="9f870-128">Next, we need a second script to process the status report cards and put the new information in the spreadsheet.</span></span> <span data-ttu-id="9f870-129">No painel de tarefas **Editor de código** , pressione **novo script** e cole o script a seguir no editor.</span><span class="sxs-lookup"><span data-stu-id="9f870-129">In the **Code Editor** task pane, press **New Script** and paste the following script into the editor.</span></span>

    ```typescript
    /**
     * This script applies the results of a Teams Adaptive Card about
     * a status update to a project management table.
     *
     * @param senderEmail - The email address of the employee updating their status.
     * @param statusReportResponse - The employee's status report.
     */
    function main(workbook: ExcelScript.Workbook,
      senderEmail: string,
      statusReportResponse: string) {

      // Get the first worksheet and the first table in that worksheet.
      let sheet = workbook.getFirstWorksheet();
      let table = sheet.getTables()[0];

      // Give the column indices names matching their expected content.
      const NAME_INDEX = 0;
      const EMAIL_INDEX = 1;
      const STATUS_REPORT_INDEX = 2;

      // Get the range and data for the whole table.
      let bodyRange = table.getRangeBetweenHeaderAndTotal();
      let tableRowCount = bodyRange.getRowCount();
      let bodyRangeValues = bodyRange.getValues();

      // Create a flag to denote success.
      let statusAdded = false;

      // Loop through the table and check each row for a matching email address.
      for (let i = 0; i < tableRowCount && !statusAdded; i++) {
        let row = bodyRangeValues[i];

        // Check if the row's email address matches.
        if (row[EMAIL_INDEX] === senderEmail) {
          // Add the Teams Adaptive Card response to the table.
          bodyRange.getCell(i, STATUS_REPORT_INDEX).setValues([
            [statusReportResponse]
          ]);
          statusAdded = true;
        }
      }

      // If successful, log the status update.
      if (statusAdded) {
        console.log(
          `Successfully added status report for ${senderEmail} containing: ${statusReportResponse}`
        );
      }
    }
    ```

7. <span data-ttu-id="9f870-130">Salve o script com o nome **salvar status**.</span><span class="sxs-lookup"><span data-stu-id="9f870-130">Save the script with the name **Save Status**.</span></span>

8. <span data-ttu-id="9f870-131">Agora, precisamos criar o fluxo.</span><span class="sxs-lookup"><span data-stu-id="9f870-131">Now, we need to create the flow.</span></span> <span data-ttu-id="9f870-132">Abrir [automatização de energia](https://flow.microsoft.com/).</span><span class="sxs-lookup"><span data-stu-id="9f870-132">Open [Power Automate](https://flow.microsoft.com/).</span></span>

    > [!TIP]
    > <span data-ttu-id="9f870-133">Se você ainda não criou um fluxo antes, confira nosso tutorial [comece a usar scripts com a automatização de energia](../../tutorials/excel-power-automate-manual.md) para aprender as noções básicas.</span><span class="sxs-lookup"><span data-stu-id="9f870-133">If you haven't created a flow before, please check out our tutorial [Start using scripts with Power Automate](../../tutorials/excel-power-automate-manual.md) to learn the basics.</span></span>

9. <span data-ttu-id="9f870-134">Criar um novo **fluxo instantâneo**.</span><span class="sxs-lookup"><span data-stu-id="9f870-134">Create a new **Instant flow**.</span></span>

10. <span data-ttu-id="9f870-135">Escolha **acionar manualmente um fluxo** das opções e pressione **criar**.</span><span class="sxs-lookup"><span data-stu-id="9f870-135">Choose **Manually trigger a flow** from the options and press **Create**.</span></span>

11. <span data-ttu-id="9f870-136">O fluxo precisa chamar o script **obter pessoas** para obter todos os funcionários com campos de status vazios.</span><span class="sxs-lookup"><span data-stu-id="9f870-136">The flow needs to call the **Get People** script to get all the employees with empty status fields.</span></span> <span data-ttu-id="9f870-137">Pressione **nova etapa** e selecione **Excel online (comercial)**.</span><span class="sxs-lookup"><span data-stu-id="9f870-137">Press **New step** and select **Excel Online (Business)**.</span></span> <span data-ttu-id="9f870-138">Em **Ações**, selecione **executar script (visualização)**.</span><span class="sxs-lookup"><span data-stu-id="9f870-138">Under **Actions**, select **Run script (preview)**.</span></span> <span data-ttu-id="9f870-139">Forneça as seguintes entradas para a etapa de fluxo:</span><span class="sxs-lookup"><span data-stu-id="9f870-139">Provide the following entries for the flow step:</span></span>

    - <span data-ttu-id="9f870-140">**Localização**: OneDrive for Business</span><span class="sxs-lookup"><span data-stu-id="9f870-140">**Location**: OneDrive for Business</span></span>
    - <span data-ttu-id="9f870-141">**Biblioteca de Documentos**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="9f870-141">**Document Library**: OneDrive</span></span>
    - <span data-ttu-id="9f870-142">**Arquivo**: task-reminders.xlsx</span><span class="sxs-lookup"><span data-stu-id="9f870-142">**File**: task-reminders.xlsx</span></span>
    - <span data-ttu-id="9f870-143">**Script**: obter pessoas</span><span class="sxs-lookup"><span data-stu-id="9f870-143">**Script**: Get People</span></span>

    ![A etapa de fluxo de script da primeira execução.](../../images/scenario-task-reminders-first-flow-step.png)

12. <span data-ttu-id="9f870-145">Em seguida, o fluxo precisa processar cada funcionário na matriz retornada pelo script.</span><span class="sxs-lookup"><span data-stu-id="9f870-145">Next, the flow needs to process each Employee in the array returned by the script.</span></span> <span data-ttu-id="9f870-146">Pressione **nova etapa** e selecione **postar um cartão adaptável a um usuário do Teams e aguarde uma resposta**.</span><span class="sxs-lookup"><span data-stu-id="9f870-146">Press **New step** and select **Post an Adaptive Card to a Teams user and wait for a response**.</span></span>

13. <span data-ttu-id="9f870-147">Para o campo **destinatário** , adicione **emails** do conteúdo dinâmico (a seleção terá o logotipo do Excel por ele).</span><span class="sxs-lookup"><span data-stu-id="9f870-147">For the **Recipient** field, add **email** from the dynamic content (the selection will have the Excel logo by it).</span></span> <span data-ttu-id="9f870-148">A adição de **email** faz com que a etapa de fluxo seja delimitada por um bloco **aplicar a cada** .</span><span class="sxs-lookup"><span data-stu-id="9f870-148">Adding **email** causes the flow step to be surrounded by an **Apply to each** block.</span></span> <span data-ttu-id="9f870-149">Isso significa que a matriz será iterada pela automatização de energia.</span><span class="sxs-lookup"><span data-stu-id="9f870-149">That means the array will be iterated over by Power Automate.</span></span>

14. <span data-ttu-id="9f870-150">O envio de um cartão adaptável exige que o JSON do cartão seja fornecido como a **mensagem**.</span><span class="sxs-lookup"><span data-stu-id="9f870-150">Sending an Adaptive Card requires the card's JSON to be provided as the **Message**.</span></span> <span data-ttu-id="9f870-151">Você pode usar o [Designer de cartão adaptável](https://adaptivecards.io/designer/) para criar cartões personalizados.</span><span class="sxs-lookup"><span data-stu-id="9f870-151">You can use the [Adaptive Card Designer](https://adaptivecards.io/designer/) to create custom cards.</span></span> <span data-ttu-id="9f870-152">Para este exemplo, use o JSON a seguir.</span><span class="sxs-lookup"><span data-stu-id="9f870-152">For this sample, use the following JSON.</span></span>  

    ```json
    {
      "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
      "type": "AdaptiveCard",
      "version": "1.0",
      "body": [
        {
          "type": "TextBlock",
          "size": "Medium",
          "weight": "Bolder",
          "text": "Update your Status Report"
        },
        {
          "type": "Image",
          "altText": "",
          "url": "https://i.imgur.com/f5RcuF3.png"
        },
        {
          "type": "TextBlock",
          "text": "This is a reminder to update your status report for this month's review. You can do so right here in this card, or by adding it directly to the spreadsheet.",
          "wrap": true
        },
        {
          "type": "Input.Text",
          "placeholder": "My status report for this month is...",
          "id": "response",
          "isMultiline": true
        }
      ],
      "actions": [
        {
          "type": "Action.Submit",
          "title": "Submit",
          "id": "submit"
        }
      ]
    }
    ```

15. <span data-ttu-id="9f870-153">Preencha os campos restantes da seguinte maneira:</span><span class="sxs-lookup"><span data-stu-id="9f870-153">Fill out the remaining fields as follows:</span></span>

    - <span data-ttu-id="9f870-154">**Mensagem de atualização**: Obrigado por enviar seu relatório de status.</span><span class="sxs-lookup"><span data-stu-id="9f870-154">**Update message**: Thank you for submitting your status report.</span></span> <span data-ttu-id="9f870-155">Sua resposta foi adicionada com êxito à planilha.</span><span class="sxs-lookup"><span data-stu-id="9f870-155">Your response has been successfully added to the spreadsheet.</span></span>
    - <span data-ttu-id="9f870-156">**Atualizar cartão**: Sim</span><span class="sxs-lookup"><span data-stu-id="9f870-156">**Should update card**: Yes</span></span>

16. <span data-ttu-id="9f870-157">No bloco **aplicar a cada** , seguindo o **cartão adaptável postar em um usuário do Teams e aguardar uma resposta**, pressione **Adicionar uma ação**.</span><span class="sxs-lookup"><span data-stu-id="9f870-157">In the **Apply to each** block, following the **Post an Adaptive Card to a Teams user and wait for a response**, press **Add an action**.</span></span> <span data-ttu-id="9f870-158">Selecione **Excel online (comercial)**.</span><span class="sxs-lookup"><span data-stu-id="9f870-158">Select **Excel Online (Business)**.</span></span> <span data-ttu-id="9f870-159">Em **Ações**, selecione **executar script (visualização)**.</span><span class="sxs-lookup"><span data-stu-id="9f870-159">Under **Actions**, select **Run script (preview)**.</span></span> <span data-ttu-id="9f870-160">Forneça as seguintes entradas para a etapa de fluxo:</span><span class="sxs-lookup"><span data-stu-id="9f870-160">Provide the following entries for the flow step:</span></span>

    - <span data-ttu-id="9f870-161">**Localização**: OneDrive for Business</span><span class="sxs-lookup"><span data-stu-id="9f870-161">**Location**: OneDrive for Business</span></span>
    - <span data-ttu-id="9f870-162">**Biblioteca de Documentos**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="9f870-162">**Document Library**: OneDrive</span></span>
    - <span data-ttu-id="9f870-163">**Arquivo**: task-reminders.xlsx</span><span class="sxs-lookup"><span data-stu-id="9f870-163">**File**: task-reminders.xlsx</span></span>
    - <span data-ttu-id="9f870-164">**Script**: salvar status</span><span class="sxs-lookup"><span data-stu-id="9f870-164">**Script**: Save Status</span></span>
    - <span data-ttu-id="9f870-165">**senderEmail**: email *(conteúdo dinâmico do Excel)*</span><span class="sxs-lookup"><span data-stu-id="9f870-165">**senderEmail**: email *(dynamic content from Excel)*</span></span>
    - <span data-ttu-id="9f870-166">**statusReportResponse**: resposta *(conteúdo dinâmico do Teams)*</span><span class="sxs-lookup"><span data-stu-id="9f870-166">**statusReportResponse**: response *(dynamic content from Teams)*</span></span>

    ![Etapa de aplicação de cada fluxo.](../../images/scenario-task-reminders-last-flow-step.png)

17. <span data-ttu-id="9f870-168">Salve o fluxo.</span><span class="sxs-lookup"><span data-stu-id="9f870-168">Save the flow.</span></span>

## <a name="running-the-flow"></a><span data-ttu-id="9f870-169">Executando o fluxo</span><span class="sxs-lookup"><span data-stu-id="9f870-169">Running the flow</span></span>

<span data-ttu-id="9f870-170">Para testar o fluxo, certifique-se de que todas as linhas de tabela com status em branco usem um endereço de email vinculado a uma conta de equipe (provavelmente, você deve usar seu próprio endereço de email durante o teste).</span><span class="sxs-lookup"><span data-stu-id="9f870-170">To test the flow, make sure any table rows with blank status use an email address tied to a Teams account (you should probably use your own email address while testing).</span></span>

<span data-ttu-id="9f870-171">Você pode selecionar **testar** no editor de fluxo ou executar o fluxo na página **meus fluxos** .</span><span class="sxs-lookup"><span data-stu-id="9f870-171">You can either select **Test** from the flow designer, or run the flow from the **My flows** page.</span></span> <span data-ttu-id="9f870-172">Depois de iniciar o fluxo e aceitar o uso das conexões necessárias, você receberá um cartão adaptável do Power Automated Teams.</span><span class="sxs-lookup"><span data-stu-id="9f870-172">After starting the flow and accepting the use of the required connections, you should receive an Adaptive Card from Power Automate through Teams.</span></span> <span data-ttu-id="9f870-173">Depois que você preencher o campo status no cartão, o fluxo continuará e atualizará a planilha com o status que você fornecer.</span><span class="sxs-lookup"><span data-stu-id="9f870-173">Once you fill out the status field in the card, the flow will continue and update the spreadsheet with the status you provide.</span></span>

### <a name="before-running-the-flow"></a><span data-ttu-id="9f870-174">Antes de executar o fluxo</span><span class="sxs-lookup"><span data-stu-id="9f870-174">Before running the flow</span></span>

![Uma planilha com um relatório de status contendo uma entrada de status ausente.](../../images/scenario-task-reminders-spreadsheet-before.png)

### <a name="receiving-the-adaptive-card"></a><span data-ttu-id="9f870-176">Recebendo o cartão adaptável</span><span class="sxs-lookup"><span data-stu-id="9f870-176">Receiving the Adaptive Card</span></span>

![Um cartão adaptável no Microsoft Teams solicitando ao funcionário uma atualização de status.](../../images/scenario-task-reminders-adaptive-card.png)

### <a name="after-running-the-flow"></a><span data-ttu-id="9f870-178">Após a execução do fluxo</span><span class="sxs-lookup"><span data-stu-id="9f870-178">After running the flow</span></span>

![Uma planilha com um relatório de status com uma entrada de status já preenchida.](../../images/scenario-task-reminders-spreadsheet-after.png)
