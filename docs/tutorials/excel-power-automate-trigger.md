---
title: Passar dados para scripts numa execução automática do fluxo no Power Automate.
description: Tutorial sobre executar os Scripts do Office para Excel na Web por meio do Power Automate quando emails são recebidos e transmitidos para o script.
ms.date: 07/14/2020
localization_priority: Priority
ms.openlocfilehash: c024891e187f22b7d10f6e9d52d262dc2ec4057f
ms.sourcegitcommit: ebd1079c7e2695ac0e7e4c616f2439975e196875
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/17/2020
ms.locfileid: "45160478"
---
# <a name="pass-data-to-scripts-in-an-automatically-run-power-automate-flow-preview"></a><span data-ttu-id="cb169-103">Passar dados para scripts em modo de execução automático no fluxo do Power Automate (visualização)</span><span class="sxs-lookup"><span data-stu-id="cb169-103">Pass data to scripts in an automatically-run Power Automate flow (preview)</span></span>

<span data-ttu-id="cb169-104">Este tutorial ensina como usar um script do Office para Excel na Web fluxo automatizado[ do Power Automate](https://flow.microsoft.com).</span><span class="sxs-lookup"><span data-stu-id="cb169-104">This tutorial teaches you how to use an Office Script for Excel on the web with an automated [Power Automate](https://flow.microsoft.com) workflow.</span></span> <span data-ttu-id="cb169-105">Seu script irá automaticamente ser executado toda vez que você receber um email, gravando informações do email em uma pasta de trabalho do Excel.</span><span class="sxs-lookup"><span data-stu-id="cb169-105">Your script will automatically run each time you receive an email, recording information from the email in an Excel workbook.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="cb169-106">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="cb169-106">Prerequisites</span></span>

[!INCLUDE [Tutorial prerequisites](../includes/power-automate-tutorial-prerequisites.md)]

> [!IMPORTANT]
> <span data-ttu-id="cb169-107">Este tutorial pressupõe que você completou a[execução de Scripts do Office no Excel na Web com o tutorial do Power Automate](excel-power-automate-manual.md).</span><span class="sxs-lookup"><span data-stu-id="cb169-107">This tutorial assumes you have completed the [Run Office Scripts in Excel on the web with Power Automate](excel-power-automate-manual.md) tutorial.</span></span>

## <a name="prepare-the-workbook"></a><span data-ttu-id="cb169-108">Preparar a pasta de trabalho</span><span class="sxs-lookup"><span data-stu-id="cb169-108">Prepare the workbook</span></span>

<span data-ttu-id="cb169-109">O Power Automate não pode usar[referências relativas](../develop/power-automate-integration.md#avoid-using-relative-references)como`Workbook.getActiveWorksheet`acessar componentes da pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="cb169-109">Power Automate can't use [relative references](../develop/power-automate-integration.md#avoid-using-relative-references) like `Workbook.getActiveWorksheet` to access workbook components.</span></span> <span data-ttu-id="cb169-110">Portanto, precisamos de uma pasta de trabalho e planilha com nomes consistentes para que o Power Automate possa consultar.</span><span class="sxs-lookup"><span data-stu-id="cb169-110">So, we need a workbook and worksheet with consistent names for Power Automate to reference.</span></span>

1. <span data-ttu-id="cb169-111">Criar um nome para a pasta de trabalho**MyWorkbook**.</span><span class="sxs-lookup"><span data-stu-id="cb169-111">Create a new workbook named **MyWorkbook**.</span></span>

2. <span data-ttu-id="cb169-112">Vá para a guia **Automatizar**e selecione**Editor de Códigos**.</span><span class="sxs-lookup"><span data-stu-id="cb169-112">Go to the **Automate** tab and select **Code Editor**.</span></span>

3. <span data-ttu-id="cb169-113">Selecione**Novo script**.</span><span class="sxs-lookup"><span data-stu-id="cb169-113">Select **New Script**.</span></span>

4. <span data-ttu-id="cb169-114">Substitua o código existente pelo seguinte script e pressione**Executar**.</span><span class="sxs-lookup"><span data-stu-id="cb169-114">Replace the existing code with the following script and press **Run**.</span></span> <span data-ttu-id="cb169-115">Isso instalará a pasta de trabalho com nomes consistentes de planilhas, tabela e tabela dinâmica.</span><span class="sxs-lookup"><span data-stu-id="cb169-115">This will setup the workbook with consistent worksheet, table, and PivotTable names.</span></span>

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
      // Add a new worksheet to store our email table
      let emailsSheet = workbook.addWorksheet("Emails");

      // Add data and create a table
      emailsSheet.getRange("A1:D1").setValues([
        ["Date", "Day of the week", "Email address", "Subject"]
      ]);
      let newTable = workbook.addTable(emailsSheet.getRange("A1:D2"), true);
      newTable.setName("EmailTable");

      // Add a new PivotTable to a new worksheet
      let pivotWorksheet = workbook.addWorksheet("SubjectPivot");
      let newPivotTable = workbook.addPivotTable("Pivot", "EmailTable", pivotWorksheet.getRange("A3:C20"));

      // Setup the pivot hierarchies
      newPivotTable.addRowHierarchy(newPivotTable.getHierarchy("Day of the week"));
      newPivotTable.addRowHierarchy(newPivotTable.getHierarchy("Email address"));
      newPivotTable.addDataHierarchy(newPivotTable.getHierarchy("Subject"));
    }
    ```

## <a name="create-an-office-script-for-your-automated-workflow"></a><span data-ttu-id="cb169-116">Criar um script do Office para o seu fluxo de trabalho automatizado.</span><span class="sxs-lookup"><span data-stu-id="cb169-116">Create an Office Script for your automated workflow</span></span>

<span data-ttu-id="cb169-117">Vamos criar um script que registre as informações de um email.</span><span class="sxs-lookup"><span data-stu-id="cb169-117">Let's create a script that logs information from an email.</span></span> <span data-ttu-id="cb169-118">Gostaríamos de saber como quais dias da semana recebemos a maioria dos emails e quantos remetentes exclusivos estão enviando esses emails.</span><span class="sxs-lookup"><span data-stu-id="cb169-118">We want to know how which days of the week we receive the most mail and how many unique senders are sending that mail.</span></span> <span data-ttu-id="cb169-119">Nossa pasta de trabalho tem uma tabela com **Data**, **Dia da semana**, **Endereços de email** e**Colunas de assunto**.</span><span class="sxs-lookup"><span data-stu-id="cb169-119">Our workbook has a table with **Date**, **Day of the week**, **Email address**, and **Subject** columns.</span></span> <span data-ttu-id="cb169-120">Nossa planilha também tem uma tabela dinâmica que está sendo dinamizada no **Dia da semana**e**Endereços de email**(essas são as hierarquias de linha).</span><span class="sxs-lookup"><span data-stu-id="cb169-120">Our worksheet also has a PivotTable that is pivoting on the **Day of the week** and **Email address** (those are the row hierarchies).</span></span> <span data-ttu-id="cb169-121">A contagem de **assuntos exclusivos** são as informações agregadas que estão sendo exibidas (a hierarquia de dados).</span><span class="sxs-lookup"><span data-stu-id="cb169-121">The count of unique **Subjects** is the aggregated information being displayed (the data hierarchy).</span></span> <span data-ttu-id="cb169-122">Faremos com que o nosso script atualize essa tabela dinâmica depois de atualizar a tabela de email.</span><span class="sxs-lookup"><span data-stu-id="cb169-122">We'll have our script refresh that PivotTable after updating the email table.</span></span>

1. <span data-ttu-id="cb169-123">Do **Editor de Código**, selecione **Novo Script**.</span><span class="sxs-lookup"><span data-stu-id="cb169-123">From within the **Code Editor**, select **New Script**.</span></span>

2. <span data-ttu-id="cb169-124">O fluxo que criaremos depois no tutorial enviará a informação do nosso script sobre cada email recebido.</span><span class="sxs-lookup"><span data-stu-id="cb169-124">The flow that we'll create later in the tutorial will send our script information about each email that's received.</span></span> <span data-ttu-id="cb169-125">O script precisa aceitar essa entrada pelos parâmetros na `main`função.</span><span class="sxs-lookup"><span data-stu-id="cb169-125">The script needs to accept that input through parameters in the `main` function.</span></span> <span data-ttu-id="cb169-126">Substitua o script padrão com o script seguinte:</span><span class="sxs-lookup"><span data-stu-id="cb169-126">Replace the default script with the following script:</span></span>

    ```TypeScript
    function main(
      workbook: ExcelScript.Workbook,
      from: string,
      dateReceived: string,
      subject: string) {

    }
    ```

3. <span data-ttu-id="cb169-127">O script precisa acessar a tabela e a tabela dinâmica da pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="cb169-127">The script needs access to the workbook's table and PivotTable.</span></span> <span data-ttu-id="cb169-128">Adicione o seguinte código ao corpo do script após a abertura`{`:</span><span class="sxs-lookup"><span data-stu-id="cb169-128">Add the following code to the body of the script, after the opening `{`:</span></span>

    ```TypeScript
    // Get the email table.
    let emailWorksheet = workbook.getWorksheet("Emails");
    let table = emailWorksheet.getTable("EmailTable");
  
    // Get the PivotTable.
    let pivotTableWorksheet = workbook.getWorksheet("SubjectPivot");
    let pivotTable = pivotTableWorksheet.getPivotTable("Pivot");
    ```

4. <span data-ttu-id="cb169-129">O `dateReceived`parâmetro é do tipo`string`.</span><span class="sxs-lookup"><span data-stu-id="cb169-129">The `dateReceived` parameter is of type `string`.</span></span> <span data-ttu-id="cb169-130">Vamos convertê-la em um[`Date`objeto](../develop/javascript-objects.md#date)para que possamos facilmente obter o dia da semana.</span><span class="sxs-lookup"><span data-stu-id="cb169-130">Let's convert that to a [`Date` object](../develop/javascript-objects.md#date) so we can easily get the day of the week.</span></span> <span data-ttu-id="cb169-131">Depois de fazer isso, será necessário mapear o valor numérico do dia para uma versão mais legível.</span><span class="sxs-lookup"><span data-stu-id="cb169-131">After doing that, we'll need to map the day's number value to a more readable version.</span></span> <span data-ttu-id="cb169-132">Adicione o seguinte código no final do script (antes do encerramento `}`):</span><span class="sxs-lookup"><span data-stu-id="cb169-132">Add the following code to the end of your script, before the closing `}`:</span></span>

    ```TypeScript
    // Parse the received date string.
    let date = new Date(dateReceived);

    // Convert number representing the day of the week into the name of the day.
    let dayText : string;
    switch (date.getDay()) {
      case 0:
        dayText = "Sunday";
        break;
      case 1:
        dayText = "Monday";
        break;
      case 2:
        dayText = "Tuesday";
        break;
      case 3:
        dayText = "Wednesday";
        break;
      case 4:
        dayText = "Thursday";
        break;
      case 5:
        dayText = "Friday";
        break;
      default:
        dayText = "Saturday";
        break;
    }
    ```

5. <span data-ttu-id="cb169-133">A cadeia`subject` pode incluir a marca de resposta "RE:".</span><span class="sxs-lookup"><span data-stu-id="cb169-133">The `subject` string may include the "RE:" reply tag.</span></span> <span data-ttu-id="cb169-134">Vamos remover isso da cadeia de caracteres para que os emails no mesmo thread tenham o mesmo assunto para a tabela.</span><span class="sxs-lookup"><span data-stu-id="cb169-134">Let's remove that from the string so that emails in the same thread have the same subject for the table.</span></span> <span data-ttu-id="cb169-135">Adicione o seguinte código no final do script (antes do encerramento `}`):</span><span class="sxs-lookup"><span data-stu-id="cb169-135">Add the following code to the end of your script, before the closing `}`:</span></span>

    ```TypeScript
    // Remove the reply tag from the email subject to group emails on the same thread.
    let subjectText = subject.replace("Re: ", "");
    subjectText = subjectText.replace("RE: ", "");
    ```

6. <span data-ttu-id="cb169-136">Agora que os dados de email foram formatados da nossa preferência, vamos adicionar uma linha à tabela de email.</span><span class="sxs-lookup"><span data-stu-id="cb169-136">Now that the email data has been formatted to our liking, let's add a row to the email table.</span></span> <span data-ttu-id="cb169-137">Adicione o seguinte código no final do script (antes do encerramento `}`):</span><span class="sxs-lookup"><span data-stu-id="cb169-137">Add the following code to the end of your script, before the closing `}`:</span></span>

    ```TypeScript
    // Add the parsed text to the table.
    table.addRow(-1, [dateReceived, dayText, from, subjectText]);
    ```

7. <span data-ttu-id="cb169-138">Por fim, vamos verificar se a tabela dinâmica está atualizada.</span><span class="sxs-lookup"><span data-stu-id="cb169-138">Finally, let's make sure the PivotTable is refreshed.</span></span> <span data-ttu-id="cb169-139">Adicione o seguinte código no final do script (antes do encerramento `}`):</span><span class="sxs-lookup"><span data-stu-id="cb169-139">Add the following code to the end of your script, before the closing `}`:</span></span>

    ```TypeScript
    // Refresh the PivotTable to include the new row.
    pivotTable.refresh();
    ```

8. <span data-ttu-id="cb169-140">Renomeie seu script **Registre o email**e pressione **Salvar script**.</span><span class="sxs-lookup"><span data-stu-id="cb169-140">Rename your script **Record Email** and press **Save script**.</span></span>

<span data-ttu-id="cb169-141">O seu script já está pronto para um fluxo de trabalho automatizado.</span><span class="sxs-lookup"><span data-stu-id="cb169-141">Your script is now ready for a Power Automate workflow.</span></span> <span data-ttu-id="cb169-142">Ele deverá ser semelhante ao script a seguir:</span><span class="sxs-lookup"><span data-stu-id="cb169-142">It should look like the following script:</span></span>

```TypeScript
function main(
  workbook: ExcelScript.Workbook,
  from: string,
  dateReceived: string,
  subject: string) {
  // Get the email table.
  let emailWorksheet = workbook.getWorksheet("Emails");
  let table = emailWorksheet.getTable("EmailTable");

  // Get the PivotTable.
  let pivotTableWorksheet = workbook.getWorksheet("Pivot");
  let pivotTable = pivotTableWorksheet.getPivotTable("SubjectPivot");

  // Parse the received date string.
  let date = new Date(dateReceived);

  // Convert number representing the day of the week into the name of the day.
  let dayText: string;
  switch (date.getDay()) {
    case 0:
      dayText = "Sunday";
      break;
    case 1:
      dayText = "Monday";
      break;
    case 2:
      dayText = "Tuesday";
      break;
    case 3:
      dayText = "Wednesday";
      break;
    case 4:
      dayText = "Thursday";
      break;
    case 5:
      dayText = "Friday";
      break;
    default:
      dayText = "Saturday";
      break;
  }

  // Remove the reply tag from the email subject to group emails on the same thread.
  let subjectText = subject.replace("Re: ", "");
  subjectText = subjectText.replace("RE: ", "");

  // Add the parsed text to the table.
  table.addRow(-1, [dateReceived, dayText, from, subjectText]);

  // Refresh the PivotTable to include the new row.
  pivotTable.refresh();
}
```

## <a name="create-an-automated-workflow-with-power-automate"></a><span data-ttu-id="cb169-143">Criar um fluxo de trabalho automatizado com o Power Automate</span><span class="sxs-lookup"><span data-stu-id="cb169-143">Create an automated workflow with Power Automate</span></span>

1. <span data-ttu-id="cb169-144">Entre no [site do Power Automate](https://flow.microsoft.com).</span><span class="sxs-lookup"><span data-stu-id="cb169-144">Sign in to the [Power Automate site](https://flow.microsoft.com).</span></span>

2. <span data-ttu-id="cb169-145">No menu exibido do lado esquerdo da tela, pressione **Criar**.</span><span class="sxs-lookup"><span data-stu-id="cb169-145">In the menu that's displayed on the left side of the screen, press **Create**.</span></span> <span data-ttu-id="cb169-146">Isso o conduzirá a uma lista de maneiras de criar novos fluxos de trabalho.</span><span class="sxs-lookup"><span data-stu-id="cb169-146">This brings you to list of ways to create new workflows.</span></span>

    ![O botão Criar no Power Automate.](../images/power-automate-tutorial-1.png)

3. <span data-ttu-id="cb169-148">Na seção **Começar no espaço em branco**, selecione **Fluxo automático**.</span><span class="sxs-lookup"><span data-stu-id="cb169-148">In the **Start from blank** section, select **Automated flow**.</span></span> <span data-ttu-id="cb169-149">Isso cria um fluxo de trabalho iniciado por um evento, como o recebimento de emails.</span><span class="sxs-lookup"><span data-stu-id="cb169-149">This creates a workflow triggered by an event, such as receiving an email.</span></span>

    ![A opção de fluxo automatizado em Power Automate.](../images/power-automate-params-tutorial-1.png)

4. <span data-ttu-id="cb169-151">Na caixa de diálogo exibida, insira o nome para seu fluxo na **caixa de texto**Nome de Fluxo.</span><span class="sxs-lookup"><span data-stu-id="cb169-151">In the dialog window that appears, enter a name for your flow in the **Flow name** text box.</span></span> <span data-ttu-id="cb169-152">Em seguida, selecione**Quando um novo email chegar**da lista de opções em **escolha o gatilho de fluxo**.</span><span class="sxs-lookup"><span data-stu-id="cb169-152">Then select **When a new email arrives** from the list of options under **Choose your flow's trigger**.</span></span> <span data-ttu-id="cb169-153">Talvez seja necessário procurar pela opção usando a caixa de pesquisa.</span><span class="sxs-lookup"><span data-stu-id="cb169-153">You may need to search for the option using the search box.</span></span> <span data-ttu-id="cb169-154">Por fim, pressione **criar**.</span><span class="sxs-lookup"><span data-stu-id="cb169-154">Finally, press **Create**.</span></span>

    ![Parte da janela Criar Uma Fluxo Automatizado no Power Automate, que mostra a opção "novo email recebido".](../images/power-automate-params-tutorial-2.png)

    > [!NOTE]
    > <span data-ttu-id="cb169-156">Este tutorial usa o Outlook.</span><span class="sxs-lookup"><span data-stu-id="cb169-156">This tutorial uses Outlook.</span></span> <span data-ttu-id="cb169-157">Sinta-se à vontade para usar o seu serviço de email preferido, embora algumas opções possam ser diferentes.</span><span class="sxs-lookup"><span data-stu-id="cb169-157">Feel free to use your preferred email service instead, though some options may be different.</span></span>

5. <span data-ttu-id="cb169-158">Pressione **Nova etapa**.</span><span class="sxs-lookup"><span data-stu-id="cb169-158">Press **New step**.</span></span>

6. <span data-ttu-id="cb169-159">Selecione a guia **Padrão** e, em seguida, selecione **Excel Online (Business)**.</span><span class="sxs-lookup"><span data-stu-id="cb169-159">Select the **Standard** tab, then select **Excel Online (Business)**.</span></span>

    ![A opção do Power Automate para Excel Online (Business)](../images/power-automate-tutorial-4.png)

7. <span data-ttu-id="cb169-161">Em **Ações**, selecione **executar script (visualização)**.</span><span class="sxs-lookup"><span data-stu-id="cb169-161">Under **Actions**, select **Run script (preview)**.</span></span>

    ![Opção de ação do Power Automate para Executar script (visualização).](../images/power-automate-tutorial-5.png)

8. <span data-ttu-id="cb169-163">Especifique as seguintes configurações para o conector **Executar Script**:</span><span class="sxs-lookup"><span data-stu-id="cb169-163">Specify the following settings for the **Run script** connector:</span></span>

    - <span data-ttu-id="cb169-164">**Localização**: OneDrive for Business</span><span class="sxs-lookup"><span data-stu-id="cb169-164">**Location**: OneDrive for Business</span></span>
    - <span data-ttu-id="cb169-165">**Biblioteca de Documentos**: OneDrive</span><span class="sxs-lookup"><span data-stu-id="cb169-165">**Document Library**: OneDrive</span></span>
    - <span data-ttu-id="cb169-166">**Arquivo**: MyWorkbook.xlsx</span><span class="sxs-lookup"><span data-stu-id="cb169-166">**File**: MyWorkbook.xlsx</span></span>
    - <span data-ttu-id="cb169-167">**De**: Gravar Email</span><span class="sxs-lookup"><span data-stu-id="cb169-167">**Script**: Record Email</span></span>
    - <span data-ttu-id="cb169-168">**De**: De *(conteúdo dinâmico do Outlook)*</span><span class="sxs-lookup"><span data-stu-id="cb169-168">**from**: From *(dynamic content from Outlook)*</span></span>
    - <span data-ttu-id="cb169-169">**DateReceived**: Hora Recebida *(conteúdo dinâmico do Outlook)*</span><span class="sxs-lookup"><span data-stu-id="cb169-169">**dateReceived**: Received Time *(dynamic content from Outlook)*</span></span>
    - <span data-ttu-id="cb169-170">**assunto**: Assunto *(conteúdo dinâmico do Outlook)*</span><span class="sxs-lookup"><span data-stu-id="cb169-170">**subject**: Subject *(dynamic content from Outlook)*</span></span>

    <span data-ttu-id="cb169-171">*Observe que os parâmetros para o script só aparecerão quando o script for selecionado.*</span><span class="sxs-lookup"><span data-stu-id="cb169-171">*Note that the parameters for the script will only appear once the script is selected.*</span></span>

    ![Opção de ação do Power Automate para Executar script (visualização).](../images/power-automate-params-tutorial-3.png)

9. <span data-ttu-id="cb169-173">Pressione **Salvar**.</span><span class="sxs-lookup"><span data-stu-id="cb169-173">Press **Save**.</span></span>

<span data-ttu-id="cb169-174">Agora, o fluxo está habilitado.</span><span class="sxs-lookup"><span data-stu-id="cb169-174">Your flow is now enabled.</span></span> <span data-ttu-id="cb169-175">O seu script será automaticamente executado sempre que você receber um email por meio do Outlook.</span><span class="sxs-lookup"><span data-stu-id="cb169-175">It will automatically run your script each time you receive an email through Outlook.</span></span>

## <a name="manage-the-script-in-power-automate"></a><span data-ttu-id="cb169-176">Gerenciar o script no Power Automate</span><span class="sxs-lookup"><span data-stu-id="cb169-176">Manage the script in Power Automate</span></span>

1. <span data-ttu-id="cb169-177">Na página principal do Power Automate, selecione **Meus fluxos**.</span><span class="sxs-lookup"><span data-stu-id="cb169-177">From the main Power Automate page, select **My flows**.</span></span>

    ![Botão Meus fluxos no Power Automate.](../images/power-automate-tutorial-7.png)

2. <span data-ttu-id="cb169-179">Selecione o seu fluxo.</span><span class="sxs-lookup"><span data-stu-id="cb169-179">Select your flow.</span></span> <span data-ttu-id="cb169-180">Aqui você pode ver o histórico de execução.</span><span class="sxs-lookup"><span data-stu-id="cb169-180">Here you can see the run history.</span></span> <span data-ttu-id="cb169-181">Você pode atualizar a página ou pressionar o botão atualizar **Executar Todos** para atualizar o histórico.</span><span class="sxs-lookup"><span data-stu-id="cb169-181">You can refresh the page or press the refresh **All runs** button to update the history.</span></span> <span data-ttu-id="cb169-182">O fluxo será disparado logo após o recebimento de um email.</span><span class="sxs-lookup"><span data-stu-id="cb169-182">The flow will trigger shortly after an email is received.</span></span> <span data-ttu-id="cb169-183">Testar o fluxo enviando a si mesmo um email.</span><span class="sxs-lookup"><span data-stu-id="cb169-183">Test the flow by sending yourself mail.</span></span>

<span data-ttu-id="cb169-184">Quando o fluxo é acionado e executa seu script com sucesso, você deverá ver as atualizações da planilha na pasta de trabalho e da tabela dinâmica.</span><span class="sxs-lookup"><span data-stu-id="cb169-184">When the flow is triggered and successfully runs your script, you should see the workbook's table and PivotTable update.</span></span>

![A tabela de email após o fluxo ter sido executado algumas vezes.](../images/power-automate-params-tutorial-4.png)

![A tabela dinâmica após o fluxo ter sido executado algumas vezes.](../images/power-automate-params-tutorial-5.png)

## <a name="next-steps"></a><span data-ttu-id="cb169-187">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="cb169-187">Next steps</span></span>

<span data-ttu-id="cb169-188">Visite [executar os Scripts do Office com o Power Automate](../develop/power-automate-integration.md) para saber mais sobre como conectar Scripts do Office com o Power Automate.</span><span class="sxs-lookup"><span data-stu-id="cb169-188">Visit [Run Office Scripts with Power Automate](../develop/power-automate-integration.md) to learn more about connecting Office Scripts with Power Automate.</span></span>

<span data-ttu-id="cb169-189">Você também pode conferir o exemplo de [lembretes automáticos de tarefas](../resources/scenarios/task-reminders.md) para aprender a combinar os Scripts do Office e Power Automate com as placas adaptáveis de equipes.</span><span class="sxs-lookup"><span data-stu-id="cb169-189">You can also check out the [Automated task reminders sample scenario](../resources/scenarios/task-reminders.md) to learn how to combine Office Scripts and Power Automate with Teams Adaptive Cards.</span></span>
