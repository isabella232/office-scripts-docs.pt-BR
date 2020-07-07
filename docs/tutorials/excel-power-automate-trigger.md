---
title: Executar automaticamente scripts com Power Automate
description: Tutorial sobre como executar scripts do Office para Excel na Web através da automatização de energia usando um gatilho externo automático (recebendo emails através do Outlook).
ms.date: 06/29/2020
localization_priority: Priority
ms.openlocfilehash: a750197d6b5ae770ad7d2e17b3ee00dc65ee8875
ms.sourcegitcommit: bf9f33c37c6f7805d6b408aa648bb9785a7cd133
ms.contentlocale: pt-BR
ms.lasthandoff: 07/06/2020
ms.locfileid: "45043416"
---
# <a name="automatically-run-scripts-with-power-automate-preview"></a><span data-ttu-id="e7aa1-103">Executar scripts automaticamente com a automatização de energia (prévia)</span><span class="sxs-lookup"><span data-stu-id="e7aa1-103">Automatically run scripts with Power Automate (preview)</span></span>

<span data-ttu-id="e7aa1-104">Este tutorial ensina como usar um script do Office para Excel na Web com [um fluxo de trabalho automatizado](https://flow.microsoft.com) automático.</span><span class="sxs-lookup"><span data-stu-id="e7aa1-104">This tutorial teaches you how to use an Office Script for Excel on the web with an automated [Power Automate](https://flow.microsoft.com) workflow.</span></span> <span data-ttu-id="e7aa1-105">O script será executado automaticamente cada vez que você receber um email, gravando informações do email em uma pasta de trabalho do Excel.</span><span class="sxs-lookup"><span data-stu-id="e7aa1-105">Your script will automatically run each time you receive an email, recording information from the email in an Excel workbook.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="e7aa1-106">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="e7aa1-106">Prerequisites</span></span>

[!INCLUDE [Tutorial prerequisites](../includes/tutorial-prerequisites.md)]

> [!IMPORTANT]
> <span data-ttu-id="e7aa1-107">Este tutorial pressupõe que você tenha concluído os [scripts executar do Office no Excel na Web com o tutorial automatizar de energia](excel-power-automate-manual.md) .</span><span class="sxs-lookup"><span data-stu-id="e7aa1-107">This tutorial assumes you have completed the [Run Office Scripts in Excel on the web with Power Automate](excel-power-automate-manual.md) tutorial.</span></span>

## <a name="prepare-the-workbook"></a><span data-ttu-id="e7aa1-108">Preparar a pasta de trabalho</span><span class="sxs-lookup"><span data-stu-id="e7aa1-108">Prepare the workbook</span></span>

<span data-ttu-id="e7aa1-109">A automatização de energia não pode usar [referências relativas](../develop/power-automate-integration.md#avoid-using-relative-references) como `Workbook.getActiveWorksheet` acessar componentes de pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="e7aa1-109">Power Automate can't use [relative references](../develop/power-automate-integration.md#avoid-using-relative-references) like `Workbook.getActiveWorksheet` to access workbook components.</span></span> <span data-ttu-id="e7aa1-110">Portanto, precisamos de uma pasta de trabalho e planilha com nomes consistentes para que a automatização de energia seja referenciada.</span><span class="sxs-lookup"><span data-stu-id="e7aa1-110">So, we need a workbook and worksheet with consistent names for Power Automate to reference.</span></span>

1. <span data-ttu-id="e7aa1-111">Crie uma nova pasta de trabalho chamada **myworkbook**.</span><span class="sxs-lookup"><span data-stu-id="e7aa1-111">Create a new workbook named **MyWorkbook**.</span></span>

2. <span data-ttu-id="e7aa1-112">Vá para a guia **automatizar** e selecione **Editor de código**.</span><span class="sxs-lookup"><span data-stu-id="e7aa1-112">Go to the **Automate** tab and select **Code Editor**.</span></span>

3. <span data-ttu-id="e7aa1-113">Selecione **novo script**.</span><span class="sxs-lookup"><span data-stu-id="e7aa1-113">Select **New Script**.</span></span>

4. <span data-ttu-id="e7aa1-114">Substitua o código existente pelo seguinte script e pressione **executar**.</span><span class="sxs-lookup"><span data-stu-id="e7aa1-114">Replace the existing code with the following script and press **Run**.</span></span> <span data-ttu-id="e7aa1-115">Isso instalará a pasta de trabalho com nomes consistentes de planilha, tabela e tabela dinâmica.</span><span class="sxs-lookup"><span data-stu-id="e7aa1-115">This will setup the workbook with consistent worksheet, table, and PivotTable names.</span></span>

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

## <a name="create-an-office-script-for-your-automated-workflow"></a><span data-ttu-id="e7aa1-116">Criar um script do Office para o fluxo de trabalho automatizado</span><span class="sxs-lookup"><span data-stu-id="e7aa1-116">Create an Office Script for your automated workflow</span></span>

<span data-ttu-id="e7aa1-117">Vamos criar um script que registre informações de um email.</span><span class="sxs-lookup"><span data-stu-id="e7aa1-117">Let's create a script that logs information from an email.</span></span> <span data-ttu-id="e7aa1-118">Queremos saber como quais dias da semana recebemos a maioria dos emails e quantos remetentes exclusivos estão enviando esse email.</span><span class="sxs-lookup"><span data-stu-id="e7aa1-118">We want to know how which days of the week we receive the most mail and how many unique senders are sending that mail.</span></span> <span data-ttu-id="e7aa1-119">Nossa pasta de trabalho tem uma tabela com **Data**, **dia da semana**, **endereço de email**e colunas de **assunto** .</span><span class="sxs-lookup"><span data-stu-id="e7aa1-119">Our workbook has a table with **Date**, **Day of the week**, **Email address**, and **Subject** columns.</span></span> <span data-ttu-id="e7aa1-120">Nossa planilha também tem uma tabela dinâmica que está sendo dinamizada no **dia da semana** e **endereço de email** (essas são as hierarquias de linha).</span><span class="sxs-lookup"><span data-stu-id="e7aa1-120">Our worksheet also has a PivotTable that is pivoting on the **Day of the week** and **Email address** (those are the row hierarchies).</span></span> <span data-ttu-id="e7aa1-121">A contagem de **assuntos** exclusivos é as informações agregadas que estão sendo exibidas (a hierarquia de dados).</span><span class="sxs-lookup"><span data-stu-id="e7aa1-121">The count of unique **Subjects** is the aggregated information being displayed (the data hierarchy).</span></span> <span data-ttu-id="e7aa1-122">Teremos o script atualizar essa tabela dinâmica depois de atualizar a tabela de email.</span><span class="sxs-lookup"><span data-stu-id="e7aa1-122">We'll have our script refresh that PivotTable after updating the email table.</span></span>

1. <span data-ttu-id="e7aa1-123">No editor de **códigos**, selecione **novo script**.</span><span class="sxs-lookup"><span data-stu-id="e7aa1-123">From within the **Code Editor**, select **New Script**.</span></span>

2. <span data-ttu-id="e7aa1-124">O fluxo que criaremos mais tarde no tutorial enviará as informações de script sobre cada email recebido.</span><span class="sxs-lookup"><span data-stu-id="e7aa1-124">The flow that we'll create later in the tutorial will send our script information about each email that's received.</span></span> <span data-ttu-id="e7aa1-125">O script precisa aceitar essa entrada através de parâmetros na `main` função.</span><span class="sxs-lookup"><span data-stu-id="e7aa1-125">The script needs to accept that input through parameters in the `main` function.</span></span> <span data-ttu-id="e7aa1-126">Substitua o script padrão pelo seguinte script:</span><span class="sxs-lookup"><span data-stu-id="e7aa1-126">Replace the default script with the following script:</span></span>

    ```TypeScript
    function main(
      workbook: ExcelScript.Workbook,
      from: string,
      dateReceived: string,
      subject: string) {

    }
    ```

3. <span data-ttu-id="e7aa1-127">O script precisa acessar a tabela e a tabela dinâmica da pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="e7aa1-127">The script needs access to the workbook's table and PivotTable.</span></span> <span data-ttu-id="e7aa1-128">Adicione o seguinte código ao corpo do script, após a abertura `{` :</span><span class="sxs-lookup"><span data-stu-id="e7aa1-128">Add the following code to the body of the script, after the opening `{`:</span></span>

    ```TypeScript
    // Get the email table.
    let emailWorksheet = workbook.getWorksheet("Emails");
    let table = emailWorksheet.getTable("EmailTable");
  
    // Get the PivotTable.
    let pivotTableWorksheet = workbook.getWorksheet("SubjectPivot");
    let pivotTable = pivotTableWorksheet.getPivotTable("Pivot");
    ```

4. <span data-ttu-id="e7aa1-129">O `dateReceived` parâmetro é do tipo `string` .</span><span class="sxs-lookup"><span data-stu-id="e7aa1-129">The `dateReceived` parameter is of type `string`.</span></span> <span data-ttu-id="e7aa1-130">Vamos convertê-lo em um [ `Date` objeto](../develop/javascript-objects.md#date) para que possamos obter facilmente o dia da semana.</span><span class="sxs-lookup"><span data-stu-id="e7aa1-130">Let's convert that to a [`Date` object](../develop/javascript-objects.md#date) so we can easily get the day of the week.</span></span> <span data-ttu-id="e7aa1-131">Depois disso, precisaremos mapear o valor do número do dia para uma versão mais legível.</span><span class="sxs-lookup"><span data-stu-id="e7aa1-131">After doing that, we'll need to map the day's number value to a more readable version.</span></span> <span data-ttu-id="e7aa1-132">Adicione o código a seguir ao final do seu script, antes de fechar `}` :</span><span class="sxs-lookup"><span data-stu-id="e7aa1-132">Add the following code to the end of your script, before the closing `}`:</span></span>

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

5. <span data-ttu-id="e7aa1-133">A `subject` cadeia de caracteres pode incluir a marca de resposta "Re:".</span><span class="sxs-lookup"><span data-stu-id="e7aa1-133">The `subject` string may include the "RE:" reply tag.</span></span> <span data-ttu-id="e7aa1-134">Vamos remover isso da cadeia de caracteres para que os emails no mesmo thread tenham o mesmo assunto para a tabela.</span><span class="sxs-lookup"><span data-stu-id="e7aa1-134">Let's remove that from the string so that emails in the same thread have the same subject for the table.</span></span> <span data-ttu-id="e7aa1-135">Adicione o código a seguir ao final do seu script, antes de fechar `}` :</span><span class="sxs-lookup"><span data-stu-id="e7aa1-135">Add the following code to the end of your script, before the closing `}`:</span></span>

    ```TypeScript
    // Remove the reply tag from the email subject to group emails on the same thread.
    let subjectText = subject.replace("Re: ", "");
    subjectText = subjectText.replace("RE: ", "");
    ```

6. <span data-ttu-id="e7aa1-136">Agora que os dados de email foram formatados para nossa preferência, vamos adicionar uma linha à tabela de email.</span><span class="sxs-lookup"><span data-stu-id="e7aa1-136">Now that the email data has been formatted to our liking, let's add a row to the email table.</span></span> <span data-ttu-id="e7aa1-137">Adicione o código a seguir ao final do seu script, antes de fechar `}` :</span><span class="sxs-lookup"><span data-stu-id="e7aa1-137">Add the following code to the end of your script, before the closing `}`:</span></span>

    ```TypeScript
    // Add the parsed text to the table.
    table.addRow(-1, [dateReceived, dayText, from, subjectText]);
    ```

7. <span data-ttu-id="e7aa1-138">Por fim, vamos garantir que a tabela dinâmica seja atualizada.</span><span class="sxs-lookup"><span data-stu-id="e7aa1-138">Finally, let's make sure the PivotTable is refreshed.</span></span> <span data-ttu-id="e7aa1-139">Adicione o código a seguir ao final do seu script, antes de fechar `}` :</span><span class="sxs-lookup"><span data-stu-id="e7aa1-139">Add the following code to the end of your script, before the closing `}`:</span></span>

    ```TypeScript
    // Refresh the PivotTable to include the new row.
    pivotTable.refresh();
    ```

8. <span data-ttu-id="e7aa1-140">Renomeie seu **email de registro** de script e pressione **Salvar script**.</span><span class="sxs-lookup"><span data-stu-id="e7aa1-140">Rename your script **Record Email** and press **Save script**.</span></span>

<span data-ttu-id="e7aa1-141">O script agora está pronto para um fluxo de trabalho automatizado de energia.</span><span class="sxs-lookup"><span data-stu-id="e7aa1-141">Your script is now ready for a Power Automate workflow.</span></span> <span data-ttu-id="e7aa1-142">Ele deve ser semelhante ao seguinte script:</span><span class="sxs-lookup"><span data-stu-id="e7aa1-142">It should look like the following script:</span></span>

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

## <a name="create-an-automated-workflow-with-power-automate"></a><span data-ttu-id="e7aa1-143">Criar um fluxo de trabalho automatizado com a automatização de energia</span><span class="sxs-lookup"><span data-stu-id="e7aa1-143">Create an automated workflow with Power Automate</span></span>

1. <span data-ttu-id="e7aa1-144">Entre no site de [visualização de energia automatizada](https://flow.microsoft.com).</span><span class="sxs-lookup"><span data-stu-id="e7aa1-144">Sign in to the [Power Automate preview site](https://flow.microsoft.com).</span></span>

2. <span data-ttu-id="e7aa1-145">No menu que é exibido no lado esquerdo da tela, pressione **criar**.</span><span class="sxs-lookup"><span data-stu-id="e7aa1-145">In the menu that's displayed on the left side of the screen, press **Create**.</span></span> <span data-ttu-id="e7aa1-146">Isso lhe permite listar maneiras de criar novos fluxos de trabalho.</span><span class="sxs-lookup"><span data-stu-id="e7aa1-146">This brings you to list of ways to create new workflows.</span></span>

    ![O botão criar na automatização de energia.](../images/power-automate-tutorial-1.png)

3. <span data-ttu-id="e7aa1-148">Na seção **Iniciar com base em branco** , selecione **fluxo automatizado**.</span><span class="sxs-lookup"><span data-stu-id="e7aa1-148">In the **Start from blank** section, select **Automated flow**.</span></span> <span data-ttu-id="e7aa1-149">Isso cria um fluxo de trabalho disparado por um evento, como receber um email.</span><span class="sxs-lookup"><span data-stu-id="e7aa1-149">This creates a workflow triggered by an event, such as receiving an email.</span></span>

    ![A opção de fluxo automatizado em energia automatizada.](../images/power-automate-params-tutorial-1.png)

4. <span data-ttu-id="e7aa1-151">Na janela de diálogo exibida, insira um nome para o fluxo na caixa de texto **nome do fluxo** .</span><span class="sxs-lookup"><span data-stu-id="e7aa1-151">In the dialog window that appears, enter a name for your flow in the **Flow name** text box.</span></span> <span data-ttu-id="e7aa1-152">Em seguida, selecione **quando um novo email chegar** da lista de opções em **escolha o disparador do fluxo**.</span><span class="sxs-lookup"><span data-stu-id="e7aa1-152">Then select **When a new email arrives** from the list of options under **Choose your flow's trigger**.</span></span> <span data-ttu-id="e7aa1-153">Talvez seja necessário pesquisar a opção usando a caixa de pesquisa.</span><span class="sxs-lookup"><span data-stu-id="e7aa1-153">You may need to search for the option using the search box.</span></span> <span data-ttu-id="e7aa1-154">Por fim, pressione **criar**.</span><span class="sxs-lookup"><span data-stu-id="e7aa1-154">Finally, press **Create**.</span></span>

    ![Parte da janela criar um fluxo automatizado em automatização de energia que mostra a opção "novo email recebido".](../images/power-automate-params-tutorial-2.png)

    > [!NOTE]
    > <span data-ttu-id="e7aa1-156">Este tutorial usa o Outlook.</span><span class="sxs-lookup"><span data-stu-id="e7aa1-156">This tutorial uses Outlook.</span></span> <span data-ttu-id="e7aa1-157">Sinta-se livre para usar seu serviço de email preferencial, embora algumas opções possam ser diferentes.</span><span class="sxs-lookup"><span data-stu-id="e7aa1-157">Feel free to use your preferred email service instead, though some options may be different.</span></span>

5. <span data-ttu-id="e7aa1-158">Pressione **nova etapa**.</span><span class="sxs-lookup"><span data-stu-id="e7aa1-158">Press **New step**.</span></span>

6. <span data-ttu-id="e7aa1-159">Selecione a guia **padrão** e, em seguida, selecione **Excel online (comercial)**.</span><span class="sxs-lookup"><span data-stu-id="e7aa1-159">Select the **Standard** tab, then select **Excel Online (Business)**.</span></span>

    ![A opção de automatização de energia para o Excel online (Business).](../images/power-automate-tutorial-4.png)

7. <span data-ttu-id="e7aa1-161">Em **ações**, selecione **Executar script (versão prévia)**.</span><span class="sxs-lookup"><span data-stu-id="e7aa1-161">Under **Actions**, select **Run script (preview)**.</span></span>

    ![A opção de ação automatizar a energia para executar script (visualização).](../images/power-automate-tutorial-5.png)

8. <span data-ttu-id="e7aa1-163">Especifique as seguintes configurações para executar o conector de **script** :</span><span class="sxs-lookup"><span data-stu-id="e7aa1-163">Specify the following settings for the **Run script** connector:</span></span>

    - <span data-ttu-id="e7aa1-164">**Local**: onedrive for Business</span><span class="sxs-lookup"><span data-stu-id="e7aa1-164">**Location**: OneDrive for Business</span></span>
    - <span data-ttu-id="e7aa1-165">**Biblioteca de documentos**: onedrive</span><span class="sxs-lookup"><span data-stu-id="e7aa1-165">**Document Library**: OneDrive</span></span>
    - <span data-ttu-id="e7aa1-166">**Arquivo**: MyWorkbook.xlsx</span><span class="sxs-lookup"><span data-stu-id="e7aa1-166">**File**: MyWorkbook.xlsx</span></span>
    - <span data-ttu-id="e7aa1-167">**Script**: gravar email</span><span class="sxs-lookup"><span data-stu-id="e7aa1-167">**Script**: Record Email</span></span>
    - <span data-ttu-id="e7aa1-168">**de**: from *(conteúdo dinâmico do Outlook)*</span><span class="sxs-lookup"><span data-stu-id="e7aa1-168">**from**: From *(dynamic content from Outlook)*</span></span>
    - <span data-ttu-id="e7aa1-169">**dateReceived**: tempo *de recebimento (conteúdo dinâmico do Outlook)*</span><span class="sxs-lookup"><span data-stu-id="e7aa1-169">**dateReceived**: Received Time *(dynamic content from Outlook)*</span></span>
    - <span data-ttu-id="e7aa1-170">**assunto**: assunto *(conteúdo dinâmico do Outlook)*</span><span class="sxs-lookup"><span data-stu-id="e7aa1-170">**subject**: Subject *(dynamic content from Outlook)*</span></span>

    <span data-ttu-id="e7aa1-171">*Observe que os parâmetros para o script só aparecerão depois que o script for selecionado.*</span><span class="sxs-lookup"><span data-stu-id="e7aa1-171">*Note that the parameters for the script will only appear once the script is selected.*</span></span>

    ![A opção de ação automatizar a energia para executar script (visualização).](../images/power-automate-params-tutorial-3.png)

9. <span data-ttu-id="e7aa1-173">Pressione **salvar**.</span><span class="sxs-lookup"><span data-stu-id="e7aa1-173">Press **Save**.</span></span>

<span data-ttu-id="e7aa1-174">Agora, o fluxo está habilitado.</span><span class="sxs-lookup"><span data-stu-id="e7aa1-174">Your flow is now enabled.</span></span> <span data-ttu-id="e7aa1-175">O script será executado automaticamente sempre que você receber um email por meio do Outlook.</span><span class="sxs-lookup"><span data-stu-id="e7aa1-175">It will automatically run your script each time you receive an email through Outlook.</span></span>

## <a name="manage-the-script-in-power-automate"></a><span data-ttu-id="e7aa1-176">Gerenciar o script em automatização de energia</span><span class="sxs-lookup"><span data-stu-id="e7aa1-176">Manage the script in Power Automate</span></span>

1. <span data-ttu-id="e7aa1-177">Na página automatizar alimentação principal, selecione **meus fluxos**.</span><span class="sxs-lookup"><span data-stu-id="e7aa1-177">From the main Power Automate page, select **My flows**.</span></span>

    ![O botão meus fluxos em automatização de energia.](../images/power-automate-tutorial-7.png)

2. <span data-ttu-id="e7aa1-179">Selecione seu fluxo.</span><span class="sxs-lookup"><span data-stu-id="e7aa1-179">Select your flow.</span></span> <span data-ttu-id="e7aa1-180">Aqui você pode ver o histórico de execução.</span><span class="sxs-lookup"><span data-stu-id="e7aa1-180">Here you can see the run history.</span></span> <span data-ttu-id="e7aa1-181">Você pode atualizar a página ou pressionar o botão atualizar **tudo em execução** para atualizar o histórico.</span><span class="sxs-lookup"><span data-stu-id="e7aa1-181">You can refresh the page or press the refresh **All runs** button to update the history.</span></span> <span data-ttu-id="e7aa1-182">O fluxo será disparado logo após o recebimento de um email.</span><span class="sxs-lookup"><span data-stu-id="e7aa1-182">The flow will trigger shortly after an email is received.</span></span> <span data-ttu-id="e7aa1-183">Teste o fluxo enviando emails por conta própria.</span><span class="sxs-lookup"><span data-stu-id="e7aa1-183">Test the flow by sending yourself mail.</span></span>

<span data-ttu-id="e7aa1-184">Quando o fluxo é acionado e executa o script com êxito, você deve ver a tabela e a atualização da tabela dinâmica da pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="e7aa1-184">When the flow is triggered and successfully runs your script, you should see the workbook's table and PivotTable update.</span></span>

![A tabela de email depois que o fluxo é executado algumas vezes.](../images/power-automate-params-tutorial-4.png)

![A tabela dinâmica após o fluxo ter sido executada algumas vezes.](../images/power-automate-params-tutorial-5.png)

## <a name="next-steps"></a><span data-ttu-id="e7aa1-187">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="e7aa1-187">Next steps</span></span>

<span data-ttu-id="e7aa1-188">Visite [executar scripts do Office com a automatização de energia](../develop/power-automate-integration.md) para saber mais sobre como conectar scripts do Office com automatização de energia.</span><span class="sxs-lookup"><span data-stu-id="e7aa1-188">Visit [Run Office Scripts with Power Automate](../develop/power-automate-integration.md) to learn more about connecting Office Scripts with Power Automate.</span></span>

<span data-ttu-id="e7aa1-189">Você também pode conferir o [cenário de exemplo de lembretes automáticos de tarefas](../resources/scenarios/task-reminders.md) para saber como combinar scripts do Office e automatizar a automação com cartões adaptáveis do teams.</span><span class="sxs-lookup"><span data-stu-id="e7aa1-189">You can also check out the [Automated task reminders sample scenario](../resources/scenarios/task-reminders.md) to learn how to combine Office Scripts and Power Automate with Teams Adaptive Cards.</span></span>
