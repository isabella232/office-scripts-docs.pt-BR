---
title: Começar a usar scripts com Power Automate
description: Um tutorial sobre como usar scripts do Office em energia automatizada através de um gatilho manual.
ms.date: 07/01/2020
localization_priority: Priority
ms.openlocfilehash: 83e072a45fc724ff2aac5bf5f3893dcb64eaf2ff
ms.sourcegitcommit: edf58aed3cd38f57e5e7227465a1ef5515e15703
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/08/2020
ms.locfileid: "45081632"
---
# <a name="start-using-scripts-with-power-automate-preview"></a>Começar a usar scripts com a automatização de energia (visualização)

Este tutorial ensina a executar um script do Office para Excel na Web através da [automatização de energia](https://flow.microsoft.com).

## <a name="prerequisites"></a>Pré-requisitos

[!INCLUDE [Tutorial prerequisites](../includes/tutorial-prerequisites.md)]

> [!IMPORTANT]
> Este tutorial pressupõe que você tenha concluído o tutorial [gravar, editar e criar scripts do Office no Excel na Web](excel-tutorial.md) .

## <a name="prepare-the-workbook"></a>Preparar a pasta de trabalho

A automatização de energia não pode usar referências relativas como `Workbook.getActiveWorksheet` acessar componentes de pasta de trabalho. Portanto, precisamos de uma pasta de trabalho e planilha com nomes consistentes que os recursos de automatização podem fazer referência.

1. Crie uma nova pasta de trabalho chamada **myworkbook**.

2. Na pasta de trabalho **myworkbook** , crie uma planilha chamada **TutorialWorksheet**.

## <a name="create-an-office-script"></a>Criar um script do Office

1. Vá para a guia **automatizar** e selecione **Editor de código**.

2. Selecione **novo script**.

3. Substitua o script padrão pelo script a seguir. Esse script adiciona a data e hora atuais às duas primeiras células da planilha **TutorialWorksheet** .

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

4. Renomeie o script para **Definir data e hora**. Pressione o nome do script para alterá-lo.

5. Salve o script pressionando **Salvar script**.

## <a name="create-an-automated-workflow-with-power-automate"></a>Criar um fluxo de trabalho automatizado com a automatização de energia

1. Entre no site de [visualização de energia automatizada](https://flow.microsoft.com).

2. No menu que é exibido no lado esquerdo da tela, pressione **criar**. Isso lhe permite listar maneiras de criar novos fluxos de trabalho.

    ![O botão criar na automatização de energia.](../images/power-automate-tutorial-1.png)

3. Na seção **Iniciar com base em branco** , selecione **fluxo instantâneo**. Isso cria um fluxo de trabalho ativado manualmente.

    ![A opção de fluxo instantâneo para a criação de um novo fluxo de trabalho.](../images/power-automate-tutorial-2.png)

4. Na janela de diálogo exibida, insira um nome para o fluxo na caixa de **texto nome do fluxo** , selecione **acionar manualmente um fluxo** na lista de opções em **escolher como acionar o fluxo**e pressione **criar**.

    ![A opção de gatilho manual para criar um novo fluxo instantâneo.](../images/power-automate-tutorial-3.png)

5. Pressione **nova etapa**.

6. Selecione a guia **padrão** e, em seguida, selecione **Excel online (comercial)**.

    ![A opção de automatização de energia para o Excel online (Business).](../images/power-automate-tutorial-4.png)

7. Em **ações**, selecione **Executar script (versão prévia)**.

    ![A opção de ação automatizar a energia para executar script (visualização).](../images/power-automate-tutorial-5.png)

8. Especifique as seguintes configurações para executar o conector de **script** :

    - **Local**: onedrive for Business
    - **Biblioteca de documentos**: onedrive
    - **Arquivo**: MyWorkbook.xlsx
    - **Script**: Definir data e hora

    ![As configurações de conector para executar um script em automatização de energia.](../images/power-automate-tutorial-6.png)

9. Pressione **salvar**.

Agora, o fluxo está pronto para ser executado através da automatização de energia. Você pode testá-lo usando o botão **testar** no editor de fluxo ou siga as etapas restantes do tutorial para executar o fluxo de sua coleção de fluxo.

## <a name="run-the-script-through-power-automate"></a>Executar o script através da automatização de energia

1. Na página automatizar alimentação principal, selecione **meus fluxos**.

    ![O botão meus fluxos em automatização de energia.](../images/power-automate-tutorial-7.png)

2. Selecione **meu fluxo de tutorial** na lista de fluxos exibida na guia **meus fluxos** . Isso mostra os detalhes do fluxo que criamos anteriormente.

3. Pressione **executar**.

    ![O botão Executar em automatização de energia.](../images/power-automate-tutorial-8.png)

4. Um painel de tarefas será exibido para executar o fluxo. Se você for solicitado a **entrar no** Excel online, pressione **continuar**.

5. Pressione o **fluxo de execução**. Isso executa o fluxo, que executa o script relacionado do Office.

6. Pressione **concluído**. Você deve ver a seção **runs** Update de acordo.

7. Atualize a página para ver os resultados da automatização de energia. Se tiver êxito, vá para a pasta de trabalho para ver as células atualizadas. Se ele falhar, verifique as configurações do fluxo e execute-o uma segunda vez.

    ![Saída automatizada de energia mostrando uma execução de fluxo bem-sucedida.](../images/power-automate-tutorial-9.png)

## <a name="next-steps"></a>Próximas etapas

Preencha os [scripts executados automaticamente com o tutorial automatizar fluxos de energia automatizada](excel-power-automate-trigger.md) . Ele ensina como transmitir dados de um serviço de fluxo de trabalho para o script do Office.
