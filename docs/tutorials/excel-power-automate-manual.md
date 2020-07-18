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
# <a name="call-scripts-from-a-manual-power-automate-flow-preview"></a>Comece a usar scripts de um fluxo manual do Power Automate (pré-visualização)

Este tutorial ensina como executar um Script do Office para o Excel na web por meio do [Power Automate](https://flow.microsoft.com).

## <a name="prerequisites"></a>Pré-requisitos

[!INCLUDE [Tutorial prerequisites](../includes/power-automate-tutorial-prerequisites.md)]

> [!IMPORTANT]
> Este tutorial pressupõe que você tenha concluído o tutorial [Registrar, editar e criar Scripts do Office no Excel na Web](excel-tutorial.md).

## <a name="prepare-the-workbook"></a>Preparar a pasta de trabalho

O Power Automate não consegue usar referências relativas como `Workbook.getActiveWorksheet` para acessar os componentes da pasta de trabalho. Portanto, precisamos de uma pasta de trabalho e de uma planilha com nomes consistentes que o Power Automate consiga consultar.

1. Crie uma pasta de trabalho intitulada **MyWorkbook**.

2. Na pasta de trabalho **MyWorkbook**, crie uma planilha intitulada **TutorialWorksheet**.

## <a name="create-an-office-script"></a>Criar um Script do Office

1. Vá para a guia **Automatizar** e selecione **Editor de Códigos**.

2. Selecione **Novo Script**.

3. Substitua o script padrão pelo script abaixo. Esse script adiciona a data e hora atuais às duas primeiras células da planilha **TutorialWorksheet**.

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

4. Renomeie o script como **Definir data e hora**. Pressione o nome do script para alterá-lo.

5. Salve o script pressionando **Salvar Script**.

## <a name="create-an-automated-workflow-with-power-automate"></a>Criar um fluxo de trabalho automatizado com o Power Automate

1. Entre no [site do Power Automate](https://flow.microsoft.com).

2. No menu exibido do lado esquerdo da tela, pressione **Criar**. Isso o conduzirá a uma lista de maneiras de criar novos fluxos de trabalho.

    ![Botão Criar no Power Automate.](../images/power-automate-tutorial-1.png)

3. Na seção **Começar no espaço em branco**, selecione **Fluxo instantâneo**. Isso irá criar um fluxo de trabalho ativado manualmente.

    ![Opção Fluxo instantâneo para criar um novo fluxo de trabalho.](../images/power-automate-tutorial-2.png)

4. Na janela da caixa de diálogo que aparece, insira um nome para o seu fluxo na caixa de texto **Nome do fluxo**; selecione **Acionar um fluxo manualmente** na lista de opções em **Escolher como acionar o fluxo**, e pressione **Criar**.

    ![Opção acionamento manual para a criação de um novo fluxo instantâneo.](../images/power-automate-tutorial-3.png)

    Observe que o fluxo acionado manualmente é apenas um entre os diversos tipos de fluxo. No tutorial a seguir, você criará um fluxo que é executado automaticamente quando você recebe um email.

5. Pressione **Nova etapa**.

6. Selecione a guia **Padrão** e, em seguida, selecione **Excel Online (Business)**.

    ![Opção Power Automate para o Excel Online (Business).](../images/power-automate-tutorial-4.png)

7. Em **Ações**, selecione **Executar script (pré-visualização)**.

    ![Opção de ação do Power Automate para Executar script (pré-visualização).](../images/power-automate-tutorial-5.png)

8. Especifique as seguintes configurações para o conector **Executar script**:

    - **Localização**: OneDrive for Business
    - **Biblioteca de Documentos**: OneDrive
    - **Arquivo**: MyWorkbook.xlsx
    - **Script**: Definir data e hora

    ![Configurações do conector para executar um script no Power Automate.](../images/power-automate-tutorial-6.png)

9. Pressione **Salvar**.

Seu fluxo agora está pronto para ser executado por meio do Power Automate. Você pode testá-lo usando o botão **Testar** no editor de fluxo ou seguir as etapas restantes do tutorial para executar o fluxo a partir da sua coleção de fluxos.

## <a name="run-the-script-through-power-automate"></a>Executar o script por meio da automação

1. Na página principal do Power Automate, selecione **Meus fluxos**.

    ![Botão Meus fluxos no Power Automate.](../images/power-automate-tutorial-7.png)

2. Selecione **Fluxo do meu tutorial** na lista de fluxos exibida na guia **Meus fluxos**. Isso irá lhe mostrar os detalhes do fluxo que criamos anteriormente.

3. Pressione **Executar**.

    ![Botão Executar no Power Automate.](../images/power-automate-tutorial-8.png)

4. Um painel de tarefas irá aparecer para executar o fluxo. Se você for solicitado a **Entrar** no Excel Online, faça o login pressionando **Continuar**.

5. Pressione **Executar o fluxo**. Isso executará o fluxo, que, por sua vez, executará o Script do Office associado.

6. Pressione **Concluído**. Você deverá ver a seção **Executar** ser atualizada de acordo.

7. Atualize a página para ver os resultados do Power Automate. Se o script tiver sido bem-sucedido, vá para a pasta de trabalho para ver as células atualizadas. Se tiver falhado, verifique as configurações do fluxo e execute-o novamente.

    ![Resultado do Power Automate mostrando um fluxo executado com sucesso.](../images/power-automate-tutorial-9.png)

## <a name="next-steps"></a>Próximas etapas

Faça o tutorial [Transferir dados para scripts em um fluxo executado automaticamente pelo Power Automate](excel-power-automate-trigger.md). O tutorial ensinará como transferir dados de um serviço de fluxo de trabalho para o seu Script do Office e executar o fluxo do Power Automate quando certos eventos ocorrerem.
