---
title: Grave, edite e crie scripts do Office no Excel na Web
description: Um tutorial sobre o básico dos scripts do Office, incluindo a gravação de scripts com o Gravador de ações e a gravação de dados em uma pasta de trabalho.
ms.date: 01/27/2020
localization_priority: Priority
ms.openlocfilehash: 1971ff2ffd80554beb6ac561677ee3384f87ca81
ms.sourcegitcommit: b075eed5a6f275274fbbf6d62633219eac416f26
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/10/2020
ms.locfileid: "42700037"
---
# <a name="record-edit-and-create-office-scripts-in-excel-on-the-web"></a>Grave, edite e crie scripts do Office no Excel na Web

Este tutorial ensinará os conceitos básicos de gravação, edição e escrita de um Script do Office para Excel na Web.

## <a name="prerequisites"></a>Pré-requisitos

[!INCLUDE [Preview note](../includes/preview-note.md)]

Antes de iniciar este tutorial, você precisará acessar os scripts do Office, que exigem o seguinte:

- [Excel na Web](https://www.office.com/launch/excel).
- Peça para o administrador [habilitar os scripts do Office da sua organização](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf), o que adiciona a guia **Automação** à faixa de opções.

> [!IMPORTANT]
> Este tutorial é destinado a pessoas com conhecimento básico ou de nível intermediário de JavaScript ou TypeScript. Se você não conhece o JavaScript, recomendamos que revise o [tutorial do Mozilla JavaScript](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction). Visite [Scripts do Office no Excel na Web](../overview/excel.md) para saber mais sobre o ambiente de scripts.

## <a name="add-data-and-record-a-basic-script"></a>Adicione dados e grave um script básico

Primeiro, precisaremos de alguns dados e um pequeno script inicial.

1. Crie uma nova pasta de trabalho no Excel para a Web.
2. Copie os seguintes dados de vendas de frutas e cole-os na planilha, começando na célula **A1**.

    |Fruta |2018 |2019 |
    |:---|:---|:---|
    |Laranjas |1.000 |1.200 |
    |Limões |800 |900 |
    |Limões-galego |600 |500 |
    |Toranjas |900 |700 |

3. Abra a guia **Automação**. Se você não vir a guia **Automação**, verifique o extravasamento da fita pressionando a seta suspensa.
4. Pressione o botão **Ações de registro**.
5. Clique nas células **A2:C2** (a linha "Laranjas") e defina a cor de preenchimento como laranja.
6. Pare a gravação pressionando o botão **Parar**.
7. Preencha o campo **Nome do script** com um nome digno de memória.
8. *Opcional:* Preencha o campo **Descrição** com uma descrição significativa. Isso é usado para fornecer contexto sobre o que o script faz. Para o tutorial, você pode usar "Linhas de códigos de cores de uma tabela".

   > [!TIP]
   > Você pode editar a descrição de um script posteriormente no painel **Detalhes do script**, localizado no menu do Editor de códigos **...**.

9. Salve o script pressionando o botão **Salvar**.

    Sua planilha deve ficar assim (não se preocupe se a cor for diferente):

    ![Uma linha de dados de vendas de frutas com a linha "Laranjas" destacada em laranja.](../images/tutorial-1.png)

## <a name="edit-an-existing-script"></a>Edite um script existente

O script anterior coloriu a linha "Laranjas" para ficar laranja. Vamos adicionar uma linha amarela aos "Limões".

1. Abra a guia **Automação**.
2. Pressione o botão **Editor de códigos**.
3. Abra o script que você gravou na seção anterior. Você deve ver algo semelhante a este código:

    ```TypeScript
    async function main(context: Excel.RequestContext) {
      // Set fill color to FFC000 for range Sheet1!A2:C2
      let workbook = context.workbook;
      let worksheets = workbook.worksheets;
      let selectedSheet = worksheets.getActiveWorksheet();
      selectedSheet.getRange("A2:C2").format.fill.color = "FFC000";
    }
    ```

    Esse código obtém a planilha atual acessando primeiro a coleção de planilhas da pasta de trabalho. Depois, defina a cor de preenchimento do intervalo **A2:C2**.

    Os intervalos são parte fundamental dos scripts do Office no Excel na Web. Um intervalo é um bloco retangular e contíguo de células que contém valores, fórmula e formatação. Eles são a estrutura básica das células através da qual você executará a maioria das tarefas de script.

4. Adicione a seguinte linha no final do script (entre onde `color` está definido e o encerramento `}`):

    ```TypeScript
    selectedSheet.getRange("A3:C3").format.fill.color = "yellow";
    ```

5. Teste o script pressionando **Executar**. Sua pasta de trabalho já deve ter esta aparência:

    ![Uma linha de dados de vendas de frutas com a linha "Laranjas" é realçada em laranja e a linha "Limões" é realçada em amarelo.](../images/tutorial-2.png)

## <a name="create-a-table"></a>Crie uma tabela

Vamos converter esses dados de vendas de frutas em uma tabela. Usaremos nosso script em todo o processo.

1. Adicione a seguinte linha no final do script (antes do encerramento `}`):

    ```TypeScript
    let table = selectedSheet.tables.add("A1:C5", true);
    ```

2. Essa chamada retorna um `Table` objeto. Vamos usar essa tabela para classificar os dados. Classificaremos os dados em ordem crescente com base nos valores na coluna "Frutas". Adicione a seguinte linha assim que criar a tabela:

    ```TypeScript
    table.sort.apply([{ key: 0, ascending: true }]);
    ```

    Seu script deve ter esta aparência:

    ```TypeScript
    async function main(context: Excel.RequestContext) {
      // Set fill color to FFC000 for range Sheet1!A2:C2
      let workbook = context.workbook;
      let worksheets = workbook.worksheets;
      let selectedSheet = worksheets.getActiveWorksheet();
      selectedSheet.getRange("A2:C2").format.fill.color = "FFC000";
      selectedSheet.getRange("A3:C3").format.fill.color = "yellow";
      let table = selectedSheet.tables.add("A1:C5", true);
      table.sort.apply([{ key: 0, ascending: true }]);
    }
    ```

    As tabelas têm um objeto `TableSort` acessado através da propriedade `Table.sort`. Você pode aplicar critérios de classificação a esse objeto. O `apply` método utiliza uma matriz de `SortField` objetos. Nesse caso, só temos um critério de classificação, por isso só usamos um. `SortField`. `key: 0` define a coluna com os valores que determinam a classificação como "0" (que nesse caso, é a primeira coluna na tabela **A** ). `ascending: true` classifica os dados em ordem crescente (em vez de ordem decrescente).

3. Execute o script. Você deve visualizar uma tabela como esta:

    ![Uma tabela de vendas de frutas sortidas.](../images/tutorial-3.png)

    > [!NOTE]
    > Se você executar novamente o script, receberá um erro. Isso ocorre porque você não pode criar uma tabela em cima de outra tabela. No entanto, você pode executar o script em uma planilha ou pasta de trabalho diferente.

### <a name="re-run-the-script"></a>Reexecute o script

1. Crie uma nova planilha na pasta de trabalho atual.
2. Copie os dados das frutas do início do tutorial e cole-os na nova planilha, começando na célula **A1**.
3. Execute o script.

## <a name="next-steps"></a>Próximas etapas

Conclua o tutorial [Ler os dados da pasta de trabalho com scripts do Office no Excel na Web.](excel-read-tutorial.md). Ele ensina como ler dados de uma pasta de trabalho com um script do Office.
