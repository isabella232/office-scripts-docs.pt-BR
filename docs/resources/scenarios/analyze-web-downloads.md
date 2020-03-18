---
title: 'Cenário de exemplo de scripts do Office: analisar downloads da Web'
description: Um exemplo que obtém dados brutos de tráfego da Internet em uma pasta de trabalho do Excel e determina o local de origem, antes de organizá-las em uma tabela.
ms.date: 02/20/2020
localization_priority: Normal
ms.openlocfilehash: 9ee12c8d4ca7c191168e3734d7cd9eadc333c165
ms.sourcegitcommit: b075eed5a6f275274fbbf6d62633219eac416f26
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/10/2020
ms.locfileid: "42700070"
---
# <a name="office-scripts-sample-scenario-analyze-web-downloads"></a>Cenário de exemplo de scripts do Office: analisar downloads da Web

Neste cenário, você está com a tarefa de analisar relatórios de download no site da sua empresa. O objetivo dessa análise é determinar se o tráfego da Web está vindo dos Estados Unidos ou em qualquer lugar do mundo.

Seus colegas carregam os dados brutos na sua pasta de trabalho. O conjunto de dados de cada semana tem sua própria planilha. Há também a planilha de **Resumo** com uma tabela e um gráfico que mostra as tendências da semana sobre a semana.

Você desenvolverá um script que analisa dados de downloads semanais na planilha ativa. Ele analisará o endereço IP associado a cada download e determinará se ele veio ou não dos EUA. A resposta será inserida na planilha como um valor booliano ("TRUE" ou "FALSE") e a formatação condicional será aplicada a essas células. Os resultados do local do endereço IP serão totalizados na planilha e copiados para a tabela Resumo.

## <a name="scripting-skills-covered"></a>Habilidades de script abordadas

- Análise de texto
- Subfunções em scripts
- Formatação condicional
- Tabelas

## <a name="demo-video"></a>Vídeo de demonstração

Este exemplo foi demonstrado como parte da chamada da comunidade de desenvolvedores dos suplementos do Office para fevereiro de 2020.

> [!VIDEO https://www.youtube.com/embed/vPEqbb7t6-Y?start=154]

## <a name="setup-instructions"></a>Instruções de configuração

1. Baixe o <a href="analyze-web-downloads.xlsx">Analyze-Web-downloads. xlsx</a> para o onedrive.

2. Abra a pasta de trabalho com o Excel para a Web.

3. Na guia **automatizar** , abra o **Editor de código**.

4. No painel de tarefas **Editor de código** , pressione **novo script** e cole o script a seguir no editor.

    ```TypeScript
      async function main(context: Excel.RequestContext) {
        let currentWorksheet = context.workbook.worksheets
          .getActiveWorksheet();
        // Get the values of the active range of the active worksheet.
        let logRange = currentWorksheet.getUsedRange().load("values");

        // Get the Summary worksheet and table.
        let summaryWorksheet = context.workbook.worksheets.getItem("Summary");
        let summaryTable = context.workbook.tables.getItem("Table1");

        // Get the range that will contain TRUE/FALSE if the IP address is from the United States (US).
        let isUSColumn = logRange
          .getLastColumn()
          .getOffsetRange(0, 1)
          .load("address");

        // Get the values of all the US IP addresses.
        let ipRange = context.workbook.worksheets
          .getItem("USIPAddresses")
          .getUsedRange()
          .load("values");
        await context.sync();

        // Remove the first row.
        let topRow = logRange.values.shift();

        // Create a new array to contain the boolean representing if this is a US IP address.
        let newCol = [[]];

        // Go through each row in worksheet and add Boolean.
        for (let i = 0; i < logRange.values.length; i++) {
          let curRowIP = logRange.values[i][1];
          if (findIP(ipRange.values, ipAddressToInteger(curRowIP)) > 0) {
            newCol.push([true]);
          } else {
            newCol.push([false]);
          }
        }

        // Remove the empty column header and add proper heading.
        newCol.shift();
        newCol.unshift(["Is US IP"]);

        // Write the result to the spreadsheet.
        isUSColumn.values = newCol;
        addSummaryData();
        applyConditionalFormatting();
        currentWorksheet.getUsedRange().format.autofitColumns();

        // Get the calculated summary data.
        let summaryRange = currentWorksheet.getRange("J2:M2").load("values");
        await context.sync();

        // Add the corresponding row to the summary table.
        summaryTable.rows.add(null, summaryRange.values);

        // Function to apply conditional formatting to the new column.
        function applyConditionalFormatting() {
          // Add conditional formatting to the new column.
          let conditionalFormatTrue = isUSColumn.conditionalFormats.add(
            Excel.ConditionalFormatType.cellValue
          );
          let conditionalFormatFalse = isUSColumn.conditionalFormats.add(
            Excel.ConditionalFormatType.cellValue
          );
          // Set TRUE to light blue and FALSE to light orange.
          conditionalFormatTrue.cellValue.format.fill.color = "#8FA8DB";
          conditionalFormatTrue.cellValue.rule = {
            formula1: "=TRUE",
            operator: "EqualTo"
          };
          conditionalFormatFalse.cellValue.format.fill.color = "#F8CCAD";
          conditionalFormatFalse.cellValue.rule = {
            formula1: "=FALSE",
            operator: "EqualTo"
          };
        }

        // Adds the summary data to the current sheet and to the summary table.
        function addSummaryData() {
          // Add a summary row and table.
          let summaryHeader = [["Year", "Week", "US", "Other"]];
          let countTrueFormula =
            "=COUNTIF(" + isUSColumn.address + ', "=TRUE")/' + (newCol.length - 1);
          let countFalseFormula =
            "=COUNTIF(" + isUSColumn.address + ', "=FALSE")/' + (newCol.length - 1);

          let summaryContent = [
            [
              '=TEXT(A2,"YYYY")',
              '=TEXTJOIN(" ", FALSE, "Wk", WEEKNUM(A2))',
              countTrueFormula,
              countFalseFormula
            ]
          ];
          let summaryHeaderRow = context.workbook.worksheets
            .getActiveWorksheet()
            .getRange("J1:M1");
          let summaryContentRow = context.workbook.worksheets
            .getActiveWorksheet()
            .getRange("J2:M2");
          summaryHeaderRow.values = summaryHeader;
          summaryContentRow.values = summaryContent;
          let formats = [[".000", ".000"]];
          summaryContentRow
            .getOffsetRange(0, 2)
            .getResizedRange(0, -2).numberFormat = formats;
        }
      }

      // Translate an IP address into an integer.
      function ipAddressToInteger(ipAddress: string) {
        // Split the IP address into octets.
        let octets = ipAddress.split(".");

        // Create a number for each octet and do the math to create the integer value of the IP address.
        let fullNum =
          // Define an arbitrary number for the last octet.
          111 +
          parseInt(octets[2]) * 256 +
          parseInt(octets[1]) * 65536 +
          parseInt(octets[0]) * 16777216;
        return fullNum;
      }

      // Return the row number where the ip address is found.
      function findIP(ipLookupTable: number[][], n: number) {
        for (let i = 0; i < ipLookupTable.length; i++) {
          if (ipLookupTable[i][0] <= n && ipLookupTable[i][1] >= n) {
            return i;
          }
        }
        return -1;
      }
    ```

5. Renomeie o script para **analisar downloads da Web** e salvá-lo.

## <a name="running-the-script"></a>Executando o script

Navegue até qualquer uma das planilhas **semana\* ** e execute o script de **análise de downloads da Web** . O script aplicará a formatação condicional e o rótulo de local na planilha atual. Ele também atualizará a planilha de **Resumo** .

### <a name="before-running-the-script"></a>Antes de executar o script

![Uma planilha que mostra dados brutos de tráfego da Web.](../../images/scenario-analyze-web-downloads-before.png)

### <a name="after-running-the-script"></a>Após executar o script

![Uma planilha que mostra informações de local de IP formatados com as linhas de tráfego da Web anteriores.](../../images/scenario-analyze-web-downloads-after.png)

![A tabela e o gráfico de resumo que resumem as planilhas nas quais o script foi executado.](../../images/scenario-analyze-web-downloads-table.png)
