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
# <a name="office-scripts-sample-scenario-analyze-web-downloads"></a><span data-ttu-id="dff42-103">Cenário de exemplo de scripts do Office: analisar downloads da Web</span><span class="sxs-lookup"><span data-stu-id="dff42-103">Office Scripts sample scenario: Analyze web downloads</span></span>

<span data-ttu-id="dff42-104">Neste cenário, você está com a tarefa de analisar relatórios de download no site da sua empresa.</span><span class="sxs-lookup"><span data-stu-id="dff42-104">In this scenario, you're tasked with analyzing download reports from your company's website.</span></span> <span data-ttu-id="dff42-105">O objetivo dessa análise é determinar se o tráfego da Web está vindo dos Estados Unidos ou em qualquer lugar do mundo.</span><span class="sxs-lookup"><span data-stu-id="dff42-105">The goal of this analysis is to determine if the web traffic is coming from the United States or elsewhere in the world.</span></span>

<span data-ttu-id="dff42-106">Seus colegas carregam os dados brutos na sua pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="dff42-106">Your colleagues upload the raw data to your workbook.</span></span> <span data-ttu-id="dff42-107">O conjunto de dados de cada semana tem sua própria planilha.</span><span class="sxs-lookup"><span data-stu-id="dff42-107">Each week's set of data has its own worksheet.</span></span> <span data-ttu-id="dff42-108">Há também a planilha de **Resumo** com uma tabela e um gráfico que mostra as tendências da semana sobre a semana.</span><span class="sxs-lookup"><span data-stu-id="dff42-108">There is also the **Summary** worksheet with a table and chart that shows week-over-week trends.</span></span>

<span data-ttu-id="dff42-109">Você desenvolverá um script que analisa dados de downloads semanais na planilha ativa.</span><span class="sxs-lookup"><span data-stu-id="dff42-109">You'll develop a script that analyzes weekly downloads data in the active worksheet.</span></span> <span data-ttu-id="dff42-110">Ele analisará o endereço IP associado a cada download e determinará se ele veio ou não dos EUA.</span><span class="sxs-lookup"><span data-stu-id="dff42-110">It will parse the IP address associated with each download and determine whether or not it came from the US.</span></span> <span data-ttu-id="dff42-111">A resposta será inserida na planilha como um valor booliano ("TRUE" ou "FALSE") e a formatação condicional será aplicada a essas células.</span><span class="sxs-lookup"><span data-stu-id="dff42-111">The answer will be inserted in the worksheet as a boolean value ("TRUE" or "FALSE") and conditional formatting will be applied to those cells.</span></span> <span data-ttu-id="dff42-112">Os resultados do local do endereço IP serão totalizados na planilha e copiados para a tabela Resumo.</span><span class="sxs-lookup"><span data-stu-id="dff42-112">The IP address location results will be totaled on the worksheet and copied to the summary table.</span></span>

## <a name="scripting-skills-covered"></a><span data-ttu-id="dff42-113">Habilidades de script abordadas</span><span class="sxs-lookup"><span data-stu-id="dff42-113">Scripting skills covered</span></span>

- <span data-ttu-id="dff42-114">Análise de texto</span><span class="sxs-lookup"><span data-stu-id="dff42-114">Text parsing</span></span>
- <span data-ttu-id="dff42-115">Subfunções em scripts</span><span class="sxs-lookup"><span data-stu-id="dff42-115">Subfunctions in scripts</span></span>
- <span data-ttu-id="dff42-116">Formatação condicional</span><span class="sxs-lookup"><span data-stu-id="dff42-116">Conditional formatting</span></span>
- <span data-ttu-id="dff42-117">Tabelas</span><span class="sxs-lookup"><span data-stu-id="dff42-117">Tables</span></span>

## <a name="demo-video"></a><span data-ttu-id="dff42-118">Vídeo de demonstração</span><span class="sxs-lookup"><span data-stu-id="dff42-118">Demo video</span></span>

<span data-ttu-id="dff42-119">Este exemplo foi demonstrado como parte da chamada da comunidade de desenvolvedores dos suplementos do Office para fevereiro de 2020.</span><span class="sxs-lookup"><span data-stu-id="dff42-119">This sample was demoed as part of the Office Add-ins developer community call for February 2020.</span></span>

> [!VIDEO https://www.youtube.com/embed/vPEqbb7t6-Y?start=154]

## <a name="setup-instructions"></a><span data-ttu-id="dff42-120">Instruções de configuração</span><span class="sxs-lookup"><span data-stu-id="dff42-120">Setup instructions</span></span>

1. <span data-ttu-id="dff42-121">Baixe o <a href="analyze-web-downloads.xlsx">Analyze-Web-downloads. xlsx</a> para o onedrive.</span><span class="sxs-lookup"><span data-stu-id="dff42-121">Download <a href="analyze-web-downloads.xlsx">analyze-web-downloads.xlsx</a> to your OneDrive.</span></span>

2. <span data-ttu-id="dff42-122">Abra a pasta de trabalho com o Excel para a Web.</span><span class="sxs-lookup"><span data-stu-id="dff42-122">Open the workbook with Excel for the web.</span></span>

3. <span data-ttu-id="dff42-123">Na guia **automatizar** , abra o **Editor de código**.</span><span class="sxs-lookup"><span data-stu-id="dff42-123">Under the **Automate** tab, open the **Code Editor**.</span></span>

4. <span data-ttu-id="dff42-124">No painel de tarefas **Editor de código** , pressione **novo script** e cole o script a seguir no editor.</span><span class="sxs-lookup"><span data-stu-id="dff42-124">In the **Code Editor** task pane, press **New Script** and paste the following script into the editor.</span></span>

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

5. <span data-ttu-id="dff42-125">Renomeie o script para **analisar downloads da Web** e salvá-lo.</span><span class="sxs-lookup"><span data-stu-id="dff42-125">Rename the script to **Analyze Web Downloads** and save it.</span></span>

## <a name="running-the-script"></a><span data-ttu-id="dff42-126">Executando o script</span><span class="sxs-lookup"><span data-stu-id="dff42-126">Running the script</span></span>

<span data-ttu-id="dff42-127">Navegue até qualquer uma das planilhas \*\*semana\* \*\* e execute o script de **análise de downloads da Web** .</span><span class="sxs-lookup"><span data-stu-id="dff42-127">Navigate to any of the **Week\*\*** worksheets and run the **Analyze Web Downloads** script.</span></span> <span data-ttu-id="dff42-128">O script aplicará a formatação condicional e o rótulo de local na planilha atual.</span><span class="sxs-lookup"><span data-stu-id="dff42-128">The script will apply the conditional formatting and location labelling on the current sheet.</span></span> <span data-ttu-id="dff42-129">Ele também atualizará a planilha de **Resumo** .</span><span class="sxs-lookup"><span data-stu-id="dff42-129">It will also update the **Summary** worksheet.</span></span>

### <a name="before-running-the-script"></a><span data-ttu-id="dff42-130">Antes de executar o script</span><span class="sxs-lookup"><span data-stu-id="dff42-130">Before running the script</span></span>

![Uma planilha que mostra dados brutos de tráfego da Web.](../../images/scenario-analyze-web-downloads-before.png)

### <a name="after-running-the-script"></a><span data-ttu-id="dff42-132">Após executar o script</span><span class="sxs-lookup"><span data-stu-id="dff42-132">After running the script</span></span>

![Uma planilha que mostra informações de local de IP formatados com as linhas de tráfego da Web anteriores.](../../images/scenario-analyze-web-downloads-after.png)

![A tabela e o gráfico de resumo que resumem as planilhas nas quais o script foi executado.](../../images/scenario-analyze-web-downloads-table.png)
