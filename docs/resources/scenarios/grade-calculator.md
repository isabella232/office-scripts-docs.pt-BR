---
title: 'Cenário de exemplo de scripts do Office: calculadora de série'
description: Um exemplo que determina a porcentagem e as classificações de uma classe de alunos.
ms.date: 02/20/2020
localization_priority: Normal
ms.openlocfilehash: 0db6f7c116594f7655bfc0adc8f5a79dbbf2a0af
ms.sourcegitcommit: b075eed5a6f275274fbbf6d62633219eac416f26
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/10/2020
ms.locfileid: "42700067"
---
# <a name="office-scripts-sample-scenario-grade-calculator"></a><span data-ttu-id="17d57-103">Cenário de exemplo de scripts do Office: calculadora de série</span><span class="sxs-lookup"><span data-stu-id="17d57-103">Office Scripts sample scenario: Grade calculator</span></span>

<span data-ttu-id="17d57-104">Neste cenário, você é um instrutor tallyingndo notas finais de todos os alunos.</span><span class="sxs-lookup"><span data-stu-id="17d57-104">In this scenario, you're an instructor tallying every student's end-of-term grades.</span></span> <span data-ttu-id="17d57-105">Você está inserindo as pontuações para suas atribuições e testes enquanto vai.</span><span class="sxs-lookup"><span data-stu-id="17d57-105">You've been entering the scores for their assignments and tests as you go.</span></span> <span data-ttu-id="17d57-106">Agora, é hora de determinar os Fates dos alunos.</span><span class="sxs-lookup"><span data-stu-id="17d57-106">Now, it is time to determine the students' fates.</span></span>

<span data-ttu-id="17d57-107">Você desenvolverá um script que totaliza as notas para cada categoria de ponto.</span><span class="sxs-lookup"><span data-stu-id="17d57-107">You'll develop a script that totals the grades for each point category.</span></span> <span data-ttu-id="17d57-108">Em seguida, ele atribuirá uma nota de carta a cada aluno com base no total.</span><span class="sxs-lookup"><span data-stu-id="17d57-108">It will then assign a letter grade to each student based on the total.</span></span> <span data-ttu-id="17d57-109">Para ajudar a garantir a precisão, você adicionará algumas verificações para ver se alguma Pontuação individual é muito baixa ou alta.</span><span class="sxs-lookup"><span data-stu-id="17d57-109">To help ensure accuracy, you'll add a couple checks to see if any individual scores are too low or high.</span></span> <span data-ttu-id="17d57-110">Se a pontuação de um aluno for menor do que zero ou maior do que o valor de ponto possível, o script sinalizará a célula com um preenchimento vermelho e não fará o total dos pontos do aluno.</span><span class="sxs-lookup"><span data-stu-id="17d57-110">If a student's score is less than zero or more than the possible point value, the script will flag the cell with a red fill and not total that student's points.</span></span> <span data-ttu-id="17d57-111">Essa será uma indicação clara de quais registros você precisa fazer uma verificação dupla.</span><span class="sxs-lookup"><span data-stu-id="17d57-111">This will be a clear indication of which records you need to double-check.</span></span> <span data-ttu-id="17d57-112">Você também adicionará alguma formatação básica às notas para que possa exibir rapidamente a parte superior e a parte inferior da classe.</span><span class="sxs-lookup"><span data-stu-id="17d57-112">You'll also add some basic formatting to the grades so you can quickly view the top and bottom of the class.</span></span>

## <a name="scripting-skills-covered"></a><span data-ttu-id="17d57-113">Habilidades de script abordadas</span><span class="sxs-lookup"><span data-stu-id="17d57-113">Scripting skills covered</span></span>

- <span data-ttu-id="17d57-114">Formatação de célula</span><span class="sxs-lookup"><span data-stu-id="17d57-114">Cell formatting</span></span>
- <span data-ttu-id="17d57-115">Verificação de erros</span><span class="sxs-lookup"><span data-stu-id="17d57-115">Error checking</span></span>
- <span data-ttu-id="17d57-116">Expressões regulares</span><span class="sxs-lookup"><span data-stu-id="17d57-116">Regular expressions</span></span>

## <a name="setup-instructions"></a><span data-ttu-id="17d57-117">Instruções de configuração</span><span class="sxs-lookup"><span data-stu-id="17d57-117">Setup instructions</span></span>

1. <span data-ttu-id="17d57-118">Baixe o <a href="grade-calculator.xlsx">grade-Calculator. xlsx</a> para o onedrive.</span><span class="sxs-lookup"><span data-stu-id="17d57-118">Download <a href="grade-calculator.xlsx">grade-calculator.xlsx</a> to your OneDrive.</span></span>

2. <span data-ttu-id="17d57-119">Abra a pasta de trabalho com o Excel para a Web.</span><span class="sxs-lookup"><span data-stu-id="17d57-119">Open the workbook with Excel for the web.</span></span>

3. <span data-ttu-id="17d57-120">Na guia **automatizar** , abra o **Editor de código**.</span><span class="sxs-lookup"><span data-stu-id="17d57-120">Under the **Automate** tab, open the **Code Editor**.</span></span>

4. <span data-ttu-id="17d57-121">No painel de tarefas **Editor de código** , pressione **novo script** e cole o script a seguir no editor.</span><span class="sxs-lookup"><span data-stu-id="17d57-121">In the **Code Editor** task pane, press **New Script** and paste the following script into the editor.</span></span>

    ```TypeScript
    async function main(context: Excel.RequestContext) {
      // Get the number of student record rows.
      let sheet = context.workbook.worksheets.getActiveWorksheet();
      let studentsRange = sheet.getUsedRange().load("values, rowCount");
      await context.sync();
      console.log("Total students: " + (studentsRange.rowCount - 1));

      // Clean up any formatting from previous runs of the script.
      studentsRange.clear(Excel.ClearApplyTo.formats);
      studentsRange.getColumn(4).getCell(0, 0).getRowsBelow(studentsRange.rowCount - 1).clear(Excel.ClearApplyTo.all);
      studentsRange.getColumn(5).getCell(0, 0).getRowsBelow(studentsRange.rowCount - 1).clear(Excel.ClearApplyTo.all);
      await context.sync();

      // Parse the headers for the maximum possible scores for each category.
      // The format is `category (score)`.
      let assignmentsMax = studentsRange.values[0][1].match(/\d+/)[0];
      let midTermMax = studentsRange.values[0][2].match(/\d+/)[0];
      let finalsMax = studentsRange.values[0][3].match(/\d+/)[0];
      console.log("Assignments max score:" + assignmentsMax);
      console.log("Mid-term max score: " + midTermMax);
      console.log("Final max score: " + finalsMax);

      // Look at every student row.
      for (let i = 1; i < studentsRange.values.length; i++) {
        let row = studentsRange.values[i];
        let total = row[1] + row[2] + row[3];
        let valid = true;

        // Look for any records that are too low or too high.
        if (row[1] < 0 || row[1] > assignmentsMax) {
          studentsRange.getCell(i, 1).format.fill.color = "Red";
          valid = false;
        }
        if (row[2] < 0 || row[2] > midTermMax) {
          studentsRange.getCell(i, 2).format.fill.color = "Red";
          valid = false;
        }
        if (row[3] < 0 || row[3] > finalsMax) {
          studentsRange.getCell(i, 3).format.fill.color = "Red";
          valid = false;
        }

        // If the scores are valid, total that student's points and assign them a letter grade.
        if (valid) {
          let grade: string;
          switch (true) {
            case total < 60:
              grade = "E";
              break;
            case total < 70:
              grade = "D";
              break;
            case total < 80:
              grade = "C";
              break;
            case total < 90:
              grade = "B";
              break;
            default:
              grade = "A";
              break;
          }

          studentsRange.getCell(i, 4).values = [[total]];
          studentsRange.getCell(i, 5).values = [[grade]];

          // Highlight excellent students and those in need of attention.
          if (grade === "A") {
            studentsRange.getCell(i, 5).format.fill.color = "Green";
          } else if (grade === "E" || grade === "D") {
            studentsRange.getCell(i, 5).format.fill.color = "Orange";
          }
        }
      }

      studentsRange.getColumn(5).format.horizontalAlignment = "Center";
    }
    ```

5. <span data-ttu-id="17d57-122">Renomeie o script para fazer a **grade** e salve-o.</span><span class="sxs-lookup"><span data-stu-id="17d57-122">Rename the script to **Grade Calculator** and save it.</span></span>

## <a name="running-the-script"></a><span data-ttu-id="17d57-123">Executando o script</span><span class="sxs-lookup"><span data-stu-id="17d57-123">Running the script</span></span>

<span data-ttu-id="17d57-124">Execute o script de **calculadora de nota** na planilha única.</span><span class="sxs-lookup"><span data-stu-id="17d57-124">Run the **Grade Calculator** script on the only worksheet.</span></span> <span data-ttu-id="17d57-125">O script totaliza as notas e atribui a cada aluno uma letra de nota.</span><span class="sxs-lookup"><span data-stu-id="17d57-125">The script will total the grades and assign each student a letter grade.</span></span> <span data-ttu-id="17d57-126">Se qualquer nota individual tiver mais pontos do que a atribuição ou o teste for importante, a classificação transgressor será marcada como vermelho e o total não será calculado.</span><span class="sxs-lookup"><span data-stu-id="17d57-126">If any individual grades have more points than the assignment or test is worth, then the offending grade is marked red and the total is not calculated.</span></span>

### <a name="before-running-the-script"></a><span data-ttu-id="17d57-127">Antes de executar o script</span><span class="sxs-lookup"><span data-stu-id="17d57-127">Before running the script</span></span>

![Uma planilha que mostra linhas de Pontuação para estudantes.](../../images/scenario-grade-calculator-before.png)

### <a name="after-running-the-script"></a><span data-ttu-id="17d57-129">Após executar o script</span><span class="sxs-lookup"><span data-stu-id="17d57-129">After running the script</span></span>

![Uma planilha que mostra os dados da Pontuação do aluno com células inválidas em totais vermelhos para linhas de aluno válidas.](../../images/scenario-grade-calculator-after.png)
