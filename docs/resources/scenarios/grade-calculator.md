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
# <a name="office-scripts-sample-scenario-grade-calculator"></a>Cenário de exemplo de scripts do Office: calculadora de série

Neste cenário, você é um instrutor tallyingndo notas finais de todos os alunos. Você está inserindo as pontuações para suas atribuições e testes enquanto vai. Agora, é hora de determinar os Fates dos alunos.

Você desenvolverá um script que totaliza as notas para cada categoria de ponto. Em seguida, ele atribuirá uma nota de carta a cada aluno com base no total. Para ajudar a garantir a precisão, você adicionará algumas verificações para ver se alguma Pontuação individual é muito baixa ou alta. Se a pontuação de um aluno for menor do que zero ou maior do que o valor de ponto possível, o script sinalizará a célula com um preenchimento vermelho e não fará o total dos pontos do aluno. Essa será uma indicação clara de quais registros você precisa fazer uma verificação dupla. Você também adicionará alguma formatação básica às notas para que possa exibir rapidamente a parte superior e a parte inferior da classe.

## <a name="scripting-skills-covered"></a>Habilidades de script abordadas

- Formatação de célula
- Verificação de erros
- Expressões regulares

## <a name="setup-instructions"></a>Instruções de configuração

1. Baixe o <a href="grade-calculator.xlsx">grade-Calculator. xlsx</a> para o onedrive.

2. Abra a pasta de trabalho com o Excel para a Web.

3. Na guia **automatizar** , abra o **Editor de código**.

4. No painel de tarefas **Editor de código** , pressione **novo script** e cole o script a seguir no editor.

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

5. Renomeie o script para fazer a **grade** e salve-o.

## <a name="running-the-script"></a>Executando o script

Execute o script de **calculadora de nota** na planilha única. O script totaliza as notas e atribui a cada aluno uma letra de nota. Se qualquer nota individual tiver mais pontos do que a atribuição ou o teste for importante, a classificação transgressor será marcada como vermelho e o total não será calculado.

### <a name="before-running-the-script"></a>Antes de executar o script

![Uma planilha que mostra linhas de Pontuação para estudantes.](../../images/scenario-grade-calculator-before.png)

### <a name="after-running-the-script"></a>Após executar o script

![Uma planilha que mostra os dados da Pontuação do aluno com células inválidas em totais vermelhos para linhas de aluno válidas.](../../images/scenario-grade-calculator-after.png)
