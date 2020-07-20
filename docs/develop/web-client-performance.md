---
title: Melhorar o desempenho dos scripts do Office
description: Crie scripts mais rápidos compreendendo a comunicação entre a pasta de trabalho do Excel e seu script.
ms.date: 06/15/2020
localization_priority: Normal
ms.openlocfilehash: 4d5b7c70f14e3fc598b95a6226e3ef8caf89f651
ms.sourcegitcommit: aec3c971c6640429f89b6bb99d2c95ea06725599
ms.translationtype: Auto
ms.contentlocale: pt-BR
ms.lasthandoff: 06/25/2020
ms.locfileid: "44878719"
---
# <a name="improve-the-performance-of-your-office-scripts"></a>Melhorar o desempenho dos scripts do Office

A finalidade dos scripts do Office é automatizar a série de tarefas realizadas com frequência para poupar tempo. Um script lento pode parecer que não acelera o fluxo de trabalho. Na maioria das vezes, seu script será perfeitamente bom e será executado conforme o esperado. No entanto, há alguns cenários do avoidable que podem afetar o desempenho.

O motivo mais comum para um script lento é a comunicação excessiva com a pasta de trabalho. O script é executado no computador local, enquanto a pasta de trabalho existe na nuvem. Em determinados momentos, o script sincroniza seus dados locais com o da pasta de trabalho. Isso significa que qualquer operação de gravação (como `workbook.addWorksheet()` ) só será aplicada à pasta de trabalho quando essa sincronização por trás da cena acontecer. Da mesma forma, qualquer operação de leitura (como `myRange.getValues()` ) só obterá dados da pasta de trabalho para o script nesses horários. Em ambos os casos, o script busca informações antes que ele atue nos dados. Por exemplo, o código a seguir registra com precisão o número de linhas no intervalo usado.

```TypeScript
let usedRange = workbook.getActiveWorksheet().getUsedRange();
let rowCount = usedRange.getRowCount();
// The script will read the range and row count from
// the workbook before logging the information.
console.log(rowCount);
```

As APIs de scripts do Office garantem que todos os dados na pasta de trabalho ou script sejam precisos e atualizados quando necessário. Você não precisa se preocupar com essas sincronizações para que o script seja executado corretamente. No entanto, um reconhecimento dessa comunicação de script para nuvem pode ajudar você a evitar chamadas de rede desnecessárias.

## <a name="performance-optimizations"></a>Otimizações de desempenho

Você pode aplicar técnicas simples para ajudar a reduzir a comunicação com a nuvem. Os padrões a seguir ajudam a acelerar seus scripts.

- Leia os dados da pasta de trabalho uma vez em vez de repetidamente em um loop.
- Remover instruções desnecessárias `console.log` .
- Evite usar blocos try/catch.

### <a name="read-workbook-data-outside-of-a-loop"></a>Ler dados de pasta de trabalho fora de um loop

Qualquer método que obtém dados da pasta de trabalho pode disparar uma chamada de rede. Em vez de fazer repetidamente a mesma chamada, você deve salvar dados localmente, sempre que possível. Isso se aplica especialmente ao lidar com loops.

Considere um script para obter a contagem de números negativos no intervalo usado de uma planilha. O script precisa iterar em todas as células do intervalo usado. Para fazer isso, é necessário o intervalo, o número de linhas e o número de colunas. Você deve armazená-las como variáveis locais antes de iniciar o loop. Caso contrário, cada iteração do loop forçará um retorno à pasta de trabalho.

```TypeScript
/**
 * This script provides the count of negative numbers that are present
 * in the used range of the current worksheet.
 */
function main(workbook: ExcelScript.Workbook) {
  // Get the working range.
  let usedRange = workbook.getActiveWorksheet().getUsedRange();

  // Save the values locally to avoid repeatedly asking the workbook.
  let usedRangeValues = usedRange.getValues();

  // Start the negative number counter.
  let negativeCount = 0;

  // Iterate over the entire range looking for negative numbers.
  for (let i = 0; i < usedRangeValues.length; i++) {
    for (let j = 0; j < usedRangeValues[i].length; j++) {
      if (usedRangeValues[i][j] < 0) {
        negativeCount++;
      }
    }
  }

  // Log the negative number count to the console.
  console.log(negativeCount);
}
```

> [!NOTE]
> Como experimento, tente substituir `usedRangeValues` no loop por `usedRange.getValues()` . Você pode notar que o script demora consideravelmente mais para ser executado ao lidar com intervalos grandes.

### <a name="remove-unnecessary-consolelog-statements"></a>Remover instruções desnecessárias `console.log`

O registro em log do console é uma ferramenta vital para [depurar seus scripts](../testing/troubleshooting.md). No entanto, ele força o script a sincronizar com a pasta de trabalho para garantir que as informações registradas estejam atualizadas. Considere remover declarações desnecessárias de registro em log (como as usadas para teste) antes de compartilhar seu script. Isso normalmente não causará um problema de desempenho perceptível, a menos que a `console.log()` instrução esteja em um loop.

### <a name="avoid-using-trycatch-blocks"></a>Evite usar blocos try/catch

Não recomendamos o uso de [ `try` / `catch` blocos](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) como parte do fluxo de controle esperado de um script. A maioria dos erros pode ser evitada verificando os objetos retornados da pasta de trabalho. Por exemplo, o script a seguir verifica se a tabela retornada pela pasta de trabalho existe antes de tentar adicionar uma linha.

```TypeScript
/**
 * This script adds a row to "MyTable", if that table is present.
 */
function main(workbook: ExcelScript.Workbook) {
  let table = workbook.getTable("MyTable");

  // Check if the table exists.
  if (table) {
    // Add the row.
    table.addRow(-1, ["2012", "Yes", "Maybe"]);
  } else {
    // Report the missing table.
    console.log("MyTable not found.");
  }
}
```

## <a name="case-by-case-help"></a>Ajuda caso a caso

Como a plataforma de scripts do Office se expande para trabalhar com [automatização de energia](https://flow.microsoft.com/), [cartões adaptáveis](https://docs.microsoft.com/adaptive-cards)e outros recursos entre produtos, os detalhes da comunicação de pasta de trabalho de script ficam mais complexos. Se você precisar de ajuda para executar o script com mais rapidez, saia pelo [estouro de pilha](https://stackoverflow.com/questions/tagged/office-scripts). Certifique-se de marcar sua pergunta com "Office-scripts" para que os especialistas possam encontrá-lo e ajudá-lo.

## <a name="see-also"></a>Confira também

- [Fundamentos de script para scripts do Office no Excel na Web](scripting-fundamentals.md)
- [MDN Web docs: loops e iteração](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Loops_and_iteration)
