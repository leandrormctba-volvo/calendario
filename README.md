# Calendário de Atividades 2026 AV

Dashboard dos eventos do calendário 2026, gerado a partir da planilha **Kanban calendário** (aba *Responsabilidades*).

🔗 **Dashboard online:** https://leandrormctba-volvo.github.io/calendario-av-2026/

## A apresentação (7 slides)

Ao abrir, aparece primeiro a logo da Associação Volvo (cerca de 2,5 s; clique ou aperte uma tecla para pular) e depois entra o primeiro slide.

1. **Abertura**: logo da Associação Volvo e o título Calendário de Atividades 2026, centralizados, com uma folha de calendário enquadrada bem suave ao fundo.
2. **Resumo do ano**: título, período e os números principais, incluindo a média de execução do ano (a mesma da planilha: média do % de cada mês, linha de JANEIRO, coluna %).
3. **Os números de 2026**: tarefas propostas, realizadas e canceladas. Passe o mouse nos boxes para ver os detalhes.
4. **Comparativo mensal**: propostas × realizadas × canceladas. Passe o mouse em um mês para ver o que foi realizado e por quem. Ao lado, o % realizado do ano e até o momento.
5. **Responsáveis pelas atividades**: tarefas de cada pessoa por situação + busca de atividades com filtro (realizadas, canceladas, pendentes, em andamento, transferidas). Clique numa atividade para ver o motivo/observação da planilha.
6. **Resumo por responsável + pesquisa de eventos**: tabela e busca de participantes por evento.
7. **Atendimentos por mês**: atendimentos de cada mês, total acumulado e destaques do ano.

**No computador:** navegue com as setas ← → do teclado ou com os botões embaixo.
**No celular (ou janela estreita):** os mesmos 7 slides, um por tela. Deslize o dedo para o lado (ou toque nas setas/bolinhas embaixo) para passar; se o slide for maior que a tela, role para baixo dentro dele. Toque nos boxes e nos gráficos para ver os detalhes.
O botão **Baixar PDF** (canto superior direito) baixa os 7 slides em um arquivo PDF.

## Como abrir

Abra o link online acima, ou o arquivo `index.html` no navegador (Chrome ou Edge).

## Como atualizar os dados

1. Salve a planilha atualizada na pasta **Downloads** (nome começando com `Kanban calend`).
2. Clique com o botão direito em `atualizar-dados.ps1` → **Executar com o PowerShell**.
   Ou, no terminal: `.\atualizar-dados.ps1 -Arquivo "C:\caminho\da\planilha.xlsb"`
3. O arquivo `data.js` é regerado. Recarregue o `index.html`.

O script precisa do Microsoft Excel instalado.

## Regras de cálculo

- **Mês**: o bloco do mês em que a tarefa está na planilha.
- **Realizada**: status FEITO. **Cancelada**: status CANCELADO ou NÃO REALIZADO.
- Tarefas sem status, mas com 100% ou com participantes registrados, contam como realizadas.
- **Participantes**: coluna PARTICIPANTES. Se estiver vazia e a observação disser "N participantes", esse número é usado.
