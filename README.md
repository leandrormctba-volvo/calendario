# Calendário de Eventos 2026

Dashboard dos eventos do calendário 2026, gerado a partir da planilha **Kanban calendário** (aba *Responsabilidades*).

## O que o dashboard mostra

- **Comparativo mensal**: tarefas propostas × realizadas × canceladas em cada mês.
- **Tarefas por responsável**: quantas tarefas cada pessoa tem, separadas por situação.
- **% realizado do ano**: tarefas realizadas ÷ tarefas propostas no ano (e até o mês atual).
- **Atendimentos**: soma de participantes de todos os eventos, e por mês.
- **Pesquisa por evento**: digite o nome do evento para ver quantos participantes teve.
- Filtros por mês, responsável e categoria.

## Como abrir

Abra o arquivo `index.html` no navegador (Chrome ou Edge).

## Como atualizar os dados

1. Salve a planilha atualizada na pasta **Downloads** (nome começando com `Kanban calend`).
2. Clique com o botão direito em `atualizar-dados.ps1` → **Executar com o PowerShell**.
   Ou, no terminal: `.\atualizar-dados.ps1 -Arquivo "C:\caminho\da\planilha.xlsb"`
3. O arquivo `data.js` é regerado. Recarregue o `index.html`.

O script precisa do Microsoft Excel instalado.

## Regras de cálculo

- **Mês**: o bloco do mês em que a tarefa está na planilha.
- **Realizada**: status FEITO. **Cancelada**: status CANCELADO ou NÃO REALIZADO.
- Tarefas sem status, mas com 100% ou com participantes registrados, contam como realizadas e aparecem com `*`.
- **Participantes**: coluna PARTICIPANTES. Se estiver vazia e a observação disser "N participantes", esse número é usado.
