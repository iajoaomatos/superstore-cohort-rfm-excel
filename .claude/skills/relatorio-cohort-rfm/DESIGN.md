# Design do relatório: vidro escuro minimalista

Dashboard minimalista de vidro fosco escuro: painéis semitransparentes sobre um fundo charcoal com brilhos ambientes suaves.

## Fundo
- Charcoal escuro (`#141518`).
- Luz ambiente: gradientes radiais desfocados e coloridos (verde-azulado, violeta e âmbar, baixa opacidade) por trás dos elementos, fixos, para dar profundidade.

## Cartões de vidro
- Fundo translúcido de baixa opacidade (7–20% branco).
- `backdrop-filter: blur(20px+)`.
- Borda de 1 px semitransparente clara.
- **Sem sombras pesadas**: a separação vem da luz ambiente e da borda.
- Só usar caixas de vidro para gráficos, tabelas e KPIs. **Texto e questões nunca ficam em caixas.**

## Tipografia
- Corpo, tabelas e leituras dos dados: sans-serif limpa e fina (IBM Plex Sans, pesos 200–300 nos títulos grandes e 300 no corpo), branco de alto contraste (`#ECEBE7`) ou cinzento frio suave (`#9FA3AD`) para texto secundário.
- **Questões (introdução a cada tabela ou gráfico)**: Playfair Display (ou semelhante), **itálico**, creme claro amarelado (`#F4EBCB`) para máxima legibilidade. **Sem caixas**: a questão fala por si.
- Análise, interpretação e conclusões: sempre **depois** de cada tabela ou gráfico, em sans-serif fina, branco ou cinzento frio, fora de caixas.

## Iconografia
- Ícones de linha (SVG, traço 1,7 px, num círculo com borda fina) em vez das palavras Facto, Cálculo, Interpretação e Hipótese.
- Cores: Facto verde-menta `#6CC2A3`, Cálculo azul frio `#9DB4E0`, Interpretação amarelo claro `#F0D98A`, Hipótese coral `#F08F7C`.
- Cada ícone tem `aria-label` e `title` com o nome; a legenda aparece no cabeçalho.

## Layout
- Largura total, muito espaço negativo, navegação lateral fina e discreta (sticky).
- **Tabelas e gráficos nunca cortados**: `table-layout: fixed`, larguras em percentagem, fontes com `clamp()`, quebra de linha em textos longos; sem scroll horizontal em desktop nem em ~400 px.

## Escalas de cor
- Heatmap de cohort: vermelho `rgb(196,84,72)` (0%) → amarelo `rgb(214,176,72)` (10%) → verde `rgb(74,160,116)` (25% ou mais); mês 0 em verde-escuro `#1E6B58`.
- Segmentos RFM: verdes para os melhores (Campeões, Fiéis), amarelos para atenção, laranjas/vermelhos para risco e cinzento para perdidos.

## Referência visual dos outputs (exemplo do utilizador)
Ficheiro: `output-relatório/Projeto - Excel-CDS.pdf` (slides do exercício original).
- **Cohort**: tabela de heatmap em linhas = cohort (AAAA-MM) e colunas = meses decorridos; a coluna do mês 0 a 100% em verde; células não observáveis vazias (forma triangular quando existe); escala vermelho para verde. A vista principal cobre os cohorts de 2014 e 2015 (24 cohorts × meses 0–23); os 42 cohorts (meses 0–12) ficam num bloco expansível. Arredondamento «metade para cima», como o Excel (12,5% → 13%).
- **RFM**: treemap dos segmentos com área proporcional ao n.º de clientes e legenda completa por baixo (etiquetas dentro dos blocos só quando cabem; nunca cortadas), depois a matriz R × F+M e a tabela de ações.
- **Texto**: uma «síntese dos resultados» em prosa curta depois de cada tabela ou gráfico, a nomear máximos, mínimos e picos concretos.
- **Slides Produtos/Localizações do exemplo** (pedidos por mês, clientes únicos por mês, pedidos por cliente, itens por pedido, tipo de envio) usam gráficos que **não** fazem parte do âmbito atual do relatório; só entram se o utilizador o pedir.
