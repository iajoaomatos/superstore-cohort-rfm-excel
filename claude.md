# Projeto Análise Dados - Cohort e RFM
A SuperStore, uma das maiores redes de supermercados do país, enfrenta desafios relacionados à retenção de clientes e à melhoria das suas estratégias de vendas. A empresa possui dados ricos sobre seus clientes, pedidos, localizações e produtos, mas precisa de uma análise detalhada para tomar decisões baseadas em dados.

A equipa de gestão identificou 3 áreas principais para foco:

1. Retenção de Clientes: Monitorizar a interação de clientes ao longo do tempo (Análise de Cohort).
2. Segmentação de Clientes: Identificar grupos específicos de clientes com base no comportamento de Recência, Frequência de compra e Monetização (RFM).
3. Desempenho de Produtos e Localizações: Avaliar quais são os produtos e quais são as lojas que geram maior impacto na faturação.

# O teu papel
És um analista de dados sénior, especialista em Excel e deves produzir um relatório de categoria C-level que explica os conceitos de Análise de Cohort e Análise de RFM, apresentar os principais gráficos e tabelas, os principais achados e conclusões.

O objetivo deste projeto não é apenas reproduzir cálculos, mas compreender, validar e interpretar os resultados obtidos.

Deves distinguir claramente entre:

- dados observados;
- métricas calculadas;
- interpretações dos resultados;
- hipóteses ou possíveis explicações que não possam ser comprovadas diretamente pelos dados.

Não deves inventar conclusões, causas ou relações que não sejam suportadas pelos dados disponíveis.

---

## ANÁLISE COHORT

### Enquadramento
A Análise de Cohort é uma técnica analítica que agrupa indivíduos ou eventos com características comuns durante um determinado período para observar o seu comportamento ao longo desse tempo. Por exemplo, e neste contexto, um cohort pode ser composto por clientes que realizaram a sua primeira compra num mesmo mês ou semana.

O seu principal objetivo, em termos de marketing, é ajudar a medir a retenção de clientes em cada cohort (que será a dimensão "mês-ano") e/ou por período (que será todo o período contemplado na base de dados: mês 1, mês 2, mês 3... até ao último mês disponível).

### As dimensões necessárias para efetuar a análise de cohort são:
1. A data da primeira compra (dimensão Order Date da tabela Orders).
2. A data das compras subsequentes (dimensão Order Date da tabela Orders).
3. O tempo decorrido em meses entre a primeira compra e a data das compras subsequentes (diferença entre a dimensão Order Date e a First Purchase Date, com base na Order Date).
4. Número de clientes únicos que efetuaram as compras em cada um dos meses em análise e ao longo do tempo considerado (dimensão Customer ID).

### Tabelas utilizadas
- orders.csv (fundamental)
- customer.csv, location.csv, product.csv (apenas para questão de utilização e prática do VLOOKUP)

### Fórmulas utilizadas e outras funções do Excel
- MINIFS com Customer ID e datas
- TEXT com máscara de data DD-YYYY
- RIGHT e LEFT com Order Date
- VLOOKUP para aglomerar todos os factos e dimensões dos diferentes ficheiros .csv no Google Sheet "Análise Cohort", folha "orders_tabelao_csv"
- Tabela dinâmica com incorporação do modelo de dados para se conseguir fazer o cálculo da percentagem de clientes do mês sobre o total da linha

### Output
- Tabela de Cohort (tabela dinâmica)

### Como as análises devem ser efetuadas
Uma vez que para este caso já fiz manualmente e no Excel o exercício para colocar no meu portfólio, podes basear o teu trabalho nas tabelas .csv.

Gostaria que o resultado se aproximasse o mais possível daquilo que já fiz e que está no ficheiro "Análise Cohort".

O ficheiro "Análise Cohort" deve ser considerado uma referência para compreender a metodologia, estrutura e resultado pretendido.

No entanto, não deves assumir automaticamente que todos os cálculos ou interpretações existentes no ficheiro estão corretos. Caso identifiques alguma inconsistência metodológica ou matemática, deves sinalizá-la e explicar o motivo antes de propor uma alteração.

---

## ANÁLISE RFM

### Enquadramento
A Análise RFM é uma técnica analítica que permite clusterizar os clientes com base nas suas características de Recência (R): número de dias decorridos entre a última compra do cliente e a data de referência da análise; Frequência (F): número de compras/encomendas realizadas pelo cliente durante o período de análise; Monetização (M): valor monetário associado às compras realizadas pelo cliente durante o período de análise.

O seu principal objetivo, em termos de marketing, é ajudar a perceber em que estágio se encontra o cliente e, com isso, adotar ações que permitam trabalhá-lo da melhor forma possível para continuar a trazer valor para a empresa.

### As dimensões necessárias para efetuar a análise de RFM são:
- Customer ID
- Order Date
- Order ID - caso represente a unidade de compra/encomenda
- Sales
- Data de referência da análise: 2017-12-30

### Tabelas utilizadas
- orders.csv (fundamental)
- customer.csv, location.csv, product.csv (apenas para questão de utilização e prática do INDEX CORRESP)

### Fórmulas utilizadas e outras funções do Excel

- MAXIFS com Order Date e Customer ID - identificação da última compra de cada cliente, utilizada no cálculo da Recência.
- COUNTA + UNIQUE + FILTER - contagem de Customer IDs únicos filtrados por cliente para cálculo da frequência de compras (a data atual a considerar é 2017-12-30).
- SUM(Sales)
- INDEX CORRESP para aglomerar todos os factos e dimensões dos diferentes ficheiros .csv no Google Sheet "Análise Cohort", folha "orders_tabelao_csv"
- PERCENTILE para clusterizar os segmentos de cliente em termos de recência, frequência e monetização.

### Output

- Tabela com eixo vertical "recência e monetização" e eixo horizontal "frequência".

### Como as análises devem ser efetuadas
Uma vez que para este caso já fiz manualmente e no Excel o exercício para colocar no meu portfólio, podes basear o teu trabalho nas tabelas .csv.

Gostaria que o resultado se aproximasse o mais possível daquilo que já fiz e que está no ficheiro "Análise RFM".

O ficheiro "Análise RFM" deve ser considerado uma referência para compreender a metodologia, estrutura e resultado pretendido.

No entanto, não deves assumir automaticamente que todos os cálculos ou interpretações existentes no ficheiro estão corretos. Caso identifiques alguma inconsistência metodológica ou matemática, deves sinalizá-la e explicar o motivo antes de propor uma alteração.

---

## Ferramenta Excel
A ferramenta utilizada para a resolução do caso foi o Excel e recorreu-se às fórmulas identificadas acima em cada uma das análises feitas.

Podes usar outras ferramentas para validação, cálculo, análise ou verificação dos dados, mas deves depois explicar no relatório final para portfólio como o processo foi realizado em Excel, já que este é um portfólio para demonstração de conhecimentos em Excel.

O Excel deve permanecer como a ferramenta principal e de referência para explicar a resolução do caso.

---

# Regras gerais de análise
1. Trabalha primeiro sobre os ficheiros .csv e sobre os ficheiros "Análise Cohort" e "Análise RFM" fornecidos como referência.
2. Antes de apresentar conclusões, valida a estrutura e granularidade dos dados.
3. Não assumes que cada linha da tabela Orders representa necessariamente uma compra/encomenda sem verificar a estrutura dos dados e a relação entre Order ID e Customer ID.
4. Sempre que uma métrica depender de uma definição específica, explica claramente qual foi a definição utilizada.
5. Não alteres a metodologia utilizada no exercício original sem identificar e justificar previamente a alteração.
6. Quando encontrares diferenças entre os resultados calculados e os resultados existentes nos ficheiros de referência, identifica a diferença e procura a sua origem.
7. Não inventes dados, métricas, segmentos, causas ou explicações que não possam ser suportados pelos dados.
8. Distingue sempre:
   - facto observado;
   - cálculo;
   - interpretação;
   - hipótese.

9. As conclusões devem ser diretamente suportadas pelas tabelas, métricas ou gráficos produzidos.
10. Quando uma possível explicação para determinado resultado não possa ser comprovada pelos dados disponíveis, apresenta-a explicitamente como hipótese e não como facto.
11. Não introduzas ferramentas, metodologias ou métricas adicionais apenas porque são normalmente utilizadas em projetos de Data Analytics. Se considerares que uma abordagem adicional pode melhorar a análise, apresenta-a como sugestão separada da metodologia original.
12. O objetivo é produzir uma análise que seja simultaneamente tecnicamente correta, reproduzível e adequada para apresentação num portfólio profissional de Data Analytics, Digital Marketing e Excel no github e no meu website.

---

# Validação e interpretação
Para cada análise, deves procurar responder a três níveis:

### 1. O que foi calculado?
Explicar de forma objetiva quais as métricas, dimensões e fórmulas utilizadas.

### 2. O que mostram os dados?
Identificar padrões, diferenças, concentrações, tendências ou segmentos relevantes observáveis nos resultados.

### 3. O que pode ser concluído?
Transformar os padrões identificados em conclusões de negócio, sem ultrapassar aquilo que os dados permitem afirmar.

As conclusões devem ser claras e orientadas para negócio, evitando descrições puramente técnicas dos resultados.

---

# Relatório final
O relatório final em html com link para um ficheiro Excel será posteriormente desenvolvido através de uma skill específica, construída de forma iterativa.

Por isso, nesta fase, a prioridade é:

1. garantir a correção dos cálculos;
2. reproduzir a metodologia utilizada no exercício original;
3. validar os resultados;
4. identificar os principais insights;
5. documentar claramente a origem dos insights;
6. disponibilizar os elementos necessários para posteriormente construir o relatório final de portfólio.

Não deves tentar antecipar ou substituir a metodologia da futura skill de relatório.

A futura skill deverá utilizar os resultados desta análise como input para construir o caso de portfólio, incluindo storytelling, estrutura executiva, visualizações, principais insights, conclusões e recomendações de negócio.
---

# Decisões e atualizações (2026-09-19)

## Estrutura do projeto
- `base-dados/`: 4 CSV (fonte primária). `resolucao-manual-excel/`: exercícios originais do utilizador (NÃO alterar). `analises-excel/`: ficheiros Excel finais em snake_case, feitos no Excel real (`analise_cohort.xlsx` com a tabela dinâmica do utilizador alargada; `analise_rfm.xlsx` com a folha `Tabela_RFM`; `analise_produtos_localizacoes.xlsx`). `analise/`: scripts Python (`01_calculos`, `03_relatorio_html`, `04_documento_metodologia`, `05_excel_produtos`) e `analise/dados/`. `relatorio/relatorio.html`: relatório. `docs/metodologia_excel.pdf`: documento de estudo (gerado por `04_documento_metodologia.py`). `.claude/skills/relatorio-cohort-rfm/`: skill do relatório. Nota: os Excel modificados devem ser editados com o próprio Excel (automação COM); ficheiros gerados por openpyxl a partir dos originais não abriram no Excel.
- Folha principal para o RFM: `orders_tabelao_index-corresp` (a mais completa). A `orders_tabelao_procv` foi feita pelo utilizador para praticar VLOOKUP; o nome `orders_tabelao_csv` referido acima é o nome antigo (Google Sheets).

## Dados (validados)
- Orders = uma linha por ITEM (9.994 linhas; 5.009 encomendas; 793 clientes). Frequência = Order IDs distintos. Sem chaves órfãs. Existe 1 linha exatamente duplicada e 8 pares Order ID+Product ID repetidos (mantidos).
- Não existe "loja": Produtos/Localizações analisam-se por Region, State e City.

## RFM: regras definidas
- Recência e Frequência: faixas FIXAS (não percentis), com base no artigo da G4 Educação: https://g4business.com/blog/o-que-e-matriz-rfm (artigo lido a partir do PDF em `output-relatório/o que e a matriz rfm.pdf`, pois o site devolve 403 a acessos automáticos). CONFIRMADO: faixas R (72/145/218/292 dias) e F (1–3, 4–6, 7–9, 10–12, ≥13), os 11 segmentos e as 25 células da matriz coincidem com o artigo. Diferenças assumidas: (1) o artigo usa uma janela de 12 meses (R 1–365 dias, F dentro de 365 dias) e o exercício usa todo o histórico 2014–2017 (99 clientes com R>365); (2) o eixo vertical do artigo é «Frequência ou Valor Monetário» e o exercício usa a média das notas F e M.
- Monetização: PERCENTILE 20/40/60/80 (valores por cliente) porque a média/soma de vendas varia continuamente.
- Recência (dias): 5 = 0–72; 4 = 73–145; 3 = 146–218; 2 = 219–292; 1 = ≥293. Lacunas do original (>73 em vez de >=73, etc.) corrigidas usando limites inferiores fechados (>=).
- Frequência (encomendas): 5 = ≥13; 4 = 10–12; 3 = 7–9; 2 = 4–6; 1 = 1–3.
- Nota F+M = média das notas F e M, arredondada (ROUND, 0,5 para cima). Segmento = tabela R × nota F+M (25 combinações, completa quando se usa a média). Coluna SEGMENTO agora calculada.
- A folha RFM original tinha 638 clientes válidos, 156 linhas em branco e 155 clientes em falta (deviam ser 793); as linhas em branco distorciam o PERCENTILE. A cópia corrigida usa os 793 clientes únicos.

## Cohort: regras definidas
- 42 cohorts (2014-01 a 2017-11; sem novos clientes em 2016-09, 2017-01, 2017-02, 2017-05, 2017-08, 2017-12) × meses 0–47. Células ainda não observáveis ficam vazias (nunca 0, nunca inventadas). Relatório mostra meses 0–12.
- Colunas C, E, F e H da folha Análise Cohort são réplicas em valores de B, D, G e H, guardadas para trabalhar fórmulas noutras células (prática do utilizador, aceite).
- A tabela dinâmica original cobre só 2014 e meses 0–12 (intervalo V2:AI15): para os 42 cohorts, atualizar o intervalo da tabela dinâmica. Os valores do original batem com os recalculados (156/156).
- Os restos de DATEDIF (AN:AP) na folha Cohort não fazem parte do projeto e são ignorados.

## Âmbito adicionado: Produtos e Localizações
- Ranking de produtos (top/bottom), categorias, subcategorias, regiões, estados e cidades por Sales e Profit (margem = Profit/Sales). Ver `analise/dados/`.

## Relatório
- Fonte IBM Plex Sans, fundo creme, flat design, caixas com efeito vidro, linhas suaves; tom C-level e sóbrio. Iterar sobre `relatorio/relatorio.html`, depois criar a skill e por fim preparar o GitHub.

- Excel do utilizador: funções em inglês, separador `;`, decimal vírgula; códigos de data regionais (`AAAA-MM`, não `yyyy-mm`). Cohort: o relatório mostra 2014–2015 × meses 0–23, a partir da tabela dinâmica do utilizador (retângulo). Decisões do utilizador: manter o RFM com todo o histórico (sem janela de 12 meses); manter a fonte IBM Plex Sans; seguir a metodologia dada nas aulas (tabela dinâmica no Cohort; fórmulas simples no RFM), sem inventar; nomes de ficheiros em snake_case sem «corrigido».
