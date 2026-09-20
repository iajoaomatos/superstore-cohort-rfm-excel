# SuperStore: retenção, segmentação de clientes e desempenho comercial (Excel)

Projeto do Curso de Analista de Dados: Excel para Analista de Dados - CDS.

Análise de Cohort, RFM e desempenho de produtos e localizações sobre 4 anos de encomendas de uma rede de supermercados (janeiro de 2014 a dezembro de 2017), resolvida em Excel e validada em Python.

## Contexto

A SuperStore, uma das maiores redes de supermercados do país, enfrenta desafios de retenção de clientes e de melhoria das suas estratégias de vendas. Tem dados ricos sobre clientes, pedidos, localizações e produtos, mas precisa de uma análise detalhada para decidir com base em dados. A gestão identificou três áreas de foco (enunciado em [objetivo-projeto-final/projeto-do-aluno.pdf](objetivo-projeto-final/projeto-do-aluno.pdf)):

1. **Retenção de clientes**: monitorizar a interação dos clientes ao longo do tempo (Análise de Cohort).
2. **Segmentação de clientes**: identificar grupos com base no comportamento de compra (RFM).
3. **Desempenho de produtos e localizações**: avaliar que produtos e lojas geram maior impacto na faturação.

## Perguntas de negócio

| Análise | Perguntas |
|---|---|
| Cohort | Qual é a retenção de clientes ao longo dos meses? Que cohorts apresentam maior retenção? |
| RFM | Quem são os «Campeões» e os «Clientes em risco»? Como se distribuem os clientes pelos segmentos? Que ações tomar para fidelizar ou recuperar? |
| Produtos e localizações | Que produtos geram mais receita e quais têm baixo desempenho? Existe relação entre o desempenho e a região? |

## Entregas

| Pedido no enunciado | Onde está |
|---|---|
| Planilha consolidada em Excel | [analises-excel/](analises-excel/): `analise_cohort.xlsx`, `analise_rfm.xlsx` e `analise_produtos_localizacoes.xlsx` |
| Relatório com principais descobertas | [relatorio/relatorio.html](relatorio/relatorio.html) |

Para ver o relatório, descarrega o repositório e abre `relatorio/relatorio.html` no browser (no GitHub vê-se apenas o código). O relatório tem uma ligação para cada um dos três ficheiros Excel.

## Principais resultados

Base: 793 clientes, 5.009 encomendas, 1.862 produtos e 2.297.201 de vendas.

- **Cohort**: a retenção média mensal dos 24 cohorts de 2014 e 2015 (meses 1 a 23) é de 11,0%. Não desce com a idade do cohort: a média dos meses 1 a 12 é de 9,7% e a dos meses 13 a 23 é de 12,4%. O melhor cohort é 2015-09, com pico de 44% no mês 11.
- **RFM**: Fiéis e Fiéis com potencial somam 67% dos clientes e 72% da receita. Os 30 Campeões (3,8% dos clientes) geram 9,4% da receita. Em risco e Não posso perder são 8,2% dos clientes e 11,7% da receita, o grupo a reativar em primeiro lugar.
- **Produtos e localizações**: Tecnologia, Mobiliário e Material de escritório pesam de forma semelhante nas vendas (36,4%, 32,3% e 31,3%), mas as margens são 17,4%, 2,5% e 17,0%. Tables tem margem negativa (−8,6%). West e East somam 61% das vendas; Central tem a menor margem das regiões.

## Metodologia (resumo)

| Análise | Como foi feita em Excel |
|---|---|
| Cohort | `MINIFS` (1.ª compra do cliente), `TEXT`, `RIGHT` e `LEFT` (meses decorridos) e tabela dinâmica com modelo de dados: contagem distinta de clientes, mostrada como % do total da linha. Escala de cor vermelho, amarelo e verde. |
| RFM | `MAXIFS` (recência), `COUNTA` + `UNIQUE` + `FILTER` (frequência, em encomendas distintas), `SUMIFS` (monetização). Notas de 1 a 5 com faixas fixas para Recência e Frequência e `PERCENTILE` para a Monetização. Segmentos com `INDEX`/`MATCH`. Tabela RFM com `COUNTIFS`. |
| Produtos e localizações | `SUMIFS` por categoria, região, estado, cidade e produto; margem = lucro ÷ vendas; ranking com `RANK`; Pareto por % acumulada; tabela dinâmica e gráficos. |

Os segmentos e as faixas RFM seguem o artigo [O que é a Matriz RFM?](https://g4business.com/blog/o-que-e-matriz-rfm), da G4 Educação.

Cada análise foi confirmada em Python sobre os CSV: os valores do Excel coincidem com os do Python (por exemplo, as 576 células da tabela de cohort e os 793 clientes do RFM).

**Decisões e correções:**
- A tabela RFM original tinha 156 linhas em branco e faltavam 155 clientes; foi refeita com os 793 clientes.
- As faixas de recência tinham lacunas nos dias 73, 146, 219 e 293; foram fechadas com `>=`.
- A frequência conta encomendas distintas, porque cada linha de `orders` é um item de encomenda.
- Não existe «loja» nos dados: o ranking de localizações é por região, estado e cidade.
- A análise usa todo o histórico de 2014 a 2017 (o artigo da G4 sugere uma janela de 12 meses).

## Estrutura do repositório

```
base-dados/                 CSV originais: orders, customers, product, location
analises-excel/             Excel finais (Cohort, RFM, Produtos e Localizações)
relatorio/relatorio.html    Relatório
analise/                    Scripts Python de validação e de geração do relatório
objetivo-projeto-final/     Enunciado do projeto
.claude/skills/             Skill usada para gerar o relatório
claude.md                   Instruções e decisões do projeto
```

## Como reproduzir

Requisitos: Python 3 com `pandas`, `numpy` e `openpyxl` (`pip install pandas numpy openpyxl`). Correr a partir da raiz do repositório:

1. `python analise/01_calculos.py`: calcula Cohort, RFM e rankings a partir dos CSV e cria `analise/dados/` (não versionada).
2. `python analise/03_relatorio_html.py`: gera `relatorio/relatorio.html`.
3. `python analise/05_excel_produtos.py`: gera a estrutura do Excel de Produtos e Localizações (as fórmulas são calculadas ao abrir o ficheiro no Excel).

## Notas sobre os ficheiros Excel

- Os cálculos estão em fórmulas visíveis e a tabela dinâmica do Cohort usa o modelo de dados do Excel (pode não converter corretamente no Google Sheets).
- Os ficheiros `analise_cohort.xlsx` e `analise_rfm.xlsx` foram criados com consultas do Power Query que apontam para os CSV no computador original. Abrem e mostram os resultados em qualquer computador; «Atualizar tudo» só funciona se as consultas forem apontadas para os CSV de `base-dados/`.
- O relatório usa a fonte IBM Plex Sans (carregada do Google Fonts).
