"""Gera docs/metodologia_excel.md: documento de estudo (o que é cada análise, factos e dimensões, fórmulas, passos no Excel,
tabelas, gráficos, interpretação com exemplo trabalhado). Todos os números vêm de analise/dados/ e dos CSV."""
import json, os
os.makedirs("analise/dados", exist_ok=True)
import pandas as pd

R = json.load(open("analise/dados/resultados.json", encoding="utf-8"))
C, F = R["cohort"], R["rfm"]
n = lambda x: f"{x:,.0f}".replace(",", ".")
pc = lambda x, d=1: f"{x*100:.{d}f}".replace(".", ",") + "%"
o = pd.read_csv("base-dados/orders.csv")
o["d"] = pd.to_datetime(o["Order Date"], format="%m/%d/%Y")
o["m"] = o.d.dt.to_period("M")

# ---- números de apoio ----
act = o.groupby("m")["Customer ID"].nunique()
cal = act.groupby(act.index.month).mean().round(0).astype(int)
vm = o.groupby(o.d.dt.month).Sales.sum()
share3 = vm.loc[[9, 11, 12]].sum() / vm.sum()
novos = o.groupby("Customer ID").d.min().dt.year.value_counts().sort_index()
fortes = sorted([x for x in C["media_1a12"] if x[1] >= 20], key=lambda x: -x[2])
seg = {s["segmento"]: s for s in F["segmentos"]}
c = C["curva_pond"]
q = F["percentis_M"]
tot = R["total_sales"]
risco = seg["Em risco"]["receita"] + seg["Não posso perder"]["receita"]
cnt = pd.read_csv("analise/dados/cohort_clientes.csv", index_col=0)
pct = pd.read_csv("analise/dados/cohort_retencao_pct.csv", index_col=0)
ex_n, ex_k, ex_v = int(cnt.loc["2014-01", "0"]), 10, int(cnt.loc["2014-01", "10"])
cat = {r["Category"]: r for r in R["cat"]}
sub = {r["Sub-Category"]: r for r in R["sub"]}
mg = lambda r: r["lucro"] / r["vendas"]

co24 = [c for c in C["cohorts"] if c["m"] < "2016"]
media_c = {c["m"]: sum(c["r"][1:24]) / 23 for c in co24}
melhor = max(co24, key=lambda c: media_c[c["m"]])
pior = min(co24, key=lambda c: media_c[c["m"]])
media_geral = sum(media_c.values()) / len(co24)
vert = [sum(c["r"][k] for c in co24) / len(co24) for k in range(24)]
kmax = max(range(1, 24), key=lambda k: vert[k])
kmin = min(range(1, 24), key=lambda k: vert[k])
h1, h2 = sum(vert[1:13]) / 12, sum(vert[13:24]) / 11
pico = max(melhor["r"][1:24])
kpico = melhor["r"].index(pico)

L_seg = "\n".join(f"| {s['segmento']} | {n(s['clientes'])} | {pc(s['pct_clientes'])} | {pc(s['pct_receita'])} | {s['R_med']:.0f} | {s['F_med']:.1f} |" for s in F["segmentos"])
L_cat = "\n".join(f"| {r['Category']} | {n(r['vendas'])} | {pc(r['pct'])} | {n(r['lucro'])} | {pc(mg(r))} |" for r in R["cat"])
L_sub = "\n".join(f"| {r['Sub-Category']} | {n(r['vendas'])} | {pc(r['pct'])} | {n(r['lucro'])} | {pc(mg(r))} |" for r in R["sub"])
L_reg = "\n".join(f"| {r['Region']} | {n(r['vendas'])} | {pc(r['pct'])} | {n(r['lucro'])} | {pc(mg(r))} |" for r in R["regiao"])
L_est = "\n".join(f"| {r['State']} | {n(r['vendas'])} | {pc(r['pct'])} |" for r in R["estado"])
L_ano = "\n".join(f"| {r['Ano']} | {n(r['vendas'])} | {pc(r['pct'])} |" for r in sorted(R["ano"], key=lambda r: r["Ano"]))

doc = f"""# Metodologia e estudo de apoio — SuperStore (Cohort, RFM, Produtos e Localizações)

Documento de estudo para perceber, **para cada tipo de análise**: o que é, que factos e dimensões são precisos, como se calcula cada um em Excel (fórmulas), como se cria a tabela e o gráfico, como se interpreta (com um exemplo trabalhado) e que pressupostos e limitações existem. O exercício foi feito manualmente em Excel; o Python (`analise/`) serviu para validar e reproduzir.

## 0. Como usar este documento

**Tipos de afirmação** (os mesmos ícones do relatório):
- **Facto**: observado diretamente nos dados.
- **Cálculo**: métrica derivada dos dados por uma fórmula.
- **Interpretação**: conclusão de negócio que os números suportam.
- **Hipótese**: explicação possível que os dados **não** conseguem confirmar.

**Facto vs dimensão** (vocabulário de análise de dados, usado neste documento):
- **Medida (o «facto» numérico)**: o que se soma, conta ou mede. Exemplos: `Sales`, `Profit`, nº de clientes distintos, nº de encomendas.
- **Dimensão**: o atributo pelo qual se agrupa ou filtra. Exemplos: mês, cliente, categoria, região.
- Uma análise é sempre «medida × dimensão(ões)»: por exemplo, «vendas (medida) por categoria (dimensão)» ou «clientes distintos (medida) por cohort e por mês decorrido (dimensões)».

**Ficheiros.**

| Análise | Ficheiro Excel | O que tem |
|---|---|---|
| Cohort | `analises-excel/analise_cohort.xlsx` | O teu ficheiro do exercício com a **tua tabela dinâmica** (contagem distinta, modelo de dados) alargada aos cohorts de 2014 e 2015 e aos meses 0–23, com a escala de cor e as médias por linha e por coluna. |
| RFM | `analises-excel/analise_rfm.xlsx` | O teu ficheiro do exercício (folha `RFM`, com as correções da secção 4.9) + a folha `Tabela_RFM` (tabela RFM construída). |
| Produtos e Localizações | `analises-excel/analise_produtos_localizacoes.xlsx` | Novo: `dados`, `Resumo`, `Categorias`, `Regioes`, `Cidades`, `Produtos`, `Top_Bottom`, `Anos`, `Tabela_dinamica`, com fórmulas e gráficos. |
| Originais | `resolucao-manual-excel/` | Os teus ficheiros originais, **não alterados**. |

**O teu Excel.** Os nomes das funções estão em **inglês**, o separador de argumentos é **`;`** e o decimal é **vírgula** (por exemplo `PERCENTILE(E:E;0,2)`). Os **códigos de formato de data** seguem o idioma regional: em português o ano é `AAAA` (por isso `TEXT(C2;"MM-AAAA")`), e a formatação de célula `aaaa-mm` mostra 2014-01. Ao passar para outro idioma ou para o Google Sheets, estes formatos podem ter de mudar (`yyyy`).

---

## 1. Dados e validação

**Tabelas.** `orders.csv` (facto), `customers.csv`, `product.csv` e `location.csv` (dimensões). Juntam-se com `VLOOKUP` (folha `orders_tabelao_procv`) ou `INDEX`/`MATCH` (folha `orders_tabelao_index-corresp`, a mais completa):
- `=VLOOKUP(E2;customers!A:C;2;FALSE)`: procura o Customer ID da linha na 1.ª coluna de `customers` e devolve a 2.ª coluna (nome).
- `=INDEX(product!C:C;MATCH(F2;product!A:A;0))`: `MATCH` acha a linha do Product ID; `INDEX` devolve a coluna pretendida dessa linha. É mais flexível que `VLOOKUP` (não exige que a chave seja a 1.ª coluna).

**Granularidade (o ponto mais importante).** Facto: `orders.csv` tem {n(len(o))} linhas mas só {n(o['Order ID'].nunique())} encomendas. Cada linha é **um item de uma encomenda** (média de 2 linhas por encomenda; máximo de 14). Por isso:
- **Frequência** = nº de Order IDs **distintos** por cliente (nunca o nº de linhas).
- **Clientes ativos num mês** = Customer IDs **distintos** nesse mês.
- Somar `Sales` e `Profit` é seguro (cada linha é um item com o seu valor).

**Verificações feitas** (Python sobre os CSV; em Excel: `COUNTA`, `UNIQUE`, `COUNTIFS`):

| Verificação | Resultado |
|---|---|
| Período das datas | 2014-01-03 a 2017-12-30 |
| Clientes / encomendas / produtos | {n(o['Customer ID'].nunique())} / {n(o['Order ID'].nunique())} / {n(o['Product ID'].nunique())} |
| Encomendas com mais de 1 cliente ou mais de 1 data | 0 |
| Chaves sem correspondência em customers, product ou location | 0 |
| Valores em falta | 0 |
| Encomendas por cliente | mínimo 1, média 6,3, máximo 17 |
| Linha totalmente duplicada | {R['dup_linha_exata']} (mantida) |
| Pares Order ID + Product ID repetidos | 8 (mantidos: podem ser itens legítimos; impacto residual) |
| Vendas totais | {n(tot)} |

**Pressupostos.** Os dados não têm «loja»: a localização é por encomenda (`City`, `State`, `Region`), por isso o «ranking de lojas» é por região, estado e cidade. A moeda não está indicada, por isso os valores aparecem sem símbolo.

---

## 2. Método geral de interpretação (vale para as três análises)

1. **Observar**: descrever o que a tabela ou gráfico mostra, sem explicar (o maior, o menor, a forma).
2. **Quantificar**: pôr números (%, contagens, diferenças). Uma conclusão sem número não é verificável.
3. **Comparar**: com um ponto de referência (a média, outro grupo, o período anterior).
4. **Concluir para o negócio** (**interpretação**): dizer o que isso significa para a decisão, sem ir além do que os números mostram.
5. **Separar as hipóteses**: se propõe uma causa, escreve-a como **hipótese** e diz que dados a confirmariam.
6. **Verificar o tamanho da amostra**: grupos pequenos (poucos clientes) não sustentam conclusões.

Cada análise abaixo termina com um **exemplo trabalhado** que aplica estes 6 passos.

---

## 3. Análise de Cohort

### 3.1 O que é e quando se usa
Agrupa os clientes pelo **mês da primeira compra** (o *cohort*) e mede, mês a mês, que percentagem desse grupo voltou a comprar. Serve para medir **retenção** e comparar grupos de clientes de épocas diferentes. Responde a: qual é a retenção ao longo dos meses, que cohorts retêm mais, e se há sazonalidade.

### 3.2 Medidas e dimensões necessárias
| Campo | Tipo | Tabela de origem | Para que serve |
|---|---|---|---|
| Customer ID | Dimensão | orders | Identifica o cliente (base da contagem distinta) |
| Order Date | Dimensão (data) | orders | Data de cada compra |
| Mês da 1.ª compra (cohort) | Dimensão derivada | calculada | Define a que grupo o cliente pertence |
| Mês da atividade | Dimensão derivada | calculada | Mês em que o cliente comprou |
| Meses decorridos | Dimensão derivada | calculada | Distância em meses entre a 1.ª compra e a compra da linha (0, 1, 2, …) |
| Clientes distintos | **Medida** | calculada | Nº de clientes diferentes com compras em cada célula |
| % do cohort | **Medida** | calculada | Clientes ativos no mês ÷ clientes do cohort (mês 0) |

### 3.3 Como se calcula cada campo em Excel (folha `Analise Cohort`)
| Campo | Fórmula | Como funciona |
|---|---|---|
| 1.ª compra do cliente | `=MINIFS(orders[Order Date];A:A;A2)` | Data mais antiga de `Order Date` entre as linhas cujo Customer ID é o da linha atual |
| Cohort (texto) | `=TEXT(C2;"MM-AAAA")` | Converte a data em texto `"10-2015"`. `MM` = mês com 2 dígitos, `AAAA` = ano com 4 (em Excel português). Não depende só da função: depende do idioma regional |
| Mês da atividade (texto) | `=TEXT(F2;"MM-AAAA")` | Mesma lógica, sobre a data da compra |
| Meses decorridos | `=(RIGHT(H2;4)-RIGHT(E2;4))*12+(LEFT(H2;2)-LEFT(E2;2))` | `RIGHT` tira o ano (4 últimos caracteres) e `LEFT` o mês (2 primeiros) do texto; diferença de anos × 12 + diferença de meses |
| Alternativa sem idioma | `=(YEAR(F2)-YEAR(C2))*12+MONTH(F2)-MONTH(C2)` | Mesmo resultado, diretamente sobre as datas; não depende do idioma |
| Colunas C, E, F, H | cópias em valores de B, D, G, H | Guardadas para poderes usar fórmulas noutras células sem alterar as originais |

### 3.4 Como se cria a tabela (a tua tabela dinâmica)
1. Preparar a «TABELA LIMPA» (colunas M:P da folha `Analise Cohort`): Customer ID, cohort (data do 1.º dia do mês), mês da atividade e meses decorridos.
2. **Inserir > Tabela Dinâmica > Adicionar estes dados ao Modelo de Dados**.
3. Campos: **Linhas** = mês da 1.ª compra (First purchase Date); **Colunas** = meses decorridos (Time passed); **Valores** = Customer ID.
4. Nas definições do campo de valor: **Resumir valores por > Contagem Distinta** (só existe com o modelo de dados).
5. **Mostrar valores como > % do total da linha**. Num modelo de dados, o total de uma contagem distinta é o nº de clientes distintos da linha, ou seja, o tamanho do cohort (o mesmo número do mês 0). Por isso cada célula é «clientes do mês ÷ clientes do cohort» e o mês 0 é sempre 100%.
6. **Formato dos valores**: percentagem com 1 decimal. Nas opções da tabela dinâmica (Esquema e Formato), «Para células vazias mostrar» = 0: uma combinação cohort × mês sem clientes fica vazia e aparece como 0,0%.
7. **Quais cohorts e meses mostrar**: nos filtros de «Rótulos de Linha» e «Rótulos de Coluna» ficam selecionados apenas 2014-01 a 2015-12 e os meses 0 a 23. Na tua versão original estavam 2014 (12 cohorts) e os meses 0 a 12; a tabela passou de 12 × 13 para 24 × 24 e ocupa `V2:AT27`.
   - Para a tabela poder crescer tive de apagar três células soltas (`AN2:AP2`, restos de um `DATEDIF`) que estavam no caminho.
8. **Médias fora da tabela dinâmica** (fórmulas simples, com uma coluna e uma linha de folga):
   - Média de cada cohort (leitura horizontal), na coluna `AV`: `=AVERAGE(X4:AT4)` (meses 1 a 23), copiada para as 24 linhas.
   - Média de cada mês (leitura vertical), na linha 29: `=AVERAGE(W4:W27)`, copiada para as colunas dos meses.
9. Atualizar a tabela (Dados > Atualizar Tudo) recalcula os valores a partir do modelo de dados.

### 3.5 Como se cria o heatmap (escala de cor)
1. Selecionar os valores dos meses 1 a 23 (sem a coluna do mês 0).
2. **Base > Formatação Condicional > Escalas de Cor > Escala de 3 cores > Mais regras**.
3. Mínimo: tipo **Número**, valor 0, cor vermelha. Ponto médio: tipo **Número**, valor 0,1, cor amarela. Máximo: tipo **Número**, valor 0,25, cor verde.
4. A coluna do mês 0 (sempre 100%) formata-se à parte, a verde-escuro, para não «esmagar» a escala.
5. Usar valores fixos (e não «mínimo/máximo») mantém as cores comparáveis entre tabelas.

### 3.6 Como se lê
- Linha = cohort; coluna *k* = *k* meses depois da 1.ª compra.
- **Mês 0 é sempre 100%** (por definição: são os clientes que criam o cohort).
- Cada célula é o % do cohort com **pelo menos uma compra** nesse mês. **Não é acumulada.**
- **0%** significa que o mês já passou e ninguém do cohort comprou.
- **Leitura horizontal** (ao longo de uma linha): como se comporta **um cohort** ao longo do tempo; a média da linha resume-o.
- **Leitura vertical** (ao longo de uma coluna): como se comportam **todos os cohorts** no mesmo mês de vida; a média da coluna resume-o.
- **Para calcular médias só se comparam cohorts completos**: aqueles que têm todos os meses da janela observados. No quadro de 2014–2015, os 24 cohorts têm os meses 1 a 23 observados (os dados vão até dez/2017), por isso todas as médias comparam cohorts completos.

### 3.7 Exemplo trabalhado de interpretação
*Pergunta: «qual é a retenção ao longo dos meses e que cohorts retêm mais?»* (só com o que está no quadro)
1. **Observar**: as células variam de 0% a mais de 40%, sem um padrão limpo de descida ao longo das linhas.
2. **Leitura horizontal (média de cada cohort, meses 1 a 23)**: vai de {pc(media_c[pior['m']])} ({pior['m']}, N = {pior['n']}) a {pc(media_c[melhor['m']])} ({melhor['m']}, N = {melhor['n']}); a média dos 24 cohorts é {pc(media_geral)}.
3. **Leitura vertical (média de cada mês, 24 cohorts)**: vai de {pc(vert[kmin])} no mês {kmin} a {pc(vert[kmax])} no mês {kmax}; a média dos meses 1–12 é {pc(h1)} e a dos meses 13–23 é {pc(h2)}.
4. **Exemplo do melhor cohort**: {melhor['m']} (N = {melhor['n']}) tem pico de {round(pico * 100)}% no mês {kpico}, ou seja {round(pico * melhor['n'])} dos {melhor['n']} clientes.
5. **Concluir** (interpretação): a retenção mensal não desce com a idade do cohort; fica perto de {pc(media_geral)} e é mais alta nos meses 13–23 do que nos meses 1–12.
6. **Amostra**: os cohorts de 2015 têm 7 a 18 clientes; um cliente vale 6% a 14% de uma célula. Por isso os melhores cohorts médios (2015) são pequenos e as diferenças entre eles devem ser lidas com cautela.

*O que fica de fora*: afirmações sobre sazonalidade do calendário, número de clientes por ano ou causas (Black Friday, Natal) não se leem neste quadro; seriam outra análise, com outros dados na tabela, ou hipóteses.

### 3.8 Porque é que o PDF é um triângulo e o quadro é um retângulo
Os valores coincidem onde ambos têm célula (conferi a linha 2014-01 inteira e as células de 2015-10 e 2015-11). A diferença é a **janela**: os dados vão a dez/2017, por isso os cohorts de 2014 e 2015 têm 24 meses completos (retângulo). O PDF limita a atividade a dez/2015, daí o triângulo. **Não é erro**: é uma escolha de janela.
- Numa tabela dinâmica com modelo de dados, um mês **ainda não observado** e um mês **observado sem clientes** ficam ambos vazios (e aparecem como 0,0%). No retângulo não há ambiguidade, porque todas as células foram observadas; num triângulo, os 0,0% da parte não observada seriam enganadores.

### 3.9 Pressupostos e limitações
- Meses de calendário, não janelas de 30 dias.
- O cohort é o mês da 1.ª compra **no histórico disponível**; clientes de 2014 podem já ter comprado antes (hipótese: não é possível saber).
- Arredondamento «metade para cima», como o Excel (12,5% → 13%).
- 42 cohorts em 48 meses: sem novos clientes em {", ".join(C['meses_sem_novos'])}.
- Cohorts pequenos (N ≤ 3: 11 dos 42) oscilam demasiado e não se comparam.
- A medida é o **% de clientes**, não a receita.

### 3.10 Origem de cada conclusão
| Conclusão do relatório | Tipo | Como foi obtida |
|---|---|---|
| Média de cada cohort nos meses 1–23: de {pc(media_c[pior['m']])} ({pior['m']}) a {pc(media_c[melhor['m']])} ({melhor['m']}); média geral {pc(media_geral)} | Cálculo | `=AVERAGE(X4:AT4)` em cada linha da tabela dinâmica; média das 24 médias |
| Média de cada mês: de {pc(vert[kmin])} (mês {kmin}) a {pc(vert[kmax])} (mês {kmax}); meses 1–12 = {pc(h1)}, meses 13–23 = {pc(h2)} | Cálculo | `=AVERAGE(W4:W27)` em cada coluna; média das colunas 1–12 e 13–23 |
| O cohort {melhor['m']} (N = {melhor['n']}) tem pico de {round(pico * 100)}% no mês {kpico} ({round(pico * melhor['n'])} dos {melhor['n']} clientes) | Facto | Leitura direta da linha do cohort; clientes = % × N |
| A retenção mensal não desce com a idade do cohort | Interpretação | Comparação das médias verticais dos meses 1–12 e 13–23 |
| As diferenças entre cohorts pequenos devem ser lidas com cautela | Interpretação | N dos cohorts de 2015 (7 a 18): 1 cliente = 6% a 14% da célula |

---

## 4. Análise RFM

### 4.1 O que é e quando se usa
Classifica cada cliente por **Recência** (há quantos dias comprou), **Frequência** (quantas vezes comprou) e **Monetização** (quanto gastou), e junta as notas em **segmentos** (Campeões, Fiéis, Em risco…) para decidir que ação tomar com cada grupo. Data de referência: **30/12/2017** (`=MAX(Order Date)`).

### 4.2 Medidas e dimensões necessárias
| Campo | Tipo | Origem | Para que serve |
|---|---|---|---|
| Customer ID | Dimensão | orders | Uma linha por cliente (793) |
| Order Date | Dimensão | orders | Base da Recência |
| Order ID | Dimensão | orders | Unidade de compra (base da Frequência) |
| Sales | **Medida** | orders | Base da Monetização |
| Data de referência | Parâmetro | `=MAX(Order Date)` | Ponto a partir do qual se contam os dias |
| Recência (dias), Frequência, Monetização | **Medidas** | calculadas | As três medidas RFM |
| Notas R, F, M (1 a 5) | Dimensões derivadas | calculadas | Discretizam as medidas |
| Segmento | Dimensão derivada | tabela de segmentos | Resultado final |

### 4.3 Como se calcula cada campo em Excel (folha `RFM`)
| Campo | Fórmula | Como funciona |
|---|---|---|
| Recência (data) | `=MAXIFS(orders…!B:B;orders…!E:E;A2)` | Última data de compra do cliente |
| Recência (dias) | `=$M$1-B2` | Data de referência menos essa data |
| Frequência | `=COUNTA(UNIQUE(FILTER(orders…!A:A;orders…!E:E=A2)))` | `FILTER` seleciona os Order IDs do cliente; `UNIQUE` elimina repetições (uma encomenda tem vários itens); `COUNTA` conta |
| Monetização | `=SUMIFS(orders…!G:G;orders…!E:E;A2)` | Soma de `Sales` do cliente |
| Nota R | `=IF(C2>=293;1;IF(C2>=219;2;IF(C2>=146;3;IF(C2>=73;4;5))))` | Faixas fixas, limites inferiores fechados |
| Nota F | `=IFS(D2>=13;5;D2>=10;4;D2>=7;3;D2>=4;2;D2>=1;1)` | Faixas fixas |
| Nota M | `=IF(E2<=PERCENTILE(E:E;0,2);1;IF(E2<=PERCENTILE(E:E;0,4);2;IF(E2<=PERCENTILE(E:E;0,6);3;IF(E2<=PERCENTILE(E:E;0,8);4;5))))` | Quintis: `PERCENTILE(E:E;0,2)` é o valor abaixo do qual estão 20% dos clientes |
| Nota F+M | `=AVERAGE(G2:H2)` | Média das notas F e M |
| Segmento | `=INDEX($S$3:$S$27;MATCH(1;($O$3:$O$27=F2)*($P$3:$P$27=ROUND(I2;0));0))` | Procura na tabela de 25 linhas a combinação nota R e nota F+M arredondada |

**Porquê faixas fixas para R e F e quintis para M?** R e F têm uma escala natural (dias, nº de compras) que o artigo da G4 Educação define com faixas fixas; M é um valor contínuo sem escala natural, por isso divide-se em quintis.

### 4.4 Faixas de nota
| Nota | Recência (dias) | Frequência (encomendas) | Monetização (soma de vendas) |
|---|---|---|---|
| 5 | 0–72 | 13 ou mais | acima de {n(q[3])} |
| 4 | 73–145 | 10–12 | {n(q[2])} a {n(q[3])} |
| 3 | 146–218 | 7–9 | {n(q[1])} a {n(q[2])} |
| 2 | 219–292 | 4–6 | {n(q[0])} a {n(q[1])} |
| 1 | 293 ou mais | 1–3 | até {n(q[0])} |

As faixas de R e F **coincidem com o artigo** da G4 Educação.

### 4.5 Tabela de segmentos (a mesma do artigo, 11 segmentos)
Linhas = nota R; colunas = nota F+M arredondada.

| | F+M 1 | F+M 2 | F+M 3 | F+M 4 | F+M 5 |
|---|---|---|---|---|---|
| **R5** | Novos | Fiéis com potencial | Fiéis com potencial | Fiéis | Campeões |
| **R4** | Promissores | Fiéis com potencial | Fiéis com potencial | Fiéis | Fiéis |
| **R3** | Quase dormentes | Quase dormentes | Necessitam atenção | Fiéis | Fiéis |
| **R2** | Perdidos | Hibernando | Em risco | Em risco | Não posso perder |
| **R1** | Perdidos | Perdidos | Em risco | Em risco | Não posso perder |

Conferi as 25 células com o gráfico do artigo, uma a uma: coincidem. O eixo vertical do artigo diz «Frequência **ou** Valor Monetário»; o exercício usa a **média das notas F e M**, arredondada com `ROUND` (0,5 para cima).

### 4.6 Como se cria a tabela RFM (folha `Tabela_RFM`) e o gráfico
Foi feita com as fórmulas mais simples: `COUNTIFS`, `SUMIFS` e `INDEX`.
1. **Coluna auxiliar** na folha `RFM`, `K`: `=ROUND(I2;0)` (nota F+M arredondada), porque o `COUNTIFS` não arredonda.
2. **Matriz de clientes** (linhas = nota de Recência 5 a 1; colunas = nota F+M 1 a 5): `=COUNTIFS(RFM!$F$2:$F$794;$A5;RFM!$K$2:$K$794;B$4)`. Conta os clientes com essa nota R **e** essa nota F+M. A soma das 25 células é 793.
3. **Segmento de cada célula**: `=INDEX(RFM!$S$3:$S$27;($A14-1)*5+B$13)`. A tabela de segmentos tem 25 linhas ordenadas por R e depois por F; a posição da célula é `(R-1)*5+F+M`.
4. **Distribuição por segmento**: `=COUNTIFS(RFM!$J$2:$J$794;A23)` (clientes), `=B23/SUM($B$23:$B$33)` (% dos clientes), `=SUMIFS(RFM!$E$2:$E$794;RFM!$J$2:$J$794;A23)` (receita) e `=D23/SUM($D$23:$D$33)` (% da receita).
5. **Heatmap da matriz**: escala de 2 cores (branco para verde) sobre a matriz de clientes.
6. **Treemap** (o gráfico do teu exemplo): selecionar a coluna dos segmentos e a dos clientes (`A22:B33`) e **Inserir > Gráficos > Gráfico de Hierarquia > Mapa de Árvore**. A área de cada bloco é proporcional ao nº de clientes; ativar «rótulos de dados» para mostrar nome e valor. **Este gráfico não consegui criá-lo por automação nesta versão do Excel**: insere-se em 3 cliques a partir da tabela do ponto 4.
7. Validei a folha no Excel: as 25 células da matriz e os 11 segmentos batem com o Python, sem diferenças (793 clientes).

### 4.7 Resultados por segmento
| Segmento | Clientes | % clientes | % receita | Recência mediana (dias) | Frequência mediana |
|---|---|---|---|---|---|
{L_seg}

### 4.8 Exemplo trabalhado de interpretação
*Pergunta: «quem são os Campeões e os clientes em risco?»*
1. **Observar**: no treemap, Fiéis com potencial e Fiéis ocupam cerca de dois terços da área (67%); Campeões é um bloco pequeno.
2. **Quantificar**: Campeões = {seg['Campeões']['clientes']} clientes ({pc(seg['Campeões']['pct_clientes'])}) e {pc(seg['Campeões']['pct_receita'])} da receita. Em risco = {seg['Em risco']['clientes']} clientes ({pc(seg['Em risco']['pct_clientes'])}); Não posso perder = {seg['Não posso perder']['clientes']}.
3. **Comparar**: os Campeões geram {str(round(seg['Campeões']['pct_receita']/seg['Campeões']['pct_clientes'], 1)).replace('.', ',')} vezes a receita que teriam se a receita fosse repartida por igual ({pc(seg['Campeões']['pct_receita'])} da receita com {pc(seg['Campeões']['pct_clientes'])} dos clientes). Em risco e Não posso perder juntos valem {n(risco)} de receita histórica ({pc(risco/tot)} do total), com recência mediana de {seg['Em risco']['R_med']:.0f} e {seg['Não posso perder']['R_med']:.0f} dias.
4. **Concluir** (interpretação): a receita concentra-se em clientes recentes e frequentes; o grupo Em risco / Não posso perder merece atenção pela receita passada e por já não comprar há mais de um ano.
5. **Hipótese**: não sabemos por que deixaram de comprar (preço, concorrência, falta de produto). Os dados só mostram que deixaram.
6. **Ação sugerida**: vem do artigo da G4 (baseado em Putler), não dos dados. Por exemplo, Em risco: e-mails personalizados e oferta de renovação; Campeões: recompensá-los e pedir que promovam a marca.

### 4.9 Correções feitas à tua folha RFM (na cópia corrigida)
1. **Lacunas nas faixas de Recência.** O original usava `<=72` e depois `>73`, `>146`, `>219`, `>293`: os dias 73, 146, 219 e 293 ficavam sem nota (a fórmula devolvia «0»). Passou a usar `>=`. Afeta {F['lacunas_original']} cliente(s).
2. **Lista de clientes.** A folha tinha **638 clientes válidos, 156 linhas em branco e 155 clientes em falta** (o correto são **793**). As linhas em branco ficavam com recência absurda (43.099 dias), frequência 0 e monetização 0, e **distorciam o `PERCENTILE`** da Monetização.
3. **Notas de Recência em texto** (`"5"`) passaram a números.
4. **Coluna SEGMENTO** estava vazia; foi calculada com a fórmula de 4.3.
5. **Cabeçalho**: «NOTA Monetização» (na verdade a média de F e M) passou a «Nota F+M (média)».
6. Recalculei no Excel e comparei com o Python cliente a cliente: **0 diferenças** em Recência, Frequência, Monetização, notas e segmento.

### 4.10 Diferenças em relação ao artigo (a decidir)
- O artigo define um **período de análise de 12 meses** (Recência de 1 a 365 dias; Frequência dentro de 365 dias). O exercício usa **todo o histórico** (2014–2017): a Frequência chega a 17 encomendas e **99 clientes (12,5%) têm mais de 365 dias de recência** e ficam com nota 1. Restringir aos últimos 12 meses mudaria os segmentos. **Não foi calculado**; fica como sugestão separada.
- O eixo vertical do artigo é «Frequência ou Valor Monetário»; o exercício usa a média F+M (secção 4.5).

### 4.11 Origem de cada conclusão
| Conclusão do relatório | Tipo | Como foi obtida |
|---|---|---|
| Fiéis ({seg['Fiéis']['clientes']}) e Fiéis com potencial ({seg['Fiéis com potencial']['clientes']}) somam {pc(seg['Fiéis']['pct_clientes'] + seg['Fiéis com potencial']['pct_clientes'], 0)} dos clientes e {pc(seg['Fiéis']['pct_receita'] + seg['Fiéis com potencial']['pct_receita'], 0)} da receita | Facto | `COUNTIFS` e `SUMIFS` por segmento |
| Os {seg['Campeões']['clientes']} Campeões geram {pc(seg['Campeões']['pct_receita'])} da receita | Facto | Idem |
| Em risco e Não posso perder: recência mediana de {seg['Em risco']['R_med']:.0f} e {seg['Não posso perder']['R_med']:.0f} dias; {n(risco)} de receita histórica | Facto | Mediana da coluna Recência (dias) e `SUMIFS` de vendas nesses segmentos |
| A receita concentra-se em clientes recentes e frequentes | Interpretação | Consequência dos números acima |

---

## 5. Produtos e Localizações (ranking de desempenho)

### 5.1 O que é e quando se usa
Um **ranking** ordena itens (produtos, categorias, regiões) por uma medida (vendas, lucro) para ver **onde está o impacto**. Junta-se a **margem** (lucro ÷ vendas) para distinguir volume de rentabilidade, e ao **Pareto** (a regra dos 80/20: poucos itens explicam a maior parte das vendas). Responde a: que produtos geram mais receita e quais têm baixo desempenho; se há relação entre desempenho e região.

### 5.2 Medidas e dimensões necessárias
| Campo | Tipo | Origem | Para que serve |
|---|---|---|---|
| Sales | **Medida** | orders | Vendas |
| Profit | **Medida** | orders | Lucro |
| Margem = Profit ÷ Sales | **Medida** derivada | calculada | Rentabilidade |
| Product ID / Product Name | Dimensão | orders + product | Ranking de produtos |
| Category / Sub-Category | Dimensão | product | Agrupamento de produtos |
| Region / State / City | Dimensão | location | Ranking de localizações (não há «loja») |
| Order Date → Ano | Dimensão derivada | `=YEAR(B2)` | Evolução anual |

Para juntar estas dimensões à tabela `orders` usa-se o tabelão (secção 1): `INDEX`/`MATCH` sobre `product` (por Product ID) e `location` (por Order ID).

### 5.3 Como se calcula cada campo em Excel (ficheiro `analise_produtos_localizacoes.xlsx`)
| Campo | Fórmula | Como funciona |
|---|---|---|
| Vendas por categoria, região, produto… | `=SUMIFS(dados!$E$2:$E$9995;dados!$I$2:$I$9995;A5)` | Soma `Sales` das linhas cuja categoria é a da célula A5 |
| Lucro | `=SUMIFS(dados!$H$2:$H$9995;dados!$I$2:$I$9995;A5)` | Idem sobre `Profit` |
| % das vendas | `=B5/SUM($B$5:$B$7)` | Vendas do item ÷ total (com `$` para fixar o intervalo do total) |
| Margem | `=D5/B5` | Lucro ÷ vendas |
| % acumulada (Pareto) | `=SUM($F$2:F2)` | Soma acumulada das % de vendas, com a lista ordenada de maior para menor |
| Ranking | `=RANK(C2;$C$2:$C$1863)` | Posição do produto por vendas (1 = maior) |
| Ano | `=YEAR(B2)` | Ano da data da encomenda |
| Produtos para 80% das vendas | `=SUMPRODUCT(--(Produtos!G2:G1863<=0,8))+1` | Conta os produtos até se chegar a 80% acumulado, mais o que o ultrapassa |
| Produtos com lucro negativo | `=COUNTIF(Produtos!D2:D1863;"<0")` | Conta lucros totais negativos |

### 5.4 Como se cria a tabela, o ranking e os gráficos
1. **Tabelão** (folha `dados`): `orders` + categoria, subcategoria, nome, região, estado e cidade, mais a coluna `Ano`.
2. **Tabelas por dimensão**: escrever a lista de valores da dimensão (categorias, regiões…) e as colunas `SUMIFS` da tabela acima. Ordenar a lista por vendas (maior para menor) **antes** de calcular o `% acumulada`.
3. **Ranking de produtos**: uma linha por Product ID; ordenar por vendas em ordem decrescente; `% acumulada` dá o Pareto.
4. **Top e bottom 10**: as 10 primeiras e as 10 últimas linhas do ranking ordenado (folha `Top_Bottom`).
5. **Tabela dinâmica** (folha `Tabela_dinamica`): Inserir > Tabela Dinâmica sobre `dados`; Linhas = Category e Sub-Category; Colunas = Region; Valores = Soma de Sales. Serve para explorar cruzamentos sem escrever fórmulas.
6. **Gráficos de barras**: selecionar a coluna de nomes e a de vendas > **Inserir > Gráfico de Barras**; para o ranking ficar do maior para o menor em cima, em «Formatar eixo» marcar **Categorias por ordem inversa**.
7. Validei o ficheiro no Excel: as vendas e os lucros de **1.862 produtos** batem com o Python, sem diferenças; o Pareto dá 414 produtos e há 301 com lucro negativo.

### 5.5 Resultados
**Por ano**
| Ano | Vendas | % |
|---|---|---|
{L_ano}

**Por categoria**
| Categoria | Vendas | % | Lucro | Margem |
|---|---|---|---|---|
{L_cat}

**Por subcategoria**
| Subcategoria | Vendas | % | Lucro | Margem |
|---|---|---|---|---|
{L_sub}

**Por região**
| Região | Vendas | % | Lucro | Margem |
|---|---|---|---|---|
{L_reg}

**10 estados com mais vendas**
| Estado | Vendas | % |
|---|---|---|
{L_est}

### 5.6 Exemplo trabalhado de interpretação
*Pergunta: «que produtos geram mais receita e quais têm baixo desempenho?»*
1. **Observar**: as três categorias têm pesos parecidos nas vendas ({", ".join(pc(r['pct']) for r in R['cat'])}).
2. **Quantificar**: o lucro é muito diferente: Tecnologia {n(cat['Technology']['lucro'])}, Material de escritório {n(cat['Office Supplies']['lucro'])} e Mobiliário {n(cat['Furniture']['lucro'])}. As margens são {pc(mg(cat['Technology']))}, {pc(mg(cat['Office Supplies']))} e {pc(mg(cat['Furniture']))}.
3. **Comparar**: o Mobiliário vende {pc(cat['Furniture']['pct'])} do total mas margem de {pc(mg(cat['Furniture']))}, cerca de um sétimo da das outras categorias. Dentro dele, Tables ({n(sub['Tables']['lucro'])}) e Bookcases ({n(sub['Bookcases']['lucro'])}) têm lucro **negativo**.
4. **Concluir** (interpretação): vendas e lucro não coincidem; o peso de uma categoria nas vendas não indica a sua rentabilidade.
5. **Hipótese**: o baixo lucro do Mobiliário pode vir de descontos ou custos; os dados deste relatório não o provam. A coluna `Discount` existe e poderia testá-lo (fora do âmbito atual).
6. **Amostra**: são milhares de linhas por categoria, por isso é sólido; para estados pequenos (poucas encomendas) seria preciso cuidado.

### 5.7 Origem de cada conclusão
| Conclusão do relatório | Tipo | Como foi obtida |
|---|---|---|
| Tecnologia, Mobiliário e Material de escritório pesam de forma semelhante nas vendas mas com lucros muito diferentes | Facto | `SUMIFS` de `Sales` e `Profit` por categoria |
| Tables, Bookcases e Supplies têm lucro negativo | Facto | `SUMIFS` de `Profit` por subcategoria |
| {R['prod_80pct']} produtos ({pc(R['prod_80pct']/R['prod_n'], 0)} do catálogo) geram 80% das vendas | Facto | % acumulada do ranking (Pareto) |
| {R['prod_lucro_neg']} dos {n(R['prod_n'])} produtos têm lucro total negativo | Facto | `COUNTIF` do lucro por produto < 0 |
| Califórnia e Nova Iorque somam {pc(R['estado'][0]['pct'] + R['estado'][1]['pct'])} das vendas; West e East somam {pc(R['regiao'][0]['pct'] + R['regiao'][1]['pct'], 0)} | Facto | `SUMIFS` por estado e região |
| Vendas e lucro não coincidem | Interpretação | Comparação de % vendas com margem por categoria |
| A baixa rentabilidade do Mobiliário pode dever-se a descontos ou custos | Hipótese | Não demonstrada: exigiria analisar `Discount` |

---

## 6. Google Drive / Google Sheets
- **Convertem-se em geral bem**: `SUMIFS`, `COUNTIFS`, `MINIFS`, `MAXIFS`, `IFS`, `EDATE`, `UNIQUE`, `FILTER`, `SORT`, `PERCENTILE`, `INDEX`/`MATCH`, `RANK`, `SUMPRODUCT`, formatação condicional por escala de cor e gráficos de barras.
- **Não se convertem (ou não de forma fiável)**: tabelas dinâmicas com **modelo de dados** (como a do Cohort, com contagem distinta) e o treemap. Se o Cohort tiver de funcionar no Google Sheets, seria preciso uma versão só com fórmulas; não a incluí porque o exercício foi feito com tabela dinâmica. As tabelas dinâmicas simples costumam converter-se.
- **Ainda não testei os ficheiros no Google Sheets**: só no Excel. Recomendo abri-los lá depois de carregar e confirmar as folhas `RFM`, `Tabela_RFM` e `Produtos`.
- **Formatos de data e separadores** dependem do idioma da folha; se aparecer um erro em fórmulas com `TEXT`, trocar `AAAA` por `yyyy` (ou o contrário).

---

## 7. Correspondência com o Python
| Passo | Script |
|---|---|
| Cálculo de cohort, RFM, produtos e localizações (validação) | `analise/01_calculos.py` |
| Relatório HTML | `analise/03_relatorio_html.py` |
| Este documento (PDF e Word) | `analise/04_documento_metodologia.py` |
| Excel de Produtos e Localizações (estrutura e fórmulas) | `analise/05_excel_produtos.py` (calculado e guardado depois no Excel) |
| Cohort e RFM | Feitos diretamente no Excel, a partir dos teus originais (que ficam intactos) |
| Resultados intermédios | `analise/dados/` |
"""
import markdown, os, subprocess, shutil, pathlib
open("analise/dados/metodologia_excel.md", "w", encoding="utf-8").write(doc)
import re


def separar_listas(texto):
    """O Markdown exige uma linha em branco antes de uma lista; garante-a."""
    out, ant = [], ""
    for lin in texto.split(chr(10)):
        lin = re.sub(r"^   ([-*] )", r"    ", lin)
        eh_item = re.match(r"^\s*([-*]|\d+\.) ", lin) is not None
        ant_item = re.match(r"^\s*([-*]|\d+\.) ", ant) is not None
        if eh_item and ant.strip() and not ant_item and not ant.lstrip().startswith("|"):
            out.append("")
        out.append(lin)
        ant = lin
    return chr(10).join(out)


corpo = markdown.markdown(separar_listas(doc), extensions=["tables", "sane_lists"])
css = """@page{size:A4;margin:16mm 14mm}body{font:10.5pt/1.5 'Segoe UI',Arial,sans-serif;color:#222}h1{font-size:20pt;border-bottom:2px solid #2f7a5e;padding-bottom:6px}
h2{font-size:15pt;color:#1e4d3e;margin-top:26px;border-bottom:1px solid #ccc;padding-bottom:3px;page-break-after:avoid}h3{font-size:12pt;color:#1e4d3e;margin-top:18px;page-break-after:avoid}
table{border-collapse:collapse;width:100%;margin:8px 0 12px;font-size:9pt;page-break-inside:auto}tr{page-break-inside:avoid}th{background:#2f4f4f;color:#fff;text-align:left;padding:4px 6px}
td{border:1px solid #ccc;padding:4px 6px;vertical-align:top}code{font-family:Consolas,monospace;font-size:8.8pt;background:#f2f0ea;padding:1px 3px;border-radius:3px;word-break:break-word}
hr{border:0;border-top:1px solid #ccc;margin:18px 0}"""
html = f"<!doctype html><html lang='pt-PT'><head><meta charset='utf-8'><title>Metodologia e estudo de apoio</title><style>{css}</style></head><body>{corpo}</body></html>"
os.makedirs("docs", exist_ok=True)
hp = pathlib.Path("analise/dados/metodologia_excel.html").resolve()
hp.write_text(html, encoding="utf-8")
edge = next(x for x in (r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe", r"C:\Program Files\Microsoft\Edge\Application\msedge.exe") if os.path.exists(x))
pdf = pathlib.Path("docs/metodologia_excel.pdf").resolve()
subprocess.run([edge, "--headless", "--disable-gpu", "--no-pdf-header-footer", f"--print-to-pdf={pdf}", hp.as_uri()], check=True, capture_output=True)
print("PDF:", pdf.exists(), len(doc.splitlines()), "linhas")
