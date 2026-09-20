---
name: relatorio-cohort-rfm
description: Gera o relatório HTML de portfólio (nível C-level) do projeto SuperStore com Análise de Cohort, RFM e rankings de produtos e localizações, a partir dos resultados em analise/dados/. Usar quando o utilizador pedir para criar, atualizar ou iterar o relatório/dashboard HTML do projeto.
---

# Relatório de portfólio: Cohort, RFM e Produtos/Localizações

Esta skill produz `relatorio/relatorio.html`, um ficheiro único e autónomo. O visual está em [DESIGN.md](DESIGN.md) e deve ser seguido à letra. O conteúdo segue as regras do `claude.md` do projeto.

## Fluxo de trabalho
1. Confirmar que `analise/dados/` está atualizado (`resultados.json`, `cohort_*.csv`, `rfm_clientes.csv`). Se os dados mudaram, correr `python analise/01_calculos.py`.
2. Gerar o relatório com `python analise/03_relatorio_html.py` (é o único sítio onde o HTML é montado; alterações de conteúdo ou de design fazem-se nesse script, nunca no HTML gerado).
3. Verificar visualmente: captura com Edge em modo headless (`msedge --headless --screenshot=... --window-size=1600,7600 file:///.../relatorio.html`) e conferir os pontos da checklist abaixo.
4. Mostrar ao utilizador e iterar. Nunca dar o relatório como final antes de o utilizador o validar.

## Estrutura do relatório (ordem fixa)
1. Cabeçalho: título, subtítulo, legenda dos ícones e KPIs. **Sem questão principal** e **sem secção «Base de dados»** (a validação dos dados vive em `docs/metodologia_excel.pdf`).
2. Cohort, RFM e Produtos e Localizações. Cada secção tem:
   1. **Questão** de partida, em texto solto, sem caixa (ver DESIGN.md).
   2. Tabela ou gráfico (cohort: só o quadro 2014–2015 × meses 0–23; RFM: treemap + tabela de ações, **sem** matriz R × F+M).
   3. Ligação ao ficheiro Excel da análise (função `xl()`): `analises-excel/analise_cohort.xlsx`, `analise_rfm.xlsx` e `analise_produtos_localizacoes.xlsx` (nomes em snake_case, sem a palavra «corrigido»).
   4. **Leitura dos dados**: interpretações e conclusões *depois* de cada tabela ou gráfico, fora de caixas.
3. **Sem secção «Metodologia»** no relatório: as diferenças em relação ao artigo G4 e as fórmulas estão em `docs/metodologia_excel.pdf` (PDF gerado por `analise/04_documento_metodologia.py`).
4. Rodapé: fonte dos dados, e uma linha abaixo: «Projeto do Curso de Analista de Dados: Excel para Analista de Dados - CDS».

## Gráficos permitidos
Só: matriz de cohort (heatmap), análise RFM (treemap por segmento, matriz R × F+M e tabela de ações) e rankings de produtos e localizações. O aspeto de cada um segue o exemplo do utilizador descrito no fim do DESIGN.md. Não acrescentar outros gráficos sem pedido do utilizador.

## Regras de conteúdo
- Cada afirmação leva um ícone que indica o seu tipo: **Facto** (observado nos dados), **Cálculo** (métrica derivada), **Interpretação** (conclusão de negócio suportada pelos dados) e **Hipótese** (explicação que os dados não comprovam). Usar ícones, nunca as palavras. A legenda dos ícones aparece uma vez no cabeçalho.
- Não inventar dados, causas ou conclusões. Todo o número no texto tem de existir em `analise/dados/`.
- O quadro do cohort é o da tabela dinâmica do utilizador: cohorts de 2014 e 2015 × meses 0–23 (retângulo, todas as células observadas). O mês 0 é sempre 100%.
- **As leituras limitam-se ao que está no quadro ou gráfico**: nada de números, causas ou períodos que não apareçam nele.
- **Cohort**: obrigatórias uma leitura **horizontal** (média de cada cohort, incluindo o melhor cohort como exemplo) e uma **vertical** (média de cada mês). Só se calculam médias com cohorts completos (todos os meses da janela observados).
- **3 a 4 conclusões por análise** (Cálculo, Facto e Interpretação); sem hipóteses nem extrapolações. Como analista, ficar pela análise e interpretação dos dados.
- Não incluir no relatório notas sobre metodologia ou janelas (vão para o documento de estudo).
- Ler com cautela os cohorts pequenos (a coluna N do quadro mostra o tamanho).
- O Excel continua a ser a ferramenta de referência: a explicação das fórmulas está em `docs/metodologia_excel.pdf`.
- Segmentos RFM e ações por segmento: fonte G4 Educação (PDF em `output-relatório/`). A matriz coincide com o artigo; as diferenças assumidas (janela de 12 meses, que o utilizador decidiu não aplicar, e média F+M) constam do documento de estudo.
- Língua: português europeu. Moeda não indicada nos dados: mostrar números sem símbolo.

## Checklist de verificação visual
- Nenhuma tabela, gráfico ou texto cortado ou com scroll horizontal, em 1600 px e em ~400 px de largura.
- Todas as questões em Playfair Display itálico, sem caixa, com contraste elevado sobre o fundo.
- Ícones em vez das palavras Facto/Cálculo/Interpretação/Hipótese.
- Cohort: mês 0 a 100% e escala vermelho, amarelo e verde legível sobre fundo escuro.
- Sem sombras pesadas; separação só por vidro, borda de 1 px e luz ambiente.

## Pendentes conhecidos (perguntar antes de assumir)
- Confirmar a fonte principal (IBM Plex Sans é a assumida).
- Decidir se se calcula a versão RFM só com os últimos 12 meses, como no artigo.
