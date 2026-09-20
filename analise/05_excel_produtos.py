"""Cria resolucao-corrigida/Análise Produtos e Localizações.xlsx com fórmulas (SUMIFS, RANK, ...) e gráficos.
Depois é aberto no Excel (analise/06_excel_com.ps1) para calcular e guardar os valores em cache."""
import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter as L

B = "base-dados/"
o = pd.read_csv(B + "orders.csv")
o["Order Date"] = pd.to_datetime(o["Order Date"], format="%m/%d/%Y")
t = o.merge(pd.read_csv(B + "product.csv"), on="Product ID").merge(pd.read_csv(B + "location.csv"), on="Order ID")
t = t.loc[o.index] if len(t) == len(o) else t   # mantém a ordem original de orders
N = len(t) + 1   # última linha de dados

wb = Workbook()
HEAD = Font(bold=True, color="FFFFFF")
FILL = PatternFill("solid", fgColor="2F4F4F")
INP = Font(color="0000FF")


def cab(ws, row, titulos):
    for i, x in enumerate(titulos, 1):
        c = ws.cell(row, i, x)
        c.font, c.fill = HEAD, FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


# ---------- dados (tabelão) ----------
d = wb.active
d.title = "dados"
cols = ["Order ID", "Order Date", "Customer ID", "Product ID", "Sales", "Quantity", "Discount", "Profit", "Category", "Sub-Category",
        "Product Name", "Region", "State", "City", "Ano"]
cab(d, 1, cols)
src = ["Order ID", "Order Date", "Customer ID", "Product ID", "Sales", "Quantity", "Discount", "Profit", "Category", "Sub-Category",
       "Product Name", "Region", "State", "City"]
for r, row in enumerate(t[src].itertuples(index=False), 2):
    for c, v in enumerate(row, 1):
        d.cell(r, c, v.to_pydatetime() if hasattr(v, "to_pydatetime") else v)
    d.cell(r, 2).number_format = "yyyy-mm-dd"
    d.cell(r, 15, f"=YEAR(B{r})")
d.freeze_panes = "A2"
for i, w in enumerate([16, 12, 12, 17, 10, 9, 9, 10, 16, 14, 40, 10, 16, 18, 7], 1):
    d.column_dimensions[L(i)].width = w

R = lambda col: f"dados!${col}$2:${col}${N}"   # intervalo fixo de uma coluna de dados
S, PR = R("E"), R("H")


def tabela(ws, r0, chave_col, valores, titulo_chave):
    """Escreve uma tabela SUMIFS: chave | Vendas | % Vendas | Lucro | Margem | % acumulada. Devolve a última linha."""
    cab(ws, r0, [titulo_chave, "Vendas", "% Vendas", "Lucro", "Margem", "% acumulada"])
    a, b = r0 + 1, r0 + len(valores)
    for i, v in enumerate(valores):
        r = a + i
        ws.cell(r, 1, v)
        ws.cell(r, 2, f"=SUMIFS({S},{R(chave_col)},A{r})").number_format = "#,##0"
        ws.cell(r, 3, f"=B{r}/SUM($B${a}:$B${b})").number_format = "0.0%"
        ws.cell(r, 4, f"=SUMIFS({PR},{R(chave_col)},A{r})").number_format = "#,##0"
        ws.cell(r, 5, f"=D{r}/B{r}").number_format = "0.0%"
        ws.cell(r, 6, f"=SUM($C${a}:C{r})").number_format = "0.0%"
    r = b + 1
    ws.cell(r, 1, "Total").font = Font(bold=True)
    ws.cell(r, 2, f"=SUM(B{a}:B{b})").number_format = "#,##0"
    ws.cell(r, 4, f"=SUM(D{a}:D{b})").number_format = "#,##0"
    ws.cell(r, 5, f"=D{r}/B{r}").number_format = "0.0%"
    return a, b


def larguras(ws, ls):
    for i, w in enumerate(ls, 1):
        ws.column_dimensions[L(i)].width = w


def barras(ws, titulo, a, b, ancora, col=2, cats=1, h=9, w=18):
    ch = BarChart()
    ch.type, ch.title, ch.height, ch.width, ch.legend = "bar", titulo, h, w, None
    ch.add_data(Reference(ws, min_col=col, min_row=a - 1, max_row=b), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=cats, min_row=a, max_row=b))
    ch.y_axis.numFmt = "#,##0"
    ch.x_axis.scaling.orientation = "maxMin"
    ch.x_axis.delete = ch.y_axis.delete = False
    ws.add_chart(ch, ancora)


def ordem(col):
    return t.groupby(col).Sales.sum().sort_values(ascending=False).index.tolist()


# ---------- Categorias ----------
c = wb.create_sheet("Categorias")
c["A1"], c["A1"].font = "Vendas, lucro e margem por categoria e subcategoria", Font(bold=True, size=13)
a1, b1 = tabela(c, 3, "I", ordem("Category"), "Categoria")
a2, b2 = tabela(c, b1 + 4, "J", ordem("Sub-Category"), "Subcategoria")
larguras(c, [22, 14, 11, 14, 11, 12])
barras(c, "Vendas por subcategoria", a2, b2, "H3")
barras(c, "Margem por subcategoria", a2, b2, "H22", col=5)

# ---------- Regiões, estados, cidades ----------
g = wb.create_sheet("Regioes")
g["A1"], g["A1"].font = "Vendas, lucro e margem por região, estado e cidade", Font(bold=True, size=13)
ar, br = tabela(g, 3, "L", ordem("Region"), "Região")
ae, be = tabela(g, br + 4, "M", ordem("State"), "Estado")
larguras(g, [22, 14, 11, 14, 11, 12])
barras(g, "Vendas por região", ar, br, "H3", h=7)
barras(g, "10 estados com mais vendas", ae, ae + 9, "H18")
cid = wb.create_sheet("Cidades")
cid["A1"], cid["A1"].font = "Vendas, lucro e margem por cidade (ordenadas por vendas)", Font(bold=True, size=13)
tabela(cid, 3, "N", ordem("City"), "Cidade")
larguras(cid, [26, 14, 11, 14, 11, 12])

# ---------- Produtos ----------
p = wb.create_sheet("Produtos")
prod = t.groupby(["Product ID", "Product Name"]).Sales.sum().sort_values(ascending=False).reset_index()
cab(p, 1, ["Product ID", "Produto", "Vendas", "Lucro", "Margem", "% Vendas", "% acumulada (Pareto)", "Ranking"])
n_p = len(prod) + 1
for i, (pid, nm, _) in enumerate(prod.itertuples(index=False), 2):
    p.cell(i, 1, pid)
    p.cell(i, 2, nm)
    p.cell(i, 3, f"=SUMIFS({S},{R('D')},A{i})").number_format = "#,##0.00"
    p.cell(i, 4, f"=SUMIFS({PR},{R('D')},A{i})").number_format = "#,##0.00"
    p.cell(i, 5, f"=IFERROR(D{i}/C{i},0)").number_format = "0.0%"
    p.cell(i, 6, f"=C{i}/SUM($C$2:$C${n_p})").number_format = "0.00%"
    p.cell(i, 7, f"=SUM($F$2:F{i})").number_format = "0.0%"
    p.cell(i, 8, f"=RANK(C{i},$C$2:$C${n_p})")
p.freeze_panes = "A2"
larguras(p, [17, 55, 13, 13, 10, 10, 16, 9])

# ---------- Top e Bottom 10 ----------
tb = wb.create_sheet("Top_Bottom")
tb["A1"], tb["A1"].font = "10 produtos com mais e com menos vendas", Font(bold=True, size=13)
for base, titulo, linhas in ((3, "Mais vendas", range(2, 12)), (16, "Menos vendas", range(n_p - 9, n_p + 1))):
    cab(tb, base, [titulo, "Vendas", "Lucro"])
    for k, r in enumerate(linhas, base + 1):
        tb.cell(k, 1, f"=Produtos!B{r}")
        tb.cell(k, 2, f"=Produtos!C{r}").number_format = "#,##0.00"
        tb.cell(k, 3, f"=Produtos!D{r}").number_format = "#,##0.00"
larguras(tb, [60, 13, 13])
barras(tb, "Top 10 produtos por vendas", 4, 13, "E3")

# ---------- Anos ----------
an = wb.create_sheet("Anos")
an["A1"], an["A1"].font = "Vendas por ano", Font(bold=True, size=13)
cab(an, 3, ["Ano", "Vendas", "% Vendas", "Lucro"])
for i, y in enumerate(range(2014, 2018), 4):
    an.cell(i, 1, y)
    an.cell(i, 2, f"=SUMIFS({S},{R('O')},A{i})").number_format = "#,##0"
    an.cell(i, 3, f"=B{i}/SUM($B$4:$B$7)").number_format = "0.0%"
    an.cell(i, 4, f"=SUMIFS({PR},{R('O')},A{i})").number_format = "#,##0"
larguras(an, [10, 14, 11, 14])
barras(an, "Vendas por ano", 4, 7, "F3", h=7, w=12)

# ---------- Resumo (1.ª folha) ----------
s = wb.create_sheet("Resumo", 0)
s["A1"], s["A1"].font = "SuperStore — Produtos e Localizações", Font(bold=True, size=14)
kp = [("Vendas totais", f"=SUM({S})", "#,##0"), ("Lucro total", f"=SUM({PR})", "#,##0"), ("Margem global", "=B4/B3", "0.0%"),
      ("Nº de produtos", f"=COUNTA(Produtos!A2:A{n_p})", "0"),
      ("Nº de produtos que geram 80% das vendas", f"=SUMPRODUCT(--(Produtos!G2:G{n_p}<=0.8))+1", "0"),
      ("% do catálogo que gera 80% das vendas", "=B7/B6", "0.0%"),
      ("Nº de produtos com lucro total negativo", f'=COUNTIF(Produtos!D2:D{n_p},"<0")', "0")]
for i, (k, f, nf) in enumerate(kp, 3):
    s.cell(i, 1, k)
    s.cell(i, 2, f).number_format = nf
s["A12"] = "Folhas"
s["A12"].font = Font(bold=True)
for i, x in enumerate(["dados: tabelão (orders + product + location) com a coluna Ano", "Categorias / Regioes / Cidades: tabelas SUMIFS e gráficos",
                       "Produtos: ranking de todos os produtos, % acumulada (Pareto) e ranking", "Top_Bottom: 10 produtos com mais e com menos vendas",
                       "Anos: vendas por ano"], 13):
    s.cell(i, 1, x)
larguras(s, [48, 16])

wb.save("resolucao-corrigida/Análise Produtos e Localizações.xlsx")
print("gravado", len(prod), "produtos")
