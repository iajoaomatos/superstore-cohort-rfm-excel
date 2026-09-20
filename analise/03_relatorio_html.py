"""Gera relatorio/relatorio.html (v2, largura total) a partir de analise/dados/."""
import json, os
from urllib.parse import quote

R = json.load(open("analise/dados/resultados.json", encoding="utf-8"))
C, F = R["cohort"], R["rfm"]
n = lambda x: f"{x:,.0f}".replace(",", ".")
pc = lambda x, d=1: f"{x*100:.{d}f}".replace(".", ",") + "%"


ICON = {
    "fac": ("Facto", '<circle cx="12" cy="12" r="9"/><path d="M8 12.5l2.7 2.7L16 9.5"/>'),
    "cál": ("Cálculo", '<path d="M17 5H7l6 7-6 7h10"/>'),
    "int": ("Interpretação", '<path d="M9 18h6M10 21h4M12 3a6 6 0 0 0-3.5 10.9c.6.5 1 1.2 1 2.1h5c0-.9.4-1.6 1-2.1A6 6 0 0 0 12 3z"/>'),
    "hip": ("Hipótese", '<circle cx="12" cy="12" r="9"/><path d="M9.5 9.5a2.5 2.5 0 1 1 3.6 2.2c-.7.4-1.1 1-1.1 1.8M12 17h.01"/>'),
}


def icon(k):
    nome, path = ICON[k]
    return f'<span class="ic {k}" role="img" aria-label="{nome}" title="{nome}"><svg viewBox="0 0 24 24">{path}</svg></span>'


def read(items):
    """Leitura dos dados: fora de caixas, com etiqueta por tipo de afirmação."""
    return '<div class="read">' + "".join(
        f'<p>{icon(k.lower()[:3])}<span>{t}</span></p>' for k, t in items) + "</div>"



def xl(*ficheiros):
    """Ligações (relativas a relatorio/) para os ficheiros Excel de cada análise."""
    ic = '<svg viewBox="0 0 24 24"><path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z"/><path d="M14 3v5h5M9 13l4 5M13 13l-4 5"/></svg>'
    return '<div class="xl">' + "".join(f'<a href="{quote("../" + p)}">{ic}{t}</a>' for t, p in ficheiros) + "</div>"


def q(txt):
    return f'<p class="q">{txt}</p>'


# ---------- cohort: escala vermelho -> amarelo -> verde ----------
STOPS = [(0.0, (196, 84, 72)), (0.10, (214, 176, 72)), (0.25, (74, 160, 116))]


def cohort_color(v, k):
    if v is None:
        return ""
    if k == 0:
        return "background:#1E6B58;font-weight:500;"
    v = min(v, 0.25)
    for (a, ca), (b, cb) in zip(STOPS, STOPS[1:]):
        if v <= b:
            t = (v - a) / (b - a)
            r, g, bl = (round(x + (y - x) * t) for x, y in zip(ca, cb))
            return f"background:rgb({r},{g},{bl});"
    return ""


def heat(cohorts, ncols, limite=None):
    """Tabela de cohorts. Arredondamento 'metade para cima' (como o Excel); células não observáveis ficam vazias."""
    h = '<table class="hm"><thead><tr><th class="l">Cohort</th><th>N</th>' + "".join(f"<th>{k}</th>" for k in range(ncols)) + "</tr></thead><tbody>"
    for c in cohorts:
        h += f'<tr><td class="l">{c["m"]}</td><td class="n">{c["n"]}</td>' + "".join(
            f'<td style="{cohort_color(v, k) if not (limite and int(c["m"][:4]) * 12 + int(c["m"][5:]) + k > limite) else ""}">{"" if v is None or (limite and int(c["m"][:4]) * 12 + int(c["m"][5:]) + k > limite) else str(int(v * 100 + 0.5)) + "%"}</td>'
            for k, v in enumerate(c["r"][:ncols])) + "</tr>"
    return h + "</tbody></table>"


hm24 = heat([c for c in C["cohorts"] if c["m"] < "2016"], 24)   # cohorts 2014-2015, meses 0-23 (como no exercício)
hm12 = heat(C["cohorts"], 13)                                     # os 42 cohorts, meses 0-12


def squarify(items, x, y, w, h):
    """Treemap: items = [(nome, valor)] por ordem decrescente; devolve (nome, x, y, w, h)."""
    tot = sum(v for _, v in items)
    items = [(k, v * w * h / tot) for k, v in items]
    out = []

    def worst(row, side):
        s = sum(a for _, a in row)
        return max(max(side * side * a / (s * s), s * s / (side * side * a)) for _, a in row)
    while items:
        side = min(w, h)
        row, i = [items[0]], 1
        while i < len(items) and worst(row + [items[i]], side) <= worst(row, side):
            row.append(items[i])
            i += 1
        items = items[i:]
        s = sum(a for _, a in row)
        if w >= h:
            cw, yy = s / h, y
            for k, a in row:
                out.append((k, x, yy, cw, a / cw))
                yy += a / cw
            x, w = x + cw, w - cw
        else:
            rh, xx = s / w, x
            for k, a in row:
                out.append((k, xx, y, a / rh, rh))
                xx += a / rh
            y, h = y + rh, h - rh
    return out

# ---------- RFM ----------
SEG = {(1, 1): "Perdidos", (1, 2): "Perdidos", (1, 3): "Em risco", (1, 4): "Em risco", (1, 5): "Não posso perder",
       (2, 1): "Perdidos", (2, 2): "Hibernando", (2, 3): "Em risco", (2, 4): "Em risco", (2, 5): "Não posso perder",
       (3, 1): "Quase dormentes", (3, 2): "Quase dormentes", (3, 3): "Necessitam atenção", (3, 4): "Fiéis", (3, 5): "Fiéis",
       (4, 1): "Promissores", (4, 2): "Fiéis com potencial", (4, 3): "Fiéis com potencial", (4, 4): "Fiéis", (4, 5): "Fiéis",
       (5, 1): "Novos", (5, 2): "Fiéis com potencial", (5, 3): "Fiéis com potencial", (5, 4): "Fiéis", (5, 5): "Campeões"}
COR = {"Campeões": "#2F7A5E", "Fiéis": "#5A9C82", "Fiéis com potencial": "#9CC8B4", "Novos": "#C9DFD3", "Promissores": "#E3E9D5",
       "Necessitam atenção": "#F3DE9E", "Quase dormentes": "#F0C98A", "Hibernando": "#E8B08A", "Em risco": "#E48C78",
       "Não posso perder": "#C4604B", "Perdidos": "#BDB7A8"}
DARK = {"Campeões", "Fiéis", "Não posso perder"}
ACAO = {"Campeões": "Recompensá-los; podem ser os primeiros a adotar novos produtos e promover a marca.",
        "Fiéis": "Oferecer produtos de maior valor, pedir avaliações e reforçar o envolvimento.",
        "Fiéis com potencial": "Programas de fidelidade ou de membros e recomendação de outros produtos.",
        "Novos": "Dar todo o apoio no arranque (onboarding) e começar a construir a relação.",
        "Promissores": "Criar notoriedade da marca e oferecer testes ou avaliações gratuitas.",
        "Necessitam atenção": "Ofertas por tempo limitado e recomendações com base em compras anteriores.",
        "Quase dormentes": "Partilhar conteúdos de valor, recomendar produtos populares com desconto e reconectar.",
        "Em risco": "E-mails personalizados para se reconectar; oferecer renovações e recursos úteis.",
        "Não posso perder": "Reconquistar com renovações ou novos produtos e conversar para evitar que vão para a concorrência.",
        "Hibernando": "Oferecer outros produtos relevantes e descontos especiais.",
        "Perdidos": "Tentar reavivar o interesse com campanhas; se não resultar, não insistir."}
seg = F["segmentos"]
mxs = max(s["pct_clientes"] for s in seg)
mxr = max(s["pct_receita"] for s in seg)
sb = "".join(
    f'<div class="srow"><div class="sn"><i style="background:{COR[s["segmento"]]}"></i>{s["segmento"]}</div>'
    f'<div class="sb"><div class="bar" style="width:{s["pct_clientes"]/mxs*100:.1f}%;background:{COR[s["segmento"]]}"></div></div><div class="sv">{pc(s["pct_clientes"])}</div>'
    f'<div class="sb"><div class="bar g" style="width:{s["pct_receita"]/mxr*100:.1f}%"></div></div><div class="sv">{pc(s["pct_receita"])}</div></div>'
    for s in seg)
tiles = squarify([(s["segmento"], s["clientes"]) for s in seg], 0, 0, 100, 50)   # área proporcional ao n.º de clientes
tm = '<div class="tm">' + "".join(
    f'<div class="tile" title="{k}: {n(next(s["clientes"] for s in seg if s["segmento"] == k))} clientes" '
    f'style="left:{x:.3f}%;top:{y * 2:.3f}%;width:{w:.3f}%;height:{h * 2:.3f}%;background:{COR[k]};color:{"#fff" if k in DARK else "#1D1E22"}">'
    + (f'<b>{k}</b><span>{pc(next(s["pct_clientes"] for s in seg if s["segmento"] == k))}</span>' if w * h >= 170 else "") + "</div>"
    for k, x, y, w, h in tiles) + "</div>"
tmleg = '<div class="tmleg">' + "".join(
    f'<span><i style="background:{COR[s["segmento"]]}"></i>{s["segmento"]} · {n(s["clientes"])} ({pc(s["pct_clientes"])})</span>' for s in seg) + "</div>"
mat = F["matriz"]
cell = lambda r, f: mat.get(str(f), {}).get(str(r), 0)
rn = {1: "≥ 293 dias", 2: "219–292", 3: "146–218", 4: "73–145", 5: "0–72"}
mx = "".join(
    f'<tr><td class="l">R{r}<small>{rn[r]}</small></td>' + "".join(
        f'<td style="background:{COR[SEG[(r, f)]]};color:{"#fff" if SEG[(r, f)] in DARK else "#1D1E22"}"><b>{cell(r, f)}</b><small>{SEG[(r, f)]}</small></td>'
        for f in range(1, 6)) + "</tr>" for r in (5, 4, 3, 2, 1))
mxt = ('<table class="mx"><thead><tr><th></th>' + "".join(f"<th>F+M = {f}</th>" for f in range(1, 6)) + f"</tr></thead><tbody>{mx}</tbody></table>")
st = "".join(
    f'<tr><td class="l"><i class="dot" style="background:{COR[s["segmento"]]}"></i>{s["segmento"]}</td><td class="r">{n(s["clientes"])}</td>'
    f'<td class="r">{pc(s["pct_clientes"])}</td><td class="r">{pc(s["pct_receita"])}</td><td class="acao">{ACAO[s["segmento"]]}</td></tr>' for s in seg)


# ---------- produtos / localizações ----------
def tbl(rows):
    return ('<table class="t"><thead><tr><th>Produto</th><th class="r">Vendas</th><th class="r">Lucro</th></tr></thead><tbody>' + "".join(
        f'<tr><td>{r["Product Name"][:60]}</td><td class="r">{n(r["vendas"])}</td>'
        f'<td class="r {"neg" if r["lucro"] < 0 else ""}">{n(r["lucro"])}</td></tr>' for r in rows) + "</tbody></table>")


def hb(rows, key, lim=None):
    rows = rows[:lim] if lim else rows
    mxv = max(r["vendas"] for r in rows)
    return "".join(
        f'<div class="hrow"><div class="hn">{r[key]}</div><div class="hb"><i style="width:{r["vendas"]/mxv*100:.1f}%"></i></div>'
        f'<div class="hv">{pc(r["pct"])}</div><div class="hm2 {"neg" if r["lucro"] < 0 else ""}">{pc(r["lucro"]/r["vendas"], 0)}</div></div>'
        for r in rows)


CSS = """
:root{--bg:#141518;--ink:#ECEBE7;--mut:#9FA3AD;--line:rgba(255,255,255,.10);--acc:#6CC2A3;--warm:#F08F7C;--cream:#F4EBCB;--glass:rgba(255,255,255,.07);--glassb:rgba(255,255,255,.14)}
*{box-sizing:border-box}html{scroll-behavior:smooth}
body{margin:0;background:var(--bg);color:var(--ink);font:300 17px/1.7 "IBM Plex Sans",system-ui,-apple-system,"Segoe UI",sans-serif;
background-image:radial-gradient(700px 520px at 8% 4%,rgba(46,160,140,.30),transparent 70%),radial-gradient(760px 560px at 96% 22%,rgba(150,90,200,.24),transparent 70%),radial-gradient(800px 560px at 40% 98%,rgba(230,150,70,.18),transparent 70%);background-attachment:fixed}
.layout{display:grid;grid-template-columns:190px minmax(0,1fr);gap:40px;padding:32px 44px 90px}
nav{position:sticky;top:32px;align-self:start;height:max-content}
nav .brand{font-size:11px;letter-spacing:.18em;text-transform:uppercase;color:var(--mut);margin-bottom:16px}
nav a{display:block;color:var(--mut);text-decoration:none;font-size:14.5px;padding:8px 0 8px 14px;border-left:1px solid var(--line)}
nav a:hover{color:var(--ink);border-left-color:var(--cream)}
main{min-width:0}
header h1{font-weight:200;font-size:clamp(32px,4vw,54px);line-height:1.1;margin:0 0 .35em;letter-spacing:-.02em}
header p.sub{color:var(--mut);font-size:19px;max-width:980px;margin:0}
.glass{background:var(--glass);backdrop-filter:blur(22px) saturate(140%);-webkit-backdrop-filter:blur(22px) saturate(140%);border:1px solid var(--glassb);border-radius:20px;padding:26px 28px}
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(190px,1fr));gap:16px;margin:36px 0 0}
.kpi b{display:block;font-weight:200;font-size:40px;letter-spacing:-.02em;line-height:1.15}.kpi span{font-size:14px;color:var(--mut)}
section{margin-top:96px;scroll-margin-top:24px}
h2{font-weight:300;font-size:14px;margin:0 0 10px;letter-spacing:.18em;text-transform:uppercase;color:var(--mut)}
h2 small{display:inline;font-size:inherit;color:var(--acc);margin-right:10px;letter-spacing:.18em}
h3{font-weight:400;font-size:17px;margin:0 0 16px;color:var(--ink)}
.q{font-family:"Playfair Display",Georgia,serif;font-style:italic;font-weight:400;color:var(--cream);font-size:clamp(24px,2.6vw,34px);line-height:1.3;margin:0 0 30px;max-width:1100px}
.q.big{font-size:clamp(26px,3vw,38px);margin:30px 0 0}
.key{display:flex;flex-wrap:wrap;gap:22px;margin-top:30px;font-size:13.5px;color:var(--mut)}.key>span{display:inline-flex;align-items:center;gap:8px}
.grid2{display:grid;grid-template-columns:repeat(auto-fit,minmax(min(100%,540px),1fr));gap:20px;margin-top:20px}
.read{margin-top:26px;max-width:1100px}.read p{margin:0 0 12px;display:flex;gap:14px;align-items:flex-start;color:#D9D8D3}
.ic{flex:none;display:inline-flex;width:28px;height:28px;align-items:center;justify-content:center;border-radius:50%;border:1px solid var(--line);margin-top:2px}
.ic svg{width:16px;height:16px;fill:none;stroke:currentColor;stroke-width:1.7;stroke-linecap:round;stroke-linejoin:round}
.key .ic{width:24px;height:24px;margin:0}.key .ic svg{width:14px;height:14px}
.fac{color:#6CC2A3}.cál{color:#9DB4E0}.int{color:#F0D98A}.hip{color:#F08F7C}
table{border-collapse:collapse;width:100%;font-size:15px}
th{font-weight:400;color:var(--mut);text-align:center;padding:8px 4px;border-bottom:1px solid var(--line);font-size:13px}th.l{text-align:left}
td{padding:4px;text-align:center;border-bottom:1px solid rgba(255,255,255,.05);font-variant-numeric:tabular-nums}td.l{text-align:left}td.n{color:var(--mut)}
.hm{table-layout:fixed}.hm th:first-child,.hm td:first-child{width:92px}.hm th:nth-child(2),.hm td:nth-child(2){width:52px}
.hm td{height:30px;font-size:clamp(11px,1vw,14px);color:#fff}.hm td.l{color:var(--ink);border-bottom-color:var(--line)}.hm td.n{color:var(--mut);border-bottom-color:var(--line)}
.legend{display:flex;flex-wrap:wrap;align-items:center;gap:10px;font-size:13px;color:var(--mut);margin-top:16px}.legend i{width:200px;max-width:40vw;height:10px;border-radius:5px;background:linear-gradient(90deg,rgb(196,84,72),rgb(214,176,72),rgb(74,160,116))}
.mx{table-layout:fixed;border-collapse:separate;border-spacing:4px}.mx td{height:76px;border:0;border-radius:10px;padding:2px;word-break:break-word}.mx td b{display:block;font-size:clamp(16px,1.6vw,22px);font-weight:400}.mx td small,.mx td.l small{display:block;font-size:clamp(10px,.9vw,12px);opacity:.9;line-height:1.25}
.mx th:first-child,.mx td.l{width:84px}.mx td.l{background:none!important;color:var(--ink);font-weight:400;border-color:transparent}.mx td.l small{color:var(--mut)}
.t td{text-align:left;border-bottom:1px solid var(--line);vertical-align:top}.t th.r,.t td.r{text-align:right}.neg{color:var(--warm)}
.acao{text-align:left!important;color:var(--mut);font-size:14.5px}.dot{display:inline-block;width:10px;height:10px;border-radius:3px;margin-right:9px}
.srow,.shead{display:grid;grid-template-columns:minmax(120px,190px) 1fr 56px 1fr 56px;gap:10px;align-items:center}
.srow{padding:8px 0;border-top:1px solid var(--line);font-size:15px}.shead{font-size:11.5px;color:var(--mut);letter-spacing:.08em;text-transform:uppercase;padding-bottom:6px}.shead span:nth-child(2),.shead span:nth-child(4){grid-column:span 2}
.sn i{display:inline-block;width:10px;height:10px;border-radius:3px;margin-right:9px}.sb{height:10px;background:rgba(255,255,255,.07);border-radius:3px}.bar{height:100%;border-radius:3px}.bar.g{background:#8A8FA0}.sv{text-align:right;color:var(--mut);font-size:14px}
.hrow{display:grid;grid-template-columns:minmax(90px,150px) 1fr 58px 48px;gap:12px;align-items:center;padding:7px 0;border-top:1px solid var(--line);font-size:15px}
.hb{height:9px;background:rgba(255,255,255,.07);border-radius:3px}.hb i{display:block;height:100%;background:var(--acc);border-radius:3px}.hv,.hm2{text-align:right;color:var(--mut);font-size:14px}
.tm{position:relative;width:100%;aspect-ratio:2/1;border-radius:12px;overflow:hidden}.tile{position:absolute;border:2px solid #141518;padding:8px 10px;overflow:hidden;line-height:1.25;display:flex;flex-direction:column;justify-content:flex-end;font-size:clamp(11px,1.15vw,15px)}.tile b{font-weight:500;overflow-wrap:anywhere}.tile span{opacity:.9}
.tmleg{display:flex;flex-wrap:wrap;gap:8px 22px;margin-top:16px;font-size:13.5px;color:var(--mut)}.tmleg i{display:inline-block;width:10px;height:10px;border-radius:3px;margin-right:8px}
details summary{cursor:pointer;color:var(--cream);font-size:16px}details[open] summary{margin-bottom:4px}
.xl{display:flex;flex-wrap:wrap;gap:10px 22px;margin:18px 0 0}.xl a{display:inline-flex;align-items:center;gap:8px;color:var(--cream);font-size:14px;text-decoration:none;border-bottom:1px solid rgba(244,235,203,.35)}.xl a:hover{border-bottom-color:var(--cream)}.xl svg{width:15px;height:15px;fill:none;stroke:currentColor;stroke-width:1.6;stroke-linecap:round;stroke-linejoin:round}
.note{font-size:14px;color:var(--mut);margin:14px 0 0;max-width:1100px}
footer{margin-top:96px;font-size:13.5px;color:var(--mut);border-top:1px solid var(--line);padding-top:18px}
@media(max-width:1000px){.layout{grid-template-columns:1fr;padding:20px 18px 60px}nav{position:static;display:flex;flex-wrap:wrap;gap:4px 18px}nav a{border:0;padding:4px 0}}
@media(max-width:640px){.srow,.shead{grid-template-columns:1fr 1fr 44px 1fr 44px;font-size:13px}.hrow{grid-template-columns:96px 1fr 54px 44px}.glass{padding:18px}}
"""


# ---------- leituras limitadas ao que aparece nos quadros ----------
co24 = [c for c in C["cohorts"] if c["m"] < "2016"]                       # os 24 cohorts do quadro, todos com os meses 1-23 observados
media_c = {c["m"]: sum(c["r"][1:24]) / 23 for c in co24}                # leitura horizontal: média de cada cohort
melhor = max(co24, key=lambda c: media_c[c["m"]])
pior = min(co24, key=lambda c: media_c[c["m"]])
media_geral = sum(media_c.values()) / len(co24)
vert = [sum(c["r"][k] for c in co24) / len(co24) for k in range(24)]     # leitura vertical: média de cada mês
kmax = max(range(1, 24), key=lambda k: vert[k])
kmin = min(range(1, 24), key=lambda k: vert[k])
h1, h2 = sum(vert[1:13]) / 12, sum(vert[13:24]) / 11
pico = max(melhor["r"][1:24])
kpico = melhor["r"].index(pico)
zeros = sum(1 for v in melhor["r"][1:24] if v == 0)
ns = [c["n"] for c in co24 if c["m"] >= "2015"]
cat_ = {r["Category"]: r for r in R["cat"]}
sub_ = {r["Sub-Category"]: r for r in R["sub"]}
reg_ = {r["Region"]: r for r in R["regiao"]}
mg = lambda r: r["lucro"] / r["vendas"]
sg_ = {x["segmento"]: x for x in F["segmentos"]}
p1 = lambda x: pc(x, 1)
leit_cohort = [
    ("Cálculo", f"Leitura horizontal (média de cada cohort nos meses 1 a 23; os 24 cohorts têm os 23 meses observados, por isso são comparáveis): vai de {p1(media_c[pior['m']])} ({pior['m']}, N = {pior['n']}) a {p1(media_c[melhor['m']])} ({melhor['m']}, N = {melhor['n']}); a média dos 24 cohorts é {p1(media_geral)}."),
    ("Cálculo", f"Leitura vertical (média de cada mês nos 24 cohorts): vai de {p1(vert[kmin])} no mês {kmin} a {p1(vert[kmax])} no mês {kmax}; a média dos meses 1 a 12 é {p1(h1)} e a dos meses 13 a 23 é {p1(h2)}."),
    ("Facto", f"O cohort com a média mais alta, {melhor['m']} (N = {melhor['n']}), tem pico de {round(pico * 100)}% no mês {kpico} ({round(pico * melhor['n'])} dos {melhor['n']} clientes) e {zeros} meses a 0%."),
    ("Interpretação", f"A retenção mensal não desce com a idade do cohort: fica perto de {p1(media_geral)} e é mais alta nos meses 13 a 23 do que nos meses 1 a 12. Os cohorts de 2015 têm N entre {min(ns)} e {max(ns)} (um cliente move a célula entre {p1(1 / max(ns))} e {p1(1 / min(ns))}), por isso as diferenças entre cohorts pequenos devem ser lidas com cautela."),
]
risco_pc = sg_["Em risco"]["pct_clientes"] + sg_["Não posso perder"]["pct_clientes"]
risco_pr = sg_["Em risco"]["pct_receita"] + sg_["Não posso perder"]["pct_receita"]
leit_rfm = [
    ("Facto", f"Fiéis ({sg_['Fiéis']['clientes']}) e Fiéis com potencial ({sg_['Fiéis com potencial']['clientes']}) somam {pc(sg_['Fiéis']['pct_clientes'] + sg_['Fiéis com potencial']['pct_clientes'], 0)} dos clientes e {pc(sg_['Fiéis']['pct_receita'] + sg_['Fiéis com potencial']['pct_receita'], 0)} da receita."),
    ("Facto", f"Os {sg_['Campeões']['clientes']} Campeões ({p1(sg_['Campeões']['pct_clientes'])} dos clientes) geram {p1(sg_['Campeões']['pct_receita'])} da receita, {str(round(sg_['Campeões']['pct_receita'] / sg_['Campeões']['pct_clientes'], 1)).replace('.', ',')} vezes o que teriam com uma repartição igual."),
    ("Facto", f"Em risco ({sg_['Em risco']['clientes']}) e Não posso perder ({sg_['Não posso perder']['clientes']}) são {p1(risco_pc)} dos clientes e {p1(risco_pr)} da receita; os {sg_['Perdidos']['clientes']} Perdidos ({p1(sg_['Perdidos']['pct_clientes'])}) geram só {p1(sg_['Perdidos']['pct_receita'])}."),
    ("Interpretação", f"A receita concentra-se nos segmentos de compra recente e boa frequência. Em risco e Não posso perder pesam mais na receita ({p1(risco_pr)}) do que no número de clientes ({p1(risco_pc)}), pelo que são o grupo a reativar em primeiro lugar, com as ações indicadas na tabela."),
]
leit_prod = [
    ("Facto", f"Tecnologia ({p1(cat_['Technology']['pct'])}), Mobiliário ({p1(cat_['Furniture']['pct'])}) e Material de escritório ({p1(cat_['Office Supplies']['pct'])}) pesam de forma semelhante nas vendas, mas as margens são {p1(mg(cat_['Technology']))}, {p1(mg(cat_['Furniture']))} e {p1(mg(cat_['Office Supplies']))}."),
    ("Facto", f"Phones ({p1(sub_['Phones']['pct'])}) e Chairs ({p1(sub_['Chairs']['pct'])}) são as subcategorias com mais vendas. Tables tem {p1(sub_['Tables']['pct'])} das vendas e margem negativa ({p1(mg(sub_['Tables']))})."),
    ("Facto", f"West ({p1(reg_['West']['pct'])}) e East ({p1(reg_['East']['pct'])}) somam {pc(reg_['West']['pct'] + reg_['East']['pct'], 0)} das vendas; Califórnia ({p1(R['estado'][0]['pct'])}) e Nova Iorque ({p1(R['estado'][1]['pct'])}) somam {p1(R['estado'][0]['pct'] + R['estado'][1]['pct'])}. Central tem a menor margem das regiões ({p1(mg(reg_['Central']))})."),
    ("Interpretação", "Vendas e lucro não coincidem: o peso de uma categoria ou de uma região nas vendas não indica a sua rentabilidade (Mobiliário, Tables e Central têm margens muito abaixo das restantes), e as vendas concentram-se em poucos estados."),
]
c_ = C["curva_pond"]
html = f'''<!doctype html><html lang="pt-PT"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>SuperStore — Cohort, RFM e Desempenho</title>
<link rel="preconnect" href="https://fonts.googleapis.com"><link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@200;300;400;500&family=Playfair+Display:ital,wght@1,400;1,500&display=swap" rel="stylesheet">
<style>{CSS}</style></head><body><div class="layout">
<nav><div class="brand">SuperStore · Excel</div><a href="#cohort">1 · Cohort</a><a href="#rfm">2 · RFM</a><a href="#produtos">3 · Produtos e locais</a></nav>
<main>
<header><h1>Retenção, segmentação de clientes e desempenho comercial</h1>
<p class="sub">Análise de Cohort, RFM e rankings de produtos e localizações sobre 4 anos de encomendas (janeiro de 2014 a dezembro de 2017). Versão para revisão.</p>

<div class="key">{"".join(f"<span>{icon(k)}{ICON[k][0]}</span>" for k in ICON)}</div>
<div class="kpis"><div class="glass kpi"><b>{n(R["total_sales"])}</b><span>Vendas totais (2014–2017)</span></div>
<div class="glass kpi"><b>793</b><span>Clientes</span></div><div class="glass kpi"><b>5.009</b><span>Encomendas</span></div>
<div class="glass kpi"><b>1.862</b><span>Produtos</span></div><div class="glass kpi"><b>{pc(R["ano"][0]["pct"], 0)}</b><span>das vendas em 2017, o melhor ano</span></div></div></header>



<section id="cohort"><h2><small>1 · Retenção de clientes</small>Análise de Cohort</h2>
{q("Qual é a retenção de clientes ao longo dos meses? Que cohorts retêm mais?")}
<div class="glass"><h3>Retenção por cohort de 2014 e 2015, meses 0 a 23 (% dos clientes do cohort ativos em cada mês)</h3><div class="scroll">{hm24}</div>
<div class="legend"><span>0%</span><i></i><span>25% ou mais</span><span style="margin-left:18px">Mês 0 = 100% (mês da 1.ª compra).</span></div></div>
<p class="note">São 42 cohorts no total (mês da primeira compra); o quadro mostra os 24 de 2014 e 2015, os únicos com dimensão para conclusões (os 18 cohorts de 2016 e 2017 somam só 62 clientes).</p>
{xl(("analise_cohort.xlsx: a tabela dinâmica do exercício, alargada, com as médias","analises-excel/analise_cohort.xlsx"))}
{read(leit_cohort)}</section>

<section id="rfm"><h2><small>2 · Segmentação de clientes</small>Análise RFM</h2>
{q("Quem são os Campeões e os Clientes em risco? Como estão distribuídos os clientes pelos segmentos? Que ações tomar para fidelizar ou recuperar?")}
<p class="note" style="margin:0 0 20px">Data de referência: 30/12/2017. Recência = dias desde a última compra; Frequência = encomendas distintas; Monetização = soma das vendas. Notas de 1 a 5, segmentos da matriz RFM da G4 Educação.</p>
<div class="glass"><h3>Distribuição dos clientes por segmento (área proporcional ao n.º de clientes)</h3>{tm}{tmleg}</div>

<div class="glass" style="margin-top:20px"><h3>Segmentos e ações sugeridas</h3><div class="scroll"><table class="t"><thead><tr><th>Segmento</th><th class="r">Clientes</th><th class="r">% clientes</th><th class="r">% receita</th><th style="text-align:left">Ação sugerida (fonte: G4 Educação)</th></tr></thead><tbody>{st}</tbody></table></div></div>
{xl(("analise_rfm.xlsx: folhas RFM e Tabela_RFM","analises-excel/analise_rfm.xlsx"))}
{read(leit_rfm)}</section>

<section id="produtos"><h2><small>3 · Produtos e localizações</small>Desempenho comercial</h2>
{q("Que produtos geram mais receita e quais têm baixo desempenho? O desempenho depende da região em que as encomendas são feitas?")}
<div class="grid2">
<div class="glass"><h3>Categorias e subcategorias <span style="color:var(--mut);font-weight:400;font-size:14px">(% das vendas · margem)</span></h3>{hb(R["cat"], "Category")}<div style="height:12px"></div>{hb(R["sub"], "Sub-Category", 8)}<p class="note">Margem = lucro / vendas.</p></div>
<div class="glass"><h3>Regiões, estados e cidades <span style="color:var(--mut);font-weight:400;font-size:14px">(% das vendas · margem)</span></h3>{hb(R["regiao"], "Region")}<div style="height:12px"></div>{hb(R["estado"], "State", 6)}<div style="height:12px"></div>{hb(R["cidade"], "City", 6)}</div></div>
<div class="grid2">
<div class="glass"><h3>10 produtos com mais vendas</h3><div class="scroll">{tbl(R["prod_top"])}</div></div>
<div class="glass"><h3>10 produtos com menos vendas</h3><div class="scroll">{tbl(R["prod_bottom"])}</div></div></div>
{xl(("analise_produtos_localizacoes.xlsx","analises-excel/analise_produtos_localizacoes.xlsx"))}
{read(leit_prod)}</section>



<footer>Fonte: orders.csv, customers.csv, product.csv, location.csv. Cálculos validados em Python e reproduzidos em Excel; fórmulas descritas na documentação do projeto. Versão para iteração.<br>Projeto do Curso de Analista de Dados: Excel para Analista de Dados - CDS</footer>
</main></div></body></html>'''
os.makedirs("relatorio", exist_ok=True)
open("relatorio/relatorio.html", "w", encoding="utf-8").write(html)
print(len(html))
