"""Cálculos de Cohort, RFM e Produtos/Localizações a partir dos CSV (validação em Python).
O processo equivalente em Excel está descrito no claude.md e no relatório."""
import pandas as pd, numpy as np, json, math, os
os.makedirs("analise/dados", exist_ok=True)   # a pasta dados não vai para o GitHub: é recriada aqui
B = "base-dados/"
o = pd.read_csv(B+"orders.csv"); c = pd.read_csv(B+"customers.csv")
p = pd.read_csv(B+"product.csv"); l = pd.read_csv(B+"location.csv")
o["Order Date"] = pd.to_datetime(o["Order Date"], format="%m/%d/%Y")
t = (o.merge(c, on="Customer ID").merge(p, on="Product ID").merge(l, on="Order ID"))
assert len(t) == len(o)
REF = pd.Timestamp("2017-12-30")
res = {}

# ---------- COHORT ----------
o["mes"] = o["Order Date"].dt.to_period("M")
o["coorte"] = o.groupby("Customer ID")["mes"].transform("min")
o["dt"] = (o.mes.dt.year-o.coorte.dt.year)*12 + (o.mes.dt.month-o.coorte.dt.month)
cnt = o.groupby(["coorte","dt"])["Customer ID"].nunique().unstack()
last = pd.Period("2017-12","M")
for co in cnt.index:                      # célula observável só se coorte+dt <= último mês
    for k in cnt.columns:
        if co + k > last: cnt.loc[co,k] = np.nan
cnt = cnt.fillna(0).where(pd.DataFrame({k:[(co+k)<=last for co in cnt.index] for k in cnt.columns},index=cnt.index)) 
pct = cnt.div(cnt[0], axis=0)
cnt.index = cnt.index.astype(str); pct.index = pct.index.astype(str)
cnt.to_csv("analise/dados/cohort_clientes.csv"); pct.round(6).to_csv("analise/dados/cohort_retencao_pct.csv")
res["cohort"] = {"cohorts":[{"m":i,"n":int(cnt.loc[i,0]),"r":[None if pd.isna(x) else round(float(x),4) for x in pct.loc[i,:23]]} for i in pct.index],
                 "meses_sem_novos":[str(m) for m in pd.period_range("2014-01","2017-12",freq="M") if str(m) not in cnt.index]}
# curva média ponderada (só coortes observáveis em t)
curva=[]
for k in range(0,13):
    obs = cnt[k].dropna(); base = cnt.loc[obs.index,0]
    curva.append(round(float(obs.sum()/base.sum()),4))
res["cohort"]["curva_pond"]=curva
# retenção média t1..t12 por coorte com t12 observável
rows=[]
for i in pct.index:
    v=pct.loc[i,1:12]
    if v.notna().all(): rows.append((i,int(cnt.loc[i,0]),round(float(v.mean()),4)))
res["cohort"]["media_1a12"]=rows

# ---------- RFM ----------
g = o.groupby("Customer ID").agg(ultima=("Order Date","max"), F=("Order ID","nunique"), M=("Sales","sum")).reset_index()
g["R"] = (REF-g.ultima).dt.days
def nR(d): return 5 if d<=72 else 4 if d<=145 else 3 if d<=218 else 2 if d<=292 else 1
def nF(f): return 5 if f>=13 else 4 if f>=10 else 3 if f>=7 else 2 if f>=4 else 1
q = [np.percentile(g.M,x) for x in (20,40,60,80)]
def nM(m): return 1 if m<=q[0] else 2 if m<=q[1] else 3 if m<=q[2] else 4 if m<=q[3] else 5
g["nR"]=g.R.map(nR); g["nF"]=g.F.map(nF); g["nM"]=g.M.map(nM)
g["FM"]=((g.nF+g.nM)/2+0.5).map(math.floor)   # média arredondada (0,5 para cima, como o ROUND do Excel)
seg = {(1,1):"Perdidos",(1,2):"Perdidos",(1,3):"Em risco",(1,4):"Em risco",(1,5):"Não posso perder",
 (2,1):"Perdidos",(2,2):"Hibernando",(2,3):"Em risco",(2,4):"Em risco",(2,5):"Não posso perder",
 (3,1):"Quase dormentes",(3,2):"Quase dormentes",(3,3):"Necessitam atenção",(3,4):"Fiéis",(3,5):"Fiéis",
 (4,1):"Promissores",(4,2):"Fiéis com potencial",(4,3):"Fiéis com potencial",(4,4):"Fiéis",(4,5):"Fiéis",
 (5,1):"Novos",(5,2):"Fiéis com potencial",(5,3):"Fiéis com potencial",(5,4):"Fiéis",(5,5):"Campeões"}
g["segmento"]=[seg[(a,b)] for a,b in zip(g.nR,g.FM)]
g.round(2).to_csv("analise/dados/rfm_clientes.csv",index=False)
sg = g.groupby("segmento").agg(clientes=("Customer ID","count"),receita=("M","sum"),R_med=("R","median"),F_med=("F","median")).sort_values("receita",ascending=False)
sg["pct_clientes"]=sg.clientes/len(g); sg["pct_receita"]=sg.receita/g.M.sum()
res["rfm"]={"n":len(g),"percentis_M":[round(x,2) for x in q],
 "notaR":g.nR.value_counts().sort_index().to_dict(),"notaF":g.nF.value_counts().sort_index().to_dict(),"notaM":g.nM.value_counts().sort_index().to_dict(),
 "segmentos":sg.round(4).reset_index().to_dict("records"),
 "matriz":g.groupby(["nR","FM"]).size().unstack(fill_value=0).to_dict()}
res["rfm"]["lacunas_original"]=int(g.R.isin([73,146,219,293]).sum())

# ---------- PRODUTOS E LOCALIZAÇÕES ----------
tot = t.Sales.sum(); res["total_sales"]=round(tot,2)
def rk(col,n=None):
    r=t.groupby(col).agg(vendas=("Sales","sum"),lucro=("Profit","sum"),encomendas=("Order ID","nunique")).sort_values("vendas",ascending=False)
    r["pct"]=r.vendas/tot; return r.head(n) if n else r
prod=rk(["Product ID","Product Name"]); prod["cum"]=prod.pct.cumsum()
prod.round(4).reset_index().to_csv("analise/dados/produtos_ranking.csv",index=False)
res["prod_top"]=prod.head(10).round(4).reset_index().to_dict("records")
res["prod_bottom"]=prod.tail(10).round(4).reset_index().to_dict("records")
res["prod_n"]=len(prod); res["prod_80pct"]=int((prod.cum<=0.8).sum()+1)
res["prod_lucro_neg"]=int((prod.lucro<0).sum())
for k,col in {"cat":["Category"],"sub":["Sub-Category"],"regiao":["Region"],"estado":["State"],"cidade":["City"],"ano":[t["Order Date"].dt.year.rename("Ano")]}.items():
    r=rk(col); res[k]=r.round(4).reset_index().to_dict("records") if k not in("estado","cidade") else r.head(10).round(4).reset_index().to_dict("records")
    r.round(4).reset_index().to_csv(f"analise/dados/vendas_{k}.csv",index=False)
res["n_estados"]=t.State.nunique(); res["n_cidades"]=t.City.nunique()
res["estado_bottom"]=rk(["State"]).tail(5).round(4).reset_index().to_dict("records")
res["dup_linha_exata"]=int(o.duplicated().sum())
json.dump(res,open("analise/dados/resultados.json","w",encoding="utf-8"),ensure_ascii=False,indent=1,default=str)
print("ok")
