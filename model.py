import numpy as np

# ---------- IRR helper ----------
def irr_monthly(cfs):
    cfs=np.array(cfs,dtype=float)
    if cfs.min()>=0 or cfs.max()<=0: return None
    lo,hi=-0.99,1.0
    def npv(r): 
        t=np.arange(len(cfs)); return np.sum(cfs/(1+r)**t)
    # bisection
    flo,fhi=npv(lo),npv(hi)
    # expand hi
    it=0
    while flo*fhi>0 and hi<5 and it<60: hi*=1.5; fhi=npv(hi); it+=1
    if flo*fhi>0: return None
    for _ in range(200):
        mid=(lo+hi)/2; fm=npv(mid)
        if abs(fm)<1: return mid
        if flo*fm<0: hi=mid; fhi=fm
        else: lo=mid; flo=fm
    return (lo+hi)/2

def annual(m): 
    return None if m is None else (1+m)**12-1

# ---------- Base parameters (from Excel SUPUESTOS/ER F2) ----------
N=432
PRICE_AVG=305651443.2/432          # 707,526 all-in (apto+parqueo, sin bodega)
COMMISION_RATE=15149205.1/305651443.2   # 4.957% ventas
# Soft costs fijos F2 (Excel ER F2), bodega income ya = 0
SOFT_FIJOS = 12314239+0+67800+5722558+42625+1579837.9+451064+2393952.1+6866201.5+9229753+2619519.9+1600000  # terreno+...
# (terreno, estudios, honorarios, permisos, req, fianzas, showroom, tramites, FHA, admin, varios, imprevistos)
ISO_PER_UNIT=1893650.6/432
TIMBRES_RATE=2434693.7/305651443.2
IUSI=675000.0
ISR_RATE=0.1875
RATE=0.08            # tasa banco anual

def run(absorcion, constr_months, price_growth, constr_budget=180_000_000, promo_pct=0.01,
        sales_start=0, constr_start=4, deliv_lag=0, label=""):
    # ----- timeline -----
    sale_months=int(np.ceil(N/absorcion))
    constr_end=constr_start+constr_months
    first_deliv=constr_end+deliv_lag          # entregas empatan con fin de obra
    # unit sale schedule
    units=[]  # (sale_month, price)
    sold=0; m=sales_start
    while sold<N:
        n=min(absorcion, N-sold)
        yr=(m)//12
        price=PRICE_AVG*((1+price_growth)**yr)
        for _ in range(n): units.append([m,price])
        sold+=n; m+=1
    # deliveries spread over 9 months from first_deliv (3 torres x 3 estimaciones)
    deliv_window=9
    # assign each unit a delivery month (>= its sale, >= first_deliv)
    for k,u in enumerate(units):
        d=first_deliv + int((k/N)*deliv_window)
        u.append(max(d,u[0]+1))
    horizon=max(u[2] for u in units)+3+6
    T=horizon+1
    inflow=np.zeros(T); opcost=np.zeros(T)
    total_rev=0.0
    for s,price,d in units:
        total_rev+=price
        # reserva
        inflow[s]+=5000
        # enganche 10% en cuotas de s+1..d
        eng=0.10*price-5000
        nm=max(d-s,1)
        for t in range(s+1,d+1):
            if t<T: inflow[t]+=eng/nm
        # hipoteca 90% a escrituración d+3
        esc=min(d+3,T-1); inflow[esc]+=0.90*price
        # comisión al vender
        opcost[s]+=COMMISION_RATE*price
    # terreno upfront
    opcost[0]+=12314239
    # soft fijos (sin terreno) distribuidos sobre construcción
    soft=SOFT_FIJOS-12314239
    for t in range(constr_start,constr_end):
        opcost[t]+=soft/constr_months
    # promoción = promo_pct * presupuesto total inversión (constr+soft+terreno+comis)
    promo_total=promo_pct*(constr_budget+SOFT_FIJOS+COMMISION_RATE*total_rev)
    for t in range(sales_start,sales_start+sale_months):
        opcost[t]+=promo_total/sale_months
    # construcción S-curve
    n=constr_months; k=0.30; mid=n/2
    raw=[1/(1+np.exp(-k*(i-mid))) for i in range(n)]
    dens=[max(raw[i]-(raw[i-1] if i>0 else 0),0) for i in range(n)]
    sden=sum(dens)
    for i in range(n):
        opcost[constr_start+i]+=constr_budget*dens[i]/sden
    # ISO/IUSI/timbres
    for s,price,d in units: opcost[d]+=ISO_PER_UNIT+TIMBRES_RATE*price
    opcost[constr_end]+=IUSI
    # ----- financing (bank) + equity -----
    pre_tax=inflow-opcost
    cash=0.0; debt=0.0; equity_in=0.0; interest_tot=0.0
    eq_flow=np.zeros(T); int_series=np.zeros(T)
    EQUITY_CAP=13000000  # socios aportan ~terreno; resto enganches+banco (alto apalancamiento F2)
    for t in range(T):
        # interest on debt
        i=debt*RATE/12; debt+=i; interest_tot+=i; int_series[t]=i
        cash+=pre_tax[t]
        if cash<0:
            need=-cash
            # equity first up to cap, then debt
            eq=min(need, max(EQUITY_CAP-equity_in,0))
            equity_in+=eq; eq_flow[t]-=eq; cash+=eq
            if cash<0:
                debt+=-cash; cash=0
        else:
            # repay debt
            if debt>0:
                pay=min(cash,debt); debt-=pay; cash-=pay
            if cash>0 and t>=first_deliv:
                eq_flow[t]+=cash; cash=0  # distribuir a socios
    # remaining cash to equity at end
    eq_flow[-1]+=cash
    # ----- P&L -----
    UAII=total_rev-(opcost.sum())   # opcost incl ISO/timbres/IUSI? yes; but ISR separate
    # recompute: opcost includes ISO/IUSI/timbres already; treat those as costs
    EBIT=total_rev-opcost.sum()+ (sum(ISO_PER_UNIT for _ in units)+IUSI+TIMBRES_RATE*total_rev)  # add back tax-like to get pre-tax oper
    # simpler: utilidad antes impuestos = rev - opcost(excl iso/timbres/iusi) - interest
    tax_like=sum(ISO_PER_UNIT for _ in units)+IUSI+TIMBRES_RATE*total_rev
    UAI=total_rev-(opcost.sum()-tax_like)-interest_tot
    ISR=max(UAI,0)*ISR_RATE
    util_neta=UAI-ISR-tax_like
    margen=util_neta/total_rev
    # apply ISR near completion for IRR
    proj_after_tax=pre_tax.copy()
    proj_after_tax[min(int(first_deliv)+6,T-1)]-=ISR
    # interest in equity flow
    eq_flow_final=eq_flow.copy()
    eq_flow_final[min(int(first_deliv)+6,T-1)]-=ISR
    tir_proy=annual(irr_monthly(proj_after_tax))
    tir_acc=annual(irr_monthly(eq_flow_final))
    return dict(label=label,absorcion=absorcion,constr_months=constr_months,price_growth=price_growth,
        sale_months=sale_months,total_rev=total_rev,util=util_neta,margen=margen,
        interest=interest_tot,isr=ISR,promo=promo_total,equity=equity_in,
        constr_end=constr_end,first_deliv=first_deliv,horizon=horizon,
        tir_proy=tir_proy,tir_acc=tir_acc, T=T,
        inflow=inflow,opcost=opcost,eq_flow=eq_flow_final)

# Calibration base (Excel: absorcion 17, ~24m constr, growth 0)
base=run(17,24,0.0,constr_budget=149297579,promo_pct=2702651/210036985,label="BASE Excel")
A=run(25,18,0.0,constr_budget=180_000_000,promo_pct=0.01,label="A Rápido 25/mes")
B=run(14,26,0.03,constr_budget=180_000_000,promo_pct=0.01,label="B Lento 14/mes +3%")

for r in (base,A,B):
    print(f"\n=== {r['label']} ===")
    print(f"  absorción {r['absorcion']}/mes · construcción {r['constr_months']}m · +precio {r['price_growth']*100:.0f}%/año")
    print(f"  meses de venta: {r['sale_months']} · fin obra mes {r['constr_end']} · 1ra entrega mes {r['first_deliv']} · horizonte {r['horizon']}m")
    print(f"  Ventas: Q{r['total_rev']/1e6:.1f}M · Utilidad: Q{r['util']/1e6:.1f}M · Margen {r['margen']*100:.1f}%")
    print(f"  Intereses: Q{r['interest']/1e6:.1f}M · ISR: Q{r['isr']/1e6:.1f}M · Promoción: Q{r['promo']/1e6:.2f}M")
    print(f"  TIR Proyecto: {r['tir_proy']*100:.1f}% · TIR Accionistas: {r['tir_acc']*100:.1f}%")

# ---------- export JSON for dashboard ----------
import json
def monthly(r):
    net=(r['inflow']-r['opcost']).tolist()
    return [round(x) for x in net]
def constr_series(r):
    # reconstruct construction monthly from opcost is mixed; recompute via model not stored -> approximate: positive opcost spikes
    return None
out={}
for key,r in (('base',base),('A',A),('B',B)):
    out[key]=dict(absorcion=r['absorcion'],constr_months=r['constr_months'],
        price_growth=r['price_growth'],sale_months=r['sale_months'],
        ventas=round(r['total_rev']),util=round(r['util']),margen=round(r['margen']*1000)/10,
        interest=round(r['interest']),isr=round(r['isr']),promo=round(r['promo']),
        constr_end=int(r['constr_end']),first_deliv=int(r['first_deliv']),horizon=int(r['horizon']),
        tir_proy=round(r['tir_proy']*1000)/10,tir_acc=round(r['tir_acc']*1000)/10,
        flujo_neto=monthly(r))
print(json.dumps(out['A'],ensure_ascii=False)[:200])
json.dump(out,open('/tmp/model_out.json','w'),ensure_ascii=False)
print('exported', list(out.keys()), 'lens', {k:len(out[k]['flujo_neto']) for k in out})
