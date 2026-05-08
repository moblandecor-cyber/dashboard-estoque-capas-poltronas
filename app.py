import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import warnings, os, io, requests
warnings.filterwarnings("ignore")

# ── Configuração ───────────────────────────────────────────────────────────────
DEFAULT_FILE_PATH  = r"C:\Users\Matheus\Downloads\Controle_Estoque_Capas_12.xlsx"
LOCAL_FILE_EXISTS  = os.path.exists(DEFAULT_FILE_PATH)
GSHEETS_ID         = "1zP5wU3VvJxH7X-y3zeleUNOaOT2URTDswxoIJjbny08"
GSHEETS_EXPORT_URL = f"https://docs.google.com/spreadsheets/d/{GSHEETS_ID}/export?format=xlsx"

# ── Paleta Profissional ────────────────────────────────────────────────────────
TERRA   = "#B5451B"   # terracota refinada
AMBER   = "#C97D2E"   # âmbar escuro
COPPER  = "#A0522D"   # cobre
SAND    = "#C8A882"   # areia
CHOCO   = "#1C1C2E"   # azul-grafite escuro (sidebar)
NAVY    = "#2E3A59"   # azul navy (acentos)
SLATE   = "#4A5568"   # cinza slate
SURFACE = "#FFFFFF"   # fundo branco puro
BG      = "#FFFFFF"   # background branco
CARD    = "#FFFFFF"   # cards brancos
BORDER  = "#E2E8F0"   # bordas cinza claro
TEXT    = "#1A202C"   # texto escuro
MUT     = "#718096"   # texto mutado
SUCCESS = "#276749"
WARN    = "#C05621"
DANGER  = "#9B2335"

SEQ = [TERRA, AMBER, NAVY, COPPER, SLATE, SAND, "#6B7280", "#374151", "#C8A882", "#2E3A59"]

_BASE = dict(
    plot_bgcolor=SURFACE,
    paper_bgcolor="rgba(0,0,0,0)",
    font=dict(family="Inter, Helvetica Neue, sans-serif", color=TEXT, size=12),
    colorway=SEQ,
    legend=dict(bgcolor="rgba(0,0,0,0)", borderwidth=0, font=dict(size=11)),
)
CL  = {**_BASE,
       "xaxis": dict(gridcolor="#F1F5F9", linecolor=BORDER, tickfont=dict(size=11)),
       "yaxis": dict(gridcolor="#F1F5F9", linecolor=BORDER, tickfont=dict(size=11))}
CLP = dict(**_BASE)   # pie / donut

# ── Page config ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="MOBLAN Decor – Gestão de Estoque",
    page_icon="🛋️",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

html, body, [class*="css"], .stApp {{
    font-family: 'Inter', sans-serif !important;
    background-color: {BG} !important;
}}
[data-testid="stAppViewContainer"] > .main {{ background-color:{BG}; }}
[data-testid="stHeader"] {{ background:transparent; }}

/* ── Sidebar ── */
[data-testid="stSidebar"] {{
    background-color:{CHOCO} !important;
    border-right:1px solid #13131F;
}}
[data-testid="stSidebar"] * {{ color:#CBD5E1 !important; }}
[data-testid="stSidebar"] h2,[data-testid="stSidebar"] h3 {{ color:#F1F5F9 !important; }}
[data-testid="stSidebar"] [data-baseweb="select"] > div {{
    background-color:#2D2D45 !important; border-color:#3D3D5C !important;
}}
[data-testid="stSidebar"] .stButton > button {{
    background:linear-gradient(135deg,{TERRA},{AMBER}) !important;
    color:white !important; border:none !important;
    border-radius:8px !important; font-weight:600 !important;
}}
[data-testid="stSidebar"] [data-testid="stFileUploader"] {{
    background:#2D2D45 !important; border:1px dashed #4A4A6A !important;
    border-radius:10px !important;
}}

/* ── KPI card ── */
.kpi {{ background:{CARD}; border-radius:12px; padding:20px 22px;
    box-shadow:0 1px 3px rgba(0,0,0,.06), 0 4px 16px rgba(0,0,0,.04);
    border:1px solid {BORDER}; margin-bottom:12px;
    position:relative; overflow:hidden; }}
.kpi::before {{ content:''; position:absolute; top:0;left:0;right:0;
    height:3px; background:linear-gradient(90deg,{TERRA},{AMBER}); }}
.kpi.warn::before   {{ background:linear-gradient(90deg,#DD6B20,#ED8936); }}
.kpi.alert::before  {{ background:linear-gradient(90deg,{DANGER},#E53E3E); }}
.kpi.success::before{{ background:linear-gradient(90deg,{SUCCESS},#48BB78); }}
.kpi.copper::before {{ background:linear-gradient(90deg,{NAVY},{SLATE}); }}
.kpi-icon  {{ font-size:1.4rem; margin-bottom:8px; display:block; }}
.kpi-label {{ font-size:.67rem; color:{MUT}; text-transform:uppercase;
    letter-spacing:1.5px; font-weight:600; margin-bottom:5px; }}
.kpi-val   {{ font-size:1.9rem; font-weight:700; color:{TEXT}; line-height:1.1; }}
.kpi-sub   {{ font-size:.71rem; color:{MUT}; margin-top:4px; }}

/* ── Header ── */
.header {{ background:{CARD}; border-radius:12px; padding:22px 28px;
    border:1px solid {BORDER};
    box-shadow:0 1px 3px rgba(0,0,0,.05); margin-bottom:20px; }}
.brand  {{ font-size:1.6rem; font-weight:700; color:{TEXT}; letter-spacing:-.3px; }}
.brand-dot {{ display:inline-block; width:10px; height:10px; border-radius:50%;
    background:linear-gradient(135deg,{TERRA},{AMBER}); margin-right:8px; }}
.brand-sub {{ font-size:.78rem; color:{MUT}; text-transform:uppercase; letter-spacing:.6px; }}

/* ── Section title ── */
.sec {{ font-size:.95rem; font-weight:600; color:{TEXT};
    border-left:3px solid {TERRA}; padding-left:10px;
    margin:24px 0 14px 0; }}

/* ── Alert box ── */
.alert-box {{ background:#FFF5F5; border-radius:8px; padding:14px 16px;
    border-left:3px solid {DANGER}; margin-bottom:8px; }}
.alert-title {{ font-weight:600; color:{DANGER}; font-size:.85rem; }}
.alert-sub   {{ color:{MUT}; font-size:.78rem; margin-top:3px; }}

/* ── Tabs ── */
.stTabs [data-baseweb="tab-list"] {{
    background:#F8FAFC; border-radius:10px; padding:4px; gap:3px;
    box-shadow:none; border:1px solid {BORDER};
}}
.stTabs [data-baseweb="tab"] {{
    border-radius:8px; padding:8px 20px; font-weight:500;
    color:{MUT}; font-size:.87rem; border:none !important;
    transition: all .15s ease;
}}
.stTabs [aria-selected="true"] {{
    background:{TERRA} !important;
    color:white !important; box-shadow:0 2px 6px rgba(181,69,27,.3) !important;
    font-weight:600 !important;
}}
hr {{ border-color:{BORDER} !important; margin:16px 0 !important; }}
</style>
""", unsafe_allow_html=True)


# ── Helpers ───────────────────────────────────────────────────────────────────
def kpi(icon, label, value, sub="", style=""):
    st.markdown(
        f'<div class="kpi {style}"><span class="kpi-icon">{icon}</span>'
        f'<div class="kpi-label">{label}</div>'
        f'<div class="kpi-val">{value}</div>'
        f'<div class="kpi-sub">{sub}</div></div>',
        unsafe_allow_html=True)

def sec(t): st.markdown(f'<div class="sec">{t}</div>', unsafe_allow_html=True)

def short(s):
    return (str(s).replace("POLTRONA ","").replace("KIT 2 ","K2 ")
            .replace("NAMORADEIRA ","NAM. ").replace("PUFF ","")
            .replace("BANQUETA ","BNQ. ").replace("CADEIRA ","CAD. "))

def categoria(modelo):
    m = str(modelo).upper()
    if "NAMORADEIRA" in m: return "Namoradeira"
    if "BANQUETA"   in m: return "Banqueta"
    if "CADEIRA"    in m: return "Cadeira"
    if "PUFF"       in m: return "Puff"
    if "KIT"        in m: return "Kit"
    return "Poltrona"


# ── Carregamento ──────────────────────────────────────────────────────────────
@st.cache_data(ttl=300, show_spinner="Carregando dados...")
def load_data(file_source):
    if isinstance(file_source, bytes):
        file_source = io.BytesIO(file_source)
    elif isinstance(file_source, str) and file_source.startswith("http"):
        r = requests.get(file_source, timeout=30)
        r.raise_for_status()
        file_source = io.BytesIO(r.content)

    xl = pd.ExcelFile(file_source)

    def clean(df, cols):
        for c in cols: df[c] = df[c].astype(str).str.strip().str.upper()
        return df
    def num(df, cols):
        for c in cols: df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
        return df
    def s_cap(s):
        s = str(s)
        if "ABAIXO" in s or "MÍNIMO" in s: return "Abaixo do Mínimo"
        if "PRODUZIR" in s: return "A Produzir"
        return "OK"
    def s_est(s):
        s = str(s)
        if "REPOR" in s: return "Repor"
        if "ABAIXO" in s or "MÍN" in s: return "Abaixo do Mínimo"
        return "OK"

    # Capas
    df_c = xl.parse("Produção - Capas", header=2)
    df_c.columns = ["Modelo","Tecido","Cor","Qtd_Costuradas","Qtd_Minima","Qtd_Ideal","A_Produzir","Status","Obs"]
    df_c = df_c[df_c["Modelo"].notna()]
    clean(df_c, ["Modelo","Tecido","Cor"])
    df_c = df_c[~df_c["Modelo"].isin(["MODELO","TOTAIS"])]
    num(df_c, ["Qtd_Costuradas","Qtd_Minima","Qtd_Ideal","A_Produzir"])
    df_c["Status_Clean"] = df_c["Status"].apply(s_cap)

    # Avulso
    df_av = xl.parse("Estoque - Avulso", header=3)
    if len(df_av.columns) == 11:
        df_av.columns = ["Modelo","Tecido","Cor","Curva_ABC","Obs","Estoque","Minimo","Ideal","A_Produzir","Status","Obs_Geral"]
    else:
        df_av.columns = ["Modelo","Tecido","Cor","Curva_ABC","Obs","Estoque","Minimo","Ideal","Status","Obs_Geral"]
    df_av = df_av[df_av["Modelo"].notna()]
    clean(df_av, ["Modelo","Tecido","Cor"])
    df_av = df_av[~df_av["Modelo"].isin(["MODELO","IDENTIFICAÇÃO"])]
    num(df_av, ["Estoque","Minimo","Ideal"])
    df_av["Status_Clean"] = df_av["Status"].apply(s_est)

    # Kit
    df_kt = xl.parse("Estoque - Kit 2", header=3)
    if len(df_kt.columns) == 11:
        df_kt.columns = ["Modelo","Tecido","Cor","Curva_ABC","Obs","Estoque","Minimo","Ideal","A_Produzir","Status","Obs_Geral"]
    else:
        df_kt.columns = ["Modelo","Tecido","Cor","Curva_ABC","Obs","Estoque","Minimo","Ideal","Status","Obs_Geral"]
    df_kt = df_kt[df_kt["Modelo"].notna()]
    clean(df_kt, ["Modelo","Tecido","Cor"])
    df_kt = df_kt[~df_kt["Modelo"].isin(["MODELO","IDENTIFICAÇÃO"])]
    num(df_kt, ["Estoque","Minimo","Ideal"])
    df_kt["Status_Clean"] = df_kt["Status"].apply(s_est)

    # Vendas
    df_v = xl.parse("Vendas ABR 26", header=2)
    df_v.columns = ["Modelo","Tecido","Cor","Total_Mes","Media_Dia","Minimo_3d","Ideal_7d","Pct_Total"]
    df_v = df_v[df_v["Modelo"].notna()]
    df_v["Modelo"] = df_v["Modelo"].astype(str).str.strip()
    df_v = df_v[~df_v["Modelo"].str.upper().isin(["MODELO","TOTAL GERAL"])]
    df_v = df_v[~df_v["Modelo"].str.contains(",", na=False)]
    clean(df_v, ["Tecido","Cor"])
    num(df_v, ["Total_Mes","Media_Dia","Minimo_3d","Ideal_7d","Pct_Total"])

    # ABC
    df_abc = xl.parse("Curva ABC", header=3)
    df_abc.columns = ["Rank","Curva","Modelo","Tecido","Cor","Unid_Abril","Pct_Total","Pct_Acum","Min_Ideal"]
    df_abc["Rank"] = pd.to_numeric(df_abc["Rank"], errors="coerce")
    df_abc = df_abc[df_abc["Rank"].notna() & df_abc["Modelo"].notna()]
    df_abc["Modelo"] = df_abc["Modelo"].astype(str).str.strip()
    clean(df_abc, ["Tecido","Cor"])
    num(df_abc, ["Unid_Abril","Pct_Total","Pct_Acum"])

    return df_c, df_av, df_kt, df_v, df_abc


# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 🛋️ MOBLAN Decor")
    st.markdown("Gestão de Estoque & Produção")
    st.markdown("---")
    st.markdown("### 📂 Fonte de Dados")
    usar_gs  = st.toggle("🔗 Google Sheets (ao vivo)", value=True)
    uploaded = st.file_uploader("ou envie um arquivo .xlsx", type=["xlsx"])

    if usar_gs and not uploaded:
        file_source = GSHEETS_EXPORT_URL; file_label = "Google Sheets – ao vivo"
    elif uploaded:
        file_source = uploaded.getvalue(); file_label = uploaded.name; usar_gs = False
    elif LOCAL_FILE_EXISTS:
        file_source = DEFAULT_FILE_PATH; file_label = DEFAULT_FILE_PATH.split("\\")[-1]
    else:
        file_source = None; file_label = "Sem arquivo"

    st.caption(f"📄 {file_label}")
    st.markdown("---")

    if file_source is None:
        st.warning("Faça upload da planilha para continuar.")
        st.stop()

    df_capas, df_avulso, df_kit, df_vendas, df_abc = load_data(file_source)

    def opts(s, ex=("NAN","NA","")):
        return sorted({str(v) for v in s.dropna() if str(v).strip() not in ex})

    st.markdown("### 🔍 Filtros")
    LIMIAR_BAIXO = st.slider("Estoque baixo (< X un)", 1, 20, 5)
    LIMIAR_PARADO = st.slider("Estoque parado (≥ X un)", 5, 50, 20)
    sel_mod   = st.multiselect("Modelo",    opts(df_avulso["Modelo"]), placeholder="Todos")
    sel_tec   = st.multiselect("Tecido",    opts(df_avulso["Tecido"]), placeholder="Todos")
    sel_cor   = st.multiselect("Cor",       opts(df_avulso["Cor"]),    placeholder="Todas")
    sel_curva = st.multiselect("Curva ABC", ["A","B","C"],             placeholder="Todas")
    sel_cat   = st.multiselect("Categoria", ["Poltrona","Namoradeira","Banqueta","Puff","Kit","Cadeira"], placeholder="Todas")

    st.markdown("---")
    if st.button("🔄 Atualizar dados", use_container_width=True):
        load_data.clear(); st.rerun()
    st.caption("📅 Estoque: 07/05/2026  |  Vendas: Abr/2026")


# ── Filtro helpers ────────────────────────────────────────────────────────────
def filt(df, mc="Modelo", tc="Tecido", cc="Cor", sc=None, kc=None):
    m = pd.Series(True, index=df.index)
    if sel_mod  and mc in df: m &= df[mc].isin(sel_mod)
    if sel_tec  and tc in df: m &= df[tc].isin(sel_tec)
    if sel_cor  and cc in df: m &= df[cc].isin(sel_cor)
    if sc:                    m &= df[sc].isin(["OK","Abaixo do Mínimo","A Produzir","Repor"])
    if sel_curva and kc:      m &= df[kc].isin(sel_curva)
    return df[m]

# Adicionar categoria e filtrar
for df in [df_avulso, df_kit]:
    df["Categoria"] = df["Modelo"].apply(categoria)

fc  = filt(df_capas,  sc="Status_Clean")
fav = filt(df_avulso, kc="Curva_ABC")
fkt = filt(df_kit,    kc="Curva_ABC")
if sel_cat:
    fav = fav[fav["Categoria"].isin(sel_cat)]
    fkt = fkt[fkt["Categoria"].isin(sel_cat)]
fv  = df_vendas.copy()
if sel_tec: fv = fv[fv["Tecido"].isin(sel_tec)]
if sel_cor: fv = fv[fv["Cor"].isin(sel_cor)]
fa  = df_abc.copy()
if sel_tec:   fa = fa[fa["Tecido"].isin(sel_tec)]
if sel_curva: fa = fa[fa["Curva"].isin(sel_curva)]

# ── KPIs globais ──────────────────────────────────────────────────────────────
tot_av      = int(fav["Estoque"].sum())
tot_kit     = int(fkt["Estoque"].sum())
tot_prod    = tot_av + tot_kit
tot_cap     = int(fc["Qtd_Costuradas"].sum())
skus_av     = len(fav)
skus_zero   = int((fav["Estoque"] == 0).sum())
skus_baixo  = int((fav["Estoque"].between(1, LIMIAR_BAIXO - 1)).sum())
skus_parado = int((fav["Estoque"] >= LIMIAR_PARADO).sum())
n_skus      = fav["Modelo"].nunique()
taxa_rupt   = skus_zero / skus_av * 100 if skus_av else 0
media_dia   = fv["Media_Dia"].sum()
cob_dias    = round(tot_av / media_dia, 1) if media_dia > 0 else 0
cap_prod    = int(fc["A_Produzir"].sum())
alertas_k   = int((fkt["Status_Clean"] != "OK").sum())
tot_vend    = int(fv["Total_Mes"].sum())
giro        = round(tot_vend / tot_av, 1) if tot_av > 0 else 0


# ── Header ────────────────────────────────────────────────────────────────────
st.markdown(f"""
<div class="header">
    <div>
        <div class="brand"><span class="brand-dot"></span>MOBLAN Decor</div>
        <div class="brand-sub">Painel de Gestão · Estoque & Produção · Maio 2026</div>
    </div>
</div>""", unsafe_allow_html=True)

# ── KPI Linha 1 – Produtos ────────────────────────────────────────────────────
c = st.columns(5)
with c[0]: kpi("🛋️", "Total Produtos em Estoque", f"{tot_prod:,}", f"Avulso: {tot_av:,}  |  Kit: {tot_kit:,}")
with c[1]: kpi("🧵", "Total Capas Costuradas",    f"{tot_cap:,}", f"{cap_prod} a produzir", "copper")
with c[2]: kpi("⚠️", "Estoque Baixo",             str(skus_baixo), f"< {LIMIAR_BAIXO} un por SKU", "warn")
with c[3]: kpi("🚨", "Produtos Zerados",           str(skus_zero),  f"{taxa_rupt:.1f}% dos SKUs", "alert")
with c[4]: kpi("📦", "Mix Ativo",                 str(n_skus),    "modelos únicos em estoque")

# ── KPI Linha 2 – Operacional ─────────────────────────────────────────────────
c2 = st.columns(5)
with c2[0]: kpi("📅", "Cobertura de Estoque",    f"{cob_dias} dias", "estoque ÷ venda/dia", "warn" if cob_dias < 7 else "")
with c2[1]: kpi("🔄", "Giro de Estoque",         f"{giro}×", "vendas Abr ÷ estoque atual")
with c2[2]: kpi("🛑", "Produtos Parados",         str(skus_parado), f"≥ {LIMIAR_PARADO} un estocadas", "warn" if skus_parado > 10 else "")
with c2[3]: kpi("🎁", "Alertas Kit 2",           str(alertas_k), "kits abaixo do mínimo", "alert" if alertas_k > 50 else "warn")
with c2[4]: kpi("🛒", "Vendas Abril 2026",       f"{tot_vend:,}", f"≈ {media_dia:.0f} un/dia", "success")

st.markdown("---")

# ══════════════════════════════════════════════════════════════════════════════
tab1, tab2, tab3, tab4 = st.tabs([
    "🏠  Visão Geral",
    "📦  Produtos Acabados",
    "🧵  Capas & Costura",
    "📊  Estratégico",
])


# ─────────────────────────────────────────────────────────────────────────────
# TAB 1 – VISÃO GERAL
# ─────────────────────────────────────────────────────────────────────────────
with tab1:
    col_a, col_b = st.columns([2, 1])

    with col_a:
        sec("Estoque por Modelo – Top 20")
        top_mod = (fav.groupby("Modelo")["Estoque"].sum()
                      .reset_index().nlargest(20,"Estoque"))
        top_mod["Mod"] = top_mod["Modelo"].apply(short)
        fig = px.bar(top_mod.sort_values("Estoque"), x="Estoque", y="Mod",
                     orientation="h", labels={"Estoque":"Unidades","Mod":""},
                     color="Estoque",
                     color_continuous_scale=[[0,"#F0E6DC"],[0.5,AMBER],[1,TERRA]])
        fig.update_layout(**CL, coloraxis_showscale=False,
                          height=480, margin=dict(t=10,b=10))
        st.plotly_chart(fig, use_container_width=True)

    with col_b:
        sec("Estoque por Categoria")
        cat_df = fav.groupby("Categoria")["Estoque"].sum().reset_index()
        fig = px.pie(cat_df, values="Estoque", names="Categoria", hole=0.55,
                     color_discrete_sequence=SEQ)
        fig.update_layout(**CLP, height=230, margin=dict(t=10,b=10,l=10,r=10))
        st.plotly_chart(fig, use_container_width=True)

        sec("Avulso vs Kit 2")
        tipo_df = pd.DataFrame({"Tipo":["Avulso","Kit 2"],"Estoque":[tot_av, tot_kit]})
        fig = px.pie(tipo_df, values="Estoque", names="Tipo", hole=0.55,
                     color_discrete_map={"Avulso": TERRA, "Kit 2": AMBER})
        fig.update_layout(**CLP, height=200, margin=dict(t=10,b=10,l=10,r=10))
        st.plotly_chart(fig, use_container_width=True)

    col_c, col_d = st.columns(2)
    with col_c:
        sec("Estoque por Cor")
        by_cor = (fav[fav["Cor"].str.len() > 2]
                    .groupby("Cor")["Estoque"].sum().reset_index()
                    .sort_values("Estoque", ascending=True))
        fig = px.bar(by_cor, x="Estoque", y="Cor", orientation="h",
                     labels={"Estoque":"Unidades","Cor":""},
                     color="Estoque",
                     color_continuous_scale=[[0,"#F0E6DC"],[0.5,COPPER],[1,TERRA]])
        fig.update_layout(**CL, coloraxis_showscale=False,
                          height=380, margin=dict(t=10,b=10))
        st.plotly_chart(fig, use_container_width=True)

    with col_d:
        sec("Estoque por Tecido")
        by_tec = (fav[fav["Tecido"].str.len() > 2]
                    .groupby("Tecido")["Estoque"].sum().reset_index()
                    .sort_values("Estoque", ascending=False))
        fig = px.bar(by_tec, x="Tecido", y="Estoque",
                     labels={"Estoque":"Unidades","Tecido":""},
                     color="Tecido", color_discrete_sequence=SEQ)
        fig.update_layout(**CL, showlegend=False, height=280, margin=dict(t=10,b=10))
        st.plotly_chart(fig, use_container_width=True)

        sec("Capas por Tecido")
        cap_tec = (fc[fc["Tecido"].str.len() > 2]
                     .groupby("Tecido")["Qtd_Costuradas"].sum().reset_index()
                     .sort_values("Qtd_Costuradas", ascending=False))
        fig = px.bar(cap_tec, x="Tecido", y="Qtd_Costuradas",
                     labels={"Qtd_Costuradas":"Capas","Tecido":""},
                     color="Tecido", color_discrete_sequence=SEQ)
        fig.update_layout(**CL, showlegend=False, height=240, margin=dict(t=10,b=10))
        st.plotly_chart(fig, use_container_width=True)


# ─────────────────────────────────────────────────────────────────────────────
# TAB 2 – PRODUTOS ACABADOS
# ─────────────────────────────────────────────────────────────────────────────
with tab2:

    # ── Estoque baixo
    sec(f"🟠 Produtos com Estoque Baixo (< {LIMIAR_BAIXO} un)")
    baixo = fav[(fav["Estoque"] > 0) & (fav["Estoque"] < LIMIAR_BAIXO)].copy()
    baixo["Mod"] = baixo["Modelo"].apply(short)
    if len(baixo):
        col_x, col_y = st.columns([2, 1])
        with col_x:
            fig = px.bar(baixo.sort_values("Estoque"),
                         x="Estoque", y="Mod", orientation="h",
                         color="Estoque",
                         color_continuous_scale=[[0,TERRA],[1,AMBER]],
                         labels={"Estoque":"Un","Mod":""})
            fig.update_layout(**CL, coloraxis_showscale=False,
                              height=max(250, len(baixo)*28), margin=dict(t=10,b=10))
            st.plotly_chart(fig, use_container_width=True)
        with col_y:
            tbl = baixo[["Modelo","Cor","Tecido","Estoque","Minimo"]].copy()
            tbl.columns = ["Modelo","Cor","Tecido","Estoque","Mínimo"]
            st.dataframe(tbl.sort_values("Estoque"), use_container_width=True,
                         hide_index=True, height=300)
    else:
        st.success("Nenhum produto com estoque baixo nos filtros atuais.")

    # ── Produtos zerados
    sec("🔴 Produtos Sem Estoque (Zerados)")
    zeros = fav[fav["Estoque"] == 0][["Modelo","Cor","Tecido","Curva_ABC","Minimo"]].copy()
    zeros.columns = ["Modelo","Cor","Tecido","Curva","Mínimo"]
    if len(zeros):
        col_z1, col_z2 = st.columns([1, 2])
        with col_z1:
            st.metric("Total de SKUs zerados", len(zeros))
            st.dataframe(zeros.head(30), use_container_width=True,
                         hide_index=True, height=350)
        with col_z2:
            zer_mod = fav[fav["Estoque"]==0].groupby("Modelo").size().reset_index(name="SKUs Zerados")
            zer_mod["Mod"] = zer_mod["Modelo"].apply(short)
            zer_mod = zer_mod.nlargest(15,"SKUs Zerados")
            fig = px.bar(zer_mod.sort_values("SKUs Zerados"),
                         x="SKUs Zerados", y="Mod", orientation="h",
                         color="SKUs Zerados",
                         color_continuous_scale=[[0,"#F0E6DC"],[1,DANGER]],
                         labels={"Mod":""})
            fig.update_layout(**CL, coloraxis_showscale=False,
                              height=420, margin=dict(t=10,b=10))
            st.plotly_chart(fig, use_container_width=True)
    else:
        st.success("Nenhum produto zerado.")

    # ── Produtos parados
    sec(f"🟡 Produtos Parados (≥ {LIMIAR_PARADO} un em estoque)")
    # Combinar estoque com vendas para detectar alto estoque + baixa saída
    est_mod = fav.groupby("Modelo")["Estoque"].sum().reset_index()
    v_upper = fv.copy()
    v_upper["Modelo_UP"] = v_upper["Modelo"].str.upper()
    def get_vendas(modelo):
        for _, r in v_upper.iterrows():
            if r["Modelo_UP"] in modelo: return r["Total_Mes"]
        return 0
    est_mod["Vendas_Mes"] = est_mod["Modelo"].apply(get_vendas)
    parado = est_mod[est_mod["Estoque"] >= LIMIAR_PARADO].copy()
    parado["Mod"] = parado["Modelo"].apply(short)
    if len(parado):
        fig = go.Figure()
        fig.add_trace(go.Bar(name="Estoque", x=parado["Mod"], y=parado["Estoque"],
                             marker_color=AMBER))
        fig.add_trace(go.Bar(name="Vendas Abr", x=parado["Mod"], y=parado["Vendas_Mes"],
                             marker_color=TERRA))
        fig.update_layout(**CL, barmode="group", height=300,
                          xaxis_tickangle=-35, margin=dict(t=10,b=10))
        st.plotly_chart(fig, use_container_width=True)

    # ── Estoque completo avulso
    sec("📋 Estoque Completo – Produtos Avulsos")
    tbl_av = fav[["Modelo","Categoria","Tecido","Cor","Curva_ABC","Estoque","Minimo","Ideal","Status_Clean"]].copy()
    tbl_av.columns = ["Modelo","Categoria","Tecido","Cor","Curva","Estoque","Mínimo","Ideal","Status"]
    st.dataframe(tbl_av.sort_values("Estoque", ascending=False),
                 use_container_width=True, hide_index=True, height=380)
    csv = tbl_av.to_csv(index=False).encode("utf-8-sig")
    st.download_button("⬇️ Exportar CSV", csv, "estoque_avulso.csv", "text/csv")


# ─────────────────────────────────────────────────────────────────────────────
# TAB 3 – CAPAS & COSTURA
# ─────────────────────────────────────────────────────────────────────────────
with tab3:
    c_kpi = st.columns(4)
    n_ab  = int((fc["Status_Clean"]=="Abaixo do Mínimo").sum())
    n_pr  = int((fc["Status_Clean"]=="A Produzir").sum())
    with c_kpi[0]: kpi("🧵", "Total Capas Costuradas", f"{tot_cap:,}", f"{len(fc)} SKUs")
    with c_kpi[1]: kpi("📋", "A Produzir",             str(cap_prod), f"{n_pr} SKUs urgentes", "alert")
    with c_kpi[2]: kpi("⚠️", "Abaixo do Mínimo",       str(n_ab),    "SKUs críticos", "warn")
    with c_kpi[3]:
        ok_pct = round((fc["Status_Clean"]=="OK").sum()/len(fc)*100, 1) if len(fc) else 0
        kpi("✅", "SKUs em Dia",  f"{ok_pct}%", "dentro do estoque mínimo", "success")

    st.markdown("---")
    col_c1, col_c2 = st.columns([2, 1])

    with col_c1:
        sec("Capas por Modelo – Costuradas vs A Produzir")
        ms = (fc.groupby("Modelo")
                .agg(Costuradas=("Qtd_Costuradas","sum"),
                     A_Produzir=("A_Produzir","sum"),
                     Minima=("Qtd_Minima","sum"))
                .reset_index().sort_values("Costuradas", ascending=False).head(20))
        ms["Mod"] = ms["Modelo"].apply(short)
        fig = go.Figure()
        fig.add_trace(go.Bar(name="Costuradas", x=ms["Mod"], y=ms["Costuradas"],
                             marker_color=AMBER, marker_line_width=0))
        fig.add_trace(go.Bar(name="A Produzir", x=ms["Mod"], y=ms["A_Produzir"],
                             marker_color=DANGER, marker_line_width=0))
        fig.add_trace(go.Scatter(name="Qtd Mínima", x=ms["Mod"], y=ms["Minima"],
                                 mode="markers", marker=dict(color=CHOCO, size=9, symbol="diamond")))
        fig.update_layout(**CL, barmode="group", height=360,
                          xaxis_tickangle=-40, margin=dict(t=10,b=10))
        st.plotly_chart(fig, use_container_width=True)

    with col_c2:
        sec("Status das Capas")
        s_cnt = fc["Status_Clean"].value_counts().reset_index()
        s_cnt.columns = ["Status","Qtd"]
        fig = px.pie(s_cnt, values="Qtd", names="Status", hole=0.60,
                     color="Status", color_discrete_map={
                         "OK": AMBER, "Abaixo do Mínimo": TERRA, "A Produzir": DANGER})
        fig.update_layout(**CLP, height=300, margin=dict(t=10,b=10,l=10,r=10))
        st.plotly_chart(fig, use_container_width=True)

    # ── Relação Produto x Capas (indicador estratégico)
    sec("⚖️ Relação Produto Acabado × Capas Disponíveis")
    prod_mod = fav.groupby("Modelo")["Estoque"].sum().reset_index().rename(columns={"Estoque":"Produtos"})
    cap_mod  = fc.groupby("Modelo")["Qtd_Costuradas"].sum().reset_index().rename(columns={"Qtd_Costuradas":"Capas"})
    relacao  = prod_mod.merge(cap_mod, on="Modelo", how="outer").fillna(0)
    relacao["Mod"] = relacao["Modelo"].apply(short)
    relacao["Razao"] = (relacao["Capas"] / relacao["Produtos"].replace(0, 0.001)).round(1)
    relacao["Situacao"] = relacao.apply(lambda r:
        "🔴 Faltam capas" if r["Capas"] < r["Produtos"] else
        ("🟢 Equilibrado"  if r["Capas"] <= r["Produtos"]*2 else "🟡 Excesso de capas"), axis=1)
    relacao = relacao[relacao["Produtos"] + relacao["Capas"] > 0].sort_values("Razao")

    fig = go.Figure()
    fig.add_trace(go.Bar(name="Produtos Acabados", x=relacao["Mod"],
                         y=relacao["Produtos"], marker_color=TERRA, marker_line_width=0))
    fig.add_trace(go.Bar(name="Capas Disponíveis", x=relacao["Mod"],
                         y=relacao["Capas"], marker_color=AMBER, marker_line_width=0))
    fig.update_layout(**CL, barmode="group", height=360,
                      xaxis_tickangle=-40, margin=dict(t=10,b=10),
                      title="Barras lado a lado: Produto Acabado vs Capas")
    st.plotly_chart(fig, use_container_width=True)

    col_r1, col_r2 = st.columns([1, 2])
    with col_r1:
        st.markdown("**Situação por Modelo**")
        sit = relacao[["Modelo","Produtos","Capas","Razao","Situacao"]].copy()
        sit.columns = ["Modelo","Produtos","Capas","Razão C/P","Situação"]
        st.dataframe(sit.sort_values("Razão C/P"), use_container_width=True,
                     hide_index=True, height=350)
    with col_r2:
        sec("Heatmap – Capas por Cor × Tecido")
        heat = (fc[(fc["Cor"].str.len()>2) & (fc["Tecido"].str.len()>2)]
                  .groupby(["Cor","Tecido"])["Qtd_Costuradas"].sum().reset_index())
        pivot = heat.pivot(index="Cor", columns="Tecido", values="Qtd_Costuradas").fillna(0)
        top20c = heat.groupby("Cor")["Qtd_Costuradas"].sum().nlargest(18).index
        pivot = pivot.loc[pivot.index.isin(top20c)]
        fig = px.imshow(pivot, text_auto=True, aspect="auto",
                        color_continuous_scale=[[0,"#F0E6DC"],[0.5,AMBER],[1,TERRA]],
                        labels=dict(x="Tecido", y="Cor", color="Capas"))
        fig.update_layout(**CLP, height=380, margin=dict(t=10,b=10))
        st.plotly_chart(fig, use_container_width=True)

    # Tabela capas em falta
    sec("🚨 Capas em Falta – SKUs a Produzir")
    al = fc[fc["Status_Clean"] != "OK"][
        ["Modelo","Tecido","Cor","Qtd_Costuradas","Qtd_Minima","A_Produzir","Status_Clean"]].copy()
    al.columns = ["Modelo","Tecido","Cor","Costuradas","Mínimo","A Produzir","Status"]
    st.dataframe(al.sort_values("A Produzir", ascending=False),
                 use_container_width=True, hide_index=True, height=300)


# ─────────────────────────────────────────────────────────────────────────────
# TAB 4 – ESTRATÉGICO
# ─────────────────────────────────────────────────────────────────────────────
with tab4:
    sec("📈 Curva ABC – Pareto de Vendas (Abril 2026)")
    pareto = fa.sort_values("Rank").copy()
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    for curva, cor in [("A", TERRA), ("B", AMBER), ("C", SAND)]:
        s = pareto[pareto["Curva"] == curva]
        fig.add_trace(go.Bar(name=f"Curva {curva}", x=s["Rank"], y=s["Unid_Abril"],
                             marker_color=cor, marker_line_width=0), secondary_y=False)
    fig.add_trace(go.Scatter(name="% Acumulado", x=pareto["Rank"], y=pareto["Pct_Acum"],
                             mode="lines", line=dict(color=CHOCO, width=2.5, dash="dot")),
                  secondary_y=True)
    fig.update_layout(**CL, barmode="stack", height=320, margin=dict(t=10,b=10))
    fig.update_yaxes(title_text="Unidades", secondary_y=False, gridcolor="#F5EAE0")
    fig.update_yaxes(title_text="% Acum.", secondary_y=True, range=[0,110], showgrid=False)
    st.plotly_chart(fig, use_container_width=True)

    # ABC Cards
    curva_sum = fa.groupby("Curva").agg(SKUs=("Modelo","count"), Un=("Unid_Abril","sum")).reset_index()
    c_abc = st.columns(3)
    desc = {"A":"Alto giro – prioridade máxima","B":"Médio giro – monitorar","C":"Baixo giro – revisar estoque"}
    cors = {"A": TERRA, "B": AMBER, "C": SAND}
    for col, curva in zip(c_abc, ["A","B","C"]):
        row = curva_sum[curva_sum["Curva"]==curva]
        if len(row):
            r = row.iloc[0]
            with col:
                st.markdown(f"""
                <div class="kpi" style="border-top:3px solid {cors[curva]}">
                    <div class="kpi-label">Curva {curva} — {desc[curva]}</div>
                    <div class="kpi-val">{int(r['SKUs'])} SKUs</div>
                    <div class="kpi-sub">{int(r['Un']):,} unidades vendidas</div>
                </div>""", unsafe_allow_html=True)

    st.markdown("---")
    col_e1, col_e2 = st.columns(2)

    with col_e1:
        sec("📅 Cobertura de Estoque por SKU – Top Vendas")
        top_v = fv.nlargest(20,"Total_Mes").copy()
        top_v["Cob"] = (tot_av / top_v["Media_Dia"].replace(0, 0.01)).round(1)
        top_v["SKU"] = top_v["Modelo"] + " · " + top_v["Cor"].str.capitalize()
        top_v["Alerta"] = top_v["Cob"].apply(
            lambda x: "Crítico" if x < 3 else ("Baixo" if x < 7 else "OK"))
        fig = px.bar(top_v.sort_values("Cob"), x="Cob", y="SKU", orientation="h",
                     color="Alerta",
                     color_discrete_map={"Crítico":DANGER,"Baixo":WARN,"OK":AMBER},
                     labels={"Cob":"Dias de cobertura","SKU":""})
        fig.add_vline(x=7, line_dash="dash", line_color=CHOCO, opacity=0.4,
                      annotation_text="Meta 7d")
        fig.update_layout(**CL, height=480, margin=dict(t=10,b=10))
        st.plotly_chart(fig, use_container_width=True)

    with col_e2:
        sec("🏆 Top 15 Vendas Abril 2026")
        top20 = fv.nlargest(15,"Total_Mes").copy()
        top20["SKU"] = top20["Modelo"] + " · " + top20["Cor"].str.capitalize()
        fig = px.bar(top20.sort_values("Total_Mes"), x="Total_Mes", y="SKU",
                     orientation="h", labels={"Total_Mes":"Unidades","SKU":""},
                     color="Total_Mes",
                     color_continuous_scale=[[0,"#F0E6DC"],[0.5,AMBER],[1,TERRA]])
        fig.update_layout(**CL, coloraxis_showscale=False, height=420, margin=dict(t=10,b=10))
        st.plotly_chart(fig, use_container_width=True)

        sec("Vendas por Tecido")
        vt = fv[fv["Tecido"].str.len()>2].groupby("Tecido")["Total_Mes"].sum().reset_index()
        fig = px.pie(vt, values="Total_Mes", names="Tecido", hole=0.55,
                     color_discrete_sequence=SEQ)
        fig.update_layout(**CLP, height=220, margin=dict(t=10,b=10,l=10,r=10))
        st.plotly_chart(fig, use_container_width=True)

    # Tabela ABC
    sec("Ranking ABC Completo")
    tbl_abc = fa[["Rank","Curva","Modelo","Tecido","Cor","Unid_Abril","Pct_Total","Pct_Acum"]].copy()
    tbl_abc.columns = ["#","Curva","Modelo","Tecido","Cor","Unid. Abril","% Total","% Acum."]
    st.dataframe(tbl_abc.sort_values("#"), use_container_width=True, hide_index=True, height=380)
    csv = tbl_abc.to_csv(index=False).encode("utf-8-sig")
    st.download_button("⬇️ Exportar ABC CSV", csv, "curva_abc.csv", "text/csv")
