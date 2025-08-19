# app.py — Streamlit: importar 1 planilha e analisar em gráficos
# - Só upload (.xlsx com 1+ abas)
# - Colunas: Segmento, Crédito, Débito, Líquido (se faltar Líquido, calcula = Crédito - Débito)
# - Sem 'args', sem caminhos fixos, sem 'get_model_order'

import re
import unicodedata
import pandas as pd
import streamlit as st
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter

st.set_page_config(page_title="Análise de Arrecadação — Gráficos", layout="wide")

PT_MONTHS = ["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"]

def strip_accents(s: str) -> str:
    if s is None: return ""
    return "".join(c for c in unicodedata.normalize("NFKD", str(s)) if not unicodedata.combining(c))

def norm_name(s: str) -> str:
    s = strip_accents(str(s)).strip().lower()
    return re.sub(r"[^a-z0-9]+","_", s).strip("_")

def fmt_brl(x, pos=None):
    try:
        return f"R$ {x:,.0f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return str(x)

def normalize_sheet(df: pd.DataFrame) -> pd.DataFrame:
    """Padroniza colunas e garante Líquido."""
    if df is None or df.empty:
        return pd.DataFrame(columns=["Segmento","Crédito","Débito","Líquido"])

    cols_map = {}
    for c in df.columns:
        n = norm_name(c)
        if n.startswith("segmento"): cols_map[c] = "Segmento"
        elif n.startswith("credito"): cols_map[c] = "Crédito"
        elif n.startswith("debito"):  cols_map[c] = "Débito"
        elif n.startswith("liquido"): cols_map[c] = "Líquido"
    df = df.rename(columns=cols_map)

    if "Segmento" not in df.columns:
        # tenta qualquer coluna textual
        for c in df.columns:
            if df[c].dtype == object:
                df = df.rename(columns={c: "Segmento"})
                break

    for c in ["Crédito","Débito","Líquido"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")

    if "Líquido" not in df.columns:
        if "Crédito" in df.columns and "Débito" in df.columns:
            df["Líquido"] = df["Crédito"].fillna(0.0) - df["Débito"].fillna(0.0)
        else:
            df["Líquido"] = pd.to_numeric(df.get("Líquido", 0.0), errors="coerce").fillna(0.0)

    if "Crédito" not in df.columns: df["Crédito"] = 0.0
    if "Débito"  not in df.columns: df["Débito"]  = 0.0

    out = df[["Segmento","Crédito","Débito","Líquido"]].copy()
    out["Segmento"] = out["Segmento"].astype(str).str.strip()
    return out.fillna(0.0)

def order_sheets_by_month(sheet_names):
    month_idx = {m:i for i,m in enumerate(PT_MONTHS)}
    scored = []
    for name in sheet_names:
        m = re.match(r"^([A-Za-z]{3})", strip_accents(name).strip(), flags=re.IGNORECASE)
        if m:
            key = strip_accents(m.group(1)).title()
            if key in month_idx:
                scored.append((month_idx[key], name))
                continue
        scored.append((999, name))
    scored.sort(key=lambda x: (x[0], sheet_names.index(x[1])))
    return [name for _, name in scored]

def chart_bar_by_segment(df: pd.DataFrame, month_label: str, metric: str):
    data = df.groupby("Segmento", as_index=False)[metric].sum()
    fig, ax = plt.subplots(figsize=(10,6))
    ax.bar(data["Segmento"], data[metric])
    ax.set_title(f"{month_label}: {metric} por segmento")
    ax.set_xticklabels(data["Segmento"], rotation=45, ha="right")
    ax.yaxis.set_major_formatter(FuncFormatter(fmt_brl))
    fig.tight_layout()
    return fig

def chart_line_totals(month_dfs: dict):
    rows = []
    for name, df in month_dfs.items():
        rows.append({"Mês": name, "Total_Líquido": df["Líquido"].sum()})
    tot = pd.DataFrame(rows)
    tot["Mês"] = pd.Categorical(tot["Mês"], categories=order_sheets_by_month(tot["Mês"].tolist()), ordered=True)
    tot = tot.sort_values("Mês")
    fig, ax = plt.subplots(figsize=(10,5.5))
    ax.plot(tot["Mês"], tot["Total_Líquido"], marker="o")
    ax.set_title("Total líquido por mês")
    ax.set_xticklabels(tot["Mês"], rotation=45, ha="right")
    ax.yaxis.set_major_formatter(FuncFormatter(fmt_brl))
    fig.tight_layout()
    return fig

def chart_stacked_top_segments(month_dfs: dict, top_n=6):
    # monta matriz Segmento x Mês (Líquido)
    frames = []
    for name, df in month_dfs.items():
        tmp = df.groupby("Segmento", as_index=False)["Líquido"].sum()
        tmp["Mês"] = name
        frames.append(tmp)
    if not frames:
        fig, ax = plt.subplots(); ax.text(0.5,0.5,"Sem dados", ha="center"); return fig
    mat = pd.concat(frames, ignore_index=True)
    mat["Mês"] = pd.Categorical(mat["Mês"], categories=order_sheets_by_month(mat["Mês"].unique().tolist()), ordered=True)
    mat = mat.pivot_table(index="Segmento", columns="Mês", values="Líquido", aggfunc="sum", fill_value=0.0)

    totals = mat.sum(axis=1).sort_values(ascending=False)
    top = totals.head(top_n).index
    mat_top = mat.loc[top]

    fig, ax = plt.subplots(figsize=(11,6.5))
    bottom = None
    for seg in mat_top.index:
        vals = mat_top.loc[seg].values
        if bottom is None:
            ax.bar(mat_top.columns, vals, label=seg)
            bottom = vals
        else:
            ax.bar(mat_top.columns, vals, bottom=bottom, label=seg)
            bottom = bottom + vals
    ax.set_title(f"Top {top_n} segmentos — Líquido (empilhado por mês)")
    ax.yaxis.set_major_formatter(FuncFormatter(fmt_brl))
    ax.legend(loc="upper left", bbox_to_anchor=(1,1))
    fig.tight_layout()
    return fig

# ---------------- UI ----------------
st.title("📈 Análise de Arrecadação — Upload único")
st.write("Envie um arquivo **.xlsx** com 1 ou várias abas. O app normaliza e gera gráficos automaticamente.")

up = st.file_uploader("Planilha (.xlsx)", type=["xlsx","xls"])
if not up:
    st.info("Envie sua planilha para começar.")
    st.stop()

try:
    raw = pd.read_excel(up, sheet_name=None)
except Exception as e:
    st.error(f"Falha ao ler o Excel: {e}")
    st.stop()

# Normaliza abas
clean = {}
for name, df in raw.items():
    nd = normalize_sheet(df)
    if not nd.empty and nd["Segmento"].notna().any():
        clean[name] = nd

if not clean:
    st.error("Não encontrei dados válidos (preciso de Segmento/Crédito/Débito/Líquido).")
    st.stop()

ordered_names = order_sheets_by_month(list(clean.keys()))
clean = {name: clean[name] for name in ordered_names}

# Sidebar
st.sidebar.header("Opções")
sel = st.sidebar.selectbox("Aba para gráfico de barras", ordered_names)
metric = st.sidebar.selectbox("Métrica", ["Líquido","Crédito","Débito"])

# Gráficos
st.subheader(f"Barras por segmento — {sel} ({metric})")
st.pyplot(chart_bar_by_segment(clean[sel], sel, metric))

st.subheader("Linha — Total líquido por mês")
st.pyplot(chart_line_totals(clean))

st.subheader("Barras empilhadas — Top segmentos (Líquido)")
st.pyplot(chart_stacked_top_segments(clean, top_n=6))
