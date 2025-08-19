# app.py — Streamlit: importar 1 planilha e analisar em gráficos
# Funciona com:
#  - 1 aba (consolidação) OU várias abas (ex.: Jan_2024, Fev_2024 etc.)
#  - Colunas esperadas: Segmento, Crédito, Débito, Líquido (se faltar Líquido, calcula = Crédito - Débito)
# Apenas upload, sem caminhos no disco.

import io
import re
import unicodedata
from pathlib import Path

import pandas as pd
import streamlit as st
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter

st.set_page_config(page_title="Análise de Arrecadação — Gráficos", layout="wide")

# ----------------- Helpers -----------------
PT_MONTHS = ["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"]

def strip_accents(s: str) -> str:
    if s is None: return ""
    return "".join(c for c in unicodedata.normalize("NFKD", str(s)) if not unicodedata.combining(c))

def norm_name(s: str) -> str:
    s = strip_accents(str(s)).strip().lower()
    s = re.sub(r"[^a-z0-9]+","_", s)
    return s.strip("_")

def fmt_brl(x, pos=None):
    try:
        return f"R$ {x:,.0f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return str(x)

def normalize_sheet(df: pd.DataFrame) -> pd.DataFrame:
    """Padroniza colunas e garante Líquido (Crédito - Débito) quando possível."""
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
        # Tenta achar uma coluna textual para virar "Segmento"
        text_col = None
        for c in df.columns:
            if df[c].dtype == object:
                text_col = c; break
        if text_col is None:
            return pd.DataFrame(columns=["Segmento","Crédito","Débito","Líquido"])
        df = df.rename(columns={text_col: "Segmento"})

    # Números
    for c in ["Crédito","Débito","Líquido"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")

    # Calcula líquido se faltou
    if "Líquido" not in df.columns:
        if "Crédito" in df.columns and "Débito" in df.columns:
            df["Líquido"] = df["Crédito"].fillna(0.0) - df["Débito"].fillna(0.0)
        else:
            df["Líquido"] = pd.to_numeric(df.get("Líquido", 0.0), errors="coerce").fillna(0.0)

    if "Crédito" not in df.columns: df["Crédito"] = 0.0
    if "Débito" not in df.columns:  df["Débito"]  = 0.0

    out = df[["Segmento","Crédito","Débito","Líquido"]].copy()
    out["Segmento"] = out["Segmento"].astype(str).str.strip()
    out[["Crédito","Débito","Líquido"]] = out[["Crédito","Débito","Líquido"]].fillna(0.0)
    return out

def guess_month_order(sheet_names):
    """Tenta ordenar abas por meses PT (Jan..Dez). Mantém nomes originais."""
    month_key = {m:i for i,m in enumerate(PT_MONTHS)}
    scored = []
    for name in sheet_names:
        # aceita "Jan", "Jan_2024", "Jan-2025", etc.
        m = re.match(r"^([A-Za-z]{3})", strip_accents(name).strip(), flags=re.IGNORECASE)
        if m:
            key = strip_accents(m.group(1)).title()
            if key in month_key:
                scored.append((month_key[key], name))
                continue
        # se não é mês reconhecido, manda pro fim mantendo ordem
        scored.append((999, name))
    scored.sort(key=lambda x: (x[0], sheet_names.index(x[1])))
    return [name for _, name in scored]

def total_by_month(sheets: dict) -> pd.DataFrame:
    """Soma Líquido por mês (uma linha por aba)."""
    rows = []
    for name, df in sheets.items():
        total = df["Líquido"].sum()
        rows.append({"Aba": name, "Total_Líquido": total})
    out = pd.DataFrame(rows)
    out["Aba"] = pd.Categorical(out["Aba"], categories=guess_month_order(out["Aba"].tolist()), ordered=True)
    return out.sort_values("Aba").reset_index(drop=True)

def pivot_segment_by_month(sheets: dict) -> pd.DataFrame:
    """Matriz Segmento x Mês (soma de Líquido)."""
    frames = []
    for name, df in sheets.items():
        tmp = df.groupby("Segmento", as_index=False)["Líquido"].sum()
        tmp["Mês"] = name
        frames.append(tmp)
    if not frames: return pd.DataFrame()
    mat = pd.concat(frames, ignore_index=True)
    mat["Mês"] = pd.Categorical(mat["Mês"], categories=guess_month_order(mat["Mês"].unique().tolist()), ordered=True)
    mat = mat.pivot_table(index="Segmento", columns="Mês", values="Líquido", aggfunc="sum", fill_value=0.0)
    return mat

def bar_by_segment(df: pd.DataFrame, title: str):
    fig, ax = plt.subplots(figsize=(10,6))
    ax.bar(df["Segmento"], df["Líquido"])
    ax.set_title(title)
    ax.set_xticklabels(df["Segmento"], rotation=45, ha="right")
    ax.yaxis.set_major_formatter(FuncFormatter(fmt_brl))
    fig.tight_layout()
    return fig

def line_total_by_month(df_month_totals: pd.DataFrame, title: str):
    fig, ax = plt.subplots(figsize=(10,5.5))
    ax.plot(df_month_totals["Aba"], df_month_totals["Total_Líquido"], marker="o")
    ax.set_title(title)
    ax.set_xticklabels(df_month_totals["Aba"], rotation=45, ha="right")
    ax.yaxis.set_major_formatter(FuncFormatter(fmt_brl))
    fig.tight_layout()
    return fig

def bar_stacked_top_segments(mat: pd.DataFrame, top_n: int = 6, title: str = "Top segmentos (Líquido)"):
    if mat.empty:
        fig, ax = plt.subplots(); ax.text(0.5,0.5,"Sem dados", ha="center"); return fig
    totals = mat.sum(axis=1).sort_values(ascending=False)
    top = totals.head(top_n).index
    mat_top = mat.loc[top]
    fig, ax = plt.subplots(figsize=(11,6.5))
    # Stacked bars por mês
    bottom = None
    for seg in mat_top.index:
        vals = mat_top.loc[seg].values
        if bottom is None:
            ax.bar(mat_top.columns, vals, label=seg)
            bottom = vals
        else:
            ax.bar(mat_top.columns, vals, bottom=bottom, label=seg)
            bottom = bottom + vals
    ax.set_title(title)
    ax.yaxis.set_major_formatter(FuncFormatter(fmt_brl))
    ax.legend(loc="upper left", bbox_to_anchor=(1,1))
    fig.tight_layout()
    return fig

# ----------------- UI -----------------
st.title("📈 Análise de Arrecadação — Upload de Planilha (Excel)")
st.write("Envie um único arquivo **.xlsx** com 1 ou várias abas mensais. O app calcula e gera **gráficos** automaticamente.")

uploaded = st.file_uploader("Planilha de arrecadação (.xlsx)", type=["xlsx","xls"])

if not uploaded:
    st.info("Envie seu arquivo para começar.")
    st.stop()

# Lê todas as abas e normaliza
try:
    raw_sheets = pd.read_excel(uploaded, sheet_name=None)
except Exception as e:
    st.error(f"Não consegui ler o Excel: {e}")
    st.stop()

norm_sheets = {}
for name, df in raw_sheets.items():
    nd = normalize_sheet(df)
    # Filtra abas vazias de verdade
    if not nd.empty and nd["Segmento"].notna().any():
        norm_sheets[name] = nd

if not norm_sheets:
    st.error("Não encontrei dados válidos (colunas Segmento/Crédito/Débito/Líquido).")
    st.stop()

# Ordena as abas por mês (quando possível)
ordered_sheet_names = guess_month_order(list(norm_sheets.keys()))
ordered_sheets = {name: norm_sheets[name] for name in ordered_sheet_names}

# Sidebar: seleção de aba e métrica
st.sidebar.header("Opções")
sel_sheet = st.sidebar.selectbox("Aba para gráfico de barras por segmento", ordered_sheet_names)
metric = st.sidebar.selectbox("Métrica", ["Líquido","Crédito","Débito"])

# 1) Barras por segmento (aba selecionada)
df_sel = ordered_sheets[sel_sheet].copy()
df_sel = df_sel.groupby("Segmento", as_index=False)[metric].sum().rename(columns={metric:"Líquido"})  # reusa função do gráfico
st.subheader(f"Barras por segmento — {sel_sheet} ({metric})")
st.pyplot(bar_by_segment(df_sel, f"{sel_sheet}: {metric} por segmento"))

# 2) Linha: total por mês (Líquido)
totals = total_by_month({k:v for k,v in ordered_sheets.items()})
st.subheader("Linha — Total líquido por mês (todas as abas)")
st.pyplot(line_total_by_month(totals, "Total líquido por mês"))

# 3) Barras empilhadas: top segmentos ao longo dos meses
mat = pivot_segment_by_month(ordered_sheets)
st.subheader("Barras empilhadas — Top segmentos ao longo dos meses (Líquido)")
st.pyplot(bar_stacked_top_segments(mat, top_n=6, title="Top 6 segmentos — Líquido (empilhado por mês)"))

# Download opcional: Excel normalizado + totais
with pd.ExcelWriter(io.BytesIO(), engine="openpyxl") as xw:
    # Monta um Excel em memória (não salva em disco)
    for name, df in ordered_sheets.items():
        df.to_excel(xw, index=False, sheet_name=name[:31])  # limite do Excel
    # Sheet com total por mês
    totals.to_excel(xw, index=False, sheet_name="Totais_por_mes")
    xw_bytes = xw.book.path  # isso não funciona, precisamos capturar o buffer
# Corrige: gerar bytes corretamente
buf = io.BytesIO()
with pd.ExcelWriter(buf, engine="openpyxl") as xw:
    for name, df in ordered_sheets.items():
        df.to_excel(xw, index=False, sheet_name=name[:31])
    totals.to_excel(xw, index=False, sheet_name="Totais_por_mes")
buf.seek(0)
st.download_button("⬇️ Baixar Excel normalizado + totais", data=buf.read(),
                   file_name="arrecadacao_normalizada.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
