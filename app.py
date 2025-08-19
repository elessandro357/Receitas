# app.py — Streamlit
# ✅ Sem caminhos fixos / Sem args / Com upload
# ✅ Modo 1: Análise simples (1 arquivo)
# ✅ Modo 2: Comparativo 2024 x 2025 (2 arquivos + ICMS opcional)
# ✅ Gráficos; no comparativo: Excel + PDF para download

import io
import re
import unicodedata
from typing import Dict, List
import pandas as pd
import streamlit as st
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter
from matplotlib.backends.backend_pdf import PdfPages

# ------------------ Config ------------------
st.set_page_config(page_title="Arrecadação — Análise e Comparativo", layout="wide")
PT_MONTHS = ["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"]
MONTHS_2024 = ["Jan_2024","Fev_2024","Mar_2024","Abr_2024","Mai_2024","Jun_2024"]
MONTHS_2025 = ["Jan_2025","Fev_2025","Mar_2025","Abr_2025","Mai_2025","Jun_2025"]
MODEL_FALLBACK = ["FPM","ICMS","FEP","ITR","CFM","FUS","CID","FEB","SNA","ADO"]

# ------------------ Helpers ------------------
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

def order_sheets_by_month(sheet_names: List[str]) -> List[str]:
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

def normalize_sheet(df: pd.DataFrame) -> pd.DataFrame:
    """Garante colunas Segmento, Crédito, Débito, Líquido; calcula Líquido se faltar."""
    if df is None or df.empty:
        return pd.DataFrame(columns=["Segmento","Crédito","Débito","Líquido"])

    # Normaliza nomes
    cols_map = {}
    for c in df.columns:
        n = norm_name(c)
        if n.startswith("segmento"): cols_map[c] = "Segmento"
        elif n.startswith("credito"): cols_map[c] = "Crédito"
        elif n.startswith("debito"):  cols_map[c] = "Débito"
        elif n.startswith("liquido"): cols_map[c] = "Líquido"
    df = df.rename(columns=cols_map)

    # Segmento
    if "Segmento" not in df.columns:
        # tenta usar a 1ª coluna textual como Segmento
        for c in df.columns:
            if df[c].dtype == object:
                df = df.rename(columns={c: "Segmento"})
                break
    if "Segmento" not in df.columns:
        return pd.DataFrame(columns=["Segmento","Crédito","Débito","Líquido"])

    # Números
    for c in ["Crédito","Débito","Líquido"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")

    # Líquido
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

def read_excel_all_sheets(uploaded_file) -> Dict[str, pd.DataFrame]:
    return pd.read_excel(uploaded_file, sheet_name=None)

# --------- ICMS parsing (opcional no comparativo) ----------
def parse_icms_map(icms_sheets: Dict[str, pd.DataFrame]) -> dict:
    """ICMS.xlsx (1ª aba) -> {year: {MonLabel: value}}"""
    if not icms_sheets: return {}
    name = list(icms_sheets.keys())[0]
    df = icms_sheets[name].copy()
    if df is None or df.empty: return {}
    df.columns = [strip_accents(str(c)).strip().lower() for c in df.columns]

    year_col = None
    for c in df.columns:
        if "arrecada" in c and "icms" in c:
            year_col = c; break
    if year_col is None:
        for c in df.columns:
            if df[c].astype(str).str.contains("2024|2025").any():
                year_col = c; break
    if year_col is None: return {}

    months = {
        "janeiro":"Jan","fevereiro":"Fev","marco":"Mar","março":"Mar","abril":"Abr","maio":"Mai",
        "junho":"Jun","julho":"Jul","agosto":"Ago","setembro":"Set","outubro":"Out","novembro":"Nov","dezembro":"Dez"
    }
    icms_map={}
    for _, row in df.iterrows():
        y = pd.to_numeric(row[year_col], errors="coerce")
        if pd.isna(y): continue
        y = int(y)
        per={}
        for c in df.columns:
            if c in months: per[months[c]] = row[c]
        icms_map[y]=per
    return icms_map

def upsert_icms_row(df: pd.DataFrame, value) -> pd.DataFrame:
    if value is None: return df
    val = float(pd.to_numeric(value, errors="coerce")) if pd.notna(value) else 0.0
    mask = df["Segmento"].astype(str).str.upper().str.strip() == "ICMS"
    if mask.any():
        i = df.index[mask][0]
        df.at[i,"Crédito"]=val; df.at[i,"Débito"]=0.0; df.at[i,"Líquido"]=val
    else:
        df = pd.concat([df, pd.DataFrame([{"Segmento":"ICMS","Crédito":val,"Débito":0.0,"Líquido":val}])], ignore_index=True)
    return df

# ----------------- Gráficos genéricos -----------------
def chart_bar_by_segment(df: pd.DataFrame, month_label: str, metric: str):
    data = df.groupby("Segmento", as_index=False)[metric].sum()
    fig, ax = plt.subplots(figsize=(10,6))
    ax.bar(data["Segmento"], data[metric])
    ax.set_title(f"{month_label}: {metric} por segmento")
    ax.set_xticklabels(data["Segmento"], rotation=45, ha="right")
    ax.yaxis.set_major_formatter(FuncFormatter(fmt_brl))
    fig.tight_layout()
    return fig

def chart_line_totals(month_dfs: dict, title="Total líquido por mês"):
    rows = [{"Mês": name, "Total_Líquido": df["Líquido"].sum()} for name, df in month_dfs.items()]
    tot = pd.DataFrame(rows)
    tot["Mês"] = pd.Categorical(tot["Mês"], categories=order_sheets_by_month(tot["Mês"].tolist()), ordered=True)
    tot = tot.sort_values("Mês")
    fig, ax = plt.subplots(figsize=(10,5.5))
    ax.plot(tot["Mês"], tot["Total_Líquido"], marker="o")
    ax.set_title(title)
    ax.set_xticklabels(tot["Mês"], rotation=45, ha="right")
    ax.yaxis.set_major_formatter(FuncFormatter(fmt_brl))
    fig.tight_layout()
    return fig

def chart_stacked_top_segments(month_dfs: dict, top_n=6, title="Top segmentos — Líquido (empilhado por mês)"):
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
    ax.set_title(title)
    ax.yaxis.set_major_formatter(FuncFormatter(fmt_brl))
    ax.legend(loc="upper left", bbox_to_anchor=(1,1))
    fig.tight_layout()
    return fig

# ----------------- Comparativo (2024 x 2025) -----------------
def find_model_order_from_2025(sheets_2025: dict) -> list:
    # prioriza aba "Resumo Jan-Jul 2025" (case-insensitive)
    target = None
    for name in sheets_2025.keys():
        if strip_accents(name).strip().lower() == strip_accents("Resumo Jan-Jul 2025").lower():
            target = name; break
    if target is not None and "Segmento" in sheets_2025[target].columns:
        order = [s for s in sheets_2025[target]["Segmento"].astype(str) if s.strip()]
        if "ICMS" not in order:
            if "FPM" in order:
                i = order.index("FPM")+1
                order = order[:i] + ["ICMS"] + order[i:]
            else:
                order.append("ICMS")
        return order

    # deduz pelas abas mensais 2025
    seen = []
    for m in MONTHS_2025:
        if m in sheets_2025 and "Segmento" in sheets_2025[m].columns:
            for s in sheets_2025[m]["Segmento"].astype(str):
                if s.strip() and s not in seen:
                    seen.append(s)
    if "ICMS" not in seen:
        if "FPM" in seen:
            i = seen.index("FPM")+1
            seen = seen[:i] + ["ICMS"] + seen[i:]
        else:
            seen.append("ICMS")
    return seen if seen else MODEL_FALLBACK

def load_month_liquid(sheets: dict, sheet_name: str) -> pd.DataFrame:
    if sheet_name not in sheets: return pd.DataFrame(columns=["Segmento","Líquido"])
    return normalize_sheet(sheets[sheet_name])[["Segmento","Líquido"]]

def compare_month(sheets_2024: dict, sheets_2025: dict, m24: str, m25: str, order: list) -> pd.DataFrame:
    df24 = load_month_liquid(sheets_2024, m24)
    df25 = load_month_liquid(sheets_2025, m25)
    all_segs = list(dict.fromkeys(order + df24["Segmento"].dropna().tolist() + df25["Segmento"].dropna().tolist()))
    base = pd.DataFrame({"Segmento": all_segs})
    merged = (
        base.merge(df24, on="Segmento", how="left").rename(columns={"Líquido":"2024_Líquido"})
            .merge(df25.rename(columns={"Líquido":"2025_Líquido"}), on="Segmento", how="left")
            .fillna({"2024_Líquido":0.0,"2025_Líquido":0.0})
    )
    merged["Dif_abs"] = merged["2025_Líquido"] - merged["2024_Líquido"]
    merged["Dif_%"] = merged.apply(lambda r: (r["Dif_abs"]/r["2024_Líquido"]*100.0) if r["2024_Líquido"] else None, axis=1)
    extras = [s for s in merged["Segmento"].tolist() if s not in order]
    final_order = order + extras
    merged["__order"] = merged["Segmento"].apply(lambda s: final_order.index(s) if s in final_order else 999)
    return merged.sort_values("__order").drop(columns="__order").reset_index(drop=True)

def sum_semester(sheets: dict, months: list) -> pd.DataFrame:
    acc=None
    for m in months:
        dfm = load_month_liquid(sheets, m)
        if acc is None: acc=dfm.copy()
        else:
            acc = acc.merge(dfm, on="Segmento", how="outer", suffixes=("","_tmp"))
            acc["Líquido"] = acc[["Líquido","Líquido_tmp"]].fillna(0.0).sum(axis=1)
            acc = acc.drop(columns=[c for c in acc.columns if c.endswith("_tmp")])
    if acc is None: acc = pd.DataFrame(columns=["Segmento","Líquido"])
    return acc

def chart_semester_bar(sem_df: pd.DataFrame):
    fig, ax = plt.subplots(figsize=(11, 6.5))
    x = range(len(sem_df))
    w = 0.4
    ax.bar([i-w/2 for i in x], sem_df["2024_Jan-Jun"], width=w, label="2024")
    ax.bar([i+w/2 for i in x], sem_df["2025_Jan-Jun"], width=w, label="2025")
    ax.set_xticks(list(x)); ax.set_xticklabels(sem_df["Segmento"], rotation=45, ha="right")
    ax.yaxis.set_major_formatter(FuncFormatter(fmt_brl)); ax.set_title("Líquido por Segmento – 1º Semestre (2024 vs 2025)")
    ax.legend(); fig.tight_layout(); return fig

def chart_top(sem_df: pd.DataFrame, title: str, top=True):
    d = sem_df.sort_values("Dif_abs", ascending=False)
    d = d.head(5) if top else d.tail(5).sort_values("Dif_abs")
    fig, ax = plt.subplots(figsize=(9,5.5))
    ax.barh(d["Segmento"], d["Dif_abs"]); ax.xaxis.set_major_formatter(FuncFormatter(fmt_brl))
    ax.set_title(title); fig.tight_layout(); return fig

def chart_month_grouped(month_name: str, df: pd.DataFrame):
    fig, ax = plt.subplots(figsize=(11,6.5))
    segs=df["Segmento"].tolist(); x=range(len(segs)); w=0.4
    ax.bar([i-w/2 for i in x], df["2024_Líquido"].tolist(), width=w, label="2024")
    ax.bar([i+w/2 for i in x], df["2025_Líquido"].tolist(), width=w, label="2025")
    ax.set_xticks(list(x)); ax.set_xticklabels(segs, rotation=45, ha="right"); ax.yaxis.set_major_formatter(FuncFormatter(fmt_brl))
    ax.set_title(f"{month_name}: Líquido por Segmento"); ax.legend(); fig.tight_layout(); return fig

def build_pdf(sem_df: pd.DataFrame, month_comp: dict) -> bytes:
    buf = io.BytesIO()
    with PdfPages(buf) as pdf:
        total24 = sem_df["2024_Jan-Jun"].sum()
        total25 = sem_df["2025_Jan-Jun"].sum()
        delta = total25 - total24
        pct = (delta/total24*100.0) if total24 else 0.0
        fig0, ax0 = plt.subplots(figsize=(11.69,8.27)); ax0.axis("off")
        y=0.9
        ax0.text(0.05,y,"Relatório Comparativo 1º Semestre – 2024 x 2025",fontsize=18,weight="bold"); y-=0.08
        ax0.text(0.05,y,f"Soma Jan–Jun 2024: {fmt_brl(total24)}",fontsize=12); y-=0.05
        ax0.text(0.05,y,f"Soma Jan–Jun 2025: {fmt_brl(total25)}",fontsize=12); y-=0.05
        ax0.text(0.05,y,f"Variação: {fmt_brl(delta)} ({pct:.2f}%)",fontsize=12)
        fig0.tight_layout(); pdf.savefig(fig0); plt.close(fig0)

        pdf.savefig(chart_semester_bar(sem_df)); plt.close()
        pdf.savefig(chart_top(sem_df,"Top 5 Crescimentos (Dif_abs – Jan–Jun)",True)); plt.close()
        pdf.savefig(chart_top(sem_df,"Top 5 Quedas (Dif_abs – Jan–Jun)",False)); plt.close()
        pdf.savefig(chart_line_totals({k:v for k,v in month_comp.items()}, "Totais Mensais – 1º Semestre"))

        for m in ["Jan","Fev","Mar","Abr","Mai","Jun"]:
            pdf.savefig(chart_month_grouped(m, month_comp[m])); plt.close()
    buf.seek(0); return buf.read()

def build_excel(sem_df: pd.DataFrame, month_comp: dict) -> bytes:
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as xw:
        sem_df.to_excel(xw, index=False, sheet_name="Resumo_1S")
        for name, df in month_comp.items():
            df.to_excel(xw, index=False, sheet_name=name)
    out.seek(0); return out.read()

# ------------------ UI ------------------
st.title("📊 Arrecadação — Análise e Comparativo")

modo = st.radio("Escolha o modo:", ["Análise simples (1 arquivo)", "Comparativo 2024 x 2025 (2 arquivos)"], horizontal=True)

if modo == "Análise simples (1 arquivo)":
    up = st.file_uploader("Planilha (.xlsx) — 1 ou várias abas mensais", type=["xlsx","xls"])
    if not up:
        st.info("Envie sua planilha para começar.")
        st.stop()

    try:
        raw = pd.read_excel(up, sheet_name=None)
    except Exception as e:
        st.error(f"Falha ao ler o Excel: {e}")
        st.stop()

    clean = {}
    for name, df in raw.items():
        nd = normalize_sheet(df)
        if not nd.empty and nd["Segmento"].notna().any():
            clean[name] = nd
    if not clean:
        st.error("Não encontrei dados válidos (preciso de Segmento/Crédito/Débito/Líquido).")
        st.stop()

    ordered = order_sheets_by_month(list(clean.keys()))
    clean = {name: clean[name] for name in ordered}

    st.sidebar.header("Opções")
    sel = st.sidebar.selectbox("Aba para barras por segmento", ordered)
    metric = st.sidebar.selectbox("Métrica", ["Líquido","Crédito","Débito"])

    st.subheader(f"Barras por segmento — {sel} ({metric})")
    st.pyplot(chart_bar_by_segment(clean[sel], sel, metric))

    st.subheader("Linha — Total líquido por mês")
    st.pyplot(chart_line_totals(clean))

    st.subheader("Barras empilhadas — Top segmentos (Líquido)")
    st.pyplot(chart_stacked_top_segments(clean, top_n=6))

    # Download Excel normalizado + totais
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        for name, df in clean.items():
            df.to_excel(xw, index=False, sheet_name=name[:31])
        # totais
        rows = [{"Mês": n, "Total_Líquido": df["Líquido"].sum()} for n, df in clean.items()]
        pd.DataFrame(rows).to_excel(xw, index=False, sheet_name="Totais_por_mes")
    buf.seek(0)
    st.download_button("⬇️ Baixar Excel normalizado + totais",
                       data=buf.read(),
                       file_name="arrecadacao_normalizada.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

else:
    c1, c2, c3 = st.columns([1,1,1])
    with c1: up2024 = st.file_uploader("Excel 2024 (abas Jan_2024..Jun_2024)", type=["xlsx","xls"], key="f2024")
    with c2: up2025 = st.file_uploader("Excel 2025 (abas Jan_2025..Jun_2025)", type=["xlsx","xls"], key="f2025")
    with c3: upICMS = st.file_uploader("ICMS.xlsx (opcional)", type=["xlsx","xls"], key="ficms")

    if not st.button("Gerar comparativo"):
        st.stop()

    if not up2024 or not up2025:
        st.error("Envie os dois arquivos: 2024 e 2025.")
        st.stop()

    try:
        sheets_2024 = read_excel_all_sheets(up2024)
        sheets_2025 = read_excel_all_sheets(up2025)
    except Exception as e:
        st.error(f"Falha ao ler os Excel: {e}")
        st.stop()

    # Normaliza apenas as abas que importam (Jan..Jun)
    norm24 = {m: normalize_sheet(sheets_2024[m]) for m in MONTHS_2024 if m in sheets_2024}
    norm25 = {m: normalize_sheet(sheets_2025[m]) for m in MONTHS_2025 if m in sheets_2025}

    if not norm24 or not norm25:
        st.error("Faltam abas mensais esperadas (Jan..Jun_2024 e Jan..Jun_2025).")
        st.stop()

    # ICMS opcional
    if upICMS is not None:
        try:
            icms_map = parse_icms_map(read_excel_all_sheets(upICMS))
        except Exception as e:
            st.warning(f"Não consegui ler ICMS.xlsx ({e}). Prosseguindo sem ICMS extra.")
            icms_map = {}
        for m in MONTHS_2024:
            if m in norm24 and 2024 in icms_map:
                mon = m.split("_")[0]
                norm24[m] = upsert_icms_row(norm24[m], icms_map[2024].get(mon))
        for m in MONTHS_2025:
            if m in norm25 and 2025 in icms_map:
                mon = m.split("_")[0]
                norm25[m] = upsert_icms_row(norm25[m], icms_map[2025].get(mon))

    # Ordem de segmentos (preferindo "Resumo Jan-Jul 2025" se existir)
    model_order = find_model_order_from_2025(sheets_2025)

    # Comparativos mensais
    month_comp = {}
    for m24, m25, label in zip(MONTHS_2024, MONTHS_2025, ["Jan","Fev","Mar","Abr","Mai","Jun"]):
        df = compare_month(norm24, norm25, m24, m25, model_order)
        month_comp[label] = df

    # Resumo semestre
    sem24 = sum_semester(norm24, MONTHS_2024).rename(columns={"Líquido":"2024_Jan-Jun"})
    sem25 = sum_semester(norm25, MONTHS_2025).rename(columns={"Líquido":"2025_Jan-Jun"})
    base = pd.DataFrame({"Segmento": list(dict.fromkeys(model_order + sem24["Segmento"].dropna().tolist() + sem25["Segmento"].dropna().tolist()))})
    sem_df = (base.merge(sem24, on="Segmento", how="left")
                   .merge(sem25, on="Segmento", how="left")
                   .fillna({"2024_Jan-Jun":0.0,"2025_Jan-Jun":0.0}))
    sem_df["Dif_abs"] = sem_df["2025_Jan-Jun"] - sem_df["2024_Jan-Jun"]
    sem_df["Dif_%"] = sem_df.apply(lambda r: (r["Dif_abs"]/r["2024_Jan-Jun"]*100.0) if r["2024_Jan-Jun"] else None, axis=1)
    extras = [s for s in sem_df["Segmento"].tolist() if s not in model_order]
    final_order = model_order + extras
    sem_df["__order"] = sem_df["Segmento"].apply(lambda s: final_order.index(s) if s in final_order else 999)
    sem_df = sem_df.sort_values("__order").drop(columns="__order").reset_index(drop=True)

    # Exibição
    st.subheader("Resumo 1º Semestre (Jan–Jun)")
    st.dataframe(sem_df, use_container_width=True)

    c1,c2 = st.columns(2)
    with c1:
        st.pyplot(chart_semester_bar(sem_df))
        st.pyplot(chart_line_totals(month_comp, "Totais Mensais – 1º Semestre"))
    with c2:
        st.pyplot(chart_top(sem_df,"Top 5 Crescimentos (Dif_abs – Jan–Jun)",True))
        st.pyplot(chart_top(sem_df,"Top 5 Quedas (Dif_abs – Jan–Jun)",False))

    for m in ["Jan","Fev","Mar","Abr","Mai","Jun"]:
        st.pyplot(chart_month_grouped(m, month_comp[m]))

    # Downloads
    xlsx_bytes = build_excel(sem_df, month_comp)
    pdf_bytes  = build_pdf(sem_df, month_comp)
    st.download_button("⬇️ Baixar Excel (Resumo_1S + Jan..Jun)", data=xlsx_bytes,
                       file_name="comparativo_receitas_2024_2025_1S.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.download_button("⬇️ Baixar PDF (gráficos)", data=pdf_bytes,
                       file_name="relatorio_comparativo_1S_2024_2025.pdf",
                       mime="application/pdf")
