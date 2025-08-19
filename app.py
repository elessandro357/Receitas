# app.py
# ------------------------------------------------------------
# Streamlit • Comparativo de Receitas (Líquido) 2024 x 2025 • Jan–Jun
# - Upload de arquivos Excel (2024 e 2025) e opcional ICMS.xlsx
# - Mantém ordem de segmentos da aba "Resumo Jan-Jul 2025" (se existir)
# - Injeta/atualiza ICMS por mês/ano se for enviado ICMS.xlsx
# - Gera Excel, gráficos (PNGs) e PDF para download
# ------------------------------------------------------------
import io
import unicodedata
from pathlib import Path

import pandas as pd
import streamlit as st
import matplotlib
matplotlib.use("Agg")  # backend headless
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter
from matplotlib.backends.backend_pdf import PdfPages

st.set_page_config(layout="wide", page_title="Comparativo 2024 x 2025 (Jan–Jun)")

# ================================
# Helpers
# ================================
MONTHS_2024 = ["Jan_2024","Fev_2024","Mar_2024","Abr_2024","Mai_2024","Jun_2024"]
MONTHS_2025 = ["Jan_2025","Fev_2025","Mar_2025","Abr_2025","Mai_2025","Jun_2025"]
MONTH_LABELS = ["Jan","Fev","Mar","Abr","Mai","Jun"]
MODEL_FALLBACK = ["FPM","ICMS","FEP","ITR","CFM","FUS","CID","FEB","SNA","ADO"]

def strip_accents(s: str) -> str:
    if s is None:
        return ""
    return "".join(c for c in unicodedata.normalize("NFKD", str(s)) if not unicodedata.combining(c))

def brl_formatter(x, pos):
    return f"R$ {x:,.0f}".replace(",", "X").replace(".", ",").replace("X", ".")

def read_excel_all_sheets(file_like) -> dict:
    """Lê todas as abas de um Excel (UploadedFile/BytesIO/path) → {sheet_name: DataFrame}"""
    return pd.read_excel(file_like, sheet_name=None)

def find_model_order(sheets_2025: dict) -> list:
    """Ordem de segmentos: prioriza aba 'Resumo Jan-Jul 2025'; fallback: une segmentos das abas Jan..Jun_2025; garante ICMS."""
    # Procura aba "Resumo Jan-Jul 2025" (case-insensitive)
    target = None
    for name in sheets_2025.keys():
        if strip_accents(name).strip().lower() == strip_accents("Resumo Jan-Jul 2025").lower():
            target = name
            break
    if target is not None:
        df = sheets_2025[target]
        if "Segmento" in df.columns:
            order = [s for s in df["Segmento"].astype(str).tolist() if s and str(s).strip()]
            if "ICMS" not in order:
                if "FPM" in order:
                    i = order.index("FPM") + 1
                    order = order[:i] + ["ICMS"] + order[i:]
                else:
                    order.append("ICMS")
            return order

    # Sem a aba resumo → deduz pelas abas mensais 2025
    seen = []
    for m in MONTHS_2025:
        if m in sheets_2025:
            dfm = sheets_2025[m]
            if "Segmento" in dfm.columns:
                for s in dfm["Segmento"].astype(str).tolist():
                    if s and str(s).strip() and s not in seen:
                        seen.append(s)
    if "ICMS" not in seen:
        if "FPM" in seen:
            i = seen.index("FPM") + 1
            seen = seen[:i] + ["ICMS"] + seen[i:]
        else:
            seen.append("ICMS")
    return seen if seen else MODEL_FALLBACK

def normalize_month_df(df: pd.DataFrame) -> pd.DataFrame:
    """Normaliza colunas para ter Segmento, Crédito, Débito, Líquido (calculando se necessário)."""
    if df is None or df.empty:
        return pd.DataFrame(columns=["Segmento","Crédito","Débito","Líquido"])

    # Renomeia colunas ignorando acentos
    rename_map = {}
    for c in df.columns:
        cl = strip_accents(str(c)).strip().lower()
        if cl.startswith("segmento"):
            rename_map[c] = "Segmento"
        elif cl.startswith("liquido") or cl.startswith("líquido"):
            rename_map[c] = "Líquido"
        elif cl.startswith("credito") or cl.startswith("crédito"):
            rename_map[c] = "Crédito"
        elif cl.startswith("debito") or cl.startswith("débito"):
            rename_map[c] = "Débito"
    df = df.rename(columns=rename_map)

    if "Segmento" not in df.columns:
        return pd.DataFrame(columns=["Segmento","Crédito","Débito","Líquido"])

    # Se possuir crédito/débito mas não tiver líquido, calcula
    if "Líquido" not in df.columns:
        cred_col = next((c for c in df.columns if strip_accents(str(c)).lower().startswith("credito")), None)
        deb_col  = next((c for c in df.columns if strip_accents(str(c)).lower().startswith("debito")), None)
        if cred_col and deb_col:
            df["Crédito"] = pd.to_numeric(df[cred_col], errors="coerce").fillna(0.0)
            df["Débito"]  = pd.to_numeric(df[deb_col],  errors="coerce").fillna(0.0)
            df["Líquido"] = df["Crédito"] - df["Débito"]
        else:
            # Sem como calcular
            df["Crédito"] = pd.to_numeric(df.get("Crédito", 0.0), errors="coerce").fillna(0.0)
            df["Débito"]  = pd.to_numeric(df.get("Débito", 0.0),  errors="coerce").fillna(0.0)
            df["Líquido"] = pd.to_numeric(df.get("Líquido", 0.0), errors="coerce").fillna(0.0)
    else:
        for col in ["Crédito","Débito","Líquido"]:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)
            else:
                df[col] = 0.0

    return df[["Segmento","Crédito","Débito","Líquido"]]

def upsert_icms_row(df: pd.DataFrame, value) -> pd.DataFrame:
    """Insere/atualiza linha ICMS com Crédito=value, Débito=0, Líquido=value (se value não for None)."""
    if value is None:
        return df
    val = float(pd.to_numeric(value, errors="coerce")) if pd.notna(value) else 0.0
    mask = df["Segmento"].astype(str).str.upper().str.strip() == "ICMS"
    if mask.any():
        idx = df.index[mask][0]
        df.at[idx, "Crédito"] = val
        df.at[idx, "Débito"]  = 0.0
        df.at[idx, "Líquido"] = val
    else:
        df = pd.concat([df, pd.DataFrame([{"Segmento":"ICMS","Crédito":val,"Débito":0.0,"Líquido":val}])], ignore_index=True)
    return df

def parse_icms_map(icms_sheets: dict) -> dict:
    """Lê o 1º sheet do ICMS.xlsx → {year: {MonLabel: value}} com colunas 'janeiro'..'dezembro'."""
    if not icms_sheets:
        return {}
    # pega 1ª aba
    name = list(icms_sheets.keys())[0]
    df = icms_sheets[name].copy()
    if df is None or df.empty:
        return {}

    # normaliza colunas
    df.columns = [strip_accents(str(c)).strip().lower() for c in df.columns]
    # acha coluna do ano
    year_col = None
    for c in df.columns:
        if "arrecada" in c and "icms" in c:
            year_col = c
            break
    if year_col is None:
        # fallback: primeira coluna com números 2024/2025
        for c in df.columns:
            if df[c].astype(str).str.contains("2024|2025").any():
                year_col = c
                break
    if year_col is None:
        return {}

    month_cols = {
        "janeiro": "Jan", "fevereiro": "Fev", "marco": "Mar", "março": "Mar",
        "abril": "Abr", "maio": "Mai", "junho": "Jun", "julho": "Jul",
        "agosto": "Ago", "setembro": "Set", "outubro": "Out",
        "novembro": "Nov", "dezembro": "Dez",
    }
    icms_map = {}
    for _, row in df.iterrows():
        try:
            year = int(pd.to_numeric(row[year_col], errors="coerce"))
        except Exception:
            continue
        per_month = {}
        for c in df.columns:
            if c in month_cols:
                per_month[month_cols[c]] = row[c]
        icms_map[year] = per_month
    return icms_map

def load_month_liquid(sheets: dict, sheet_name: str) -> pd.DataFrame:
    if sheet_name not in sheets:
        return pd.DataFrame(columns=["Segmento","Líquido"])
    df = normalize_month_df(sheets[sheet_name])
    return df[["Segmento","Líquido"]]

def compare_month(sheets_2024: dict, sheets_2025: dict, m24: str, m25: str, order: list) -> pd.DataFrame:
    df24 = load_month_liquid(sheets_2024, m24)
    df25 = load_month_liquid(sheets_2025, m25)
    all_segments = list(dict.fromkeys(order + df24["Segmento"].dropna().tolist() + df25["Segmento"].dropna().tolist()))
    base = pd.DataFrame({"Segmento": all_segments})
    merged = (
        base.merge(df24, on="Segmento", how="left")
            .rename(columns={"Líquido":"2024_Líquido"})
            .merge(df25.rename(columns={"Líquido":"2025_Líquido"}), on="Segmento", how="left")
            .fillna({"2024_Líquido":0.0,"2025_Líquido":0.0})
    )
    merged["Dif_abs"] = merged["2025_Líquido"] - merged["2024_Líquido"]
    merged["Dif_%"] = merged.apply(lambda r: (r["Dif_abs"] / r["2024_Líquido"] * 100.0) if r["2024_Líquido"] else None, axis=1)
    # Ordena por ordem do modelo + extras
    extras = [s for s in merged["Segmento"].tolist() if s not in order]
    final_order = order + extras
    merged["__order"] = merged["Segmento"].apply(lambda s: final_order.index(s) if s in final_order else 999)
    merged = merged.sort_values("__order").drop(columns="__order").reset_index(drop=True)
    return merged

def sum_semester(sheets: dict, months: list) -> pd.DataFrame:
    acc = None
    for m in months:
        dfm = load_month_liquid(sheets, m)
        if acc is None:
            acc = dfm.copy()
        else:
            acc = acc.merge(dfm, on="Segmento", how="outer", suffixes=("","_tmp"))
            acc["Líquido"] = acc[["Líquido","Líquido_tmp"]].fillna(0.0).sum(axis=1)
            acc = acc.drop(columns=[c for c in acc.columns if c.endswith("_tmp")])
    if acc is None:
        acc = pd.DataFrame(columns=["Segmento","Líquido"])
    return acc

# ============ Charts ============
def chart_semester_bar(sem_df: pd.DataFrame):
    fig, ax = plt.subplots(figsize=(11, 6.5))
    segments = sem_df["Segmento"].tolist()
    x = list(range(len(segments)))
    width = 0.4
    ax.bar([i - width/2 for i in x], sem_df["2024_Jan-Jun"].tolist(), width=width, label="2024")
    ax.bar([i + width/2 for i in x], sem_df["2025_Jan-Jun"].tolist(), width=width, label="2025")
    ax.set_xticks(x)
    ax.set_xticklabels(segments, rotation=45, ha="right")
    ax.yaxis.set_major_formatter(FuncFormatter(brl_formatter))
    ax.set_title("Líquido por Segmento – 1º Semestre (2024 vs 2025)")
    ax.legend()
    fig.tight_layout()
    return fig

def chart_top_changes(sem_df: pd.DataFrame, title: str, top: bool = True):
    df = sem_df.sort_values("Dif_abs", ascending=False).reset_index(drop=True)
    data = df.head(5) if top else df.tail(5).sort_values("Dif_abs")
    fig, ax = plt.subplots(figsize=(9, 5.5))
    ax.barh(data["Segmento"], data["Dif_abs"])
    ax.xaxis.set_major_formatter(FuncFormatter(brl_formatter))
    ax.set_title(title)
    fig.tight_layout()
    return fig

def chart_monthly_totals(month_comp: dict):
    totals24 = []
    totals25 = []
    for m in MONTH_LABELS:
        df = month_comp[m]
        totals24.append(df["2024_Líquido"].sum())
        totals25.append(df["2025_Líquido"].sum())
    fig, ax = plt.subplots(figsize=(10.5, 5.5))
    ax.plot(MONTH_LABELS, totals24, marker="o", label="2024")
    ax.plot(MONTH_LABELS, totals25, marker="o", label="2025")
    ax.yaxis.set_major_formatter(FuncFormatter(brl_formatter))
    ax.set_title("Totais Mensais – 1º Semestre (2024 vs 2025)")
    ax.legend()
    fig.tight_layout()
    return fig

def chart_month_grouped_bars(month_name: str, df: pd.DataFrame):
    fig, ax = plt.subplots(figsize=(11, 6.5))
    segs = df["Segmento"].tolist()
    x = list(range(len(segs)))
    width = 0.4
    ax.bar([i - width/2 for i in x], df["2024_Líquido"].tolist(), width=width, label="2024")
    ax.bar([i + width/2 for i in x], df["2025_Líquido"].tolist(), width=width, label="2025")
    ax.set_xticks(x)
    ax.set_xticklabels(segs, rotation=45, ha="right")
    ax.yaxis.set_major_formatter(FuncFormatter(brl_formatter))
    ax.set_title(f"{month_name}: Líquido por Segmento (2024 vs 2025)")
    ax.legend()
    fig.tight_layout()
    return fig

def build_pdf(sem_df: pd.DataFrame, month_comp: dict) -> bytes:
    buf = io.BytesIO()
    with PdfPages(buf) as pdf:
        # Capa
        total24 = sem_df["2024_Jan-Jun"].sum()
        total25 = sem_df["2025_Jan-Jun"].sum()
        delta = total25 - total24
        pct = (delta / total24 * 100.0) if total24 else 0.0

        fig0, ax0 = plt.subplots(figsize=(11.69, 8.27))
        ax0.axis("off")
        y = 0.9
        ax0.text(0.05, y, "Relatório Comparativo 1º Semestre – 2024 x 2025", fontsize=18, weight="bold")
        y -= 0.08
        ax0.text(0.05, y, f"Soma Jan–Jun 2024: {brl_formatter(total24, None)}", fontsize=12)
        y -= 0.05
        ax0.text(0.05, y, f"Soma Jan–Jun 2025: {brl_formatter(total25, None)}", fontsize=12)
        y -= 0.05
        ax0.text(0.05, y, f"Variação: {brl_formatter(delta, None)} ({pct:.2f}%)", fontsize=12)
        fig0.tight_layout()
        pdf.savefig(fig0)
        plt.close(fig0)

        # Gráficos principais
        pdf.savefig(chart_semester_bar(sem_df))
        pdf.savefig(chart_top_changes(sem_df, "Top 5 Crescimentos (Dif_abs – Jan–Jun)", top=True))
        pdf.savefig(chart_top_changes(sem_df, "Top 5 Quedas (Dif_abs – Jan–Jun)", top=False))
        pdf.savefig(chart_monthly_totals(month_comp))

        # Um gráfico por mês
        for label in MONTH_LABELS:
            pdf.savefig(chart_month_grouped_bars(label, month_comp[label]))
    buf.seek(0)
    return buf.read()

def build_excel(sem_df: pd.DataFrame, month_comp: dict) -> bytes:
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as xw:
        sem_df.to_excel(xw, index=False, sheet_name="Resumo_1S")
        for name, df in month_comp.items():
            df.to_excel(xw, index=False, sheet_name=name)
    out.seek(0)
    return out.read()

# ================================
# UI
# ================================
st.title("📊 Comparativo de Receitas (Líquido) — 2024 x 2025 — Jan–Jun")
st.caption("Envie os arquivos abaixo. O ICMS.xlsx é opcional e, se enviado, será injetado/atualizado nos meses correspondentes.")

col1, col2, col3 = st.columns([1,1,1])
with col1:
    file_2024 = st.file_uploader("Excel 2024 (com abas Jan_2024..Jun_2024)", type=["xlsx","xls"], key="file2024")
with col2:
    file_2025 = st.file_uploader("Excel 2025 (ordem/abas Jan_2025..Jun_2025)", type=["xlsx","xls"], key="file2025")
with col3:
    file_icms = st.file_uploader("ICMS.xlsx (opcional)", type=["xlsx","xls"], key="fileicms")

process_btn = st.button("Gerar comparativo (inclui gráficos e PDF)")

if process_btn:
    if not file_2024 or not file_2025:
        st.error("Envie os dois arquivos: 2024 e 2025.")
        st.stop()

    try:
        sheets_2024 = read_excel_all_sheets(file_2024)
        sheets_2025 = read_excel_all_sheets(file_2025)
    except Exception as e:
        st.error(f"Falha ao ler os Excel enviados: {e}")
        st.stop()

    # Ordem de segmentos (modelo 2025)
    model_order = find_model_order(sheets_2025)

    # Se veio ICMS.xlsx, injeta ICMS nas abas mensais correspondentes
    if file_icms is not None:
        try:
            icms_sheets = read_excel_all_sheets(file_icms)
            icms_map = parse_icms_map(icms_sheets)
        except Exception as e:
            st.warning(f"Não consegui ler ICMS.xlsx ({e}). Vou seguir sem ICMS adicional.")
            icms_map = {}
        # Atualiza 2024
        for sheet in MONTHS_2024:
            if sheet in sheets_2024 and 2024 in icms_map:
                mon = sheet.split("_")[0]  # Jan, Fev, ...
                dfm = normalize_month_df(sheets_2024[sheet])
                dfm = upsert_icms_row(dfm, icms_map[2024].get(mon))
                sheets_2024[sheet] = dfm
        # Atualiza 2025
        for sheet in MONTHS_2025:
            if sheet in sheets_2025 and 2025 in icms_map:
                mon = sheet.split("_")[0]
                dfm = normalize_month_df(sheets_2025[sheet])
                dfm = upsert_icms_row(dfm, icms_map[2025].get(mon))
                sheets_2025[sheet] = dfm

    # Comparativos mensais
    month_comparisons = {}
    for m24, m25, label in zip(MONTHS_2024, MONTHS_2025, MONTH_LABELS):
        df = compare_month(sheets_2024, sheets_2025, m24, m25, model_order)
        month_comparisons[label] = df

    # Resumo semestre
    sem24 = sum_semester(sheets_2024, MONTHS_2024).rename(columns={"Líquido":"2024_Jan-Jun"})
    sem25 = sum_semester(sheets_2025, MONTHS_2025).rename(columns={"Líquido":"2025_Jan-Jun"})
    base = pd.DataFrame({"Segmento": list(dict.fromkeys(model_order + sem24["Segmento"].dropna().tolist() + sem25["Segmento"].dropna().tolist()))})
    sem_df = base.merge(sem24, on="Segmento", how="left").merge(sem25, on="Segmento", how="left").fillna({"2024_Jan-Jun":0.0,"2025_Jan-Jun":0.0})
    sem_df["Dif_abs"] = sem_df["2025_Jan-Jun"] - sem_df["2024_Jan-Jun"]
    sem_df["Dif_%"] = sem_df.apply(lambda r: (r["Dif_abs"] / r["2024_Jan-Jun"] * 100.0) if r["2024_Jan-Jun"] else None, axis=1)
    # Ordena
    extras = [s for s in sem_df["Segmento"].tolist() if s not in model_order]
    final_order = model_order + extras
    sem_df["__order"] = sem_df["Segmento"].apply(lambda s: final_order.index(s) if s in final_order else 999)
    sem_df = sem_df.sort_values("__order").drop(columns="__order").reset_index(drop=True)

    # Exibição
    st.subheader("Resumo 1º Semestre (Jan–Jun)")
    st.dataframe(sem_df, use_container_width=True)

    # Gráficos em tela
    c1, c2 = st.columns(2)
    with c1:
        st.pyplot(chart_semester_bar(sem_df))
        st.pyplot(chart_monthly_totals(month_comparisons))
    with c2:
        st.pyplot(chart_top_changes(sem_df, "Top 5 Crescimentos (Dif_abs – Jan–Jun)", top=True))
        st.pyplot(chart_top_changes(sem_df, "Top 5 Quedas (Dif_abs – Jan–Jun)", top=False))

    for label in MONTH_LABELS:
        st.pyplot(chart_month_grouped_bars(label, month_comparisons[label]))

    # Downloads (Excel + PDF)
    excel_bytes = build_excel(sem_df, month_comparisons)
    pdf_bytes   = build_pdf(sem_df, month_comparisons)
    st.download_button("⬇️ Baixar Excel (Resumo_1S + Jan..Jun)", data=excel_bytes,
                       file_name="comparativo_receitas_2024_2025_1S.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.download_button("⬇️ Baixar PDF (gráficos)", data=pdf_bytes,
                       file_name="relatorio_comparativo_1S_2024_2025.pdf", mime="application/pdf")
