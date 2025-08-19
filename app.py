# app.py — Streamlit • Comparativo 2024 x 2025 (Jan–Jun) com upload
# - Upload de Excel 2024, Excel 2025 e opcional ICMS.xlsx
# - Robustez: não quebra se faltar a aba "Resumo Jan-Jul 2025" (deduz ordem)
# - Gera Excel (Resumo_1S + meses) e PDF com gráficos para download
import io
import unicodedata
from pathlib import Path

import pandas as pd
import streamlit as st
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter
from matplotlib.backends.backend_pdf import PdfPages

# ===== Config =====
st.set_page_config(layout="wide", page_title="Comparativo 2024 x 2025 (Jan–Jun)")
MONTHS_2024 = ["Jan_2024","Fev_2024","Mar_2024","Abr_2024","Mai_2024","Jun_2024"]
MONTHS_2025 = ["Jan_2025","Fev_2025","Mar_2025","Abr_2025","Mai_2025","Jun_2025"]
MONTH_LABELS = ["Jan","Fev","Mar","Abr","Mai","Jun"]
MODEL_FALLBACK = ["FPM","ICMS","FEP","ITR","CFM","FUS","CID","FEB","SNA","ADO"]

# ===== Utils =====
def strip_accents(s: str) -> str:
    if s is None: return ""
    return "".join(c for c in unicodedata.normalize("NFKD", str(s)) if not unicodedata.combining(c))

def brl(x, pos=None):
    return f"R$ {x:,.0f}".replace(",", "X").replace(".", ",").replace("X", ".")

def read_excel_all_sheets(file_like) -> dict:
    return pd.read_excel(file_like, sheet_name=None)

def find_model_order(sheets_2025: dict) -> list:
    """Ordem priorizando 'Resumo Jan-Jul 2025'; se não existir, deduz das abas Jan..Jun_2025; garante ICMS na ordem."""
    target = None
    for name in sheets_2025.keys():
        if strip_accents(name).strip().lower() == strip_accents("Resumo Jan-Jul 2025").lower():
            target = name; break
    if target is not None:
        df = sheets_2025[target]
        if "Segmento" in df.columns:
            order = [s for s in df["Segmento"].astype(str).tolist() if s and str(s).strip()]
            if "ICMS" not in order:
                if "FPM" in order:
                    i = order.index("FPM")+1
                    order = order[:i] + ["ICMS"] + order[i:]
                else:
                    order.append("ICMS")
            return order
    # Deduz pelas abas mensais
    seen = []
    for m in MONTHS_2025:
        if m in sheets_2025 and "Segmento" in sheets_2025[m].columns:
            for s in sheets_2025[m]["Segmento"].astype(str).tolist():
                if s and s.strip() and s not in seen:
                    seen.append(s)
    if "ICMS" not in seen:
        if "FPM" in seen:
            i = seen.index("FPM")+1
            seen = seen[:i] + ["ICMS"] + seen[i:]
        else:
            seen.append("ICMS")
    return seen if seen else MODEL_FALLBACK

def normalize_month_df(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame(columns=["Segmento","Crédito","Débito","Líquido"])
    rename = {}
    for c in df.columns:
        cl = strip_accents(str(c)).strip().lower()
        if cl.startswith("segmento"): rename[c]="Segmento"
        elif cl.startswith("liquido"): rename[c]="Líquido"
        elif cl.startswith("credito"): rename[c]="Crédito"
        elif cl.startswith("debito"):  rename[c]="Débito"
    df = df.rename(columns=rename)
    if "Segmento" not in df.columns:
        return pd.DataFrame(columns=["Segmento","Crédito","Débito","Líquido"])
    # calcula líquido se necessário
    if "Líquido" not in df.columns:
        cred = next((c for c in df.columns if strip_accents(str(c)).lower().startswith("credito")), None)
        deb  = next((c for c in df.columns if strip_accents(str(c)).lower().startswith("debito")), None)
        if cred and deb:
            df["Crédito"] = pd.to_numeric(df[cred], errors="coerce").fillna(0.0)
            df["Débito"]  = pd.to_numeric(df[deb],  errors="coerce").fillna(0.0)
            df["Líquido"] = df["Crédito"] - df["Débito"]
        else:
            for col in ["Crédito","Débito","Líquido"]:
                df[col] = pd.to_numeric(df.get(col, 0.0), errors="coerce").fillna(0.0)
    else:
        for col in ["Crédito","Débito","Líquido"]:
            df[col] = pd.to_numeric(df.get(col, 0.0), errors="coerce").fillna(0.0)
    return df[["Segmento","Crédito","Débito","Líquido"]]

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

def parse_icms_map(icms_sheets: dict) -> dict:
    if not icms_sheets: return {}
    name = list(icms_sheets.keys())[0]
    df = icms_sheets[name].copy()
    if df is None or df.empty: return {}
    df.columns = [strip_accents(str(c)).strip().lower() for c in df.columns]
    year_col = None
    for c in df.columns:
        if "arrecada" in c and "icms" in c: year_col = c; break
    if year_col is None:
        for c in df.columns:
            if df[c].astype(str).str.contains("2024|2025").any(): year_col = c; break
    if year_col is None: return {}
    months = {
        "janeiro":"Jan","fevereiro":"Fev","marco":"Mar","março":"Mar","abril":"Abr","maio":"Mai",
        "junho":"Jun","julho":"Jul","agosto":"Ago","setembro":"Set","outubro":"Out","novembro":"Nov","dezembro":"Dez"
    }
    icms_map={}
    for _,row in df.iterrows():
        y = pd.to_numeric(row[year_col], errors="coerce")
        if pd.isna(y): continue
        y = int(y)
        per={}
        for c in df.columns:
            if c in months: per[months[c]] = row[c]
        icms_map[y]=per
    return icms_map

def load_month_liquid(sheets: dict, sheet_name: str) -> pd.DataFrame:
    if sheet_name not in sheets: return pd.DataFrame(columns=["Segmento","Líquido"])
    return normalize_month_df(sheets[sheet_name])[["Segmento","Líquido"]]

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

# ===== Charts =====
def chart_semester_bar(sem_df: pd.DataFrame):
    fig, ax = plt.subplots(figsize=(11, 6.5))
    x = range(len(sem_df))
    w = 0.4
    ax.bar([i-w/2 for i in x], sem_df["2024_Jan-Jun"], width=w, label="2024")
    ax.bar([i+w/2 for i in x], sem_df["2025_Jan-Jun"], width=w, label="2025")
    ax.set_xticks(list(x)); ax.set_xticklabels(sem_df["Segmento"], rotation=45, ha="right")
    ax.yaxis.set_major_formatter(FuncFormatter(brl)); ax.set_title("Líquido por Segmento – 1º Semestre (2024 vs 2025)")
    ax.legend(); fig.tight_layout(); return fig

def chart_top(df: pd.DataFrame, title: str, top=True):
    d = df.sort_values("Dif_abs", ascending=False)
    d = d.head(5) if top else d.tail(5).sort_values("Dif_abs")
    fig, ax = plt.subplots(figsize=(9,5.5))
    ax.barh(d["Segmento"], d["Dif_abs"]); ax.xaxis.set_major_formatter(FuncFormatter(brl))
    ax.set_title(title); fig.tight_layout(); return fig

def chart_monthly_totals(month_comp: dict):
    t24=[month_comp[m]["2024_Líquido"].sum() for m in MONTH_LABELS]
    t25=[month_comp[m]["2025_Líquido"].sum() for m in MONTH_LABELS]
    fig, ax = plt.subplots(figsize=(10.5,5.5))
    ax.plot(MONTH_LABELS, t24, marker="o", label="2024"); ax.plot(MONTH_LABELS, t25, marker="o", label="2025")
    ax.yaxis.set_major_formatter(FuncFormatter(brl)); ax.set_title("Totais Mensais – 1º Semestre (2024 vs 2025)")
    ax.legend(); fig.tight_layout(); return fig

def chart_month_grouped(month_name: str, df: pd.DataFrame):
    fig, ax = plt.subplots(figsize=(11,6.5))
    segs=df["Segmento"].tolist(); x=range(len(segs)); w=0.4
    ax.bar([i-w/2 for i in x], df["2024_Líquido"].tolist(), width=w, label="2024")
    ax.bar([i+w/2 for i in x], df["2025_Líquido"].tolist(), width=w, label="2025")
    ax.set_xticks(list(x)); ax.set_xticklabels(segs, rotation=45, ha="right"); ax.yaxis.set_major_formatter(FuncFormatter(brl))
    ax.set_title(f"{month_name}: Líquido por Segmento"); ax.legend(); fig.tight_layout(); return fig

def build_pdf(sem_df: pd.DataFrame, month_comp: dict) -> bytes:
    buf = io.BytesIO()
    with PdfPages(buf) as pdf:
        total24 = sem_df["2024_Jan-Jun"].sum()
        total25 = sem_df["2025_Jan-Jun"].sum()
        delta = total25 - total24
        pct = (delta/total24*100.0) if total24 else 0.0
        # Capa
        fig0, ax0 = plt.subplots(figsize=(11.69,8.27)); ax0.axis("off")
        y=0.9
        ax0.text(0.05,y,"Relatório Comparativo 1º Semestre – 2024 x 2025",fontsize=18,weight="bold"); y-=0.08
        ax0.text(0.05,y,f"Soma Jan–Jun 2024: {brl(total24)}",fontsize=12); y-=0.05
        ax0.text(0.05,y,f"Soma Jan–Jun 2025: {brl(total25)}",fontsize=12); y-=0.05
        ax0.text(0.05,y,f"Variação: {brl(delta)} ({pct:.2f}%)",fontsize=12)
        fig0.tight_layout(); pdf.savefig(fig0); plt.close(fig0)
        # Páginas
        pdf.savefig(chart_semester_bar(sem_df)); plt.close()
        pdf.savefig(chart_top(sem_df,"Top 5 Crescimentos (Dif_abs – Jan–Jun)",True)); plt.close()
        pdf.savefig(chart_top(sem_df,"Top 5 Quedas (Dif_abs – Jan–Jun)",False)); plt.close()
        pdf.savefig(chart_monthly_totals(month_comp)); plt.close()
        for m in MONTH_LABELS:
            pdf.savefig(chart_month_grouped(m, month_comp[m])); plt.close()
    buf.seek(0); return buf.read()

def build_excel(sem_df: pd.DataFrame, month_comp: dict) -> bytes:
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as xw:
        sem_df.to_excel(xw, index=False, sheet_name="Resumo_1S")
        for name, df in month_comp.items():
            df.to_excel(xw, index=False, sheet_name=name)
    out.seek(0); return out.read()

# ===== UI / Flow =====
def main():
    st.title("📊 Comparativo de Receitas — 2024 x 2025 — Jan–Jun")
    st.caption("Envie os dois arquivos. ICMS.xlsx é opcional; se enviado, injeta/atualiza a linha ICMS por mês/ano.")

    c1,c2,c3 = st.columns([1,1,1])
    with c1: file_2024 = st.file_uploader("Excel 2024 (abas Jan_2024..Jun_2024)", type=["xlsx","xls"], key="file2024")
    with c2: file_2025 = st.file_uploader("Excel 2025 (abas Jan_2025..Jun_2025)", type=["xlsx","xls"], key="file2025")
    with c3: file_icms = st.file_uploader("ICMS.xlsx (opcional)", type=["xlsx","xls"], key="fileicms")

    if not st.button("Gerar comparativo"):
        st.stop()

    if not file_2024 or not file_2025:
        st.error("Envie os dois arquivos: 2024 e 2025.")
        st.stop()

    try:
        sheets_2024 = read_excel_all_sheets(file_2024)
        sheets_2025 = read_excel_all_sheets(file_2025)
    except Exception as e:
        st.error(f"Falha ao ler os Excel enviados: {e}")
        st.stop()

    # Ordem de segmentos do modelo 2025 (robusto)
    model_order = find_model_order(sheets_2025)

    # ICMS opcional
    if file_icms is not None:
        try:
            icms_map = parse_icms_map(read_excel_all_sheets(file_icms))
        except Exception as e:
            st.warning(f"Não consegui ler ICMS.xlsx ({e}). Prosseguindo sem ICMS extra.")
            icms_map = {}
        # 2024
        for sheet in MONTHS_2024:
            if sheet in sheets_2024 and 2024 in icms_map:
                mon = sheet.split("_")[0]
                dfm = normalize_month_df(sheets_2024[sheet])
                sheets_2024[sheet] = upsert_icms_row(dfm, icms_map[2024].get(mon))
        # 2025
        for sheet in MONTHS_2025:
            if sheet in sheets_2025 and 2025 in icms_map:
                mon = sheet.split("_")[0]
                dfm = normalize_month_df(sheets_2025[sheet])
                sheets_2025[sheet] = upsert_icms_row(dfm, icms_map[2025].get(mon))

    # Comparativos mensais
    month_comp = {}
    for m24, m25, label in zip(MONTHS_2024, MONTHS_2025, MONTH_LABELS):
        month_comp[label] = compare_month(sheets_2024, sheets_2025, m24, m25, model_order)

    # Resumo 1º semestre
    sem24 = sum_semester(sheets_2024, MONTHS_2024).rename(columns={"Líquido":"2024_Jan-Jun"})
    sem25 = sum_semester(sheets_2025, MONTHS_2025).rename(columns={"Líquido":"2025_Jan-Jun"})
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

    # Exibe e permite baixar
    st.subheader("Resumo 1º Semestre (Jan–Jun)")
    st.dataframe(sem_df, use_container_width=True)

    c1,c2 = st.columns(2)
    with c1:
        st.pyplot(chart_semester_bar(sem_df))
        st.pyplot(chart_monthly_totals(month_comp))
    with c2:
        st.pyplot(chart_top(sem_df,"Top 5 Crescimentos (Dif_abs – Jan–Jun)",True))
        st.pyplot(chart_top(sem_df,"Top 5 Quedas (Dif_abs – Jan–Jun)",False))

    for m in MONTH_LABELS:
        st.pyplot(chart_month_grouped(m, month_comp[m]))

    excel_bytes = build_excel(sem_df, month_comp)
    pdf_bytes   = build_pdf(sem_df, month_comp)
    st.download_button("⬇️ Baixar Excel (Resumo_1S + Jan..Jun)", data=excel_bytes,
                       file_name="comparativo_receitas_2024_2025_1S.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.download_button("⬇️ Baixar PDF (gráficos)", data=pdf_bytes,
                       file_name="relatorio_comparativo_1S_2024_2025.pdf",
                       mime="application/pdf")

if __name__ == "__main__":
    main()
