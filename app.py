# app.py
import os
import io
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(page_title="Comparativo de Receitas • Resumo", layout="wide")

# ==========================
# Helpers
# ==========================
def format_brl(v):
    try:
        return f"R$ {float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return str(v)

def find_year_column(cols, year):
    """Tenta achar a coluna do ano priorizando JAN–JUL, depois JAN–JUN/1S."""
    year = str(year)
    cols = [str(c) for c in cols if year in str(c)]
    priority = ["JAN-JUL", "JAN/JUL", "JAN JUL", "JUL", "JAN-JUN", "JAN/JUN", "JAN JUN", "1S", "JAN", "JUN"]
    for key in priority:
        for c in cols:
            if key in c.upper():
                return c
    return cols[0] if cols else None

@st.cache_data(show_spinner=False)
def load_resumo(file_or_path, sheet_guess: str = "Resumo_1S") -> pd.DataFrame:
    """Aceita caminho de arquivo OU BytesIO (upload)."""
    xls = pd.ExcelFile(file_or_path)
    # Detecta aba de resumo
    if sheet_guess in xls.sheet_names:
        sheet = sheet_guess
    else:
        sheet = next((s for s in xls.sheet_names if "RESUM" in s.upper()), xls.sheet_names[0])

    df = pd.read_excel(xls, sheet_name=sheet)
    df.columns = [str(c).strip() for c in df.columns]

    # Detecta coluna de segmento/categoria
    seg_col = next((c for c in ["Segmento","SEGMENTO","Categoria","Receita","Natureza","Descrição"] if c in df.columns), df.columns[0])

    # Detecta colunas de ano
    c2024 = find_year_column(df.columns, 2024)
    c2025 = find_year_column(df.columns, 2025)
    if c2024 is None or c2025 is None:
        raise ValueError("Aba de resumo precisa ter colunas de 2024 e 2025 (ex.: '2024_Jan-Jun' / '2025_Jan-Jul').")

    work = df[[seg_col, c2024, c2025]].copy()
    work.columns = ["segment", "y2024", "y2025"]
    for c in ["y2024", "y2025"]:
        work[c] = pd.to_numeric(work[c], errors="coerce").fillna(0.0)

    # Diferença absoluta (apenas R$)
    work["diff_abs"] = work["y2025"] - work["y2024"]
    return work.sort_values("y2025", ascending=False).reset_index(drop=True)

def make_download(df: pd.DataFrame, excel: bool = False):
    if excel:
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine="xlsxwriter") as w:
            df.to_excel(w, index=False, sheet_name="Resumo_Filtered")
        return out.getvalue(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    return df.to_csv(index=False).encode("utf-8-sig"), "text/csv"

def get_data_source(default_path: str):
    """
    Prioriza upload (main ou sidebar) com persistência em session_state.
    Retorna BytesIO (upload) OU caminho existente. Se nada, retorna None.
    """
    up_main = st.file_uploader(
        "Faça upload do Excel (.xlsx) com a aba de Resumo",
        type=["xlsx"], key="uploader_main",
        help="Ex.: 'Resumo_1S' com colunas de 2024 e 2025 (Jan–Jun/Jan–Jul)."
    )
    up_side = st.sidebar.file_uploader("Ou envie pela sidebar (opcional)", type=["xlsx"], key="uploader_side")

    uploaded = up_main or up_side
    if uploaded:
        st.session_state["uploaded_bytes"] = uploaded.read()
        st.session_state["uploaded_name"] = uploaded.name

    if "uploaded_bytes" in st.session_state:
        st.success(f"Usando arquivo enviado: {st.session_state.get('uploaded_name', 'upload.xlsx')}")
        return io.BytesIO(st.session_state["uploaded_bytes"])

    if default_path and isinstance(default_path, str) and os.path.exists(default_path):
        st.info(f"Usando arquivo padrão: {default_path}")
        return default_path

    return None

# ==========================
# Fonte de dados
# ==========================
st.sidebar.title("📁 Fonte de Dados")
default_path = os.environ.get("RECEITAS_FILE", "data/comparativo_receitas_2024_2025_1S_COM_ICMS.xlsx")

col_left, col_right = st.columns([3, 1])
with col_right:
    if st.button("Remover arquivo enviado", use_container_width=True):
        st.session_state.pop("uploaded_bytes", None)
        st.session_state.pop("uploaded_name", None)
        st.rerun()

data_source = get_data_source(default_path)

if data_source is None:
    st.title("📊 Comparativo de Receitas (Resumo)")
    st.warning("Nenhum arquivo encontrado. Envie um `.xlsx` com a aba de resumo (ex.: `Resumo_1S`) contendo colunas de 2024 e 2025.")
    st.stop()

# ==========================
# Carregamento
# ==========================
try:
    df = load_resumo(data_source)
except Exception as e:
    st.error(f"Falha ao ler o arquivo/aba de resumo. Detalhes: {e}")
    st.stop()

# ==========================
# Controles
# ==========================
st.title("📊 Comparativo de Receitas (Resumo)")
st.caption("Envie sua planilha acima. Foco apenas em valores em reais (sem %). Use os filtros para focar no que interessa.")

segments = df["segment"].tolist()
col1, col2, col3 = st.columns([2, 1, 1])
with col1:
    selected = st.multiselect("Segmentos", options=segments, default=segments)
with col2:
    sort_by = st.selectbox("Ordenar por", ["2025 (↓)", "Diferença (↓)", "2024 (↓)", "Alfabética (A→Z)"])
with col3:
    top_n = st.number_input("Top N", min_value=1, max_value=len(segments), value=min(10, len(segments)))

# Filtro + ordenação
fdf = df[df["segment"].isin(selected)].copy()
if sort_by == "2025 (↓)":
    fdf = fdf.sort_values("y2025", ascending=False)
elif sort_by == "2024 (↓)":
    fdf = fdf.sort_values("y2024", ascending=False)
elif sort_by == "Diferença (↓)":
    fdf = fdf.sort_values("diff_abs", ascending=False)
else:
    fdf = fdf.sort_values("segment", ascending=True)
fdf = fdf.head(top_n)

# ==========================
# Totais (cards)
# ==========================
total_2024 = float(fdf["y2024"].sum())
total_2025 = float(fdf["y2025"].sum())
total_diff = total_2025 - total_2024

m1, m2, m3 = st.columns(3)
with m1:
    st.metric("Total 2024 (período)", format_brl(total_2024))
with m2:
    st.metric("Total 2025 (período)", format_brl(total_2025))
with m3:
    # delta fica visual, mas é apenas R$ (sem %)
    st.metric("Diferença (2025 - 2024)", format_brl(total_diff))

# ==========================
# Gráficos (com rótulos em R$)
# ==========================
left, right = st.columns(2)

with left:
    st.subheader("Barras Agrupadas (2024 x 2025)")
    plot_df = fdf.melt(id_vars=["segment"], value_vars=["y2024", "y2025"], var_name="year", value_name="value")
    plot_df["year"] = plot_df["year"].map({"y2024": "2024", "y2025": "2025"})
    # rótulos formatados
    plot_df["label_brl"] = plot_df["value"].map(format_brl)
    fig1 = px.bar(
        plot_df, x="segment", y="value", color="year", barmode="group",
        labels={"segment": "Segmento", "value": "Receita (R$)", "year": "Ano"},
        text="label_brl"
    )
    fig1.update_traces(textposition="outside", cliponaxis=False)
    fig1.update_layout(xaxis_tickangle=-30, yaxis_title=None, legend_title_text="Ano",
                       margin=dict(l=10, r=10, t=40, b=10))
    fig1.update_yaxes(tickprefix="R$ ")
    st.plotly_chart(fig1, use_container_width=True, theme="streamlit")

with right:
    st.subheader("Barras Divergentes (2024 vs 2025)")
    txt2024 = [format_brl(v) for v in fdf["y2024"]]
    txt2025 = [format_brl(v) for v in fdf["y2025"]]
    b2024 = go.Bar(x=-fdf["y2024"], y=fdf["segment"], name="2024", orientation="h",
                   text=txt2024, textposition="outside")
    b2025 = go.Bar(x=fdf["y2025"], y=fdf["segment"], name="2025", orientation="h",
                   text=txt2025, textposition="outside")
    fig2 = go.Figure(data=[b2024, b2025])
    fig2.update_layout(barmode="relative", margin=dict(l=10, r=10, t=40, b=10))
    fig2.update_xaxes(title_text="Receita (R$) — 2024 à esquerda", tickprefix="R$ ")
    st.plotly_chart(fig2, use_container_width=True, theme="streamlit")

st.subheader("Diferença Absoluta em R$")
fdf2 = fdf.sort_values("diff_abs", ascending=False).copy()
fdf2["label_brl"] = fdf2["diff_abs"].map(format_brl)
fig3 = px.bar(
    fdf2, x="segment", y="diff_abs",
    labels={"segment": "Segmento", "diff_abs": "Diferença (R$)"},
    text="label_brl"
)
fig3.update_traces(textposition="outside", cliponaxis=False)
fig3.update_layout(xaxis_tickangle=-30, yaxis_title=None, margin=dict(l=10, r=10, t=40, b=10))
fig3.update_yaxes(tickprefix="R$ ")
st.plotly_chart(fig3, use_container_width=True, theme="streamlit")

# ==========================
# Tabela + downloads (sem %)
# ==========================
st.subheader("Tabela Filtrada (somente R$)")
show = fdf[["segment", "y2024", "y2025", "diff_abs"]].rename(columns={
    "segment": "Segmento",
    "y2024": "2024",
    "y2025": "2025",
    "diff_abs": "Diferença_R$"
})
st.dataframe(
    show.style.format({
        "2024": format_brl,
        "2025": format_brl,
        "Diferença_R$": format_brl
    }),
    use_container_width=True, hide_index=True
)

csv_data, csv_mime = make_download(show, excel=False)
xlsx_data, xlsx_mime = make_download(show, excel=True)
col_dl1, col_dl2, _ = st.columns([1, 1, 6])
with col_dl1:
    st.download_button("Baixar CSV", data=csv_data, file_name="resumo_filtrado.csv", mime=csv_mime)
with col_dl2:
    st.download_button("Baixar Excel", data=xlsx_data, file_name="resumo_filtrado.xlsx", mime=xlsx_mime)

st.caption("Os totais exibidos consideram o conjunto filtrado na seleção de segmentos.")
