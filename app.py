# app.py
import os
import io
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(page_title="Comparativo de Receitas • Resumo 1S", layout="wide")

# --------------------------
# Helpers
# --------------------------
def format_brl(v):
    try:
        return f"R$ {float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return str(v)

def find_year_column(cols, year):
    year = str(year)
    # procura por colunas que contenham "2024" ou "2025"
    candidates = [c for c in cols if year in str(c)]
    # prioriza colunas que mencionem Jan-Jun ou algo parecido
    pref = [c for c in candidates if "JAN" in str(c).upper() or "JUN" in str(c).upper() or "1S" in str(c).upper()]
    if pref:
        return pref[0]
    return candidates[0] if candidates else None

@st.cache_data(show_spinner=False)
def load_resumo(path: str, sheet_guess: str = "Resumo_1S") -> pd.DataFrame:
    xls = pd.ExcelFile(path)
    # tenta achar uma planilha que contenha "Resumo"
    sheet = None
    if sheet_guess in xls.sheet_names:
        sheet = sheet_guess
    else:
        for s in xls.sheet_names:
            if "RESUM" in s.upper():  # Resumo / Resumo_1S
                sheet = s
                break
        if sheet is None:
            # cai na primeira aba
            sheet = xls.sheet_names[0]

    df = pd.read_excel(path, sheet_name=sheet)
    # normaliza cabeçalhos
    df.columns = [str(c).strip() for c in df.columns]

    # encontra colunas
    # tenta detectar a coluna de categoria/segmento
    seg_col = None
    for cand in ["Segmento", "SEGMENTO", "Categoria", "Receita", "Natureza", "Descrição"]:
        if cand in df.columns:
            seg_col = cand
            break
    if seg_col is None:
        # usa a primeira coluna como categoria
        seg_col = df.columns[0]

    c2024 = find_year_column(df.columns, 2024)
    c2025 = find_year_column(df.columns, 2025)

    if c2024 is None or c2025 is None:
        raise ValueError("Não encontrei colunas de 2024 e 2025. Verifique se a planilha resumo possui colunas como '2024_Jan-Jun' e '2025_Jan-Jun'.")

    # seleciona e renomeia
    work = df[[seg_col, c2024, c2025]].copy()
    work.columns = ["segment", "y2024", "y2025"]

    # cast numérico
    for c in ["y2024", "y2025"]:
        work[c] = pd.to_numeric(work[c], errors="coerce").fillna(0.0)

    # diferenças
    work["diff_abs"] = work["y2025"] - work["y2024"]
    work["diff_pct"] = (work["diff_abs"] / work["y2024"].replace(0, pd.NA)) * 100
    work["diff_pct"] = work["diff_pct"].fillna(0.0)

    # ordena por maior receita 2025
    work = work.sort_values("y2025", ascending=False).reset_index(drop=True)
    return work

def make_download(df: pd.DataFrame, filename: str, excel: bool = False):
    if excel:
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name="Resumo_Filtered")
        return out.getvalue()
    else:
        return df.to_csv(index=False).encode("utf-8-sig")

# --------------------------
# Sidebar (arquivo)
# --------------------------
st.sidebar.title("📁 Fonte de Dados")
default_path = os.environ.get("RECEITAS_FILE", "data/comparativo_receitas_2024_2025_1S_COM_ICMS.xlsx")
uploaded = st.sidebar.file_uploader("Carregar Excel (.xlsx) com a aba 'Resumo'", type=["xlsx"])

if uploaded is not None:
    path = uploaded
else:
    path = default_path

# --------------------------
# Carregamento
# --------------------------
try:
    df = load_resumo(path)
except Exception as e:
    st.error(f"Falha ao ler o arquivo/aba de resumo. Detalhes: {e}")
    st.stop()

# --------------------------
# Controles
# --------------------------
st.title("📊 Comparativo de Receitas (Resumo 1º Semestre)")
st.caption("Clique na legenda para ocultar/mostrar uma série no gráfico. Use os filtros para focar só no que interessa.")

segments = df["segment"].tolist()
col1, col2, col3, col4 = st.columns([2,1,1,1])

with col1:
    selected = st.multiselect("Segmentos", options=segments, default=segments)
with col2:
    sort_by = st.selectbox("Ordenar por", ["2025 (↓)", "Diferença (↓)", "2024 (↓)", "Alfabética (A→Z)"])
with col3:
    top_n = st.number_input("Top N", min_value=1, max_value=len(segments), value=min(10, len(segments)))
with col4:
    pct_labels = st.checkbox("Exibir % no gráfico de Diferença", value=True)

# botões rápidos
cba, cbb = st.columns(2)
with cba:
    if st.button("Selecionar tudo"):
        selected = segments
with cbb:
    if st.button("Limpar seleção"):
        selected = []

fdf = df[df["segment"].isin(selected)].copy()

# ordenação
if sort_by == "2025 (↓)":
    fdf = fdf.sort_values("y2025", ascending=False)
elif sort_by == "2024 (↓)":
    fdf = fdf.sort_values("y2024", ascending=False)
elif sort_by == "Diferença (↓)":
    fdf = fdf.sort_values("diff_abs", ascending=False)
else:
    fdf = fdf.sort_values("segment", ascending=True)

fdf = fdf.head(top_n)

# --------------------------
# Gráficos
# --------------------------
left, right = st.columns(2)

with left:
    st.subheader("Barras Agrupadas (2024 x 2025)")
    plot_df = fdf.melt(id_vars=["segment"], value_vars=["y2024", "y2025"], var_name="year", value_name="value")
    plot_df["year"] = plot_df["year"].map({"y2024": "2024 (Jan-Jun)", "y2025": "2025 (Jan-Jun)"})
    fig1 = px.bar(
        plot_df, x="segment", y="value", color="year", barmode="group",
        labels={"segment": "Segmento", "value": "Receita (R$)", "year": "Ano"}
    )
    fig1.update_layout(xaxis_tickangle=-30, yaxis_title=None, legend_title_text="Ano", margin=dict(l=10,r=10,t=40,b=10))
    st.plotly_chart(fig1, use_container_width=True, theme="streamlit")

with right:
    st.subheader("Barras Divergentes (2024 vs 2025)")
    b2024 = go.Bar(
        x=-fdf["y2024"], y=fdf["segment"], name="2024 (Jan-Jun)", orientation="h"
    )
    b2025 = go.Bar(
        x=fdf["y2025"], y=fdf["segment"], name="2025 (Jan-Jun)", orientation="h"
    )
    fig2 = go.Figure(data=[b2024, b2025])
    fig2.update_layout(barmode="relative", margin=dict(l=10,r=10,t=40,b=10))
    fig2.update_xaxes(title_text="Receita (R$) — valores de 2024 à esquerda")
    st.plotly_chart(fig2, use_container_width=True, theme="streamlit")

st.subheader("Diferença Absoluta (2025 - 2024)")
fdf2 = fdf.sort_values("diff_abs", ascending=False).copy()
fig3 = px.bar(
    fdf2, x="segment", y="diff_abs", labels={"segment": "Segmento", "diff_abs": "Diferença (R$)"}
)
if pct_labels:
    fig3.update_traces(text=[f"{p:.1f}%" for p in fdf2["diff_pct"]], textposition="outside")
fig3.update_layout(xaxis_tickangle=-30, yaxis_title=None, margin=dict(l=10,r=10,t=40,b=10))
st.plotly_chart(fig3, use_container_width=True, theme="streamlit")

# --------------------------
# Tabela + downloads
# --------------------------
st.subheader("Tabela Filtrada")
show = fdf[["segment", "y2024", "y2025", "diff_abs", "diff_pct"]].rename(columns={
    "segment": "Segmento",
    "y2024": "2024_Jan-Jun",
    "y2025": "2025_Jan-Jun",
    "diff_abs": "Dif_abs",
    "diff_pct": "Dif_%"
})
st.dataframe(
    show.style.format({
        "2024_Jan-Jun": format_brl,
        "2025_Jan-Jun": format_brl,
        "Dif_abs": format_brl,
        "Dif_%": lambda v: f"{v:.1f}%"
    }),
    use_container_width=True, hide_index=True
)

csv_data = make_download(show, "resumo_filtrado.csv", excel=False)
xlsx_data = make_download(show, "resumo_filtrado.xlsx", excel=True)

col_dl1, col_dl2, col_dl3 = st.columns([1,1,4])
with col_dl1:
    st.download_button("Baixar CSV", data=csv_data, file_name="resumo_filtrado.csv", mime="text/csv")
with col_dl2:
    st.download_button("Baixar Excel", data=xlsx_data, file_name="resumo_filtrado.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.caption("Dica: o gráfico aceita esconder/mostrar séries clicando na legenda.")
