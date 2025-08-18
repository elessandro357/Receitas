# compare_semesters.py
# ------------------------------------------------------------
# Compara receitas (Líquido) 2024 x 2025, meses Jan–Jun,
# gera:
# 1) Excel: Resumo_1S + abas Jan..Jun (comparativo)
# 2) PNGs de gráficos em out/charts/
# 3) PDF consolidado com os gráficos: out/relatorio_comparativo_1S_2024_2025.pdf
# Mantém a ordem de Segmento do modelo (aba "Resumo Jan-Jul 2025").
# ------------------------------------------------------------
import argparse
from pathlib import Path
import pandas as pd
import matplotlib
matplotlib.use("Agg")  # backend para ambiente headless (CI)
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter
from matplotlib.backends.backend_pdf import PdfPages

DEFAULT_FILE_2024 = Path("data/arrecadacao_liquida_jan_dez_2024.xlsx")
DEFAULT_FILE_2025 = Path("data/arrecadacao_jan_jul_2025.xlsx")
DEFAULT_OUTPUT_XLSX = Path("out/comparativo_receitas_2024_2025_1S.xlsx")
DEFAULT_OUTPUT_DIR  = Path("out")
MONTHS_2024 = ["Jan_2024","Fev_2024","Mar_2024","Abr_2024","Mai_2024","Jun_2024"]
MONTHS_2025 = ["Jan_2025","Fev_2025","Mar_2025","Abr_2025","Mai_2025","Jun_2025"]
MONTH_LABELS = ["Jan","Fev","Mar","Abr","Mai","Jun"]

def brl_formatter(x, pos):
    return f"R$ {x:,.0f}".replace(",", "X").replace(".", ",").replace("X", ".")

def get_model_order(file_2025: Path) -> list:
    df = pd.read_excel(file_2025, sheet_name="Resumo Jan-Jul 2025")
    order = []
    for seg in df["Segmento"].tolist():
        if isinstance(seg, str) and seg.strip() and seg not in order:
            order.append(seg.strip())
    return order

def load_month_liquid(path: Path, sheet: str) -> pd.DataFrame:
    """Retorna DataFrame com: Segmento, Líquido (ou calcula = Crédito - Débito)."""
    try:
        df = pd.read_excel(path, sheet_name=sheet)
    except Exception:
        return pd.DataFrame(columns=["Segmento","Líquido"])
    # Normaliza nomes
    rename_map = {}
    for c in df.columns:
        cl = str(c).strip().lower()
        if cl.startswith("segmento"):
            rename_map[c] = "Segmento"
        elif cl.startswith("líquido") or cl.startswith("liquido"):
            rename_map[c] = "Líquido"
    df = df.rename(columns=rename_map)
    if not {"Segmento","Líquido"}.issubset(df.columns):
        cred_col = next((c for c in df.columns if str(c).lower().startswith("crédito") or str(c).lower().startswith("credito")), None)
        deb_col  = next((c for c in df.columns if str(c).lower().startswith("débito") or str(c).lower().startswith("debito")), None)
        if "Segmento" in df.columns and cred_col and deb_col:
            df["Líquido"] = df[cred_col].astype(float) - df[deb_col].astype(float)
        else:
            return pd.DataFrame(columns=["Segmento","Líquido"])
    return df[["Segmento","Líquido"]]

def compare_month(file_2024: Path, file_2025: Path, month_2024: str, month_2025: str, model_order: list) -> pd.DataFrame:
    df24 = load_month_liquid(file_2024, month_2024)
    df25 = load_month_liquid(file_2025, month_2025)
    # Base com ordem do modelo primeiro
    all_segments = list(dict.fromkeys(model_order + df24["Segmento"].dropna().tolist() + df25["Segmento"].dropna().tolist()))
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
    extras = [s for s in merged["Segmento"].tolist() if s not in model_order]
    final_order = model_order + extras
    merged["__order"] = merged["Segmento"].apply(lambda s: final_order.index(s) if s in final_order else 999)
    merged = merged.sort_values("__order").drop(columns="__order").reset_index(drop=True)
    return merged

def sum_semester(file_path: Path, months: list) -> pd.DataFrame:
    acc = None
    for m in months:
        dfm = load_month_liquid(file_path, m)
        if acc is None:
            acc = dfm.copy()
        else:
            acc = acc.merge(dfm, on="Segmento", how="outer", suffixes=("","_tmp"))
            acc["Líquido"] = acc[["Líquido","Líquido_tmp"]].fillna(0.0).sum(axis=1)
            acc = acc.drop(columns=[c for c in acc.columns if c.endswith("_tmp")])
    if acc is None:
        acc = pd.DataFrame(columns=["Segmento","Líquido"])
    return acc

# ---------------------- Gráficos ----------------------
def plot_semester_bar(sem_df: pd.DataFrame, model_order: list, out_dir: Path):
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
    path = out_dir / "01_semestre_bar.png"
    fig.savefig(path, dpi=150)
    plt.close(fig)
    return path

def plot_top_changes(sem_df: pd.DataFrame, out_dir: Path):
    df = sem_df.copy()
    df = df.sort_values("Dif_abs", ascending=False).reset_index(drop=True)
    top5 = df.head(5)
    bottom5 = df.tail(5).sort_values("Dif_abs")  # mais quedas
    # Crescimentos
    fig1, ax1 = plt.subplots(figsize=(9, 5.5))
    ax1.barh(top5["Segmento"], top5["Dif_abs"])
    ax1.xaxis.set_major_formatter(FuncFormatter(brl_formatter))
    ax1.set_title("Top 5 Crescimentos (Dif_abs – Jan–Jun)")
    fig1.tight_layout()
    p1 = out_dir / "02_top5_altas.png"
    fig1.savefig(p1, dpi=150)
    plt.close(fig1)
    # Quedas
    fig2, ax2 = plt.subplots(figsize=(9, 5.5))
    ax2.barh(bottom5["Segmento"], bottom5["Dif_abs"])
    ax2.xaxis.set_major_formatter(FuncFormatter(brl_formatter))
    ax2.set_title("Top 5 Quedas (Dif_abs – Jan–Jun)")
    fig2.tight_layout()
    p2 = out_dir / "03_top5_quedas.png"
    fig2.savefig(p2, dpi=150)
    plt.close(fig2)
    return [p1, p2]

def plot_monthly_totals(month_comp: dict, out_dir: Path):
    # month_comp: {"Jan": df, ...}
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
    path = out_dir / "04_totais_mensais_linha.png"
    fig.savefig(path, dpi=150)
    plt.close(fig)
    return path

def plot_month_grouped_bars(month_name: str, df: pd.DataFrame, out_dir: Path):
    # df: Segmento | 2024_Líquido | 2025_Líquido
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
    path = out_dir / f"05_{month_name.lower()}_barras.png"
    fig.savefig(path, dpi=150)
    plt.close(fig)
    return path

def build_pdf(charts_paths, out_pdf: Path, sem_header_text: str = ""):
    with PdfPages(out_pdf) as pdf:
        # Capa / Sumário
        fig0, ax0 = plt.subplots(figsize=(11.69, 8.27))  # A4 landscape approx
        ax0.axis("off")
        y = 0.9
        ax0.text(0.05, y, "Relatório Comparativo 1º Semestre – 2024 x 2025", fontsize=18, weight="bold")
        y -= 0.07
        ax0.text(0.05, y, sem_header_text, fontsize=11)
        y -= 0.03
        ax0.text(0.05, y, "Conteúdo: Totais do semestre, Top 5 altas/quedas, Totais mensais, Gráficos por mês.", fontsize=10)
        fig0.tight_layout()
        pdf.savefig(fig0)
        plt.close(fig0)
        # Demais páginas (cada gráfico em uma página)
        for p in charts_paths:
            img = plt.imread(p)
            fig, ax = plt.subplots(figsize=(11.69, 8.27))
            ax.imshow(img)
            ax.axis("off")
            pdf.savefig(fig)
            plt.close(fig)

# ---------------------- Main ----------------------
def main():
    parser = argparse.ArgumentParser(description="Comparativo receitas 2024 x 2025 (Jan–Jun) com gráficos")
    parser.add_argument("--file-2024", type=Path, default=DEFAULT_FILE_2024)
    parser.add_argument("--file-2025", type=Path, default=DEFAULT_FILE_2025)
    parser.add_argument("--out-xlsx", type=Path, default=DEFAULT_OUTPUT_XLSX)
    parser.add_argument("--out-dir", type=Path, default=DEFAULT_OUTPUT_DIR)
    args = parser.parse_args()

    args.out_dir.mkdir(parents=True, exist_ok=True)
    charts_dir = args.out_dir / "charts"
    charts_dir.mkdir(parents=True, exist_ok=True)

    model_order = get_model_order(args.file_2025)

    # Comparativos mensais
    month_pairs = list(zip(MONTHS_2024, MONTHS_2025, MONTH_LABELS))
    month_comparisons = {}
    for m24, m25, label in month_pairs:
        df = compare_month(args.file_2024, args.file_2025, m24, m25, model_order)
        month_comparisons[label] = df

    # Resumo 1º semestre
    sem24 = sum_semester(args.file_2024, MONTHS_2024).rename(columns={"Líquido":"2024_Jan-Jun"})
    sem25 = sum_semester(args.file_2025, MONTHS_2025).rename(columns={"Líquido":"2025_Jan-Jun"})
    base = pd.DataFrame({"Segmento": list(dict.fromkeys(model_order + sem24["Segmento"].dropna().tolist() + sem25["Segmento"].dropna().tolist()))})
    sem = base.merge(sem24, on="Segmento", how="left").merge(sem25, on="Segmento", how="left").fillna({"2024_Jan-Jun":0.0,"2025_Jan-Jun":0.0})
    sem["Dif_abs"] = sem["2025_Jan-Jun"] - sem["2024_Jan-Jun"]
    sem["Dif_%"] = sem.apply(lambda r: (r["Dif_abs"] / r["2024_Jan-Jun"] * 100.0) if r["2024_Jan-Jun"] else None, axis=1)
    # Ordena por ordem do modelo + extras
    extras = [s for s in sem["Segmento"].tolist() if s not in model_order]
    final_order = model_order + extras
    sem["__order"] = sem["Segmento"].apply(lambda s: final_order.index(s) if s in final_order else 999)
    sem = sem.sort_values("__order").drop(columns="__order").reset_index(drop=True)

    # Salva Excel
    with pd.ExcelWriter(args.out_xlsx, engine="openpyxl") as xw:
        sem.to_excel(xw, index=False, sheet_name="Resumo_1S")
        for name, df in month_comparisons.items():
            df.to_excel(xw, index=False, sheet_name=name)

    # Gera gráficos (PNGs)
    charts = []
    charts.append(plot_semester_bar(sem, model_order, charts_dir))
    charts += plot_top_changes(sem, charts_dir)
    charts.append(plot_monthly_totals(month_comparisons, charts_dir))
    for label in MONTH_LABELS:
        charts.append(plot_month_grouped_bars(label, month_comparisons[label], charts_dir))

    # PDF consolidado
    total24 = sem["2024_Jan-Jun"].sum()
    total25 = sem["2025_Jan-Jun"].sum()
    delta = total25 - total24
    pct = (delta / total24 * 100.0) if total24 else 0.0
    header = (
        f"Soma Jan–Jun 2024: {brl_formatter(total24, None)} | "
        f"Soma Jan–Jun 2025: {brl_formatter(total25, None)} | "
        f"Variação: {brl_formatter(delta, None)} ({pct:.2f}%)"
    )
    out_pdf = args.out_dir / "relatorio_comparativo_1S_2024_2025.pdf"
    build_pdf(charts, out_pdf, header)

    print(f"[OK] Excel: {args.out_xlsx.resolve()}")
    print(f"[OK] PDF:   {out_pdf.resolve()}")
    print(f"[OK] PNGs:  {charts_dir.resolve()}")

if __name__ == "__main__":
    main()
