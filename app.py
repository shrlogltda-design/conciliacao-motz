"""
Dashboard de Conciliação MOTZ - Streamlit (v3.3)
Upload de PDFs Repom + MOTZ (XLSX) + ATUA (XLS) → conciliação → visualização
v3.3: 1 linha por transferência + 22 colunas oficiais + multi-select + cores célula-por-célula
"""
import streamlit as st
import pandas as pd
import subprocess
import tempfile
import os
import shutil
from pathlib import Path
from datetime import datetime, timedelta
import hashlib
import io
import sys
import plotly.express as px

st.set_page_config(
    page_title="Conciliação MOTZ",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

COLORS = {
    "OK":            {"bg": "#D5E8C1", "fg": "#2E5410", "border": "#3B6D11", "plot": "#3B6D11"},
    "ATUA MAIOR":    {"bg": "#F8CCCC", "fg": "#7A1F1F", "border": "#A32D2D", "plot": "#A32D2D"},
    "ATUA MENOR":    {"bg": "#CDE3F7", "fg": "#0E4577", "border": "#185FA5", "plot": "#185FA5"},
    "NAO ENCONTRADO":{"bg": "#FCE9B6", "fg": "#5E3704", "border": "#854F0B", "plot": "#BF7F1C"},
    "SALDO ABERTO":  {"bg": "#DDD9FB", "fg": "#2A205F", "border": "#3C3489", "plot": "#4B3FB3"},
}

st.markdown(f"""
<style>
    .main .block-container {{ padding-top: 2rem; padding-bottom: 3rem; max-width: 1400px; }}
    h1 {{ font-size: 22px !important; font-weight: 500 !important; margin-bottom: 0 !important; }}
    h2 {{ font-size: 16px !important; font-weight: 500 !important; }}
    h3 {{ font-size: 14px !important; font-weight: 500 !important; }}
    .stMetric {{ background: var(--secondary-background-color); border-radius: 10px; padding: 12px 16px; }}
    .stMetric label {{ font-size: 12px !important; color: #6B6B66 !important; }}
    .stMetric [data-testid="stMetricValue"] {{ font-size: 22px !important; font-weight: 500 !important; }}
    .stDataFrame {{ font-size: 12px; }}
    [data-testid="stFileUploader"] section {{ border-radius: 10px; padding: 14px; }}
    .stButton > button {{ border-radius: 8px; font-size: 13px; padding: 6px 14px; }}
    .status-ok {{ background: {COLORS['OK']['bg']}; color: {COLORS['OK']['fg']}; padding: 2px 8px; border-radius: 999px; font-size: 11px; font-weight: 500; }}
    .status-maior {{ background: {COLORS['ATUA MAIOR']['bg']}; color: {COLORS['ATUA MAIOR']['fg']}; padding: 2px 8px; border-radius: 999px; font-size: 11px; font-weight: 500; }}
    .status-menor {{ background: {COLORS['ATUA MENOR']['bg']}; color: {COLORS['ATUA MENOR']['fg']}; padding: 2px 8px; border-radius: 999px; font-size: 11px; font-weight: 500; }}
    .status-ne {{ background: {COLORS['NAO ENCONTRADO']['bg']}; color: {COLORS['NAO ENCONTRADO']['fg']}; padding: 2px 8px; border-radius: 999px; font-size: 11px; font-weight: 500; }}
    .status-aberto {{ background: {COLORS['SALDO ABERTO']['bg']}; color: {COLORS['SALDO ABERTO']['fg']}; padding: 2px 8px; border-radius: 999px; font-size: 11px; font-weight: 500; }}
    div[data-testid="stButton"] > button[kind="secondary"] {{
        width: 100%; border-radius: 10px; padding: 12px 14px; font-size: 13px; font-weight: 500;
        text-align: left; border: 1.5px solid transparent; transition: all 0.15s ease; min-height: 62px;
    }}
    div[data-testid="stButton"] > button[kind="secondary"]:hover {{ transform: translateY(-1px); box-shadow: 0 2px 6px rgba(0,0,0,0.08); }}
    .card-ok button {{ background: {COLORS['OK']['bg']} !important; color: {COLORS['OK']['fg']} !important; border-color: {COLORS['OK']['border']}66 !important; }}
    .card-maior button {{ background: {COLORS['ATUA MAIOR']['bg']} !important; color: {COLORS['ATUA MAIOR']['fg']} !important; border-color: {COLORS['ATUA MAIOR']['border']}66 !important; }}
    .card-menor button {{ background: {COLORS['ATUA MENOR']['bg']} !important; color: {COLORS['ATUA MENOR']['fg']} !important; border-color: {COLORS['ATUA MENOR']['border']}66 !important; }}
    .card-ne button {{ background: {COLORS['NAO ENCONTRADO']['bg']} !important; color: {COLORS['NAO ENCONTRADO']['fg']} !important; border-color: {COLORS['NAO ENCONTRADO']['border']}66 !important; }}
    .card-aberto button {{ background: {COLORS['SALDO ABERTO']['bg']} !important; color: {COLORS['SALDO ABERTO']['fg']} !important; border-color: {COLORS['SALDO ABERTO']['border']}66 !important; }}
    .card-active button {{ border-width: 2.5px !important; box-shadow: 0 0 0 3px rgba(0,0,0,0.05) !important; }}
    .filter-hint {{ background: #F5F5F0; border-left: 3px solid #378ADD; padding: 8px 12px; border-radius: 4px; font-size: 12px; color: #4A4A44; margin-bottom: 12px; }}
</style>
""", unsafe_allow_html=True)

st.markdown("# Conciliação MOTZ consolidada")
st.caption("PDFs Repom × MOTZ (XLSX) × Cobrança ATUA (XLS) — cruzamento automático")

def parse_rs(v):
    if v is None or (isinstance(v, float) and pd.isna(v)) or v == "":
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).replace("R$", "").replace("\xa0", " ").strip()
    neg = s.startswith("-") or s.startswith("−")
    s = s.lstrip("-−").strip().replace(".", "").replace(",", ".")
    try:
        n = float(s)
        return -n if neg else n
    except Exception:
        return 0.0

def parse_date_br(v):
    if v is None or v == "" or pd.isna(v):
        return None
    if isinstance(v, datetime):
        return v
    s = str(v).strip()
    if s in ("", "01/01/0001"):
        return None
    for fmt in ("%d/%m/%Y %H:%M:%S", "%d/%m/%Y"):
        try:
            return datetime.strptime(s, fmt)
        except Exception:
            pass
    return None

def fmt_mi(n):
    if n is None:
        return "—"
    abs_n = abs(n)
    sig = "-" if n < 0 else ""
    if abs_n >= 1e6:
        return f"{sig}R$ {abs_n/1e6:,.2f} mi".replace(",", "X").replace(".", ",").replace("X", ".")
    if abs_n >= 1e3:
        return f"{sig}R$ {abs_n/1e3:,.2f} mil".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"{sig}R$ {abs_n:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def fmt_rs(n):
    if n is None:
        return "—"
    return f"R$ {n:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def colorir_linhas_tabela(df_pd):
    """Colore CÉLULA POR CÉLULA (igual à planilha MOTZ)."""
    def _pintar(coluna, row):
        status = str(row.get("Status", "")).strip().upper()
        sit_saldo = str(row.get("Situação Saldo", "")).strip().lower()
        saldo_aberto = "aberto" in sit_saldo

        if status == "OK":
            c_status = COLORS["OK"]
        elif status == "ATUA MAIOR":
            diff = abs(row.get("Diferença MOTZ×ATUA") or 0)
            c_status = COLORS["ATUA MAIOR"] if diff > 100 else COLORS["ATUA MENOR"]
        elif status == "ATUA MENOR":
            c_status = COLORS["ATUA MENOR"]
        elif "ENCONTRADO" in status:
            c_status = COLORS["NAO ENCONTRADO"]
        else:
            c_status = None

        if coluna == "Situação Saldo" and saldo_aberto:
            c = COLORS["SALDO ABERTO"]
            return f"background-color: {c['bg']}; color: {c['fg']}; font-weight: 500;"

        if c_status and coluna in ("Status", "Diferença MOTZ×ATUA", "Vlr. Saldo"):
            return f"background-color: {c_status['bg']}; color: {c_status['fg']}; font-weight: 500;"

        return ""

    def _estilo(row):
        return [_pintar(col, row) for col in row.index]

    return df_pd.style.apply(_estilo, axis=1)


def rodar_conciliacao(pdfs_bytes, motz_bytes, atua_bytes, motz_name, atua_name):
    script_path = Path(__file__).parent / "scripts" / "conciliacao.py"
    if not script_path.exists():
        raise FileNotFoundError("scripts/conciliacao.py não encontrado.")
    tmpdir = tempfile.mkdtemp(prefix="motz_")
    try:
        uploads = Path(tmpdir) / "uploads"
        uploads.mkdir()
        motz_path = uploads / motz_name
        motz_path.write_bytes(motz_bytes)
        atua_path = uploads / atua_name
        atua_path.write_bytes(atua_bytes)
        pdf_paths = []
        seen_hashes = set()
        for name, data in pdfs_bytes:
            h = hashlib.md5(data).hexdigest()
            if h in seen_hashes:
                continue
            seen_hashes.add(h)
            p = uploads / name
            p.write_bytes(data)
            pdf_paths.append(str(p))
        output_path = Path(tmpdir) / "conciliacao_final.xlsx"
        cmd = [sys.executable, str(script_path), "--motz", str(motz_path), "--atua", str(atua_path), "--pdfs", *pdf_paths, "--output", str(output_path)]
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=300)
        if result.returncode != 0:
            raise RuntimeError(f"Erro:\n{result.stdout}\n{result.stderr}")
        if not output_path.exists():
            raise RuntimeError(f"Arquivo não gerado.\n{result.stdout}\n{result.stderr}")
        return output_path.read_bytes(), result.stdout
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


# Ordem e nomes oficiais das 22 colunas
COLUNAS_OFICIAIS = [
    "Cliente", "Contrato", "TITULO (NFe)", "nr_ctrc ATUA", "Nº Carta Frete",
    "Motorista", "Nº Romaneio", "Data Emissão",
    "Vlr. Frete Líquido", "Vlr. Adiantamento", "Vlr. Saldo", "Soma Adto+Saldo",
    "vl_quebra_avaria", "Diverg. Interna (Quebra/descontos) MOTZ",
    "vl_total ATUA", "Diferença MOTZ×ATUA", "Status",
    "Data Emissão Repom", "Data Transferência", "Valor Transferido",
    "Situação Adto", "Situação Saldo",
]

COLUNAS_VALOR = {
    "Vlr. Frete Líquido", "Vlr. Adiantamento", "Vlr. Saldo", "Soma Adto+Saldo",
    "vl_quebra_avaria", "Diverg. Interna (Quebra/descontos) MOTZ",
    "vl_total ATUA", "Diferença MOTZ×ATUA", "Valor Transferido",
}

COLUNAS_DATA = {"Data Emissão", "Data Emissão Repom", "Data Transferência"}


def processar_xlsx(xlsx_bytes):
    """Lê o XLSX de conciliação preservando 1 linha por transferência."""
    xl = pd.ExcelFile(io.BytesIO(xlsx_bytes))
    sheet = next((s for s in xl.sheet_names if "concilia" in s.lower()), xl.sheet_names[0])
    df = pd.read_excel(io.BytesIO(xlsx_bytes), sheet_name=sheet)

    import re as _re
    def find_col(patterns):
        for pat in patterns:
            for c in df.columns:
                if _re.search(pat, str(c), _re.IGNORECASE):
                    return c
        return None

    MAPA = {
        "Cliente":                                  find_col([r"^Cliente"]),
        "Contrato":                                 find_col([r"^Contrato"]),
        "TITULO (NFe)":                             find_col([r"TITULO.*NFe", r"^TITULO", r"^NFe"]),
        "nr_ctrc ATUA":                             find_col([r"nr_ctrc.*ATUA", r"^nr_ctrc", r"^CTRC$"]),
        "Nº Carta Frete":                           find_col([r"N.*Carta.*Frete", r"Carta.Frete"]),
        "Motorista":                                find_col([r"^Motorista"]),
        "Nº Romaneio":                              find_col([r"N.*Romaneio", r"^Romaneio"]),
        "Data Emissão":                             find_col([r"Data Emiss[aã]o$", r"^Data Emiss[aã]o[^R]*$"]),
        "Vlr. Frete Líquido":                       find_col([r"Vlr.*Frete.*L[ií]quido", r"Frete L[ií]quido"]),
        "Vlr. Adiantamento":                        find_col([r"Vlr.*Adiantamento", r"^Adiantamento"]),
        "Vlr. Saldo":                               find_col([r"Vlr\. Saldo", r"^Saldo$"]),
        "Soma Adto+Saldo":                          find_col([r"Soma.*Adto.*Saldo", r"Adto.*Saldo"]),
        "vl_quebra_avaria":                         find_col([r"vl_quebra.avaria", r"quebra.*avaria"]),
        "Diverg. Interna (Quebra/descontos) MOTZ":  find_col([r"Diverg.*Interna", r"Diverg.*MOTZ"]),
        "vl_total ATUA":                            find_col([r"vl_total.*ATUA", r"vl_total"]),
        "Diferença MOTZ×ATUA":                      find_col([r"Diferen.a.*MOTZ.*ATUA", r"Diferen.a.*ATUA"]),
        "Status":                                   find_col([r"^Status"]),
        "Data Emissão Repom":                       find_col([r"Data.*Emiss.*Repom", r"Emiss.*Repom"]),
        "Data Transferência":                       find_col([r"Data.*Transfer", r"Transfer[êe]ncia"]),
        "Valor Transferido":                        find_col([r"Valor Transferido", r"Vlr.*Transferido"]),
        "Situação Adto":                            find_col([r"Situa..o Adto", r"Situa.*Adto"]),
        "Situação Saldo":                           find_col([r"Situa..o Saldo", r"Situa.*Saldo"]),
    }

    if not MAPA["Contrato"] or not MAPA["Status"]:
        raise ValueError(f"Planilha não reconhecida. Colunas encontradas: {list(df.columns)}")

    linhas = []
    for _, row in df.iterrows():
        contrato = str(row[MAPA["Contrato"]]).strip() if pd.notna(row[MAPA["Contrato"]]) else ""
        if not contrato or contrato == "nan":
            continue
        nova = {}
        for col_oficial, col_origem in MAPA.items():
            if col_origem is None:
                nova[col_oficial] = None if col_oficial in COLUNAS_VALOR or col_oficial in COLUNAS_DATA else ""
                continue
            val = row[col_origem]
            if col_oficial in COLUNAS_VALOR:
                nova[col_oficial] = parse_rs(val) if pd.notna(val) and str(val).strip() not in ("", "nan") else None
            elif col_oficial in COLUNAS_DATA:
                nova[col_oficial] = parse_date_br(val)
            else:
                nova[col_oficial] = str(val).strip() if pd.notna(val) else ""
        linhas.append(nova)

    df_out = pd.DataFrame(linhas)
    for col in COLUNAS_OFICIAIS:
        if col not in df_out.columns:
            df_out[col] = None
    df_out = df_out[COLUNAS_OFICIAIS]
    return df_out


# ============================================================
# UPLOAD
# ============================================================
with st.container(border=True):
    st.markdown("### 📂 Arquivos de entrada")
    st.caption("Suba os 3 tipos de arquivo da conciliação. O sistema roda o script da skill e gera a planilha consolidada automaticamente.")
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown("**PDFs Repom**")
        pdfs = st.file_uploader("Transferências bancárias", type=["pdf"], accept_multiple_files=True, key="pdfs", label_visibility="collapsed")
        if pdfs:
            st.caption(f"✓ {len(pdfs)} PDF(s)")
    with col2:
        st.markdown("**Arquivo MOTZ**")
        motz = st.file_uploader("export*.xlsx", type=["xlsx"], key="motz", label_visibility="collapsed")
        if motz:
            st.caption(f"✓ {motz.name}")
    with col3:
        st.markdown("**Cobrança ATUA**")
        atua = st.file_uploader("*cobranca*.xls", type=["xls", "xlsx"], key="atua", label_visibility="collapsed")
        if atua:
            st.caption(f"✓ {atua.name}")
    col_b1, col_b2, _ = st.columns([1, 1, 3])
    with col_b1:
        rodar_btn = st.button("🔄 Rodar conciliação", type="primary", use_container_width=True, disabled=not (pdfs and motz and atua))
    with col_b2:
        carregar_existente = st.button("📥 Carregar XLSX pronto", use_container_width=True)

if carregar_existente:
    st.session_state["modo_xlsx_pronto"] = True

if st.session_state.get("modo_xlsx_pronto"):
    with st.container(border=True):
        st.markdown("### Carregar planilha de conciliação já gerada")
        xlsx_pronto = st.file_uploader("conciliacao_motz_completa.xlsx", type=["xlsx"], key="xlsx_pronto")
        if xlsx_pronto:
            try:
                df = processar_xlsx(xlsx_pronto.read())
                st.session_state["df"] = df
                st.session_state["origem"] = f"Planilha carregada: {xlsx_pronto.name}"
                st.success(f"✓ {len(df)} transferências carregadas")
            except Exception as e:
                st.error(f"Erro ao processar: {e}")

if rodar_btn and pdfs and motz and atua:
    with st.spinner("Rodando conciliação... 30s-2min."):
        try:
            pdfs_data = [(f.name, f.read()) for f in pdfs]
            motz_data = motz.read()
            atua_data = atua.read()
            xlsx_bytes, log = rodar_conciliacao(pdfs_data, motz_data, atua_data, motz.name, atua.name)
            df = processar_xlsx(xlsx_bytes)
            st.session_state["df"] = df
            st.session_state["xlsx_bytes"] = xlsx_bytes
            st.session_state["origem"] = f"Conciliação rodada às {datetime.now().strftime('%H:%M:%S')} · {len(pdfs_data)} PDFs + {motz.name} + {atua.name}"
            st.session_state["log"] = log
            st.success(f"✓ Conciliação concluída · {len(df)} transferências")
        except Exception as e:
            st.error(f"Erro na conciliação:\n\n{str(e)}")
            st.stop()# ============================================================
# Dashboard
# ============================================================
if "df" in st.session_state:
    df = st.session_state["df"]
    st.divider()

    if "status_click" not in st.session_state:
        st.session_state["status_click"] = None

    col_a, col_b = st.columns([3, 1])
    with col_a:
        st.caption(f"🟢 {st.session_state.get('origem', '')}")
    with col_b:
        if "xlsx_bytes" in st.session_state:
            st.download_button(
                "⬇️ Baixar XLSX consolidado",
                data=st.session_state["xlsx_bytes"],
                file_name=f"conciliacao_motz_{datetime.now().strftime('%Y-%m-%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

    st.info(
        "**1 linha por transferência Repom** — igual à planilha da skill. "
        "Se um contrato teve 2 transferências (ex: adto + saldo), aparece 2 vezes. "
        "As colunas coloridas destacam divergências e saldo em aberto.",
        icon="ℹ️",
    )

    # Filtros
    with st.container(border=True):
        datas_validas = df["Data Emissão"].dropna()
        if len(datas_validas) > 0:
            date_min = datas_validas.min().date()
            date_max = datas_validas.max().date()
        else:
            date_min = date_max = datetime.now().date()

        col_f1, col_f2, col_f3, col_f4 = st.columns([1, 1, 1, 2])
        with col_f1:
            date_from = st.date_input("De", value=date_min, min_value=date_min, max_value=date_max)
        with col_f2:
            date_to = st.date_input("Até", value=date_max, min_value=date_min, max_value=date_max)
        with col_f3:
            status_filter = st.selectbox(
                "Status",
                ["Todos", "OK", "ATUA MAIOR", "ATUA MENOR", "NÃO ENCONTRADO", "Saldo aberto"],
                key="status_dropdown",
            )
        with col_f4:
            busca = st.text_input("Buscar", placeholder="contrato, CTRC, motorista, NFe...")

    # Aplicar filtros de data
    df_f = df.copy()
    if date_from:
        df_f = df_f[df_f["Data Emissão"].apply(lambda d: pd.isna(d) or d.date() >= date_from)]
    if date_to:
        df_f = df_f[df_f["Data Emissão"].apply(lambda d: pd.isna(d) or d.date() <= date_to)]

    df_periodo = df_f.copy()

    filtro_ativo = st.session_state.get("status_click")
    if status_filter != "Todos":
        filtro_ativo = status_filter

    if filtro_ativo == "Saldo aberto":
        df_f = df_f[df_f["Situação Saldo"] == "Aberto"]
    elif filtro_ativo and filtro_ativo != "Todos":
        df_f = df_f[df_f["Status"] == filtro_ativo]

    if busca:
        b = busca.lower()
        mask = (
            df_f["Contrato"].astype(str).str.lower().str.contains(b, na=False) |
            df_f["nr_ctrc ATUA"].astype(str).str.lower().str.contains(b, na=False) |
            df_f["Motorista"].astype(str).str.lower().str.contains(b, na=False) |
            df_f["TITULO (NFe)"].astype(str).str.lower().str.contains(b, na=False) |
            df_f["Cliente"].astype(str).str.lower().str.contains(b, na=False) |
            df_f["Nº Carta Frete"].astype(str).str.lower().str.contains(b, na=False) |
            df_f["Nº Romaneio"].astype(str).str.lower().str.contains(b, na=False)
        )
        df_f = df_f[mask]

    # ============================================================
    # KPIs (deduplica por contrato para somas)
    # ============================================================
    total_linhas = len(df_periodo)
    df_unicos = df_periodo.drop_duplicates(subset=["Contrato"]) if total_linhas > 0 else df_periodo
    total = len(df_unicos)

    ok_n = (df_unicos["Status"] == "OK").sum()
    maior_n = (df_unicos["Status"] == "ATUA MAIOR").sum()
    menor_n = (df_unicos["Status"] == "ATUA MENOR").sum()
    ne_n = (df_unicos["Status"] == "NÃO ENCONTRADO").sum()
    aberto_n = (df_unicos["Situação Saldo"] == "Aberto").sum()

    soma_motz = df_unicos["Vlr. Frete Líquido"].fillna(0).sum()
    soma_atua = df_unicos["vl_total ATUA"].fillna(0).sum()
    soma_transf = df_periodo["Valor Transferido"].fillna(0).sum()
    contratos_com_transf = df_periodo[df_periodo["Valor Transferido"].fillna(0) > 0]["Contrato"].nunique()
    soma_saldo_aberto = df_unicos[df_unicos["Situação Saldo"] == "Aberto"]["Vlr. Saldo"].fillna(0).sum()

    indice = (ok_n / total * 100) if total else 0

    col_k1, col_k2, col_k3, col_k4, col_k5 = st.columns(5)
    with col_k1:
        st.metric("Índice conciliação", f"{indice:.1f}%".replace(".", ","), f"{ok_n} de {total} OK")
    with col_k2:
        st.metric("Soma MOTZ", fmt_mi(soma_motz), help="Frete líquido (sem duplicar por transferência)")
    with col_k3:
        diff = soma_atua - soma_motz
        st.metric("Soma ATUA", fmt_mi(soma_atua), delta=fmt_mi(diff), delta_color="inverse")
    with col_k4:
        st.metric("Transferido Repom", fmt_mi(soma_transf), f"{contratos_com_transf} contratos c/ PDF")
    with col_k5:
        st.metric("Saldo em aberto", fmt_mi(soma_saldo_aberto), f"{aberto_n} contratos")

    # ============================================================
    # Distribuição clicável (cards + pizza)
    # ============================================================
    st.markdown("### Distribuição por status")
    st.caption("👆 Clique em um card ou numa fatia do gráfico para filtrar a tabela. Clique de novo para limpar.")

    def pct(n):
        return f"{n/total*100:.1f}%".replace(".", ",") if total else "0,0%"

    cards = [
        ("OK", ok_n, "card-ok", "🟢"),
        ("ATUA MAIOR", maior_n, "card-maior", "🔴"),
        ("ATUA MENOR", menor_n, "card-menor", "🔵"),
        ("NÃO ENCONTRADO", ne_n, "card-ne", "🟡"),
        ("Saldo aberto", aberto_n, "card-aberto", "🟣"),
    ]
    cols_cards = st.columns(5)
    filtro_atual_clique = st.session_state.get("status_click")
    for idx, (label, n, css_class, emoji) in enumerate(cards):
        with cols_cards[idx]:
            ativo = (filtro_atual_clique == label)
            classe_final = f"{css_class} card-active" if ativo else css_class
            st.markdown(f'<div class="{classe_final}">', unsafe_allow_html=True)
            if st.button(f"{emoji}  **{label}**\n\n{n} · {pct(n)}", key=f"card_{label}", use_container_width=True):
                if st.session_state.get("status_click") == label:
                    st.session_state["status_click"] = None
                else:
                    st.session_state["status_click"] = label
                st.rerun()
            st.markdown('</div>', unsafe_allow_html=True)

    if filtro_atual_clique and status_filter == "Todos":
        st.markdown(
            f'<div class="filter-hint">🎯 Filtro ativo pelo card: <b>{filtro_atual_clique}</b> '
            f'· Exibindo {len(df_f)} de {total_linhas} linhas. Clique no mesmo card para limpar.</div>',
            unsafe_allow_html=True,
        )

    col_g1, col_g2 = st.columns([1, 1])
    with col_g1:
        st.markdown("**Distribuição visual**")
        dist_data = pd.DataFrame([
            {"Status": "OK", "Qtd": ok_n, "Cor": COLORS["OK"]["plot"]},
            {"Status": "ATUA MAIOR", "Qtd": maior_n, "Cor": COLORS["ATUA MAIOR"]["plot"]},
            {"Status": "ATUA MENOR", "Qtd": menor_n, "Cor": COLORS["ATUA MENOR"]["plot"]},
            {"Status": "NÃO ENCONTRADO", "Qtd": ne_n, "Cor": COLORS["NAO ENCONTRADO"]["plot"]},
            {"Status": "Saldo aberto", "Qtd": aberto_n, "Cor": COLORS["SALDO ABERTO"]["plot"]},
        ])
        dist_data = dist_data[dist_data["Qtd"] > 0]
        if len(dist_data) > 0:
            fig = px.pie(
                dist_data, values="Qtd", names="Status", color="Status",
                color_discrete_map={r["Status"]: r["Cor"] for _, r in dist_data.iterrows()},
                hole=0.4,
            )
            fig.update_traces(
                textposition="inside",
                textinfo="percent+label",
                hovertemplate="<b>%{label}</b><br>%{value} contratos<br>%{percent}<extra></extra>",
                pull=[0.08 if s == filtro_atual_clique else 0 for s in dist_data["Status"]],
            )
            fig.update_layout(
                height=280,
                margin=dict(t=10, b=10, l=10, r=10),
                showlegend=True,
                legend=dict(orientation="v", yanchor="middle", y=0.5, xanchor="left", x=1.05, font=dict(size=11)),
            )
            evento = st.plotly_chart(fig, use_container_width=True, on_select="rerun", key="pie_chart")
            if evento and evento.get("selection") and evento["selection"].get("points"):
                ponto = evento["selection"]["points"][0]
                status_clicado = ponto.get("label")
                if status_clicado:
                    if st.session_state.get("status_click") == status_clicado:
                        st.session_state["status_click"] = None
                    else:
                        st.session_state["status_click"] = status_clicado
                    st.rerun()
        else:
            st.caption("Sem dados para plotar")

    with col_g2:
        st.markdown("**Frete líquido emitido por dia**")
        df_chart = df_unicos.dropna(subset=["Data Emissão"]).copy()
        if len(df_chart) > 0:
            df_chart["Dia"] = df_chart["Data Emissão"].dt.date
            daily = df_chart.groupby("Dia")["Vlr. Frete Líquido"].sum().reset_index()
            daily.columns = ["Data", "Frete Líquido"]
            st.bar_chart(daily, x="Data", y="Frete Líquido", height=280, color="#378ADD")
        else:
            st.caption("Sem dados de data para o período")

    # ============================================================
    # Tabela com 22 colunas + multi-select
    # ============================================================
    st.markdown(f"**Transferências · {len(df_f)} linhas exibidas** " +
                (f"(filtro: {filtro_ativo})" if filtro_ativo and filtro_ativo != "Todos" else ""))

    PADRAO_VISIVEIS = [
        "Cliente", "Contrato", "TITULO (NFe)", "Motorista",
        "Data Emissão", "Vlr. Frete Líquido", "Vlr. Saldo",
        "vl_total ATUA", "Diferença MOTZ×ATUA", "Status",
        "Data Transferência", "Valor Transferido",
        "Situação Adto", "Situação Saldo",
    ]

    with st.expander("⚙️ Escolher colunas visíveis", expanded=False):
        colunas_escolhidas = st.multiselect(
            "Selecione as colunas que deseja ver (todas as 22 disponíveis):",
            options=COLUNAS_OFICIAIS,
            default=PADRAO_VISIVEIS,
            key="colunas_multiselect",
        )
        col_rst1, col_rst2, _ = st.columns([1, 1, 3])
        with col_rst1:
            if st.button("✅ Mostrar todas"):
                st.session_state["colunas_multiselect"] = COLUNAS_OFICIAIS
                st.rerun()
        with col_rst2:
            if st.button("↺ Padrão"):
                st.session_state["colunas_multiselect"] = PADRAO_VISIVEIS
                st.rerun()

    if not colunas_escolhidas:
        colunas_escolhidas = PADRAO_VISIVEIS

    colunas_ordenadas = [c for c in COLUNAS_OFICIAIS if c in colunas_escolhidas]

    df_show = df_f.copy().sort_values(
        ["Data Emissão", "Contrato"],
        ascending=[False, True],
        na_position="last",
    )
    df_tabela = df_show[colunas_ordenadas].reset_index(drop=True)

    formatadores = {}
    for col in colunas_ordenadas:
        if col in COLUNAS_VALOR:
            formatadores[col] = lambda v: (
                f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                if pd.notna(v) and v != 0 else ("R$ 0,00" if v == 0 else "—")
            )
        elif col in COLUNAS_DATA:
            formatadores[col] = lambda d: d.strftime("%d/%m/%Y") if pd.notna(d) and isinstance(d, datetime) else "—"

    styler = colorir_linhas_tabela(df_tabela)
    if formatadores:
        styler = styler.format(formatadores)

    st.dataframe(styler, use_container_width=True, hide_index=True, height=520)

    with st.expander("🎨 Legenda de cores", expanded=False):
        st.markdown(f"""
        As cores seguem **exatamente** a planilha MOTZ original (célula por célula):

        - <span class="status-ok">🟢 OK</span> — colunas **Status**, **Diferença MOTZ×ATUA** e **Vlr. Saldo** em verde
        - <span class="status-maior">🔴 ATUA MAIOR > R$100</span> — mesmas colunas em vermelho (diferença crítica)
        - <span class="status-menor">🔵 ATUA MAIOR até R$100 / ATUA MENOR</span> — mesmas colunas em azul (diferença pequena)
        - <span class="status-ne">🟡 NÃO ENCONTRADO</span> — mesmas colunas em amarelo
        - <span class="status-aberto">🟣 Situação Saldo = Aberto</span> — apenas a coluna **Situação Saldo** em roxo

        Assim você vê ao mesmo tempo: status da conferência MOTZ×ATUA + pendência de saldo.
        """, unsafe_allow_html=True)

    csv = df_f.to_csv(index=False).encode("utf-8")
    st.download_button(
        "⬇️ Baixar tabela filtrada (CSV · todas as colunas)",
        csv,
        f"conciliacao_filtrado_{datetime.now().strftime('%Y-%m-%d')}.csv",
        "text/csv",
    )

else:
    st.info(
        "👆 **Comece subindo os 3 arquivos** (PDFs Repom, MOTZ XLSX, ATUA XLS) e clique em **Rodar conciliação**. Ou use **Carregar XLSX pronto** se você já tem a planilha consolidada gerada.",
        icon="📤",
    )
    with st.expander("ℹ️ Sobre esta ferramenta"):
        st.markdown("""
        Este aplicativo executa a skill `conciliacao-motz` diretamente no servidor.

        **Novidades v3.3:**
        - 📋 1 linha por transferência (igual à planilha original)
        - 📊 22 colunas oficiais da planilha MOTZ
        - ⚙️ Multi-select para escolher colunas visíveis
        - 🎨 Cores célula por célula (Status + Diferença + Vlr. Saldo coloridas, Situação Saldo roxa)
        - 👆 Cards de status e gráfico de pizza clicáveis
        """)

st.divider()
st.caption("Dashboard Conciliação MOTZ · skill conciliacao-motz · Streamlit Cloud · v3.3")
