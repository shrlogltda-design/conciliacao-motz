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
            st.stop()
