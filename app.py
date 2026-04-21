"""
Dashboard de Conciliação MOTZ - Streamlit
Upload de PDFs Repom + MOTZ (XLSX) + ATUA (XLS) → conciliação → visualização
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

# ============================================================
# Configuração da página
# ============================================================
st.set_page_config(
    page_title="Conciliação MOTZ",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# CSS para aproximar do visual do dashboard HTML
st.markdown("""
<style>
    .main .block-container { padding-top: 2rem; padding-bottom: 3rem; max-width: 1200px; }
    h1 { font-size: 22px !important; font-weight: 500 !important; margin-bottom: 0 !important; }
    h2 { font-size: 16px !important; font-weight: 500 !important; }
    h3 { font-size: 14px !important; font-weight: 500 !important; }
    .stMetric { background: var(--secondary-background-color); border-radius: 10px; padding: 12px 16px; }
    .stMetric label { font-size: 12px !important; color: #6B6B66 !important; }
    .stMetric [data-testid="stMetricValue"] { font-size: 22px !important; font-weight: 500 !important; }
    .stDataFrame { font-size: 13px; }
    [data-testid="stFileUploader"] section { border-radius: 10px; padding: 14px; }
    .stButton > button { border-radius: 8px; font-size: 13px; padding: 6px 14px; }
    .status-ok { background: #EAF3DE; color: #3B6D11; padding: 2px 8px; border-radius: 999px; font-size: 11px; font-weight: 500; }
    .status-maior { background: #E6F1FB; color: #185FA5; padding: 2px 8px; border-radius: 999px; font-size: 11px; font-weight: 500; }
    .status-menor { background: #FCEBEB; color: #A32D2D; padding: 2px 8px; border-radius: 999px; font-size: 11px; font-weight: 500; }
    .status-ne { background: #FAEEDA; color: #854F0B; padding: 2px 8px; border-radius: 999px; font-size: 11px; font-weight: 500; }
    .status-aberto { background: #EEEDFE; color: #3C3489; padding: 2px 8px; border-radius: 999px; font-size: 11px; font-weight: 500; }
</style>
""", unsafe_allow_html=True)


# ============================================================
# Cabeçalho
# ============================================================
st.markdown("# Conciliação MOTZ consolidada")
st.caption("PDFs Repom × MOTZ (XLSX) × Cobrança ATUA (XLS) — cruzamento automático")


# ============================================================
# Utilitários
# ============================================================
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


# ============================================================
# Executar a skill de conciliação
# ============================================================
def rodar_conciliacao(pdfs_bytes, motz_bytes, atua_bytes, motz_name, atua_name):
    """
    Roda o script scripts/conciliacao.py em um diretório temporário.
    Retorna o caminho do XLSX gerado.
    """
    script_path = Path(__file__).parent / "scripts" / "conciliacao.py"
    if not script_path.exists():
        raise FileNotFoundError(
            "scripts/conciliacao.py não encontrado. Copie o script da skill "
            "/mnt/skills/user/conciliacao-motz/scripts/ para esta pasta."
        )

    tmpdir = tempfile.mkdtemp(prefix="motz_")
    try:
        uploads = Path(tmpdir) / "uploads"
        uploads.mkdir()

        # Salvar MOTZ
        motz_path = uploads / motz_name
        motz_path.write_bytes(motz_bytes)

        # Salvar ATUA
        atua_path = uploads / atua_name
        atua_path.write_bytes(atua_bytes)

        # Salvar PDFs (deduplicar por hash MD5 como a skill faz)
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

        # Rodar a skill
        cmd = [
            "python3", str(script_path),
            "--motz", str(motz_path),
            "--atua", str(atua_path),
            "--pdfs", *pdf_paths,
            "--output", str(output_path),
        ]
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=300)
        if result.returncode != 0:
            raise RuntimeError(
                f"Erro ao rodar conciliação:\n\nSTDOUT:\n{result.stdout}\n\nSTDERR:\n{result.stderr}"
            )

        if not output_path.exists():
            raise RuntimeError(
                f"Script executou mas não gerou o arquivo.\n\n{result.stdout}\n{result.stderr}"
            )

        # Ler o XLSX gerado em memória
        xlsx_bytes = output_path.read_bytes()
        return xlsx_bytes, result.stdout

    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


# ============================================================
# Processar XLSX → contratos únicos (SEM duplicar valores)
# Segue a skill: cada transferência PDF = 1 linha, mas agregamos
# ============================================================
def processar_xlsx(xlsx_bytes):
    """Lê o XLSX de conciliação e agrega por contrato único."""
    # Encontrar a aba certa
    xl = pd.ExcelFile(io.BytesIO(xlsx_bytes))
    sheet = next((s for s in xl.sheet_names if "concilia" in s.lower()), xl.sheet_names[0])
    df = pd.read_excel(io.BytesIO(xlsx_bytes), sheet_name=sheet)

    # Mapear colunas com flexibilidade
    import re as _re
    def find_col(patterns):
        for pat in patterns:
            for c in df.columns:
                if _re.search(pat, str(c), _re.IGNORECASE):
                    return c
        return None

    COL = {
        "cliente": find_col([r"^Cliente"]),
        "contrato": find_col([r"^Contrato"]),
        "nfe": find_col([r"TITULO", r"NFe"]),
        "ctrc": find_col([r"^CTRC$"]),
        "motorista": find_col([r"^Motorista"]),
        "data_emissao": find_col([r"Data Emiss[aã]o$"]),
        "frete_liq": find_col([r"Frete L[ií]quido", r"Vlr.*Frete"]),
        "adiantamento": find_col([r"Adiantamento"]),
        "saldo": find_col([r"^Vlr\. Saldo"]),
        "vl_total_atua": find_col([r"vl_total.*ATUA", r"vl_total"]),
        "diff": find_col([r"Diferen.a.*ATUA", r"Diferen.a MOTZ"]),
        "status": find_col([r"^Status"]),
        "valor_transf": find_col([r"Valor Transferido"]),
        "sit_saldo": find_col([r"Situa..o Saldo"]),
        "sit_adto": find_col([r"Situa..o Adto"]),
    }

    if not COL["contrato"] or not COL["status"]:
        raise ValueError(
            f"Planilha não reconhecida. Colunas encontradas: {list(df.columns)}"
        )

    # Agregar por contrato único
    contratos = {}
    for _, row in df.iterrows():
        c = str(row[COL["contrato"]]).strip() if pd.notna(row[COL["contrato"]]) else ""
        if not c or c == "nan":
            continue
        valor_transf = parse_rs(row[COL["valor_transf"]]) if COL["valor_transf"] else 0

        if c not in contratos:
            vl_atua_raw = row[COL["vl_total_atua"]] if COL["vl_total_atua"] else None
            diff_raw = row[COL["diff"]] if COL["diff"] else None
            contratos[c] = {
                "Cliente": str(row[COL["cliente"]]) if COL["cliente"] and pd.notna(row[COL["cliente"]]) else "",
                "Contrato": c,
                "NFe": str(row[COL["nfe"]]) if COL["nfe"] and pd.notna(row[COL["nfe"]]) else "",
                "CTRC": str(row[COL["ctrc"]]) if COL["ctrc"] and pd.notna(row[COL["ctrc"]]) else "",
                "Motorista": str(row[COL["motorista"]]) if COL["motorista"] and pd.notna(row[COL["motorista"]]) else "",
                "Data Emissão": parse_date_br(row[COL["data_emissao"]]) if COL["data_emissao"] else None,
                "Frete Líquido": parse_rs(row[COL["frete_liq"]]) if COL["frete_liq"] else 0,
                "Adiantamento": parse_rs(row[COL["adiantamento"]]) if COL["adiantamento"] else 0,
                "Saldo": parse_rs(row[COL["saldo"]]) if COL["saldo"] else 0,
                "vl_total ATUA": parse_rs(vl_atua_raw) if vl_atua_raw is not None and str(vl_atua_raw).strip() not in ("", "nan") else None,
                "Diferença": parse_rs(diff_raw) if diff_raw is not None and str(diff_raw).strip() not in ("", "nan") else None,
                "Status": str(row[COL["status"]]).strip() if pd.notna(row[COL["status"]]) else "",
                "Situação Saldo": str(row[COL["sit_saldo"]]).strip() if COL["sit_saldo"] and pd.notna(row[COL["sit_saldo"]]) else "",
                "qtd_transf": 0,
                "Transferido": 0.0,
            }
        if valor_transf > 0:
            contratos[c]["qtd_transf"] += 1
            contratos[c]["Transferido"] += valor_transf

    return pd.DataFrame(list(contratos.values()))


# ============================================================
# UPLOAD
# ============================================================
with st.container(border=True):
    st.markdown("### 📂 Arquivos de entrada")
    st.caption(
        "Suba os 3 tipos de arquivo da conciliação. O sistema roda o script da skill "
        "e gera a planilha consolidada automaticamente."
    )

    col1, col2, col3 = st.columns(3)

    with col1:
        st.markdown("**PDFs Repom**")
        pdfs = st.file_uploader(
            "Transferências bancárias",
            type=["pdf"],
            accept_multiple_files=True,
            key="pdfs",
            label_visibility="collapsed",
        )
        if pdfs:
            st.caption(f"✓ {len(pdfs)} PDF(s)")

    with col2:
        st.markdown("**Arquivo MOTZ**")
        motz = st.file_uploader(
            "export*.xlsx",
            type=["xlsx"],
            key="motz",
            label_visibility="collapsed",
        )
        if motz:
            st.caption(f"✓ {motz.name}")

    with col3:
        st.markdown("**Cobrança ATUA**")
        atua = st.file_uploader(
            "*cobranca*.xls",
            type=["xls", "xlsx"],
            key="atua",
            label_visibility="collapsed",
        )
        if atua:
            st.caption(f"✓ {atua.name}")

    col_b1, col_b2, _ = st.columns([1, 1, 3])
    with col_b1:
        rodar_btn = st.button("🔄 Rodar conciliação", type="primary", use_container_width=True, disabled=not (pdfs and motz and atua))
    with col_b2:
        carregar_existente = st.button("📥 Carregar XLSX pronto", use_container_width=True)


# ============================================================
# Opção alternativa: carregar XLSX já gerado
# ============================================================
if carregar_existente:
    st.session_state["modo_xlsx_pronto"] = True

if st.session_state.get("modo_xlsx_pronto"):
    with st.container(border=True):
        st.markdown("### Carregar planilha de conciliação já gerada")
        xlsx_pronto = st.file_uploader(
            "conciliacao_motz_completa.xlsx",
            type=["xlsx"],
            key="xlsx_pronto",
        )
        if xlsx_pronto:
            try:
                df = processar_xlsx(xlsx_pronto.read())
                st.session_state["df"] = df
                st.session_state["origem"] = f"Planilha carregada: {xlsx_pronto.name}"
                st.success(f"✓ {len(df)} contratos carregados")
            except Exception as e:
                st.error(f"Erro ao processar: {e}")


# ============================================================
# Rodar conciliação
# ============================================================
if rodar_btn and pdfs and motz and atua:
    with st.spinner("Rodando conciliação... isso pode levar 30s-2min dependendo do tamanho dos arquivos."):
        try:
            pdfs_data = [(f.name, f.read()) for f in pdfs]
            motz_data = motz.read()
            atua_data = atua.read()

            xlsx_bytes, log = rodar_conciliacao(
                pdfs_data, motz_data, atua_data, motz.name, atua.name
            )
            df = processar_xlsx(xlsx_bytes)
            st.session_state["df"] = df
            st.session_state["xlsx_bytes"] = xlsx_bytes
            st.session_state["origem"] = (
                f"Conciliação rodada às {datetime.now().strftime('%H:%M:%S')} · "
                f"{len(pdfs_data)} PDFs + {motz.name} + {atua.name}"
            )
            st.session_state["log"] = log

            st.success(f"✓ Conciliação concluída · {len(df)} contratos únicos")

        except Exception as e:
            st.error(f"Erro na conciliação:\n\n{str(e)}")
            st.stop()


# ============================================================
# Dashboard (quando houver dados)
# ============================================================
if "df" in st.session_state:
    df = st.session_state["df"]
    st.divider()

    # Cabeçalho do dashboard
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

    # Nota sobre agregação
    st.info(
        "**Agregação por contrato único.** A skill gera 1 linha por transferência Repom. "
        "Este dashboard agrupa por contrato para não duplicar Frete Líquido e vl_total ATUA. "
        "O Valor Transferido é somado quando há múltiplas transferências.",
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
            )
        with col_f4:
            busca = st.text_input("Buscar", placeholder="contrato, CTRC, motorista, NFe...")

    # Aplicar filtros
    df_f = df.copy()
    if date_from:
        df_f = df_f[df_f["Data Emissão"].apply(lambda d: d is None or d.date() >= date_from)]
    if date_to:
        df_f = df_f[df_f["Data Emissão"].apply(lambda d: d is None or d.date() <= date_to)]

    df_periodo = df_f.copy()  # para KPIs do período (sem filtro de status)

    if status_filter == "Saldo aberto":
        df_f = df_f[df_f["Situação Saldo"] == "Aberto"]
    elif status_filter != "Todos":
        df_f = df_f[df_f["Status"] == status_filter]

    if busca:
        b = busca.lower()
        mask = (
            df_f["Contrato"].str.lower().str.contains(b, na=False) |
            df_f["CTRC"].str.lower().str.contains(b, na=False) |
            df_f["Motorista"].str.lower().str.contains(b, na=False) |
            df_f["NFe"].str.lower().str.contains(b, na=False) |
            df_f["Cliente"].str.lower().str.contains(b, na=False)
        )
        df_f = df_f[mask]

    # ============================================================
    # KPIs
    # ============================================================
    total = len(df_periodo)
    ok_n = (df_periodo["Status"] == "OK").sum()
    maior_n = (df_periodo["Status"] == "ATUA MAIOR").sum()
    menor_n = (df_periodo["Status"] == "ATUA MENOR").sum()
    ne_n = (df_periodo["Status"] == "NÃO ENCONTRADO").sum()
    aberto_n = (df_periodo["Situação Saldo"] == "Aberto").sum()

    soma_motz = df_periodo["Frete Líquido"].sum()
    soma_atua = df_periodo["vl_total ATUA"].fillna(0).sum()
    soma_transf = df_periodo["Transferido"].sum()
    soma_saldo_aberto = df_periodo[df_periodo["Situação Saldo"] == "Aberto"]["Saldo"].sum()

    indice = (ok_n / total * 100) if total else 0

    col_k1, col_k2, col_k3, col_k4, col_k5 = st.columns(5)
    with col_k1:
        st.metric("Índice conciliação", f"{indice:.1f}%".replace(".", ","), f"{ok_n} de {total} OK")
    with col_k2:
        st.metric("Soma MOTZ", fmt_mi(soma_motz), help="Frete líquido sem duplicar")
    with col_k3:
        diff = soma_atua - soma_motz
        st.metric("Soma ATUA", fmt_mi(soma_atua), delta=fmt_mi(diff), delta_color="inverse")
    with col_k4:
        st.metric("Transferido Repom", fmt_mi(soma_transf), f"{(df_periodo['qtd_transf']>0).sum()} com PDF")
    with col_k5:
        st.metric("Saldo em aberto", fmt_mi(soma_saldo_aberto), f"{aberto_n} contratos")

    # ============================================================
    # Gráfico + Distribuição de status
    # ============================================================
    col_g1, col_g2 = st.columns([1, 2])

    with col_g1:
        st.markdown("**Distribuição por status**")
        def pct(n):
            return f"{n/total*100:.1f}%".replace(".", ",") if total else "0,0%"
        st.markdown(f"""
        <div style="font-size: 13px; line-height: 2;">
            <span class="status-ok">OK</span> &nbsp; {ok_n} · {pct(ok_n)}<br>
            <span class="status-maior">ATUA maior</span> &nbsp; {maior_n} · {pct(maior_n)}<br>
            <span class="status-menor">ATUA menor</span> &nbsp; {menor_n} · {pct(menor_n)}<br>
            <span class="status-ne">Não encontrado</span> &nbsp; {ne_n} · {pct(ne_n)}<br>
            <span class="status-aberto">Saldo aberto</span> &nbsp; {aberto_n} · {pct(aberto_n)}
        </div>
        """, unsafe_allow_html=True)

    with col_g2:
        st.markdown("**Frete líquido emitido por dia**")
        df_chart = df_periodo.dropna(subset=["Data Emissão"]).copy()
        if len(df_chart) > 0:
            df_chart["Dia"] = df_chart["Data Emissão"].dt.date
            daily = df_chart.groupby("Dia")["Frete Líquido"].sum().reset_index()
            daily.columns = ["Data", "Frete Líquido"]
            st.bar_chart(daily, x="Data", y="Frete Líquido", height=220, color="#378ADD")
        else:
            st.caption("Sem dados de data para o período")

    # ============================================================
    # Tabela
    # ============================================================
    st.markdown(f"**Contratos · {len(df_f)} exibidos** " +
                (f"(filtro: {status_filter})" if status_filter != "Todos" else ""))

    df_show = df_f.copy().sort_values("Data Emissão", ascending=False, na_position="last")
    df_show["Data"] = df_show["Data Emissão"].apply(lambda d: d.strftime("%d/%m/%Y") if d else "—")
    df_show["Transf."] = df_show.apply(
        lambda r: fmt_rs(r["Transferido"]) + (f" ({int(r['qtd_transf'])}×)" if r["qtd_transf"] > 1 else "") if r["qtd_transf"] > 0 else "—",
        axis=1,
    )
    df_show["Saldo Status"] = df_show.apply(
        lambda r: fmt_rs(r["Saldo"]) + " (aberto)" if r["Situação Saldo"] == "Aberto" else "Pago",
        axis=1,
    )

    st.dataframe(
        df_show[[
            "Data", "Contrato", "Motorista", "CTRC",
            "Frete Líquido", "vl_total ATUA", "Diferença",
            "Transf.", "Status", "Saldo Status",
        ]],
        use_container_width=True,
        hide_index=True,
        column_config={
            "Frete Líquido": st.column_config.NumberColumn(format="R$ %.2f"),
            "vl_total ATUA": st.column_config.NumberColumn(format="R$ %.2f"),
            "Diferença": st.column_config.NumberColumn(format="R$ %.2f"),
        },
        height=480,
    )

    # Export dos filtrados
    csv = df_show.to_csv(index=False).encode("utf-8")
    st.download_button(
        "⬇️ Baixar tabela filtrada (CSV)",
        csv,
        f"conciliacao_filtrado_{datetime.now().strftime('%Y-%m-%d')}.csv",
        "text/csv",
    )

else:
    # Tela inicial sem dados
    st.info(
        "👆 **Comece subindo os 3 arquivos** (PDFs Repom, MOTZ XLSX, ATUA XLS) e clique em "
        "**Rodar conciliação**. Ou use **Carregar XLSX pronto** se você já tem a planilha consolidada gerada.",
        icon="📤",
    )

    with st.expander("ℹ️ Sobre esta ferramenta"):
        st.markdown("""
        Este aplicativo executa a skill `conciliacao-motz` diretamente no servidor, fazendo o cruzamento entre três fontes:

        1. **PDFs Repom** — transferências bancárias (chave: Contrato)
        2. **Arquivo MOTZ** — relatório de cartas-frete (chave: Nº formulário = Contrato do PDF)
        3. **Cobrança ATUA** — API Contabilidade (chave: nr_nf = NF cliente do MOTZ)

        **Saídas:**
        - Planilha XLSX com formatação condicional (verde/vermelho/azul/amarelo/roxo)
        - Dashboard interativo com KPIs, filtros e busca
        - Exportação filtrada em CSV

        **Limite:** 200 MB por arquivo (configurável no Streamlit).
        """)

# Footer
st.divider()
st.caption("Dashboard Conciliação MOTZ · skill conciliacao-motz · Streamlit Cloud")
