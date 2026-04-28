# ============================================================
# 🆕 BOTÃO XLS BAIXA TÍTULO DE CRÉDITO ATUA
# ============================================================
# Cole este bloco LOGO ABAIXO do botão "⬇️ XLSX do último processamento"
# (dentro do bloco `if "df" in st.session_state:`)
# ============================================================

def _gerar_csv_baixa_atua(df):
    """Gera CSV no formato exato da planilha modelo do ATUA pra baixa de títulos."""
    import io
    
    df_f = df.copy()
    
    # Filtrar: tem Data Transferência E Valor Transferido > 0
    if "Data Transferência" in df_f.columns:
        df_f = df_f[df_f["Data Transferência"].notna()]
    df_f = df_f[df_f["Valor Transferido"].fillna(0) > 0]
    
    if df_f.empty:
        return None, 0
    
    # Decidir Tipo Parcela: A (adiantamento) ou S (saldo)
    # Cada linha do dashboard é UMA transferência. Comparamos o valor
    # transferido com Vlr. Adiantamento e Vlr. Saldo do contrato.
    def _decidir_tipo(row):
        vt = float(row.get("Valor Transferido", 0) or 0)
        va = float(row.get("Vlr. Adiantamento", 0) or 0)
        vs = float(row.get("Vlr. Saldo", 0) or 0)
        # Tolerância de R$ 0,50 pra arredondamento
        if abs(vt - va) <= 0.50:
            return "A"
        if abs(vt - vs) <= 0.50:
            return "S"
        # Fallback: o que estiver mais perto
        return "A" if abs(vt - va) < abs(vt - vs) else "S"
    
    df_f["_tipo"] = df_f.apply(_decidir_tipo, axis=1)
    
    # Formatadores
    def _fmt_br(v):
        if pd.isna(v) or v == "" or v is None:
            return ""
        try:
            return f"{float(v):.2f}".replace(".", ",")
        except (ValueError, TypeError):
            return ""
    
    def _fmt_quebra(v):
        try:
            f = float(v) if pd.notna(v) else 0
            if f > 0:
                return f"{f:.2f}".replace(".", ",")
        except (ValueError, TypeError):
            pass
        return ""
    
    def _fmt_int(v):
        if pd.isna(v) or v == "":
            return ""
        try:
            return str(int(float(v)))
        except (ValueError, TypeError):
            return str(v).strip()
    
    # Montar saída
    out = pd.DataFrame({
        "Nr. CTRC": df_f.get("nr_ctrc ATUA", pd.Series(dtype=str)).apply(_fmt_int),
        "Serie": "1",
        "Valor Pago": df_f["Valor Transferido"].apply(_fmt_br),
        "Valor Desconto": "",
        "Valor de Juros": "",
        "Valor Desconto Quebra": df_f.get(
            "Diverg. Interna (Quebra/descontos) MOTZ",
            pd.Series([0] * len(df_f))
        ).apply(_fmt_quebra),
        "Valor Acres. Quebra": "",
        "Tipo Parcela (A = Adiantamento / S = Saldo)": df_f["_tipo"],
        "Nr. Fatura": df_f.get("TITULO (NFe)", pd.Series(dtype=str)).apply(_fmt_int),
    })
    
    # Remover linhas sem CTRC ou sem Fatura (não dá pra baixar no ATUA)
    out = out[(out["Nr. CTRC"] != "") & (out["Nr. Fatura"] != "")]
    
    if out.empty:
        return None, 0
    
    # Exportar
    csv_str = out.to_csv(sep=";", index=False, lineterminator="\r\n")
    return csv_str.encode("latin-1", errors="replace"), len(out)


# ---- Botão na UI ----
st.divider()
st.markdown("##### 📥 Baixa de Títulos no ATUA")
st.caption(
    "Gera o CSV no formato exato da planilha modelo do ATUA pra subir e quitar os "
    "títulos. Inclui apenas linhas com **Data Transferência** e **Valor Transferido** "
    "preenchidos. Cada transferência (adto e saldo) vira uma linha separada."
)

csv_atua, qtd_linhas = _gerar_csv_baixa_atua(df)

if csv_atua:
    st.download_button(
        f"📥 XLS Baixa Título de Crédito ATUA  ·  {qtd_linhas} linhas",
        data=csv_atua,
        file_name=f"baixa_titulos_atua_{datetime.now().strftime('%Y-%m-%d_%H%M')}.csv",
        mime="text/csv",
        use_container_width=True,
        help="CSV com separador ; e decimal vírgula, encoding latin-1, formato idêntico ao modelo do ATUA.",
    )
else:
    st.warning(
        "Nenhuma linha disponível pra baixa — verifique se há transferências com "
        "Data Transferência E Valor Transferido preenchidos."
    )
