# =========================================================
# ESTUFAGEM
# =========================================================

elif relatorio == "Estufagem Individual":

    TEMPLATE_PATH = "Estufagem.xlsx"

    st.title("📋 Estufagem Individual")

    # =====================================================
    # CAMPOS SUPERIORES
    # =====================================================

    col1, col2, col3 = st.columns(3)

    with col1:

        instrucao = st.text_input(
            "Instrução de Embarque"
        )

        produtor = st.text_input(
            "Produtor"
        )

        container = st.text_input(
            "Nº Container"
        )

        nota_fiscal = st.text_input(
            "Nota Fiscal"
        )

    with col2:

        lote = st.text_input(
            "Lote"
        )

        qtd_fardos = st.text_input(
            "Qtd Fardos"
        )

        qtd_fardos_peso = st.text_input(
            "Qtd Fardos Peso"
        )

        tara_porta = st.text_input(
            "Tara Porta Cntr"
        )

    with col3:

        max_gross = st.text_input(
            "Max Gross"
        )

        lacre = st.text_input(
            "Lacre"
        )

        terminal = st.text_input(
            "Terminal"
        )

        inicio = st.text_input(
            "Começo Estufagem"
        )

    # =====================================================
    # CAMPOS INFERIORES
    # =====================================================

    col4, col5 = st.columns(2)

    with col4:

        termino = st.text_input(
            "Término Estufagem"
        )

    with col5:

        data_hora = st.text_input(
            "Data/Hora Início"
        )

        data_hora_termino = st.text_input(
            "Data/Hora Término"
        )

    # =====================================================
    # OBSERVAÇÃO
    # =====================================================

    observacao = st.text_area(
        "OBS.",
        height=180
    )

    # =====================================================
    # GERAR RELATÓRIO
    # =====================================================

    if st.button(
        "📥 Gerar Relatório Estufagem"
    ):

        wb = load_workbook(TEMPLATE_PATH)

        ws = wb["Estufagem individual"]

        # =================================================
        # PREENCHIMENTO
        # =================================================

        ws["B2"] = instrucao
        ws["B3"] = produtor

        ws["B4"] = container
        ws["D4"] = qtd_fardos

        ws["B5"] = nota_fiscal
        ws["D5"] = qtd_fardos_peso

        ws["B6"] = lote
        ws["D6"] = max_gross

        ws["B7"] = tara_porta
        ws["D7"] = lacre

        ws["B8"] = terminal
        ws["D8"] = data_hora

        ws["B9"] = inicio
        ws["D9"] = data_hora_termino

        ws["B10"] = termino

        ws["B26"] = observacao

        # =================================================
        # FORMATAÇÃO
        # =================================================

        ws["B26"].alignment = Alignment(
            wrap_text=True,
            vertical="top"
        )

        st.success(
            "Relatório gerado com sucesso!"
        )

        download_excel(
            wb,
            f"relatorio_estufagem_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
