import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font
from io import BytesIO
from datetime import datetime

# =========================================================
# CONFIGURAÇÃO
# =========================================================

st.set_page_config(
    page_title="Sistema de Relatórios",
    layout="wide"
)

# =========================================================
# MENU LATERAL
# =========================================================

st.sidebar.title("📋 Relatórios")

relatorio = st.sidebar.radio(
    "Selecione o relatório",
    [
        "Peso Caminhão - Chegada",
        "Peso Caminhão - Saída",
        "Estufagem Individual"
    ]
)

# =========================================================
# FUNÇÃO DOWNLOAD
# =========================================================

def download_excel(wb, nome_arquivo):

    output = BytesIO()

    wb.save(output)

    output.seek(0)

    st.download_button(
        label="⬇️ Baixar Relatório",
        data=output,
        file_name=nome_arquivo,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# =========================================================
# RELATÓRIO CHEGADA
# =========================================================

if relatorio == "Peso Caminhão - Chegada":

    TEMPLATE_PATH = "Relatorios.xlsx"

    st.title("📋 Peso Caminhão - Chegada")

    # =====================================================
    # CAMPOS SUPERIORES
    # =====================================================

    col1, col2, col3 = st.columns(3)

    with col1:
        produtor = st.text_input("Produtor")
        terminal = st.text_input("Terminal")

    with col2:
        instrucao = st.text_input("Instrução")

        data_relatorio = st.date_input(
            "Data",
            datetime.today()
        )

    with col3:
        inspetor = st.text_input("Inspetor")

    observacao = st.text_area(
        "Observação",
        height=120
    )

    st.divider()

    # =====================================================
    # TABELA
    # =====================================================

    headers = [
        "DATA",
        "NOTA FISCAL",
        "LOTE",
        "QTD FARDOS",
        "PESO BRUTO",
        "PESO LÍQUIDO",
        "TARA",
        "PLACA",
        "PESO CAMINHÃO",
        "TARA CAMINHÃO",
        "PESO BRUTO CARGA",
        "PESO LÍQUIDO CARGA"
    ]

    if "linhas_chegada" not in st.session_state:

        st.session_state.linhas_chegada = [
            {
                "data": "",
                "nota": "",
                "lote": "",
                "fardos": "",
                "peso_bruto": "",
                "peso_liquido": "",
                "tara": "",
                "placa": "",
                "peso_caminhao": "",
                "tara_caminhao": "",
                "peso_bruto_carga": "",
                "peso_liquido_carga": ""
            }
        ]

    if st.button("➕ Adicionar Linha"):

        st.session_state.linhas_chegada.append(
            {
                "data": "",
                "nota": "",
                "lote": "",
                "fardos": "",
                "peso_bruto": "",
                "peso_liquido": "",
                "tara": "",
                "placa": "",
                "peso_caminhao": "",
                "tara_caminhao": "",
                "peso_bruto_carga": "",
                "peso_liquido_carga": ""
            }
        )

    cols = st.columns(len(headers))

    for col, h in zip(cols, headers):
        col.markdown(f"**{h}**")

    for i, linha in enumerate(st.session_state.linhas_chegada):

        cols = st.columns(len(headers))

        keys = list(linha.keys())

        for idx, key in enumerate(keys):

            linha[key] = cols[idx].text_input(
                "",
                value=linha[key],
                key=f"chegada_{key}_{i}"
            )

    # =====================================================
    # GERAR RELATÓRIO
    # =====================================================

    if st.button("📥 Gerar Relatório Chegada"):

        wb = load_workbook(TEMPLATE_PATH)

        ws = wb["Peso Caminhão - Chegada"]

        ws["B2"] = produtor
        ws["B3"] = terminal

        ws["F2"] = instrucao
        ws["F3"] = data_relatorio.strftime("%d/%m/%Y")

        ws["I2"] = inspetor

        ws["I2"].font = Font(
            name="Arial",
            size=12,
            bold=True,
            color="000000"
        )

        ws["I2"].alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

        ws["B26"] = observacao

        ws["B26"].alignment = Alignment(
            wrap_text=True,
            vertical="top"
        )

        start_row = 6

        for index, linha in enumerate(
            st.session_state.linhas_chegada
        ):

            row = start_row + index

            ws.cell(row=row, column=1).value = linha["data"]
            ws.cell(row=row, column=2).value = linha["nota"]
            ws.cell(row=row, column=3).value = linha["lote"]
            ws.cell(row=row, column=4).value = linha["fardos"]
            ws.cell(row=row, column=5).value = linha["peso_bruto"]
            ws.cell(row=row, column=6).value = linha["peso_liquido"]
            ws.cell(row=row, column=7).value = linha["tara"]
            ws.cell(row=row, column=8).value = linha["placa"]
            ws.cell(row=row, column=9).value = linha["peso_caminhao"]
            ws.cell(row=row, column=10).value = linha["tara_caminhao"]
            ws.cell(row=row, column=11).value = linha["peso_bruto_carga"]
            ws.cell(row=row, column=12).value = linha["peso_liquido_carga"]

        st.success("Relatório gerado com sucesso!")

        download_excel(
            wb,
            f"relatorio_chegada_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )

# =========================================================
# RELATÓRIO SAÍDA
# =========================================================

elif relatorio == "Peso Caminhão - Saída":

    TEMPLATE_PATH = "Saida.xlsx"

    st.title("📋 Peso Caminhão - Saída")

    col1, col2, col3 = st.columns(3)

    with col1:
        produtor = st.text_input("Produtor")
        terminal = st.text_input("Terminal")

    with col2:
        instrucao = st.text_input("Instrução")

        data_relatorio = st.date_input(
            "Data",
            datetime.today()
        )

    with col3:

        limite_peso_cntr = st.text_input(
            "Limite de Peso pro Cntr"
        )

        limite_total_instrucao = st.text_input(
            "Limite de Peso Total da Instrução"
        )

    observacao = st.text_area(
        "Observação",
        height=120
    )

    headers = [
        "CONTAINER",
        "NOTA FISCAL",
        "LOTE",
        "TOTAL FARDOS",
        "TARA CNTR",
        "MAX GROSS",
        "BRUTO CARGA",
        "HORÁRIO INÍCIO",
        "HORÁRIO FINAL",
        "PLACA"
    ]

    if "linhas_saida" not in st.session_state:

        st.session_state.linhas_saida = [
            {
                "container": "",
                "nota": "",
                "lote": "",
                "fardos": "",
                "tara_cntr": "",
                "max_gross": "",
                "bruto_carga": "",
                "hora_inicio": "",
                "hora_fim": "",
                "placa": ""
            }
        ]

    if st.button("➕ Adicionar Linha Saída"):

        st.session_state.linhas_saida.append(
            {
                "container": "",
                "nota": "",
                "lote": "",
                "fardos": "",
                "tara_cntr": "",
                "max_gross": "",
                "bruto_carga": "",
                "hora_inicio": "",
                "hora_fim": "",
                "placa": ""
            }
        )

    cols = st.columns(len(headers))

    for col, h in zip(cols, headers):
        col.markdown(f"**{h}**")

    for i, linha in enumerate(st.session_state.linhas_saida):

        cols = st.columns(len(headers))

        keys = list(linha.keys())

        for idx, key in enumerate(keys):

            linha[key] = cols[idx].text_input(
                "",
                value=linha[key],
                key=f"saida_{key}_{i}"
            )

    if st.button("📥 Gerar Relatório Saída"):

        wb = load_workbook(TEMPLATE_PATH)

        ws = wb["Peso Caminhão - Saída"]

        ws["B2"] = produtor
        ws["B3"] = terminal

        ws["F2"] = instrucao
        ws["F3"] = data_relatorio.strftime("%d/%m/%Y")

        ws["I2"] = limite_peso_cntr
        ws["I3"] = limite_total_instrucao

        ws["B26"] = observacao

        ws["B26"].alignment = Alignment(
            wrap_text=True,
            vertical="top"
        )

        start_row = 5

        for index, linha in enumerate(
            st.session_state.linhas_saida
        ):

            row = start_row + index

            ws.cell(row=row, column=1).value = linha["container"]
            ws.cell(row=row, column=2).value = linha["nota"]
            ws.cell(row=row, column=3).value = linha["lote"]
            ws.cell(row=row, column=4).value = linha["fardos"]
            ws.cell(row=row, column=5).value = linha["tara_cntr"]
            ws.cell(row=row, column=6).value = linha["max_gross"]
            ws.cell(row=row, column=7).value = linha["bruto_carga"]
            ws.cell(row=row, column=8).value = linha["hora_inicio"]
            ws.cell(row=row, column=9).value = linha["hora_fim"]
            ws.cell(row=row, column=10).value = linha["placa"]

        st.success("Relatório gerado com sucesso!")

        download_excel(
            wb,
            f"relatorio_saida_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )

# =========================================================
# ESTUFAGEM
# =========================================================

elif relatorio == "Estufagem Individual":

    TEMPLATE_PATH = "Estufagem.xlsx"

    st.title("📋 Estufagem Individual")

    instrucao = st.text_input(
        "Instrução de Embarque"
    )

    produtor = st.text_input("Produtor")

    col1, col2 = st.columns(2)

    with col1:

        container = st.text_input("Nº Container")

        tara_porta = st.text_input(
            "Tara Porta Cntr"
        )

        terminal = st.text_input("Terminal")

        inicio = st.text_input(
            "Começo Estufagem"
        )

        termino = st.text_input(
            "Término Estufagem"
        )

    with col2:

        qtd_fardos = st.text_input(
            "Qtd Fardos"
        )

        max_gross = st.text_input(
            "Max Gross"
        )

        lacre = st.text_input("Lacre")

        data_hora = st.text_input(
            "Data/Hora Início"
        )

        data_hora_termino = st.text_input(
            "Data/Hora Término"
        )

    observacao = st.text_area(
        "Observação",
        height=180
    )

    if st.button(
        "📥 Gerar Relatório Estufagem"
    ):

        wb = load_workbook(TEMPLATE_PATH)

        ws = wb["Estufagem individual"]

        ws["B2"] = instrucao
        ws["B3"] = produtor

        ws["B4"] = container
        ws["D4"] = qtd_fardos

        ws["B5"] = tara_porta
        ws["D5"] = max_gross

        ws["B6"] = terminal
        ws["D6"] = lacre

        ws["B7"] = inicio
        ws["D7"] = data_hora

        ws["B8"] = termino
        ws["D8"] = data_hora_termino

        ws["B26"] = observacao

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
