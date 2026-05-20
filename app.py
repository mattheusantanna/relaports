# app.py
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Side, Font
from io import BytesIO
from datetime import datetime

st.set_page_config(
    page_title="Relatório de Peso de Caminhão",
    layout="wide"
)

TEMPLATE_PATH = "Relatorios.xlsx"

st.title("📋 Relatório de Peso de Caminhão / Peso de Chegada")

# =========================
# DADOS SUPERIORES
# =========================

col1, col2, col3 = st.columns(3)

with col1:
    produtor = st.text_input("Produtor")
    terminal = st.text_input("Terminal")

with col2:
    instrucao = st.text_input("Instrução")
    data_relatorio = st.date_input("Data", datetime.today())

with col3:
    inspetor = st.text_input("Inspetor")
    placa = st.text_input("Placa do Caminhão")

st.divider()

# =========================
# TABELA DE DADOS
# =========================

st.subheader("Dados pertencentes à nota fiscal")

if "linhas" not in st.session_state:
    st.session_state.linhas = [
        {
            "data": "",
            "nota": "",
            "lote": "",
            "fardos": "",
            "peso_bruto": "",
            "peso_liquido": "",
            "tara": "",
            "peso_caminhao": ""
        }
    ]

# Adicionar linha
if st.button("➕ Adicionar Linha"):
    st.session_state.linhas.append(
        {
            "data": "",
            "nota": "",
            "lote": "",
            "fardos": "",
            "peso_bruto": "",
            "peso_liquido": "",
            "tara": "",
            "peso_caminhao": ""
        }
    )

# Cabeçalho visual
cols = st.columns([1,1,1,1,1,1,1,1])
headers = [
    "DATA",
    "NOTA FISCAL",
    "LOTE",
    "QTD FARDOS",
    "PESO BRUTO",
    "PESO LÍQUIDO",
    "TARA",
    "PESO CAMINHÃO"
]

for col, h in zip(cols, headers):
    col.markdown(f"**{h}**")

# Linhas editáveis
for i, linha in enumerate(st.session_state.linhas):

    cols = st.columns([1,1,1,1,1,1,1,1])

    linha["data"] = cols[0].text_input(
        "",
        value=linha["data"],
        key=f"data_{i}"
    )

    linha["nota"] = cols[1].text_input(
        "",
        value=linha["nota"],
        key=f"nota_{i}"
    )

    linha["lote"] = cols[2].text_input(
        "",
        value=linha["lote"],
        key=f"lote_{i}"
    )

    linha["fardos"] = cols[3].text_input(
        "",
        value=linha["fardos"],
        key=f"fardos_{i}"
    )

    linha["peso_bruto"] = cols[4].text_input(
        "",
        value=linha["peso_bruto"],
        key=f"peso_bruto_{i}"
    )

    linha["peso_liquido"] = cols[5].text_input(
        "",
        value=linha["peso_liquido"],
        key=f"peso_liquido_{i}"
    )

    linha["tara"] = cols[6].text_input(
        "",
        value=linha["tara"],
        key=f"tara_{i}"
    )

    linha["peso_caminhao"] = cols[7].text_input(
        "",
        value=linha["peso_caminhao"],
        key=f"peso_caminhao_{i}"
    )

st.divider()

# =========================
# GERAR EXCEL
# =========================

def gerar_excel():

    wb = load_workbook(TEMPLATE_PATH)
    ws = wb["Peso Caminhão - Chegada"]

    # =========================
    # CAMPOS SUPERIORES
    # =========================

    ws["B2"] = produtor
    ws["F2"] = instrucao
    ws["I2"] = inspetor

    ws["B3"] = terminal
    ws["F3"] = data_relatorio.strftime("%d/%m/%Y")
    ws["H5"] = placa

    # =========================
    # DADOS TABELA
    # =========================

    start_row = 6

    thin = Side(border_style="thin", color="000000")

    for index, linha in enumerate(st.session_state.linhas):

        row = start_row + index

        ws.cell(row=row, column=1).value = linha["data"]
        ws.cell(row=row, column=2).value = linha["nota"]
        ws.cell(row=row, column=3).value = linha["lote"]
        ws.cell(row=row, column=4).value = linha["fardos"]
        ws.cell(row=row, column=5).value = linha["peso_bruto"]
        ws.cell(row=row, column=6).value = linha["peso_liquido"]
        ws.cell(row=row, column=7).value = linha["tara"]
        ws.cell(row=row, column=9).value = linha["peso_caminhao"]

        # Estilização
        for col in range(1, 13):
            cell = ws.cell(row=row, column=col)

            cell.border = Border(
                left=thin,
                right=thin,
                top=thin,
                bottom=thin
            )

            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True
            )

            cell.font = Font(name="Arial", size=10)

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    return output

# =========================
# DOWNLOAD
# =========================

if st.button("📥 Gerar Relatório"):

    arquivo_excel = gerar_excel()

    st.success("Relatório gerado com sucesso!")

    st.download_button(
        label="⬇️ Baixar Excel",
        data=arquivo_excel,
        file_name=f"relatorio_peso_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
