import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font
from io import BytesIO
from datetime import datetime

# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURAÇÃO
# ─────────────────────────────────────────────────────────────────────────────

st.set_page_config(page_title="Sistema de Relatórios", layout="wide")

# ─────────────────────────────────────────────────────────────────────────────
# CONSTANTES
# ─────────────────────────────────────────────────────────────────────────────

RELATORIOS = [
    "Peso Caminhão - Chegada",
    "Peso Caminhão - Saída",
    "Estufagem Individual",
]

LINHA_CHEGADA = {
    "data": "", "nota": "", "lote": "", "fardos": "",
    "peso_bruto": "", "peso_liquido": "", "tara": "", "placa": "",
    "peso_caminhao": "", "tara_caminhao": "", "peso_bruto_carga": "", "peso_liquido_carga": "",
}

LINHA_SAIDA = {
    "container": "", "nota": "", "lote": "", "fardos": "",
    "tara_cntr": "", "max_gross": "", "bruto_carga": "",
    "hora_inicio": "", "hora_fim": "", "placa": "",
}

LINHA_ESTUFAGEM = {
    "nota_fiscal": "", "lote": "", "qtd_fardos": "", "peso": "", "obs": "",
}

HEADERS_CHEGADA = [
    "DATA", "NOTA FISCAL", "LOTE", "QTD FARDOS", "PESO BRUTO", "PESO LÍQUIDO",
    "TARA", "PLACA", "PESO CAMINHÃO", "TARA CAMINHÃO", "PESO BRUTO CARGA", "PESO LÍQUIDO CARGA",
]

HEADERS_SAIDA = [
    "CONTAINER", "NOTA FISCAL", "LOTE", "TOTAL FARDOS", "TARA CNTR",
    "MAX GROSS", "BRUTO CARGA", "HORÁRIO INÍCIO", "HORÁRIO FINAL", "PLACA",
]

HEADERS_ESTUFAGEM = ["NOTA FISCAL", "LOTE", "QTD FARDOS", "PESO", "OBS."]

# ─────────────────────────────────────────────────────────────────────────────
# UTILITÁRIOS
# ─────────────────────────────────────────────────────────────────────────────

def init_session(*keys_defaults: tuple):
    """Inicializa chaves no session_state caso ainda não existam."""
    for key, default in keys_defaults:
        if key not in st.session_state:
            st.session_state[key] = default


def download_excel(wb: object, nome_arquivo: str):
    """Salva o workbook em memória e exibe botão de download."""
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    st.download_button(
        label="⬇️ Baixar Relatório",
        data=output,
        file_name=nome_arquivo,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


def render_tabela(prefix: str, linhas: list, headers: list):
    """Renderiza cabeçalhos e campos de entrada para uma tabela dinâmica."""
    cols = st.columns(len(headers))
    for col, h in zip(cols, headers):
        col.markdown(f"**{h}**")

    for i, linha in enumerate(linhas):
        cols = st.columns(len(headers))
        for idx, key in enumerate(linha):
            linha[key] = cols[idx].text_input(
                label="",
                value=linha[key],
                key=f"{prefix}_{key}_{i}",
                label_visibility="collapsed",
            )


def botao_adicionar(label: str, state_key: str, modelo: dict):
    """Botão genérico para adicionar uma nova linha à tabela."""
    if st.button(f"➕ {label}"):
        st.session_state[state_key].append(dict(modelo))
        st.rerun()


def preencher_linhas(ws, linhas: list, start_row: int, colunas: list):
    """Escreve as linhas de dados nas células do worksheet a partir de start_row."""
    for i, linha in enumerate(linhas):
        row = start_row + i
        for col_idx, key in enumerate(colunas, start=1):
            ws.cell(row=row, column=col_idx).value = linha.get(key, "")


# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────────────────────────────

st.sidebar.title("📋 Relatórios")
relatorio = st.sidebar.radio("Selecione o relatório", RELATORIOS)

# ─────────────────────────────────────────────────────────────────────────────
# RELATÓRIO: CHEGADA
# ─────────────────────────────────────────────────────────────────────────────

if relatorio == "Peso Caminhão - Chegada":

    init_session(("linhas_chegada", [dict(LINHA_CHEGADA)]))

    st.title("📋 Peso Caminhão — Chegada")

    col1, col2, col3 = st.columns(3)
    with col1:
        produtor = st.text_input("Produtor", key="chegada_produtor")
        terminal = st.text_input("Terminal", key="chegada_terminal")
    with col2:
        instrucao      = st.text_input("Instrução", key="chegada_instrucao")
        data_relatorio = st.date_input("Data", datetime.today(), key="chegada_data")
    with col3:
        inspetor = st.text_input("Inspetor", key="chegada_inspetor")

    observacao = st.text_area("Observação", height=120, key="chegada_observacao")
    st.divider()

    botao_adicionar("Adicionar Linha", "linhas_chegada", LINHA_CHEGADA)
    render_tabela("chegada", st.session_state.linhas_chegada, HEADERS_CHEGADA)

    if st.button("📥 Gerar Relatório Chegada"):
        wb = load_workbook("Relatorios.xlsx")
        ws = wb["Peso Caminhão - Chegada"]

        ws["B2"] = produtor
        ws["B3"] = terminal
        ws["F2"] = instrucao
        ws["F3"] = data_relatorio.strftime("%d/%m/%Y")
        ws["I2"] = inspetor
        ws["I2"].font      = Font(name="Arial", size=12, bold=True)
        ws["I2"].alignment = Alignment(horizontal="center", vertical="center")
        ws["B26"]           = observacao
        ws["B26"].alignment = Alignment(wrap_text=True, vertical="top")

        preencher_linhas(
            ws, st.session_state.linhas_chegada, start_row=6,
            colunas=list(LINHA_CHEGADA.keys()),
        )

        st.success("Relatório gerado com sucesso!")
        download_excel(wb, f"relatorio_chegada_{datetime.now():%Y%m%d_%H%M%S}.xlsx")

# ─────────────────────────────────────────────────────────────────────────────
# RELATÓRIO: SAÍDA
# ─────────────────────────────────────────────────────────────────────────────

elif relatorio == "Peso Caminhão - Saída":

    init_session(("linhas_saida", [dict(LINHA_SAIDA)]))

    st.title("📋 Peso Caminhão — Saída")

    col1, col2, col3 = st.columns(3)
    with col1:
        produtor = st.text_input("Produtor", key="saida_produtor")
        terminal = st.text_input("Terminal", key="saida_terminal")
    with col2:
        instrucao      = st.text_input("Instrução", key="saida_instrucao")
        data_relatorio = st.date_input("Data", datetime.today(), key="saida_data")
    with col3:
        limite_peso_cntr        = st.text_input("Limite de Peso por Cntr", key="saida_lim_cntr")
        limite_total_instrucao  = st.text_input("Limite de Peso Total da Instrução", key="saida_lim_total")

    observacao = st.text_area("Observação", height=120, key="saida_observacao")
    st.divider()

    botao_adicionar("Adicionar Linha Saída", "linhas_saida", LINHA_SAIDA)
    render_tabela("saida", st.session_state.linhas_saida, HEADERS_SAIDA)

    if st.button("📥 Gerar Relatório Saída"):
        wb = load_workbook("Saida.xlsx")
        ws = wb["Peso Caminhão - Saída"]

        ws["B2"] = produtor
        ws["B3"] = terminal
        ws["F2"] = instrucao
        ws["F3"] = data_relatorio.strftime("%d/%m/%Y")
        ws["I2"] = limite_peso_cntr
        ws["I3"] = limite_total_instrucao
        ws["B26"]           = observacao
        ws["B26"].alignment = Alignment(wrap_text=True, vertical="top")

        preencher_linhas(
            ws, st.session_state.linhas_saida, start_row=5,
            colunas=list(LINHA_SAIDA.keys()),
        )

        st.success("Relatório gerado com sucesso!")
        download_excel(wb, f"relatorio_saida_{datetime.now():%Y%m%d_%H%M%S}.xlsx")

# ─────────────────────────────────────────────────────────────────────────────
# RELATÓRIO: ESTUFAGEM
# ─────────────────────────────────────────────────────────────────────────────

elif relatorio == "Estufagem Individual":

    init_session(("linhas_estufagem", [dict(LINHA_ESTUFAGEM)]))

    st.title("📋 Estufagem Individual")

    col1, col2 = st.columns(2)
    with col1:
        instrucao = st.text_input("Instrução de Embarque", key="est_instrucao")
        produtor  = st.text_input("Produtor", key="est_produtor")
        container = st.text_input("Nº Container", key="est_container")
        terminal  = st.text_input("Terminal", key="est_terminal")
    with col2:
        tara_porta = st.text_input("Tara Porta Cntr", key="est_tara_porta")
        max_gross  = st.text_input("Max Gross", key="est_max_gross")
        lacre      = st.text_input("Lacre", key="est_lacre")

    st.markdown("**⏱️ Datas e Horários**")
    col3, col4, col5, col6 = st.columns(4)
    with col3:
        inicio = st.text_input("Começo da Estufagem", key="est_inicio")
    with col4:
        data_hora_inicio = st.text_input("Data/Hora Início", key="est_dh_inicio")
    with col5:
        termino = st.text_input("Término da Estufagem", key="est_termino")
    with col6:
        data_hora_termino = st.text_input("Data/Hora Término", key="est_dh_termino")

    st.divider()
    st.subheader("📦 Dados da Estufagem")

    botao_adicionar("Adicionar Linha Estufagem", "linhas_estufagem", LINHA_ESTUFAGEM)
    render_tabela("estufagem", st.session_state.linhas_estufagem, HEADERS_ESTUFAGEM)

    inspetor   = st.text_input("Inspetor", key="est_inspetor")
    observacao = st.text_area("Observação Final", height=180, key="est_observacao")

    if st.button("📥 Gerar Relatório Estufagem"):
        wb = load_workbook("Estufagem.xlsx")
        ws = wb["Estufagem individual"]

        mapeamento = {
            "B2": instrucao,   "B3": produtor,
            "B4": container,   "D4": tara_porta,
            "B5": max_gross,   "D5": lacre,
            "B6": terminal,    "D6": data_hora_inicio,
            "B7": inicio,      "D7": data_hora_inicio,
            "B8": termino,     "D8": data_hora_termino,
            "A29": inspetor,
        }
        for cell, value in mapeamento.items():
            ws[cell] = value

        ws["B26"]           = observacao
        ws["B26"].alignment = Alignment(wrap_text=True, vertical="top")

        preencher_linhas(
            ws, st.session_state.linhas_estufagem, start_row=11,
            colunas=list(LINHA_ESTUFAGEM.keys()),
        )

        st.success("Relatório gerado com sucesso!")
        download_excel(wb, f"relatorio_estufagem_{datetime.now():%Y%m%d_%H%M%S}.xlsx")
