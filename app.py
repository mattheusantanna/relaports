import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from io import BytesIO
from datetime import datetime

# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURAÇÃO
# ─────────────────────────────────────────────────────────────────────────────

st.set_page_config(page_title="Sistema de Relatórios", layout="wide")

# ─────────────────────────────────────────────────────────────────────────────
# ESTILOS EXCEL
# ─────────────────────────────────────────────────────────────────────────────

HEADER_FILL  = PatternFill("solid", start_color="1F4E79")
SUBHDR_FILL  = PatternFill("solid", start_color="2E75B6")
COL_HDR_FILL = PatternFill("solid", start_color="BDD7EE")
WHITE_FILL   = PatternFill("solid", start_color="FFFFFF")
ALT_FILL     = PatternFill("solid", start_color="DEEAF1")
NOTE_FILL    = PatternFill("solid", start_color="FFF2CC")

_thin = Side(style="thin", color="000000")
BORDER = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)


def _scell(ws, row, col, value="", bold=False, fill=None,
           h_align="center", font_size=9, wrap=True):
    c = ws.cell(row=row, column=col, value=value)
    on_dark = fill in (HEADER_FILL, SUBHDR_FILL)
    c.font = Font(name="Arial", size=font_size, bold=bold,
                  color="FFFFFF" if on_dark else "000000")
    c.alignment = Alignment(horizontal=h_align, vertical="center", wrap_text=wrap)
    if fill:
        c.fill = fill
    c.border = BORDER
    return c


def _merge(ws, r1, c1, r2, c2, value="", bold=False, fill=None,
           h_align="center", font_size=9):
    ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
    c = ws.cell(row=r1, column=c1, value=value)
    on_dark = fill in (HEADER_FILL, SUBHDR_FILL)
    c.font = Font(name="Arial", size=font_size, bold=bold,
                  color="FFFFFF" if on_dark else "000000")
    c.alignment = Alignment(horizontal=h_align, vertical="center", wrap_text=True)
    if fill:
        c.fill = fill
        for r in range(r1, r2 + 1):
            for cc in range(c1, c2 + 1):
                ws.cell(row=r, column=cc).fill = fill
    c.border = BORDER
    return c


def _col_widths(ws, widths):
    for col, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(col)].width = w


def _data_rows(ws, start, end, ncols):
    for i, r in enumerate(range(start, end + 1)):
        fill = ALT_FILL if i % 2 else WHITE_FILL
        for c in range(1, ncols + 1):
            _scell(ws, r, c, fill=fill)
        ws.row_dimensions[r].height = 15


def _wb_to_bytes(wb) -> BytesIO:
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

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
# UTILITÁRIOS UI
# ─────────────────────────────────────────────────────────────────────────────

def init_session(*keys_defaults: tuple):
    for key, default in keys_defaults:
        if key not in st.session_state:
            st.session_state[key] = default


def download_excel(buf: BytesIO, nome_arquivo: str):
    st.download_button(
        label="⬇️ Baixar Relatório",
        data=buf,
        file_name=nome_arquivo,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


def render_tabela(prefix: str, linhas: list, headers: list):
    cols = st.columns(len(headers))
    for col, h in zip(cols, headers):
        col.markdown(f"**{h}**")
    for i, linha in enumerate(linhas):
        cols = st.columns(len(headers))
        for idx, key in enumerate(linha):
            linha[key] = cols[idx].text_input(
                label="", value=linha[key],
                key=f"{prefix}_{key}_{i}",
                label_visibility="collapsed",
            )


def botao_adicionar(label: str, state_key: str, modelo: dict):
    if st.button(f"➕ {label}"):
        st.session_state[state_key].append(dict(modelo))
        st.rerun()


def preencher_linhas(ws, linhas: list, start_row: int, colunas: list):
    for i, linha in enumerate(linhas):
        row = start_row + i
        for col_idx, key in enumerate(colunas, start=1):
            ws.cell(row=row, column=col_idx).value = linha.get(key, "")

# ─────────────────────────────────────────────────────────────────────────────
# GERADOR EXCEL — CHEGADA
# ─────────────────────────────────────────────────────────────────────────────

def gerar_excel_chegada(produtor, terminal, instrucao, data, inspetor,
                        linhas, observacao) -> BytesIO:
    wb = Workbook()
    ws = wb.active
    ws.title = "Peso Caminhão - Chegada"

    _merge(ws, 1, 1, 1, 12,
           value="RELATÓRIO DE PESO DE CAMINHÃO / PESO DE CHEGADA",
           fill=HEADER_FILL, bold=True, font_size=12)
    ws.row_dimensions[1].height = 24

    # linha de meta-dados
    meta = [("Produtor:", produtor), ("Terminal:", terminal),
            ("Instrução:", instrucao), ("Data:", data), ("Inspetor:", inspetor)]
    col = 1
    for lbl, val in meta:
        _scell(ws, 2, col,     lbl, bold=True, fill=SUBHDR_FILL, h_align="left", font_size=8)
        _scell(ws, 2, col + 1, val, fill=WHITE_FILL, h_align="left", font_size=9)
        col += 2
    # preenche col restante se houver
    for c in range(col, 13):
        _scell(ws, 2, c, fill=WHITE_FILL)
    ws.row_dimensions[2].height = 16

    # cabeçalhos da tabela (linhas 3–4)
    for r in range(3, 5):
        for c in range(1, 13):
            _scell(ws, r, c, fill=COL_HDR_FILL)

    _merge(ws, 3, 1,  4, 1,  "DATA",             fill=COL_HDR_FILL, bold=True, font_size=8)
    _merge(ws, 3, 2,  4, 2,  "NOTA FISCAL",       fill=COL_HDR_FILL, bold=True, font_size=8)
    _merge(ws, 3, 3,  4, 3,  "LOTE",              fill=COL_HDR_FILL, bold=True, font_size=8)
    _merge(ws, 3, 4,  4, 4,  "QTD\nFARDOS",       fill=COL_HDR_FILL, bold=True, font_size=8)
    _merge(ws, 3, 5,  3, 6,  "DADOS DA NOTA FISCAL", fill=COL_HDR_FILL, bold=True, font_size=8)
    _scell(ws, 4, 5,  "PESO BRUTO",   bold=True, fill=COL_HDR_FILL, font_size=8)
    _scell(ws, 4, 6,  "PESO LÍQUIDO", bold=True, fill=COL_HDR_FILL, font_size=8)
    _merge(ws, 3, 7,  4, 7,  "TARA TOTAL\nDO FARDO",  fill=COL_HDR_FILL, bold=True, font_size=8)
    _merge(ws, 3, 8,  4, 8,  "PLACA DO\nCAMINHÃO",    fill=COL_HDR_FILL, bold=True, font_size=8)
    _merge(ws, 3, 9,  3, 12, "DADOS DO CAMINHÃO / PESO DE CHEGADA",
           fill=COL_HDR_FILL, bold=True, font_size=8)
    _scell(ws, 4, 9,  "PESO BRUTO\n(CAMINHÃO CHEIO)", bold=True, fill=COL_HDR_FILL, font_size=8)
    _scell(ws, 4, 10, "TARA DO\nCAMINHÃO",            bold=True, fill=COL_HDR_FILL, font_size=8)
    _scell(ws, 4, 11, "PESO BRUTO\nDA CARGA",          bold=True, fill=COL_HDR_FILL, font_size=8)
    _scell(ws, 4, 12, "PESO LÍQUIDO\nDA CARGA",        bold=True, fill=COL_HDR_FILL, font_size=8)
    ws.row_dimensions[3].height = 18
    ws.row_dimensions[4].height = 28

    # linhas de dados
    _data_rows(ws, 5, 5 + max(len(linhas) - 1, 19), 12)
    preencher_linhas(ws, linhas, start_row=5, colunas=list(LINHA_CHEGADA.keys()))

    # observação
    obs_row = 5 + max(len(linhas), 20)
    _merge(ws, obs_row, 1, obs_row + 1, 12,
           value=f"OBS: {observacao}", fill=WHITE_FILL, h_align="left", font_size=9)
    ws.row_dimensions[obs_row].height = 30

    _col_widths(ws, [10, 14, 10, 9, 12, 12, 12, 14, 14, 12, 14, 14])
    return _wb_to_bytes(wb)

# ─────────────────────────────────────────────────────────────────────────────
# GERADOR EXCEL — SAÍDA
# ─────────────────────────────────────────────────────────────────────────────

def gerar_excel_saida(produtor, terminal, instrucao, data,
                      lim_cntr, lim_total, linhas, observacao) -> BytesIO:
    wb = Workbook()
    ws = wb.active
    ws.title = "Peso Caminhão - Saída"

    _merge(ws, 1, 1, 1, 10,
           value="RELATÓRIO DE PESAGEM DE CONTAINER / PESO SAÍDA",
           fill=HEADER_FILL, bold=True, font_size=12)
    ws.row_dimensions[1].height = 24

    meta = [("Produtor:", produtor), ("Terminal:", terminal),
            ("Instrução:", instrucao), ("Data:", data)]
    col = 1
    for lbl, val in meta:
        _scell(ws, 2, col,     lbl, bold=True, fill=SUBHDR_FILL, h_align="left", font_size=8)
        _scell(ws, 2, col + 1, val, fill=WHITE_FILL, h_align="left", font_size=9)
        col += 2
    for c in range(col, 11):
        _scell(ws, 2, c, fill=WHITE_FILL)
    ws.row_dimensions[2].height = 16

    _merge(ws, 3, 1, 3, 5,
           value=f"Limite de Peso Total da Instrução: {lim_total}",
           fill=SUBHDR_FILL, bold=True, h_align="left", font_size=8)
    _merge(ws, 3, 6, 3, 10,
           value=f"Limite de Peso por Cntr: {lim_cntr}",
           fill=SUBHDR_FILL, bold=True, h_align="left", font_size=8)
    ws.row_dimensions[3].height = 16

    hdrs = ["CONTAINER", "NOTA FISCAL", "LOTE", "TOTAL FARDOS", "TARA CNTR",
            "MAX GROSS", "BRUTO CARGA", "HORÁRIO INÍCIO", "HORÁRIO FINAL", "PLACA"]
    for c, h in enumerate(hdrs, 1):
        _scell(ws, 4, c, h, bold=True, fill=COL_HDR_FILL, font_size=8)
    ws.row_dimensions[4].height = 22

    _data_rows(ws, 5, 5 + max(len(linhas) - 1, 19), 10)
    preencher_linhas(ws, linhas, start_row=5, colunas=list(LINHA_SAIDA.keys()))

    obs_row = 5 + max(len(linhas), 20)
    _merge(ws, obs_row, 1, obs_row + 1, 10,
           value=f"OBS: {observacao}", fill=WHITE_FILL, h_align="left", font_size=9)
    ws.row_dimensions[obs_row].height = 30

    _col_widths(ws, [16, 14, 10, 12, 12, 12, 14, 14, 14, 12])
    return _wb_to_bytes(wb)

# ─────────────────────────────────────────────────────────────────────────────
# GERADOR EXCEL — ESTUFAGEM
# ─────────────────────────────────────────────────────────────────────────────

def gerar_excel_estufagem(instrucao, produtor, container, tara_porta,
                           max_gross, lacre, terminal, inicio, dh_inicio,
                           termino, dh_termino, inspetor, linhas, observacao) -> BytesIO:
    wb = Workbook()
    ws = wb.active
    ws.title = "Estufagem Individual"

    _merge(ws, 1, 1, 1, 5,
           value="RELATÓRIO DE ESTUFAGEM INDIVIDUAL CONTAINER",
           fill=HEADER_FILL, bold=True, font_size=12)
    ws.row_dimensions[1].height = 24

    def meta_row(r, lbl1, val1, lbl2="", val2=""):
        _scell(ws, r, 1, lbl1, bold=True, fill=SUBHDR_FILL, h_align="left", font_size=8)
        _scell(ws, r, 2, val1, fill=WHITE_FILL, h_align="left", font_size=9)
        if lbl2:
            _scell(ws, r, 3, lbl2, bold=True, fill=SUBHDR_FILL, h_align="left", font_size=8)
            _scell(ws, r, 4, val2, fill=WHITE_FILL, h_align="left", font_size=9)
            _scell(ws, r, 5, "",   fill=WHITE_FILL)
        else:
            for c in range(3, 6):
                _scell(ws, r, c, fill=WHITE_FILL)
        ws.row_dimensions[r].height = 16

    meta_row(2, "Instrução de Embarque:", instrucao)
    meta_row(3, "Produtor:",              produtor)
    meta_row(4, "Nº do Container:",       container,  "Tara da Porta do Cntr:", tara_porta)
    meta_row(5, "Max. Gross:",            max_gross,  "Lacre:",                 lacre)
    meta_row(6, "Terminal:",              terminal)
    meta_row(7, "Começo da Estufagem:",   inicio,     "Data/Hora Início:",      dh_inicio)
    meta_row(8, "Término da Estufagem:",  termino,    "Data/Hora Término:",     dh_termino)

    hdrs = ["NOTA FISCAL", "LOTE", "QTD FARDOS", "PESO", "OBS."]
    for c, h in enumerate(hdrs, 1):
        _scell(ws, 9, c, h, bold=True, fill=COL_HDR_FILL, font_size=8)
    ws.row_dimensions[9].height = 18

    _data_rows(ws, 10, 10 + max(len(linhas) - 1, 14), 5)
    preencher_linhas(ws, linhas, start_row=10, colunas=list(LINHA_ESTUFAGEM.keys()))

    obs_row = 10 + max(len(linhas), 15)
    _merge(ws, obs_row, 1, obs_row + 2, 5,
           value=f"Observação: {observacao}", fill=WHITE_FILL, h_align="left", font_size=9)
    for r in range(obs_row, obs_row + 3):
        ws.row_dimensions[r].height = 14

    ins_row = obs_row + 3
    _merge(ws, ins_row, 1, ins_row, 5,
           value=f"INSPETOR: {inspetor}", fill=WHITE_FILL, h_align="center",
           font_size=9, bold=True)
    ws.row_dimensions[ins_row].height = 18

    note_row = ins_row + 2
    _merge(ws, note_row, 1, note_row + 4, 5,
           value=("ATENÇÃO, INSPETORES\n\n"
                  "• Coloquem a data e o horário de INÍCIO E TÉRMINO da estufagem.\n"
                  "• Assinem seus nomes.\n"
                  "• Anexar etiquetas nos romaneios e marcar os saldos corretamente."),
           fill=NOTE_FILL, h_align="left", font_size=9)
    for r in range(note_row, note_row + 5):
        ws.row_dimensions[r].height = 14

    _col_widths(ws, [20, 14, 12, 14, 26])
    return _wb_to_bytes(wb)

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
        produtor = st.text_input("Produtor",   key="chegada_produtor")
        terminal = st.text_input("Terminal",   key="chegada_terminal")
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
        buf = gerar_excel_chegada(
            produtor, terminal, instrucao,
            data_relatorio.strftime("%d/%m/%Y"),
            inspetor, st.session_state.linhas_chegada, observacao,
        )
        st.success("Relatório gerado com sucesso!")
        download_excel(buf, f"relatorio_chegada_{datetime.now():%Y%m%d_%H%M%S}.xlsx")

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
        lim_cntr  = st.text_input("Limite de Peso por Cntr",           key="saida_lim_cntr")
        lim_total = st.text_input("Limite de Peso Total da Instrução",  key="saida_lim_total")

    observacao = st.text_area("Observação", height=120, key="saida_observacao")
    st.divider()

    botao_adicionar("Adicionar Linha Saída", "linhas_saida", LINHA_SAIDA)
    render_tabela("saida", st.session_state.linhas_saida, HEADERS_SAIDA)

    if st.button("📥 Gerar Relatório Saída"):
        buf = gerar_excel_saida(
            produtor, terminal, instrucao,
            data_relatorio.strftime("%d/%m/%Y"),
            lim_cntr, lim_total,
            st.session_state.linhas_saida, observacao,
        )
        st.success("Relatório gerado com sucesso!")
        download_excel(buf, f"relatorio_saida_{datetime.now():%Y%m%d_%H%M%S}.xlsx")

# ─────────────────────────────────────────────────────────────────────────────
# RELATÓRIO: ESTUFAGEM
# ─────────────────────────────────────────────────────────────────────────────

elif relatorio == "Estufagem Individual":

    init_session(("linhas_estufagem", [dict(LINHA_ESTUFAGEM)]))

    st.title("📋 Estufagem Individual")

    col1, col2 = st.columns(2)
    with col1:
        instrucao = st.text_input("Instrução de Embarque", key="est_instrucao")
        produtor  = st.text_input("Produtor",              key="est_produtor")
        container = st.text_input("Nº Container",          key="est_container")
        terminal  = st.text_input("Terminal",              key="est_terminal")
    with col2:
        tara_porta = st.text_input("Tara Porta Cntr", key="est_tara_porta")
        max_gross  = st.text_input("Max Gross",        key="est_max_gross")
        lacre      = st.text_input("Lacre",            key="est_lacre")

    st.markdown("**⏱️ Datas e Horários**")
    col3, col4, col5, col6 = st.columns(4)
    with col3:
        inicio = st.text_input("Começo da Estufagem", key="est_inicio")
    with col4:
        dh_inicio = st.text_input("Data/Hora Início",  key="est_dh_inicio")
    with col5:
        termino = st.text_input("Término da Estufagem", key="est_termino")
    with col6:
        dh_termino = st.text_input("Data/Hora Término", key="est_dh_termino")

    st.divider()
    st.subheader("📦 Dados da Estufagem")

    botao_adicionar("Adicionar Linha Estufagem", "linhas_estufagem", LINHA_ESTUFAGEM)
    render_tabela("estufagem", st.session_state.linhas_estufagem, HEADERS_ESTUFAGEM)

    inspetor   = st.text_input("Inspetor",        key="est_inspetor")
    observacao = st.text_area("Observação Final", key="est_observacao", height=180)

    if st.button("📥 Gerar Relatório Estufagem"):
        buf = gerar_excel_estufagem(
            instrucao, produtor, container, tara_porta,
            max_gross, lacre, terminal, inicio, dh_inicio,
            termino, dh_termino, inspetor,
            st.session_state.linhas_estufagem, observacao,
        )
        st.success("Relatório gerado com sucesso!")
        download_excel(buf, f"relatorio_estufagem_{datetime.now():%Y%m%d_%H%M%S}.xlsx")
