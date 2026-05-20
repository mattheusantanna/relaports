import streamlit as st
import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ─────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Relatórios de Pesagem",
    page_icon="⚖️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────
# GLOBAL CSS
# ─────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Sora:wght@300;400;600;700&family=JetBrains+Mono:wght@400;500&display=swap');

html, body, [class*="css"] { font-family: 'Sora', sans-serif; }

/* Sidebar */
[data-testid="stSidebar"] {
    background: linear-gradient(160deg, #0d1b2a 0%, #1b2e45 100%);
    border-right: 1px solid rgba(255,255,255,0.07);
}
[data-testid="stSidebar"] * { color: #e0eaf5 !important; }
[data-testid="stSidebar"] .stRadio label { font-size: 0.88rem; }

/* Main background */
[data-testid="stAppViewContainer"] > .main {
    background: #f0f4f9;
}

/* Cards */
.card {
    background: white;
    border-radius: 14px;
    padding: 1.4rem 1.6rem;
    margin-bottom: 1.2rem;
    box-shadow: 0 2px 12px rgba(0,0,0,0.06);
    border: 1px solid #e4ecf4;
}
.card-title {
    font-size: 0.72rem;
    font-weight: 700;
    letter-spacing: 0.12em;
    text-transform: uppercase;
    color: #6b8cae;
    margin-bottom: 0.9rem;
    padding-bottom: 0.5rem;
    border-bottom: 2px solid #e8f0fa;
}
.section-header {
    background: linear-gradient(90deg, #1b3c6e 0%, #2563a8 100%);
    color: white !important;
    padding: 0.65rem 1.2rem;
    border-radius: 8px;
    font-size: 0.78rem;
    font-weight: 700;
    letter-spacing: 0.1em;
    text-transform: uppercase;
    margin: 1.2rem 0 0.8rem 0;
}
.hero-title {
    font-size: 1.9rem;
    font-weight: 700;
    color: #0d1b2a;
    line-height: 1.2;
}
.hero-sub {
    font-size: 0.88rem;
    color: #6b8cae;
    margin-top: 0.3rem;
    font-weight: 400;
}
.badge {
    display: inline-block;
    background: #e8f0fa;
    color: #2563a8;
    font-size: 0.7rem;
    font-weight: 600;
    padding: 0.18rem 0.7rem;
    border-radius: 20px;
    letter-spacing: 0.05em;
    margin-left: 0.5rem;
    vertical-align: middle;
}
.stDownloadButton > button {
    background: linear-gradient(135deg, #1b3c6e, #2563a8) !important;
    color: white !important;
    border: none !important;
    border-radius: 10px !important;
    font-family: 'Sora', sans-serif !important;
    font-weight: 600 !important;
    font-size: 0.9rem !important;
    padding: 0.6rem 1.6rem !important;
    box-shadow: 0 4px 14px rgba(37,99,168,0.35) !important;
    transition: all 0.2s !important;
    width: 100%;
}
.stDownloadButton > button:hover {
    transform: translateY(-1px) !important;
    box-shadow: 0 6px 20px rgba(37,99,168,0.45) !important;
}
/* number inputs monospace */
input[type="number"], input[type="text"] {
    font-family: 'JetBrains Mono', monospace !important;
    font-size: 0.85rem !important;
}
/* Remove red borders on empty required */
div[data-baseweb="input"] { border-radius: 8px !important; }
.stTextInput > div > div > input,
.stNumberInput > div > div > input,
.stDateInput > div > div > input {
    border-radius: 8px !important;
    border-color: #d0dcea !important;
}
/* Tab styling */
.stTabs [data-baseweb="tab-list"] {
    gap: 6px;
    background: transparent;
}
.stTabs [data-baseweb="tab"] {
    background: #e8f0fa;
    border-radius: 8px 8px 0 0;
    font-weight: 600;
    font-size: 0.78rem;
    letter-spacing: 0.05em;
    color: #2563a8;
}
.stTabs [aria-selected="true"] {
    background: white !important;
    color: #0d1b2a !important;
    border-bottom: 2px solid #2563a8 !important;
}
hr { border-color: #e4ecf4 !important; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# SIDEBAR – Layout Selector
# ─────────────────────────────────────────────
with st.sidebar:
    st.markdown("### ⚖️ Relatórios de Pesagem")
    st.markdown("<hr style='margin:0.6rem 0 1rem'>", unsafe_allow_html=True)
    st.markdown("**Selecione o modelo de relatório:**")

    layout = st.radio(
        label="layout",
        options=[
            "🚛  Peso Caminhão / Chegada",
            "📦  Pesagem de Container / Saída",
            "🏗️  Estufagem Individual",
        ],
        label_visibility="collapsed",
    )

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("""
    <div style="font-size:0.72rem;color:#8aa8c8;line-height:1.7">
    Preencha os campos na área principal.<br>
    Ao finalizar, clique em <b>Exportar Excel</b> para baixar o arquivo formatado.
    </div>
    """, unsafe_allow_html=True)

# ─────────────────────────────────────────────
# SESSION STATE defaults
# ─────────────────────────────────────────────
def init_state(key, default):
    if key not in st.session_state:
        st.session_state[key] = default

# Rows state for each layout
init_state("rows_caminhao", [{}])
init_state("rows_container", [{}])
init_state("rows_estufagem", [{}])

# ─────────────────────────────────────────────
# EXCEL HELPERS
# ─────────────────────────────────────────────
HEADER_FILL  = PatternFill("solid", start_color="1F4E79")
SUBHDR_FILL  = PatternFill("solid", start_color="2E75B6")
COL_HDR_FILL = PatternFill("solid", start_color="BDD7EE")
WHITE_FILL   = PatternFill("solid", start_color="FFFFFF")
ALT_FILL     = PatternFill("solid", start_color="DEEAF1")

thin = Side(style="thin", color="000000")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

def scell(ws, row, col, value="", bold=False, fill=None, h_align="center",
          font_color="000000", font_size=9, wrap=True):
    c = ws.cell(row=row, column=col, value=value)
    on_dark = fill in (HEADER_FILL, SUBHDR_FILL)
    c.font = Font(name="Arial", size=font_size, bold=bold,
                  color="FFFFFF" if on_dark else font_color)
    c.alignment = Alignment(horizontal=h_align, vertical="center", wrap_text=wrap)
    if fill:
        c.fill = fill
    c.border = BORDER
    return c

def merge_range(ws, r1, c1, r2, c2, value="", bold=False, fill=None,
                h_align="center", font_size=9):
    ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
    c = ws.cell(row=r1, column=c1, value=value)
    on_dark = fill in (HEADER_FILL, SUBHDR_FILL)
    c.font = Font(name="Arial", size=font_size, bold=bold,
                  color="FFFFFF" if on_dark else font_color if not on_dark else "FFFFFF")
    c.alignment = Alignment(horizontal=h_align, vertical="center", wrap_text=True)
    if fill:
        c.fill = fill
        for r in range(r1, r2+1):
            for cc in range(c1, c2+1):
                ws.cell(row=r, column=cc).fill = fill
    c.border = BORDER
    return c

def set_col_widths(ws, widths):
    for col, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(col)].width = w


# ═══════════════════════════════════════════════════════════════
# LAYOUT 1 – Peso Caminhão / Chegada
# ═══════════════════════════════════════════════════════════════
def render_caminhao():
    st.markdown('<div class="hero-title">Peso Caminhão / Peso de Chegada</div>', unsafe_allow_html=True)
    st.markdown('<div class="hero-sub">Relatório de controle de peso por caminhão na chegada</div>', unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)

    # ── Meta fields ──────────────────────────────────────────
    st.markdown('<div class="card"><div class="card-title">Informações Gerais</div>', unsafe_allow_html=True)
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        produtor = st.text_input("Produtor", key="cam_produtor", placeholder="Nome do produtor")
    with c2:
        terminal = st.text_input("Terminal", key="cam_terminal", placeholder="Terminal")
    with c3:
        data_rel = st.date_input("Data", key="cam_data")
    with c4:
        instrucao = st.text_input("Instrução", key="cam_instrucao", placeholder="Nº instrução")
    inspetor = st.text_input("Inspetor", key="cam_inspetor", placeholder="Nome do inspetor")
    st.markdown("</div>", unsafe_allow_html=True)

    # ── Data rows ────────────────────────────────────────────
    st.markdown('<div class="section-header">Lançamentos</div>', unsafe_allow_html=True)

    rows = st.session_state.rows_caminhao
    updated_rows = []

    for i, row in enumerate(rows):
        with st.expander(f"📋 Linha {i+1}", expanded=(i == len(rows)-1)):
            ca, cb, cc, cd = st.columns([2,2,1,1])
            with ca:
                nota_fiscal = st.text_input("Nota Fiscal", value=row.get("nota_fiscal",""), key=f"cam_nf_{i}")
            with cb:
                placa = st.text_input("Placa do Caminhão", value=row.get("placa",""), key=f"cam_placa_{i}")
            with cc:
                lote = st.text_input("Lote", value=row.get("lote",""), key=f"cam_lote_{i}")
            with cd:
                qtd_fardos = st.number_input("Qtd. Fardos", value=row.get("qtd_fardos", 0), min_value=0, key=f"cam_qtd_{i}")

            ce, cf, cg = st.columns(3)
            with ce:
                peso_bruto_nf = st.number_input("Peso Bruto NF (kg)", value=row.get("peso_bruto_nf", 0.0), min_value=0.0, format="%.2f", key=f"cam_pbnf_{i}")
            with cf:
                peso_liq_nf = st.number_input("Peso Líquido NF (kg)", value=row.get("peso_liq_nf", 0.0), min_value=0.0, format="%.2f", key=f"cam_plnf_{i}")
            with cg:
                tara_fardo = st.number_input("Tara Total do Fardo (kg)", value=row.get("tara_fardo", 0.0), min_value=0.0, format="%.2f", key=f"cam_tara_{i}")

            ch, ci, cj, ck = st.columns(4)
            with ch:
                pb_caminhao = st.number_input("Peso Bruto Caminhão Cheio (kg)", value=row.get("pb_caminhao", 0.0), min_value=0.0, format="%.2f", key=f"cam_pbc_{i}")
            with ci:
                tara_caminhao = st.number_input("Tara do Caminhão (kg)", value=row.get("tara_caminhao", 0.0), min_value=0.0, format="%.2f", key=f"cam_tc_{i}")
            with cj:
                pb_carga = pb_caminhao - tara_caminhao
                st.metric("Peso Bruto da Carga (calc.)", f"{pb_carga:,.2f} kg")
            with ck:
                pl_carga = st.number_input("Peso Líquido da Carga (kg)", value=row.get("pl_carga", 0.0), min_value=0.0, format="%.2f", key=f"cam_plc_{i}")

            obs = st.text_input("OBS", value=row.get("obs",""), key=f"cam_obs_{i}")

            remove_col, _ = st.columns([1,5])
            with remove_col:
                if len(rows) > 1 and st.button("🗑 Remover linha", key=f"cam_rm_{i}"):
                    st.session_state.rows_caminhao.pop(i)
                    st.rerun()

            updated_rows.append({
                "nota_fiscal": nota_fiscal, "placa": placa, "lote": lote,
                "qtd_fardos": qtd_fardos, "peso_bruto_nf": peso_bruto_nf,
                "peso_liq_nf": peso_liq_nf, "tara_fardo": tara_fardo,
                "pb_caminhao": pb_caminhao, "tara_caminhao": tara_caminhao,
                "pb_carga": pb_carga, "pl_carga": pl_carga, "obs": obs,
            })

    st.session_state.rows_caminhao = updated_rows

    if st.button("＋ Adicionar linha", key="cam_add"):
        st.session_state.rows_caminhao.append({})
        st.rerun()

    obs_geral = st.text_area("Observações Gerais", key="cam_obs_geral", height=80)

    # ── Export ───────────────────────────────────────────────
    st.divider()
    if st.button("📊 Gerar Excel", type="primary", key="cam_gen"):
        buf = gerar_excel_caminhao(
            produtor, terminal, str(data_rel), instrucao, inspetor,
            st.session_state.rows_caminhao, obs_geral
        )
        st.download_button(
            "⬇ Baixar Excel",
            data=buf,
            file_name="relatorio_peso_caminhao.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="cam_dl"
        )


def gerar_excel_caminhao(produtor, terminal, data, instrucao, inspetor, rows, obs_geral):
    wb = Workbook()
    ws = wb.active
    ws.title = "Peso Caminhão - Chegada"

    merge_range(ws, 1,1, 1,13, value="RELATÓRIO DE PESO DE CAMINHÃO / PESO DE CHEGADA",
                fill=HEADER_FILL, bold=True, font_size=12, h_align="center")
    ws.row_dimensions[1].height = 24

    meta = [("Produtor:", produtor), ("Terminal:", terminal),
            ("Data:", data), ("Instrução:", instrucao), ("Inspetor:", inspetor)]
    col = 1
    for lbl, val in meta:
        scell(ws, 2, col, lbl, bold=True, fill=SUBHDR_FILL, h_align="left", font_size=8)
        scell(ws, 2, col+1, val, fill=WHITE_FILL, h_align="left", font_size=9)
        col += 2
    ws.row_dimensions[2].height = 16
    # fill remaining col if odd
    for c in range(col, 14):
        scell(ws, 2, c, fill=WHITE_FILL)

    # Headers
    for r in range(3,5):
        for c in range(1,14):
            scell(ws, r, c, fill=COL_HDR_FILL)

    merge_range(ws, 3,1, 4,1, value="DATA", fill=COL_HDR_FILL, bold=True, font_size=8)
    merge_range(ws, 3,2, 4,2, value="NOTA FISCAL", fill=COL_HDR_FILL, bold=True, font_size=8)
    merge_range(ws, 3,3, 4,3, value="LOTE", fill=COL_HDR_FILL, bold=True, font_size=8)
    merge_range(ws, 3,4, 4,4, value="QTD.\nFARDOS", fill=COL_HDR_FILL, bold=True, font_size=8)
    merge_range(ws, 3,5, 3,6, value="DADOS DA NOTA FISCAL", fill=COL_HDR_FILL, bold=True, font_size=8)
    scell(ws, 4, 5, value="PESO BRUTO", bold=True, fill=COL_HDR_FILL, font_size=8)
    scell(ws, 4, 6, value="PESO LÍQUIDO", bold=True, fill=COL_HDR_FILL, font_size=8)
    merge_range(ws, 3,7, 4,7, value="TARA TOTAL\nDO FARDO", fill=COL_HDR_FILL, bold=True, font_size=8)
    merge_range(ws, 3,8, 4,8, value="PLACA DO\nCAMINHÃO", fill=COL_HDR_FILL, bold=True, font_size=8)
    merge_range(ws, 3,9, 3,12, value="DADOS DO CAMINHÃO / PESO DE CHEGADA",
                fill=COL_HDR_FILL, bold=True, font_size=8)
    scell(ws, 4, 9,  value="PESO BRUTO\n(CAMINHÃO CHEIO)", bold=True, fill=COL_HDR_FILL, font_size=8)
    scell(ws, 4,10, value="TARA DO\nCAMINHÃO", bold=True, fill=COL_HDR_FILL, font_size=8)
    scell(ws, 4,11, value="PESO BRUTO\nDA CARGA", bold=True, fill=COL_HDR_FILL, font_size=8)
    scell(ws, 4,12, value="PESO LÍQUIDO\nDA CARGA", bold=True, fill=COL_HDR_FILL, font_size=8)
    merge_range(ws, 3,13, 4,13, value="OBS", fill=COL_HDR_FILL, bold=True, font_size=8)
    ws.row_dimensions[3].height = 18
    ws.row_dimensions[4].height = 28

    for i, row in enumerate(rows):
        r = 5 + i
        fill = ALT_FILL if i % 2 else WHITE_FILL
        scell(ws, r, 1,  data,                      fill=fill, font_size=9)
        scell(ws, r, 2,  row.get("nota_fiscal",""),  fill=fill, font_size=9)
        scell(ws, r, 3,  row.get("lote",""),         fill=fill, font_size=9)
        scell(ws, r, 4,  row.get("qtd_fardos",0),    fill=fill, font_size=9)
        scell(ws, r, 5,  row.get("peso_bruto_nf",0), fill=fill, font_size=9)
        scell(ws, r, 6,  row.get("peso_liq_nf",0),   fill=fill, font_size=9)
        scell(ws, r, 7,  row.get("tara_fardo",0),    fill=fill, font_size=9)
        scell(ws, r, 8,  row.get("placa",""),        fill=fill, font_size=9)
        scell(ws, r, 9,  row.get("pb_caminhao",0),   fill=fill, font_size=9)
        scell(ws, r,10,  row.get("tara_caminhao",0), fill=fill, font_size=9)
        scell(ws, r,11,  row.get("pb_carga",0),      fill=fill, font_size=9)
        scell(ws, r,12,  row.get("pl_carga",0),      fill=fill, font_size=9)
        scell(ws, r,13,  row.get("obs",""),          fill=fill, font_size=9)
        ws.row_dimensions[r].height = 15

    last_r = 5 + len(rows)
    merge_range(ws, last_r,1, last_r+1,13, value=f"OBS: {obs_geral}",
                fill=WHITE_FILL, h_align="left", font_size=9)
    ws.row_dimensions[last_r].height = 30
    ws.row_dimensions[last_r+1].height = 20

    set_col_widths(ws, [10,14,10,10,12,12,12,14,14,12,14,14,16])

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ═══════════════════════════════════════════════════════════════
# LAYOUT 2 – Pesagem Container / Saída
# ═══════════════════════════════════════════════════════════════
def render_container():
    st.markdown('<div class="hero-title">Pesagem de Container / Peso Saída</div>', unsafe_allow_html=True)
    st.markdown('<div class="hero-sub">Registro de pesagem de containers na saída</div>', unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)

    st.markdown('<div class="card"><div class="card-title">Informações Gerais</div>', unsafe_allow_html=True)
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        produtor = st.text_input("Produtor", key="cnt_produtor")
    with c2:
        terminal = st.text_input("Terminal", key="cnt_terminal")
    with c3:
        data_rel = st.date_input("Data", key="cnt_data")
    with c4:
        instrucao = st.text_input("Instrução", key="cnt_instrucao")

    c5, c6 = st.columns(2)
    with c5:
        lim_total = st.number_input("Limite de Peso Total da Instrução (kg)", key="cnt_lim_total", min_value=0.0, format="%.2f")
    with c6:
        lim_cntr = st.number_input("Limite de Peso por Container (kg)", key="cnt_lim_cntr", min_value=0.0, format="%.2f")
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="section-header">Lançamentos por Container</div>', unsafe_allow_html=True)

    rows = st.session_state.rows_container
    updated_rows = []

    for i, row in enumerate(rows):
        with st.expander(f"📦 Container {i+1}", expanded=(i == len(rows)-1)):
            ca, cb, cc = st.columns(3)
            with ca:
                container = st.text_input("Nº Container", value=row.get("container",""), key=f"cnt_cont_{i}")
            with cb:
                nota_fiscal = st.text_input("Nota Fiscal", value=row.get("nota_fiscal",""), key=f"cnt_nf_{i}")
            with cc:
                lote = st.text_input("Lote", value=row.get("lote",""), key=f"cnt_lote_{i}")

            cd, ce, cf = st.columns(3)
            with cd:
                total_fardos = st.number_input("Total Fardos", value=row.get("total_fardos",0), min_value=0, key=f"cnt_tf_{i}")
            with ce:
                tara_cntr = st.number_input("Tara Cntr (kg)", value=row.get("tara_cntr",0.0), min_value=0.0, format="%.2f", key=f"cnt_tara_{i}")
            with cf:
                max_gross = st.number_input("Max Gross (kg)", value=row.get("max_gross",0.0), min_value=0.0, format="%.2f", key=f"cnt_mg_{i}")

            cg, ch, ci, cj = st.columns(4)
            with cg:
                hor_inicio = st.text_input("Horário Início", value=row.get("hor_inicio",""), key=f"cnt_hi_{i}", placeholder="HH:MM")
            with ch:
                hor_final = st.text_input("Horário Final", value=row.get("hor_final",""), key=f"cnt_hf_{i}", placeholder="HH:MM")
            with ci:
                data_cnt = st.date_input("Data", key=f"cnt_d_{i}")
            with cj:
                bruto_carga = st.number_input("Bruto Carga (kg)", value=row.get("bruto_carga",0.0), min_value=0.0, format="%.2f", key=f"cnt_bc_{i}")

            if len(rows) > 1 and st.button("🗑 Remover container", key=f"cnt_rm_{i}"):
                st.session_state.rows_container.pop(i)
                st.rerun()

            updated_rows.append({
                "container": container, "nota_fiscal": nota_fiscal, "lote": lote,
                "total_fardos": total_fardos, "tara_cntr": tara_cntr, "max_gross": max_gross,
                "hor_inicio": hor_inicio, "hor_final": hor_final,
                "data_cnt": str(data_cnt), "bruto_carga": bruto_carga,
            })

    st.session_state.rows_container = updated_rows

    if st.button("＋ Adicionar container", key="cnt_add"):
        st.session_state.rows_container.append({})
        st.rerun()

    obs_geral = st.text_area("Observações Gerais", key="cnt_obs_geral", height=80)

    st.divider()
    if st.button("📊 Gerar Excel", type="primary", key="cnt_gen"):
        buf = gerar_excel_container(
            produtor, terminal, str(data_rel), instrucao,
            lim_total, lim_cntr,
            st.session_state.rows_container, obs_geral
        )
        st.download_button(
            "⬇ Baixar Excel",
            data=buf,
            file_name="relatorio_pesagem_container.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="cnt_dl"
        )


def gerar_excel_container(produtor, terminal, data, instrucao, lim_total, lim_cntr, rows, obs_geral):
    wb = Workbook()
    ws = wb.active
    ws.title = "Pesagem Container - Saída"

    merge_range(ws, 1,1, 1,10, value="RELATÓRIO DE PESAGEM DE CONTAINER / PESO SAÍDA",
                fill=HEADER_FILL, bold=True, font_size=12, h_align="center")
    ws.row_dimensions[1].height = 24

    meta = [("Produtor:", produtor), ("Terminal:", terminal), ("Data:", data), ("Instrução:", instrucao)]
    col = 1
    for lbl, val in meta:
        scell(ws, 2, col, lbl, bold=True, fill=SUBHDR_FILL, h_align="left", font_size=8)
        scell(ws, 2, col+1, val, fill=WHITE_FILL, h_align="left", font_size=9)
        col += 2
    ws.row_dimensions[2].height = 16

    scell(ws, 3, 1, f"Limite de Peso Total da Instrução: {lim_total} kg",
          bold=True, fill=SUBHDR_FILL, h_align="left", font_size=8)
    for c in range(2,7):
        scell(ws, 3, c, fill=SUBHDR_FILL)
    scell(ws, 3, 7, f"Limite de Peso por Cntr: {lim_cntr} kg",
          bold=True, fill=SUBHDR_FILL, h_align="left", font_size=8)
    for c in range(8,11):
        scell(ws, 3, c, fill=SUBHDR_FILL)
    ws.row_dimensions[3].height = 16

    hdrs = ["CONTAINER","NOTA FISCAL","LOTE","TOTAL FARDOS",
            "TARA CNTR","MAX GROSS","HORÁRIO INÍCIO","HORÁRIO FINAL","DATA","BRUTO CARGA"]
    for c, h in enumerate(hdrs, 1):
        scell(ws, 4, c, h, bold=True, fill=COL_HDR_FILL, font_size=8)
    ws.row_dimensions[4].height = 22

    for i, row in enumerate(rows):
        r = 5 + i
        fill = ALT_FILL if i % 2 else WHITE_FILL
        vals = [row.get("container",""), row.get("nota_fiscal",""), row.get("lote",""),
                row.get("total_fardos",0), row.get("tara_cntr",0), row.get("max_gross",0),
                row.get("hor_inicio",""), row.get("hor_final",""),
                row.get("data_cnt",""), row.get("bruto_carga",0)]
        for c, v in enumerate(vals, 1):
            scell(ws, r, c, v, fill=fill, font_size=9)
        ws.row_dimensions[r].height = 15

    last_r = 5 + len(rows)
    merge_range(ws, last_r,1, last_r+1,10, value=f"OBS: {obs_geral}",
                fill=WHITE_FILL, h_align="left", font_size=9)
    ws.row_dimensions[last_r].height = 30

    set_col_widths(ws, [16,14,10,12,12,12,14,14,10,14])

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ═══════════════════════════════════════════════════════════════
# LAYOUT 3 – Estufagem Individual Container
# ═══════════════════════════════════════════════════════════════
def render_estufagem():
    st.markdown('<div class="hero-title">Estufagem Individual de Container</div>', unsafe_allow_html=True)
    st.markdown('<div class="hero-sub">Relatório de estufagem por container individual</div>', unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)

    st.markdown('<div class="card"><div class="card-title">Dados do Container</div>', unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        instrucao = st.text_input("Instrução de Embarque", key="est_instrucao")
        produtor = st.text_input("Produtor", key="est_produtor")
        n_container = st.text_input("Nº do Container", key="est_ncont")
        tara_porta = st.number_input("Tara da Porta do Cntr (kg)", key="est_tara", min_value=0.0, format="%.2f")
    with c2:
        qtd_fardos = st.number_input("QTD. de Fardos", key="est_qtdfardos", min_value=0)
        max_gross = st.number_input("Max. Gross (kg)", key="est_maxgross", min_value=0.0, format="%.2f")
        terminal = st.text_input("Terminal", key="est_terminal")
        lacre = st.text_input("Lacre", key="est_lacre")

    c3, c4 = st.columns(2)
    with c3:
        inicio_data = st.date_input("Data Início da Estufagem", key="est_dt_inicio")
        inicio_hora = st.text_input("Hora Início", key="est_hr_inicio", placeholder="HH:MM")
    with c4:
        fim_data = st.date_input("Data Término da Estufagem", key="est_dt_fim")
        fim_hora = st.text_input("Hora Término", key="est_hr_fim", placeholder="HH:MM")
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="section-header">Itens Estufados</div>', unsafe_allow_html=True)

    rows = st.session_state.rows_estufagem
    updated_rows = []

    for i, row in enumerate(rows):
        ca, cb, cc, cd, ce = st.columns([2,2,1,2,3])
        with ca:
            nota_fiscal = st.text_input("Nota Fiscal", value=row.get("nota_fiscal",""),
                                        key=f"est_nf_{i}", label_visibility="visible" if i==0 else "collapsed")
        with cb:
            lote = st.text_input("Lote", value=row.get("lote",""),
                                 key=f"est_lote_{i}", label_visibility="visible" if i==0 else "collapsed")
        with cc:
            qtd = st.number_input("Qtd", value=row.get("qtd",0), min_value=0,
                                  key=f"est_qtd_{i}", label_visibility="visible" if i==0 else "collapsed")
        with cd:
            peso = st.number_input("Peso (kg)", value=row.get("peso",0.0), min_value=0.0,
                                   format="%.2f", key=f"est_peso_{i}",
                                   label_visibility="visible" if i==0 else "collapsed")
        with ce:
            obs = st.text_input("OBS", value=row.get("obs",""),
                                key=f"est_obs_{i}", label_visibility="visible" if i==0 else "collapsed")

        updated_rows.append({
            "nota_fiscal": nota_fiscal, "lote": lote,
            "qtd": qtd, "peso": peso, "obs": obs,
        })

    st.session_state.rows_estufagem = updated_rows

    col_add, col_rm, _ = st.columns([1,1,4])
    with col_add:
        if st.button("＋ Adicionar item", key="est_add"):
            st.session_state.rows_estufagem.append({})
            st.rerun()
    with col_rm:
        if len(rows) > 1 and st.button("− Remover último", key="est_rm"):
            st.session_state.rows_estufagem.pop()
            st.rerun()

    # Totals summary
    total_fardos_calc = sum(r.get("qtd",0) for r in st.session_state.rows_estufagem)
    total_peso_calc   = sum(r.get("peso",0.0) for r in st.session_state.rows_estufagem)
    st.markdown(f"""
    <div style="background:#e8f4e8;border-radius:10px;padding:0.8rem 1.2rem;margin:0.8rem 0;
                border-left:4px solid #28a745;font-size:0.85rem">
        <b>Totais:</b> &nbsp;&nbsp;
        Fardos: <b>{total_fardos_calc}</b> &nbsp;&nbsp;|&nbsp;&nbsp;
        Peso Total: <b>{total_peso_calc:,.2f} kg</b>
    </div>
    """, unsafe_allow_html=True)

    obs_geral = st.text_area("Observação", key="est_obs_geral", height=80)
    inspetor = st.text_input("Nome do Inspetor", key="est_inspetor")

    st.divider()
    if st.button("📊 Gerar Excel", type="primary", key="est_gen"):
        buf = gerar_excel_estufagem(
            instrucao, produtor, n_container, qtd_fardos,
            tara_porta, max_gross, terminal, lacre,
            str(inicio_data), inicio_hora, str(fim_data), fim_hora,
            st.session_state.rows_estufagem, obs_geral, inspetor
        )
        st.download_button(
            "⬇ Baixar Excel",
            data=buf,
            file_name="relatorio_estufagem_individual.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="est_dl"
        )


def gerar_excel_estufagem(instrucao, produtor, n_container, qtd_fardos,
                           tara_porta, max_gross, terminal, lacre,
                           inicio_data, inicio_hora, fim_data, fim_hora,
                           rows, obs_geral, inspetor):
    wb = Workbook()
    ws = wb.active
    ws.title = "Estufagem Individual"

    merge_range(ws, 1,1, 1,5, value="RELATÓRIO DE ESTUFAGEM INDIVIDUAL CONTAINER",
                fill=HEADER_FILL, bold=True, font_size=12, h_align="center")
    ws.row_dimensions[1].height = 24

    def meta_row(ws, r, lbl1, val1, lbl2="", val2=""):
        scell(ws, r, 1, lbl1, bold=True, fill=SUBHDR_FILL, h_align="left", font_size=8)
        scell(ws, r, 2, val1, fill=WHITE_FILL, h_align="left", font_size=9)
        if lbl2:
            scell(ws, r, 3, lbl2, bold=True, fill=SUBHDR_FILL, h_align="left", font_size=8)
            scell(ws, r, 4, val2, fill=WHITE_FILL, h_align="left", font_size=9)
            scell(ws, r, 5, "", fill=WHITE_FILL)
        else:
            for c in range(3,6):
                scell(ws, r, c, fill=WHITE_FILL)
        ws.row_dimensions[r].height = 16

    meta_row(ws, 2, "Instrução de Embarque:", instrucao)
    meta_row(ws, 3, "Produtor:", produtor)
    meta_row(ws, 4, "Nº do Container:", n_container, "QTD. de Fardos:", str(qtd_fardos))
    meta_row(ws, 5, "Tara da Porta do Cntr:", str(tara_porta), "Max. Gross:", str(max_gross))
    meta_row(ws, 6, "Terminal:", terminal, "Lacre:", lacre)
    meta_row(ws, 7, "Começo da Estufagem:", f"{inicio_data} {inicio_hora}",
             "Data/Hora Início:", f"{inicio_data} {inicio_hora}")
    meta_row(ws, 8, "Término da Estufagem:", f"{fim_data} {fim_hora}",
             "Data/Hora Término:", f"{fim_data} {fim_hora}")

    hdrs = ["NOTA FISCAL","LOTE","QTD FARDOS","PESO","OBS"]
    for c, h in enumerate(hdrs, 1):
        scell(ws, 9, c, h, bold=True, fill=COL_HDR_FILL, font_size=8)
    ws.row_dimensions[9].height = 18

    for i, row in enumerate(rows):
        r = 10 + i
        fill = ALT_FILL if i % 2 else WHITE_FILL
        scell(ws, r, 1, row.get("nota_fiscal",""), fill=fill, font_size=9)
        scell(ws, r, 2, row.get("lote",""),        fill=fill, font_size=9)
        scell(ws, r, 3, row.get("qtd",0),          fill=fill, font_size=9)
        scell(ws, r, 4, row.get("peso",0.0),       fill=fill, font_size=9)
        scell(ws, r, 5, row.get("obs",""),         fill=fill, font_size=9)
        ws.row_dimensions[r].height = 14

    last_r = 10 + len(rows)
    merge_range(ws, last_r, 1, last_r+2, 5, value=f"Observação: {obs_geral}",
                fill=WHITE_FILL, h_align="left", font_size=9)
    for r in [last_r, last_r+1, last_r+2]:
        ws.row_dimensions[r].height = 14

    ins_r = last_r + 3
    merge_range(ws, ins_r,1, ins_r,5,
                value=f"INSPETOR: {inspetor}",
                fill=WHITE_FILL, h_align="center", font_size=9, bold=True)
    ws.row_dimensions[ins_r].height = 18

    note_r = ins_r + 2
    NOTE_FILL = PatternFill("solid", start_color="FFF2CC")
    merge_range(ws, note_r,1, note_r+4,5,
                value="ATENÇÃO, INSPETORES\n\n"
                      "• Coloquem a data e o horário de INÍCIO E TÉRMINO da estufagem.\n"
                      "• Assinem seus nomes.\n"
                      "• Anexar etiquetas nos romaneios e marcar os saldos corretamente.",
                fill=NOTE_FILL, h_align="left", font_size=9)
    for r in range(note_r, note_r+5):
        ws.row_dimensions[r].height = 14

    set_col_widths(ws, [20,14,12,14,26])

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ─────────────────────────────────────────────
# ROUTER
# ─────────────────────────────────────────────
if "Caminhão" in layout:
    render_caminhao()
elif "Container" in layout:
    render_container()
else:
    render_estufagem()