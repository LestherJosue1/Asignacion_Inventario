import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io

# ── Configuración de página ───────────────────────────────────────────────────
st.set_page_config(
    page_title="Completación de Dispos",
    page_icon="📦",
    layout="wide"
)

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

    html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

    .block-container { padding: 2rem 3rem; max-width: 1400px; }

    .app-header {
        background: linear-gradient(135deg, #1a2f4e 0%, #0d4f3c 100%);
        border-radius: 12px;
        padding: 2rem 2.5rem;
        margin-bottom: 2rem;
        color: white;
    }
    .app-header h1 { margin: 0; font-size: 1.8rem; font-weight: 700; letter-spacing: -0.5px; }
    .app-header p  { margin: 0.4rem 0 0; font-size: 0.9rem; opacity: 0.75; }

    .section-title {
        font-size: 0.75rem;
        font-weight: 700;
        letter-spacing: 0.12em;
        text-transform: uppercase;
        color: #6b7280;
        margin: 1.5rem 0 0.75rem;
        padding-bottom: 0.4rem;
        border-bottom: 1px solid #e5e7eb;
    }

    .kpi-grid { display: flex; gap: 1rem; margin: 1.2rem 0; }
    .kpi-card {
        flex: 1;
        border-radius: 10px;
        padding: 1.1rem 1.3rem;
        border: 1px solid #e5e7eb;
    }
    .kpi-card.green  { background: #f0fdf4; border-color: #bbf7d0; }
    .kpi-card.yellow { background: #fefce8; border-color: #fde68a; }
    .kpi-card.red    { background: #fff1f2; border-color: #fecdd3; }
    .kpi-card.blue   { background: #eff6ff; border-color: #bfdbfe; }
    .kpi-label { font-size: 0.72rem; font-weight: 600; text-transform: uppercase;
                 letter-spacing: 0.08em; color: #6b7280; margin-bottom: 0.3rem; }
    .kpi-value { font-size: 1.9rem; font-weight: 700; line-height: 1; color: #111827; }
    .kpi-sub   { font-size: 0.78rem; color: #6b7280; margin-top: 0.25rem; }

    .stDataFrame { border-radius: 8px; overflow: hidden; }

    div[data-testid="stNumberInput"] label { font-size: 0.82rem; font-weight: 500; }
    div[data-testid="stSelectbox"]   label { font-size: 0.82rem; font-weight: 500; }

    .stButton > button {
        border-radius: 8px;
        font-weight: 600;
        font-size: 0.9rem;
        padding: 0.55rem 1.4rem;
        border: none;
    }
    .stButton > button[kind="primary"] {
        background: linear-gradient(135deg, #1a2f4e, #0d4f3c);
        color: white;
    }
    .download-btn > button {
        background: #16a34a !important;
        color: white !important;
        width: 100%;
        padding: 0.7rem !important;
    }

    .tip-box {
        background: #f8fafc;
        border-left: 3px solid #3b82f6;
        border-radius: 0 8px 8px 0;
        padding: 0.75rem 1rem;
        font-size: 0.82rem;
        color: #374151;
        margin: 0.75rem 0;
    }
</style>
""", unsafe_allow_html=True)

# ── Header ────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="app-header">
    <h1>📦 Completación de Dispos</h1>
    <p>Asignación de inventario con prioridad configurable · Llave: ESTILO_EQ + DTITULAR</p>
</div>
""", unsafe_allow_html=True)

# ── Defaults de prioridad ─────────────────────────────────────────────────────
DEFAULT_ENTREGA = {
    '01-EXPEDITE': 1,
    '02-PAST DUE': 2,
    '03-DUE':      3,
    '04-AHEAD':    4,
    '05-AHEAD':    5,
}
DEFAULT_LOTSIZE = {
    'D-2200':  1,
    'C-2600':  2,
    'F-1000':  3,
    'B-3300':  4,
    'A-4000':  5,
    'H-0400':  6,
    'G-0850':  7,
    'E-1600':  8,
    'I-0300':  9,
    'MQM':    10,
    'P-0700': 11,
}

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### ⚙️ Configuración")

    st.markdown('<p class="section-title">Modo de prioridad</p>', unsafe_allow_html=True)
    modo = st.selectbox(
        "¿Cuál dimensión domina?",
        options=['ENTREGA', 'LOTSIZE', 'IGUAL'],
        help="ENTREGA: la entrega decide primero, LOTSIZE desempata. LOTSIZE: al revés. IGUAL: se suman sin multiplicador."
    )

    st.markdown('<p class="section-title">Prioridad ENTREGA</p>', unsafe_allow_html=True)
    st.markdown('<div class="tip-box">Menor número = mayor prioridad. Usa 99 para dejar al final.</div>', unsafe_allow_html=True)
    entrega_pesos = {}
    for val, default in DEFAULT_ENTREGA.items():
        entrega_pesos[val] = st.number_input(val, min_value=1, max_value=99, value=default, key=f"e_{val}")

    st.markdown('<p class="section-title">Prioridad LOTSIZE</p>', unsafe_allow_html=True)
    lotsize_pesos = {}
    for val, default in DEFAULT_LOTSIZE.items():
        lotsize_pesos[val] = st.number_input(val, min_value=1, max_value=99, value=default, key=f"l_{val}")

# ── Funciones ─────────────────────────────────────────────────────────────────
def leer_excel(file):
    for header_row in range(10):
        try:
            df = pd.read_excel(file, header=header_row)
            df.columns = df.columns.str.strip().str.upper()
            if 'DISPO' in df.columns and 'ESTILO_EQ' in df.columns:
                return df, header_row
        except Exception:
            continue
    return None, None

def procesar(df, entrega_pesos, lotsize_pesos, modo):
    df = df.copy()
    df.columns = df.columns.str.strip().str.upper()
    df = df.dropna(subset=['ESTILO_EQ', 'DTITULAR', 'LBS_C', 'INV']).reset_index(drop=True)

    df['_PESO_E'] = df['ENTREGA'].map(entrega_pesos).fillna(99).astype(int)
    df['_PESO_L'] = df.get('LOTSIZE', pd.Series(99, index=df.index)).map(lotsize_pesos).fillna(99).astype(int) \
                   if 'LOTSIZE' in df.columns else 99

    if modo == 'ENTREGA':
        df['_ORDEN'] = df['_PESO_E'] * 100 + df['_PESO_L']
    elif modo == 'LOTSIZE':
        df['_ORDEN'] = df['_PESO_L'] * 100 + df['_PESO_E']
    else:
        df['_ORDEN'] = df['_PESO_E'] + df['_PESO_L']

    df['_IDX'] = range(len(df))
    df_sorted = df.sort_values(['ESTILO_EQ', 'DTITULAR', '_ORDEN', '_IDX'])

    lbs_asignado, inv_restante_col = [], []
    for _, grupo in df_sorted.groupby(['ESTILO_EQ', 'DTITULAR'], sort=False):
        inv_disp = grupo['INV'].iloc[0]
        acumulado = 0
        for lbs_c in grupo['LBS_C']:
            disponible = max(0, inv_disp - acumulado)
            asignado = min(lbs_c, disponible)
            lbs_asignado.append(asignado)
            acumulado += lbs_c
            inv_restante_col.append(max(0, inv_disp - acumulado))

    df_sorted['LBS_ASIGNADO'] = lbs_asignado
    df_sorted['INV_RESTANTE'] = inv_restante_col
    df_sorted['LBS_FALTANTE'] = (df_sorted['LBS_C'] - df_sorted['LBS_ASIGNADO']).clip(lower=0)
    df_sorted['PCT_CUBIERTO'] = (df_sorted['LBS_ASIGNADO'] / df_sorted['LBS_C'].replace(0, np.nan)).fillna(0)
    df_sorted['STATUS_DISPO'] = df_sorted['PCT_CUBIERTO'].apply(
        lambda p: '✅ COMPLETA' if p >= 1 else ('⚠️ PARCIAL' if p > 0 else '❌ SIN INVENTARIO')
    )

    return df_sorted.sort_values('_IDX').drop(columns=['_PESO_E', '_PESO_L', '_ORDEN', '_IDX'])

def generar_excel(df_result, entrega_pesos, lotsize_pesos, modo):
    COLOR_H_ORIG = '2F4F7F'
    COLOR_H_NEW  = '1D6A40'
    COLOR_H_CFG  = '7B3F00'
    COLOR_OK     = 'C6EFCE'
    COLOR_WARN   = 'FFEB9C'
    COLOR_ERR    = 'FFC7CE'
    FW = Font(color='FFFFFF', bold=True, name='Arial', size=10)
    FN = Font(name='Arial', size=9)
    FB = Font(name='Arial', size=9, bold=True)
    thin = Side(style='thin', color='CCCCCC')
    brd  = Border(left=thin, right=thin, top=thin, bottom=thin)

    new_cols  = ['LBS_ASIGNADO', 'LBS_FALTANTE', 'INV_RESTANTE', 'PCT_CUBIERTO', 'STATUS_DISPO']
    orig_cols = [c for c in df_result.columns if c not in new_cols]
    df_out    = df_result[orig_cols + new_cols].copy()

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        df_out.to_excel(writer, sheet_name='DETALLE', index=False)

        total      = len(df_result)
        completas  = (df_result['PCT_CUBIERTO'] >= 1).sum()
        parciales  = ((df_result['PCT_CUBIERTO'] > 0) & (df_result['PCT_CUBIERTO'] < 1)).sum()
        sin_inv    = (df_result['PCT_CUBIERTO'] == 0).sum()
        cob_global = df_result['LBS_ASIGNADO'].sum() / df_result['LBS_C'].sum()

        pd.DataFrame({
            'Métrica': ['Total filas', '✅ Completas', '⚠️ Parciales', '❌ Sin inventario',
                        'LBS necesarias', 'LBS asignadas', 'LBS faltantes',
                        '% Cobertura global', 'Modo prioridad', 'Generado'],
            'Valor':   [total, completas, parciales, sin_inv,
                        round(df_result['LBS_C'].sum(), 1),
                        round(df_result['LBS_ASIGNADO'].sum(), 1),
                        round(df_result['LBS_FALTANTE'].sum(), 1),
                        f'{cob_global:.1%}', modo,
                        datetime.now().strftime('%Y-%m-%d %H:%M')]
        }).to_excel(writer, sheet_name='RESUMEN', index=False)

        cfg_e = pd.DataFrame(list(entrega_pesos.items()), columns=['ENTREGA', 'PESO'])
        cfg_l = pd.DataFrame(list(lotsize_pesos.items()), columns=['LOTSIZE', 'PESO'])
        cfg_e.to_excel(writer, sheet_name='CONFIG', index=False, startrow=1, startcol=0)
        cfg_l.to_excel(writer, sheet_name='CONFIG', index=False, startrow=1, startcol=3)

    buf.seek(0)
    wb = load_workbook(buf)

    # Formatear DETALLE
    ws = wb['DETALLE']
    n_orig = len(orig_cols)
    for col_idx in range(1, len(df_out.columns) + 1):
        c = ws.cell(row=1, column=col_idx)
        c.fill = PatternFill('solid', start_color=COLOR_H_NEW if col_idx > n_orig else COLOR_H_ORIG)
        c.font = FW
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        c.border = brd
    ws.row_dimensions[1].height = 30
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = ws.dimensions

    s_idx = df_out.columns.get_loc('STATUS_DISPO') + 1
    p_idx = df_out.columns.get_loc('PCT_CUBIERTO') + 1
    f_ok  = PatternFill('solid', start_color=COLOR_OK)
    f_warn= PatternFill('solid', start_color=COLOR_WARN)
    f_err = PatternFill('solid', start_color=COLOR_ERR)

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        status = row[s_idx - 1].value
        fill = f_ok if status == '✅ COMPLETA' else f_warn if status == '⚠️ PARCIAL' else f_err
        for cell in row:
            cell.font = FN
            cell.border = brd
            if cell.column > n_orig:
                cell.fill = fill
                cell.alignment = Alignment(horizontal='center')
        row[p_idx - 1].number_format = '0.0%'

    for i, col_name in enumerate(df_out.columns, 1):
        w = 18 if col_name in ('DISPO','ESTILO_EQ','ENTREGA','LOTSIZE','STATUS_DISPO') \
            else 14 if col_name in ('LBS_C','LBS_ASIGNADO','LBS_FALTANTE','INV','INV_RESTANTE') \
            else 12 if col_name == 'PCT_CUBIERTO' else 11
        ws.column_dimensions[get_column_letter(i)].width = w

    # Formatear RESUMEN
    ws_r = wb['RESUMEN']
    ws_r.column_dimensions['A'].width = 26
    ws_r.column_dimensions['B'].width = 20
    for row in ws_r.iter_rows(min_row=1, max_row=ws_r.max_row):
        for cell in row:
            cell.border = brd
            cell.font   = FW if cell.row == 1 else FN
            if cell.row == 1:
                cell.fill = PatternFill('solid', start_color=COLOR_H_ORIG)

    # Formatear CONFIG
    ws_c = wb['CONFIG']
    ws_c['A1'] = 'PRIORIDAD ENTREGA'
    ws_c['D1'] = 'PRIORIDAD LOTSIZE'
    for cell in [ws_c['A1'], ws_c['D1']]:
        cell.fill = PatternFill('solid', start_color=COLOR_H_CFG)
        cell.font = FW
        cell.alignment = Alignment(horizontal='center')
    ws_c.merge_cells('A1:B1')
    ws_c.merge_cells('D1:E1')
    for col in ['A','B','D','E']:
        ws_c.column_dimensions[col].width = 16
    ws_c.column_dimensions['C'].width = 4
    for row in ws_c.iter_rows(min_row=2, max_row=ws_c.max_row):
        for cell in row:
            cell.font   = FB if cell.row == 2 else FN
            cell.border = brd
            if cell.row == 2:
                cell.fill = PatternFill('solid', start_color='D9D9D9')

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out

# ── Layout principal ──────────────────────────────────────────────────────────
col_upload, col_main = st.columns([1, 2.5], gap="large")

with col_upload:
    st.markdown('<p class="section-title">Archivo de entrada</p>', unsafe_allow_html=True)
    uploaded = st.file_uploader("Sube tu archivo Excel", type=['xlsx','xls'], label_visibility="collapsed")

    if uploaded:
        df_raw, header_row = leer_excel(uploaded)
        if df_raw is None:
            st.error("No se encontraron los encabezados. Verifica el archivo.")
        else:
            st.success(f"**{uploaded.name}**")
            st.caption(f"{len(df_raw):,} filas · {df_raw['DISPO'].nunique():,} dispos · encabezados en fila {header_row + 1}")

            cols_faltantes = [c for c in ['ENTREGA','DISPO','ESTILO_EQ','DTITULAR','LBS_C','INV']
                              if c not in df_raw.columns]
            if cols_faltantes:
                st.warning(f"Columnas faltantes: {cols_faltantes}")
            else:
                st.markdown('<p class="section-title">Vista previa</p>', unsafe_allow_html=True)
                st.dataframe(
                    df_raw[['DISPO','ENTREGA','ESTILO_EQ','DTITULAR','LBS_C','INV']].head(8),
                    use_container_width=True, hide_index=True
                )

                if st.button("▶ Procesar", type="primary", use_container_width=True):
                    with st.spinner("Calculando asignación..."):
                        st.session_state['df_result'] = procesar(df_raw, entrega_pesos, lotsize_pesos, modo)
                        st.session_state['modo'] = modo

with col_main:
    if 'df_result' in st.session_state:
        df_r = st.session_state['df_result']
        total     = len(df_r)
        completas = (df_r['PCT_CUBIERTO'] >= 1).sum()
        parciales = ((df_r['PCT_CUBIERTO'] > 0) & (df_r['PCT_CUBIERTO'] < 1)).sum()
        sin_inv   = (df_r['PCT_CUBIERTO'] == 0).sum()
        cob       = df_r['LBS_ASIGNADO'].sum() / df_r['LBS_C'].sum()

        st.markdown('<p class="section-title">Resumen de asignación</p>', unsafe_allow_html=True)
        st.markdown(f"""
        <div class="kpi-grid">
            <div class="kpi-card green">
                <div class="kpi-label">✅ Completas</div>
                <div class="kpi-value">{completas:,}</div>
                <div class="kpi-sub">{completas/total:.1%} del total</div>
            </div>
            <div class="kpi-card yellow">
                <div class="kpi-label">⚠️ Parciales</div>
                <div class="kpi-value">{parciales:,}</div>
                <div class="kpi-sub">{parciales/total:.1%} del total</div>
            </div>
            <div class="kpi-card red">
                <div class="kpi-label">❌ Sin inventario</div>
                <div class="kpi-value">{sin_inv:,}</div>
                <div class="kpi-sub">{sin_inv/total:.1%} del total</div>
            </div>
            <div class="kpi-card blue">
                <div class="kpi-label">📦 Cobertura global</div>
                <div class="kpi-value">{cob:.1%}</div>
                <div class="kpi-sub">Modo: {st.session_state['modo']}</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        st.markdown('<p class="section-title">Detalle de resultados</p>', unsafe_allow_html=True)

        # Filtros rápidos
        fc1, fc2, fc3 = st.columns(3)
        with fc1:
            filtro_status = st.multiselect(
                "Status", ['✅ COMPLETA', '⚠️ PARCIAL', '❌ SIN INVENTARIO'],
                default=['✅ COMPLETA', '⚠️ PARCIAL', '❌ SIN INVENTARIO'],
                label_visibility="collapsed"
            )
        with fc2:
            opciones_entrega = sorted(df_r['ENTREGA'].dropna().unique())
            filtro_entrega = st.multiselect("ENTREGA", opciones_entrega, default=opciones_entrega, label_visibility="collapsed")
        with fc3:
            if 'LOTSIZE' in df_r.columns:
                opciones_lotsize = sorted(df_r['LOTSIZE'].dropna().unique())
                filtro_lotsize = st.multiselect("LOTSIZE", opciones_lotsize, default=opciones_lotsize, label_visibility="collapsed")
            else:
                filtro_lotsize = None

        df_vis = df_r[df_r['STATUS_DISPO'].isin(filtro_status) & df_r['ENTREGA'].isin(filtro_entrega)]
        if filtro_lotsize and 'LOTSIZE' in df_r.columns:
            df_vis = df_vis[df_vis['LOTSIZE'].isin(filtro_lotsize)]

        cols_show = ['DISPO','ENTREGA','ESTILO_EQ','DTITULAR','LBS_C','INV',
                     'LBS_ASIGNADO','LBS_FALTANTE','PCT_CUBIERTO','STATUS_DISPO']
        if 'LOTSIZE' in df_r.columns:
            cols_show.insert(2, 'LOTSIZE')
        cols_show = [c for c in cols_show if c in df_r.columns]

        st.dataframe(
            df_vis[cols_show].style.format({'PCT_CUBIERTO': '{:.1%}', 'LBS_C': '{:,.1f}',
                                            'LBS_ASIGNADO': '{:,.1f}', 'LBS_FALTANTE': '{:,.1f}'}),
            use_container_width=True, hide_index=True, height=420
        )
        st.caption(f"Mostrando {len(df_vis):,} de {total:,} filas")

        st.markdown('<p class="section-title">Descargar resultado</p>', unsafe_allow_html=True)
        ts = datetime.now().strftime('%Y%m%d%H%M')
        excel_bytes = generar_excel(df_r, entrega_pesos, lotsize_pesos, st.session_state['modo'])

        st.markdown('<div class="download-btn">', unsafe_allow_html=True)
        st.download_button(
            label=f"⬇️  Descargar Excel — INVENTARIO_DISPOS_{ts}.xlsx",
            data=excel_bytes,
            file_name=f"INVENTARIO_DISPOS_{ts}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        st.markdown('</div>', unsafe_allow_html=True)

    else:
        st.markdown("""
        <div style="display:flex; align-items:center; justify-content:center; height:400px;
                    border: 2px dashed #d1d5db; border-radius:12px; color:#9ca3af; flex-direction:column; gap:0.5rem;">
            <div style="font-size:2.5rem">📂</div>
            <div style="font-weight:600; font-size:1rem;">Sube tu archivo y presiona Procesar</div>
            <div style="font-size:0.82rem;">Los resultados aparecerán aquí</div>
        </div>
        """, unsafe_allow_html=True)
