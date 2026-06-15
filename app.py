import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io

# ── Configuración de página ───────────────────────────────────────────────────
st.set_page_config(page_title="Completación de Dispos", page_icon="📦", layout="wide")

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
    .block-container { padding: 2rem 3rem; max-width: 1400px; }
    .app-header {
        background: linear-gradient(135deg, #1a2f4e 0%, #0d4f3c 100%);
        border-radius: 12px; padding: 2rem 2.5rem; margin-bottom: 2rem; color: white;
    }
    .app-header h1 { margin: 0; font-size: 1.8rem; font-weight: 700; letter-spacing: -0.5px; }
    .app-header p  { margin: 0.4rem 0 0; font-size: 0.9rem; opacity: 0.75; }
    .section-title {
        font-size: 0.75rem; font-weight: 700; letter-spacing: 0.12em;
        text-transform: uppercase; color: #6b7280;
        margin: 1.5rem 0 0.75rem; padding-bottom: 0.4rem; border-bottom: 1px solid #e5e7eb;
    }
    .kpi-grid { display: flex; gap: 1rem; margin: 1.2rem 0; flex-wrap: wrap; }
    .kpi-card {
        flex: 1; min-width: 120px; border-radius: 10px;
        padding: 1.1rem 1.3rem; border: 1px solid #e5e7eb;
    }
    .kpi-card.green  { background: #f0fdf4; border-color: #bbf7d0; }
    .kpi-card.yellow { background: #fefce8; border-color: #fde68a; }
    .kpi-card.red    { background: #fff1f2; border-color: #fecdd3; }
    .kpi-card.blue   { background: #eff6ff; border-color: #bfdbfe; }
    .kpi-card.gray   { background: #f9fafb; border-color: #e5e7eb; }
    .kpi-label { font-size: 0.72rem; font-weight: 600; text-transform: uppercase;
                 letter-spacing: 0.08em; color: #6b7280; margin-bottom: 0.3rem; }
    .kpi-value { font-size: 1.9rem; font-weight: 700; line-height: 1; color: #111827; }
    .kpi-sub   { font-size: 0.78rem; color: #6b7280; margin-top: 0.25rem; }
    .tip-box {
        background: #f8fafc; border-left: 3px solid #3b82f6;
        border-radius: 0 8px 8px 0; padding: 0.75rem 1rem;
        font-size: 0.82rem; color: #374151; margin: 0.75rem 0;
    }
    .warn-box {
        background: #fffbeb; border-left: 3px solid #f59e0b;
        border-radius: 0 8px 8px 0; padding: 0.75rem 1rem;
        font-size: 0.82rem; color: #374151; margin: 0.75rem 0;
    }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="app-header">
    <h1>📦 Completación de Dispos</h1>
    <p>Asignación de inventario con prioridad configurable · Llave: ESTILO_EQ + DTITULAR · Solo líneas ACTIVAS reciben inventario</p>
</div>
""", unsafe_allow_html=True)

# ── Defaults ──────────────────────────────────────────────────────────────────
DEFAULT_ENTREGA = {
    '01-EXPEDITE': 1, '02-PAST DUE': 2, '03-DUE': 3, '04-AHEAD': 4, '05-AHEAD': 5,
}
DEFAULT_LOTSIZE = {
    'C-2600': 1, 'D-2200': 2, 'B-3300': 3, 'A-4000': 4, 'F-1000': 5,
}
CASCADA_COLOR_NOMBRES = {
    '0': 'BLANCO',   '1': 'AMARILLO', '2': 'NARANJA', '3': 'ROJO',
    '4': 'MORADO',   '5': 'ROYAL',    '6': 'VERDE',   '7': 'CAFÉ',
    '8': 'GRIS',     '9': 'NEGRO',
}

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### ⚙️ Configuración")

    # ── Inventario ────────────────────────────────────────────────────────────
    st.markdown('<p class="section-title">Inventario disponible</p>', unsafe_allow_html=True)
    usar_plan_ins = st.checkbox(
        "Incluir PLAN_INS_DIA1 (INV + PLAN_INS_DIA1)",
        value=False,
        help="Suma el inventario planeado de entrada al inventario actual antes de asignar."
    )

    # ── Modo prioridad ────────────────────────────────────────────────────────
    st.markdown('<p class="section-title">Modo de prioridad</p>', unsafe_allow_html=True)
    modo = st.selectbox(
        "¿Cuál dimensión domina?",
        options=['ENTREGA', 'LOTSIZE', 'IGUAL'],
        help="ENTREGA: la entrega decide primero, LOTSIZE desempata. LOTSIZE: al revés. IGUAL: se suman."
    )

    # ── ENTREGA ───────────────────────────────────────────────────────────────
    st.markdown('<p class="section-title">Prioridad ENTREGA</p>', unsafe_allow_html=True)
    st.markdown('<div class="tip-box">Menor número = mayor prioridad.</div>', unsafe_allow_html=True)
    entrega_pesos = {}
    for val, default in DEFAULT_ENTREGA.items():
        entrega_pesos[val] = st.number_input(val, min_value=1, max_value=99,
                                             value=default, key=f"e_{val}")

    # ── LOTSIZE ───────────────────────────────────────────────────────────────
    st.markdown('<p class="section-title">Prioridad LOTSIZE</p>', unsafe_allow_html=True)
    lotsize_pesos = {}
    for val, default in DEFAULT_LOTSIZE.items():
        lotsize_pesos[val] = st.number_input(val, min_value=1, max_value=99,
                                             value=default, key=f"l_{val}")

    # ── Cascada de color ──────────────────────────────────────────────────────
    st.markdown('<p class="section-title">Cascada de color (COLOR_A dígito 3)</p>',
                unsafe_allow_html=True)
    st.markdown(
        '<div class="warn-box">🔴 INACTIVO = la línea <b>no recibe inventario</b>.<br>'
        '🟢 ACTIVO = participa normalmente en la asignación.</div>',
        unsafe_allow_html=True
    )
    cascada_activo = {}
    for digito, nombre in CASCADA_COLOR_NOMBRES.items():
        cascada_activo[digito] = st.selectbox(
            f"{digito} — {nombre}",
            options=[1, 0],
            index=0,
            format_func=lambda x: "🟢 ACTIVO" if x == 1 else "🔴 INACTIVO",
            key=f"color_{digito}"
        )

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


def extraer_digito3(color_a):
    try:
        s = str(int(float(color_a))) if str(color_a).replace('.', '').isdigit() else str(color_a)
        return s[2] if len(s) >= 3 else None
    except Exception:
        return None


def procesar(df, entrega_pesos, lotsize_pesos, modo, cascada_activo, usar_plan_ins=False):
    df = df.copy()
    df.columns = df.columns.str.strip().str.upper()
    df = df.dropna(subset=['ESTILO_EQ', 'DTITULAR', 'LBS_C', 'INV']).reset_index(drop=True)

    # ── Inventario efectivo ───────────────────────────────────────────────────
    if usar_plan_ins and 'PLAN_INS_DIA1' in df.columns:
        df['INV_EFECTIVO'] = df['INV'].fillna(0) + df['PLAN_INS_DIA1'].fillna(0)
    else:
        df['INV_EFECTIVO'] = df['INV'].fillna(0)

    # ── Cascada de color ──────────────────────────────────────────────────────
    if 'COLOR_A' in df.columns:
        df['_DIGITO3']     = df['COLOR_A'].apply(extraer_digito3)
        df['COLOR_NOMBRE'] = df['_DIGITO3'].map(CASCADA_COLOR_NOMBRES).fillna('DESCONOCIDO')
        df['ACTIVO']       = df['_DIGITO3'].map(cascada_activo).fillna(0).astype(int)
    else:
        df['_DIGITO3']     = None
        df['COLOR_NOMBRE'] = 'SIN COLOR_A'
        df['ACTIVO']       = 1

    # ── Pesos de prioridad ────────────────────────────────────────────────────
    df['_PESO_E'] = df['ENTREGA'].map(entrega_pesos).fillna(99).astype(int)
    df['_PESO_L'] = df['LOTSIZE'].map(lotsize_pesos).fillna(99).astype(int) \
                   if 'LOTSIZE' in df.columns else 99

    if modo == 'ENTREGA':
        df['_ORDEN'] = df['_PESO_E'] * 100 + df['_PESO_L']
    elif modo == 'LOTSIZE':
        df['_ORDEN'] = df['_PESO_L'] * 100 + df['_PESO_E']
    else:
        df['_ORDEN'] = df['_PESO_E'] + df['_PESO_L']

    df['_IDX'] = range(len(df))

    # ACTIVAS primero dentro del grupo, luego por orden de prioridad
    df_sorted = df.sort_values(
        ['ESTILO_EQ', 'DTITULAR', 'ACTIVO', '_ORDEN', '_IDX'],
        ascending=[True, True, False, True, True]
    )

    # ── Asignación acumulativa ────────────────────────────────────────────────
    lbs_asignado, inv_restante_col = [], []

    for _, grupo in df_sorted.groupby(['ESTILO_EQ', 'DTITULAR'], sort=False):
        inv_disp  = grupo['INV_EFECTIVO'].iloc[0]
        acumulado = 0
        for _, fila in grupo.iterrows():
            lbs_c  = fila['LBS_C']
            activo = fila['ACTIVO']
            if activo == 1:
                disponible = max(0, inv_disp - acumulado)
                asignado   = min(lbs_c, disponible)
                acumulado += lbs_c
            else:
                asignado = 0  # no recibe y no consume inventario
            lbs_asignado.append(asignado)
            inv_restante_col.append(max(0, inv_disp - acumulado))

    df_sorted['LBS_ASIGNADO'] = lbs_asignado
    df_sorted['INV_RESTANTE'] = inv_restante_col
    df_sorted['LBS_FALTANTE'] = (df_sorted['LBS_C'] - df_sorted['LBS_ASIGNADO']).clip(lower=0)

    # PCT por línea
    df_sorted['PCT_LINEA'] = (
        df_sorted['LBS_ASIGNADO'] / df_sorted['LBS_C'].replace(0, np.nan)
    ).fillna(0)

    # STATUS a nivel DISPO — evalúa solo líneas ACTIVAS
    activas_mask = df_sorted['ACTIVO'] == 1
    pct_activas  = df_sorted[activas_mask].groupby('DISPO')['PCT_LINEA'].min().rename('_min_pct')
    df_sorted    = df_sorted.merge(pct_activas, on='DISPO', how='left')
    df_sorted['_min_pct'] = df_sorted['_min_pct'].fillna(-1)

    def status_dispo(row):
        if row['ACTIVO'] == 0:
            return '⛔ INACTIVA'
        p = row['_min_pct']
        if p < 0:   return '⛔ INACTIVA'
        if p >= 1:  return '✅ COMPLETA'
        if p > 0:   return '⚠️ PARCIAL'
        return '❌ SIN INVENTARIO'

    df_sorted['STATUS_DISPO'] = df_sorted.apply(status_dispo, axis=1)

    # PCT a nivel DISPO (solo activas)
    dispo_lbs_c  = df_sorted[activas_mask].groupby('DISPO')['LBS_C'].sum()
    dispo_lbs_a  = df_sorted[activas_mask].groupby('DISPO')['LBS_ASIGNADO'].sum()
    pct_dispo    = (dispo_lbs_a / dispo_lbs_c.replace(0, np.nan)).fillna(0).rename('PCT_DISPO')
    df_sorted    = df_sorted.merge(pct_dispo, on='DISPO', how='left')
    df_sorted['PCT_DISPO'] = df_sorted['PCT_DISPO'].fillna(0)

    drop_cols = ['_PESO_E', '_PESO_L', '_ORDEN', '_IDX', '_DIGITO3', '_min_pct']
    return df_sorted.sort_values('_IDX').drop(columns=drop_cols)



def generar_reporte(df_result):
    """Genera la hoja REPORTE con los 4 cuadros solicitados."""
    completas = df_result[
        (df_result['STATUS_DISPO'] == '✅ COMPLETA') & (df_result['ACTIVO'] == 1)
    ]

    # Cuadro 1: STATUS COMPLETA — LBS_C por ENTREGA
    c1 = completas.groupby('ENTREGA')['LBS_C'].sum().reset_index()
    c1.columns = ['ENTREGA', 'LBS_C']
    c1 = c1.sort_values('LBS_C', ascending=False)
    c1.loc[len(c1)] = ['TOTAL', c1['LBS_C'].sum()]

    # Cuadro 2: STATUS COMPLETA — LBS_C por LOTSIZE
    if 'LOTSIZE' in completas.columns:
        c2 = completas.groupby('LOTSIZE')['LBS_C'].sum().reset_index()
        c2.columns = ['LOTSIZE', 'LBS_C']
        c2 = c2.sort_values('LBS_C', ascending=False)
        c2.loc[len(c2)] = ['TOTAL', c2['LBS_C'].sum()]
    else:
        c2 = pd.DataFrame({'LOTSIZE': ['Sin datos'], 'LBS_C': [0]})

    # Cuadro 3: STATUS COMPLETA — LBS_C por COLOR_NOMBRE
    if 'COLOR_NOMBRE' in completas.columns:
        c3 = completas.groupby('COLOR_NOMBRE')['LBS_C'].sum().reset_index()
        c3.columns = ['COLOR_NOMBRE', 'LBS_C']
        c3 = c3.sort_values('LBS_C', ascending=False)
        c3.loc[len(c3)] = ['TOTAL', c3['LBS_C'].sum()]
    else:
        c3 = pd.DataFrame({'COLOR_NOMBRE': ['Sin datos'], 'LBS_C': [0]})

    # Cuadro 4: Resumen por ENTREGA - LOTSIZE - DISPO - ESTILO_EQ - DTITULAR - LBS_C
    group_cols = [c for c in ['ENTREGA', 'LOTSIZE', 'DISPO', 'ESTILO_EQ', 'DTITULAR']
                  if c in df_result.columns]
    c4 = df_result[df_result['ACTIVO'] == 1].groupby(group_cols, dropna=False)['LBS_C'].sum().reset_index()
    c4 = c4.sort_values(['ENTREGA', 'LBS_C'], ascending=[True, False])

    return c1, c2, c3, c4


def escribir_reporte_excel(ws, c1, c2, c3, c4, FW, FN, FB, brd,
                            COLOR_H_ORIG, COLOR_H_NEW, COLOR_H_CFG):
    """Escribe los 4 cuadros en la hoja REPORTE."""
    COLOR_TOTAL = 'D1FAE5'
    ft_total = Font(name='Arial', size=9, bold=True, color='065F46')

    def write_table(ws, title, df, start_row, start_col, header_color):
        # Título del cuadro
        title_cell = ws.cell(row=start_row, column=start_col, value=title)
        title_cell.font = FW
        title_cell.fill = PatternFill('solid', start_color=header_color)
        title_cell.alignment = Alignment(horizontal='center', vertical='center')
        title_cell.border = brd
        end_col = start_col + len(df.columns) - 1
        if end_col > start_col:
            ws.merge_cells(
                start_row=start_row, start_column=start_col,
                end_row=start_row, end_column=end_col
            )

        # Headers
        for ci, col_name in enumerate(df.columns, start=start_col):
            c = ws.cell(row=start_row + 1, column=ci, value=col_name)
            c.font = FB
            c.fill = PatternFill('solid', start_color='D9D9D9')
            c.alignment = Alignment(horizontal='center')
            c.border = brd

        # Datos
        for ri, (_, row_data) in enumerate(df.iterrows(), start=start_row + 2):
            is_total = any(str(v).upper() == 'TOTAL' for v in row_data.values)
            for ci, val in enumerate(row_data.values, start=start_col):
                cell = ws.cell(row=ri, column=ci, value=val)
                cell.border = brd
                if is_total:
                    cell.font = ft_total
                    cell.fill = PatternFill('solid', start_color=COLOR_TOTAL)
                    cell.alignment = Alignment(horizontal='center')
                else:
                    cell.font = FN
                    if isinstance(val, (int, float)):
                        cell.number_format = '#,##0.0'
                        cell.alignment = Alignment(horizontal='right')

        return start_row + len(df) + 3  # siguiente fila disponible

    # ── Cuadro 1 (col A) ─────────────────────────────────────────────────────
    next_row = write_table(ws, '📋 Cuadro 1 — Completas por ENTREGA',
                           c1, 1, 1, COLOR_H_ORIG)

    # ── Cuadro 2 (col A, debajo de Cuadro 1) ─────────────────────────────────
    next_row = write_table(ws, '📋 Cuadro 2 — Completas por LOTSIZE',
                           c2, next_row, 1, COLOR_H_ORIG)

    # ── Cuadro 3 (col A, debajo de Cuadro 2) ─────────────────────────────────
    write_table(ws, '📋 Cuadro 3 — Completas por COLOR',
                c3, next_row, 1, COLOR_H_ORIG)

    # ── Cuadro 4 (col E) — Resumen detallado ─────────────────────────────────
    write_table(ws, '📋 Cuadro 4 — Resumen por ENTREGA / LOTSIZE / DISPO / ESTILO / TITULAR',
                c4, 1, 5, COLOR_H_NEW)

    # Anchos
    for col in ['A', 'B', 'C']:
        ws.column_dimensions[col].width = 18
    ws.column_dimensions['D'].width = 4
    for col in ['E', 'F', 'G', 'H', 'I', 'J']:
        ws.column_dimensions[col].width = 16

def generar_excel(df_result, entrega_pesos, lotsize_pesos, cascada_activo, modo, usar_plan_ins):
    COLOR_H_ORIG = '2F4F7F'
    COLOR_H_NEW  = '1D6A40'
    COLOR_H_CFG  = '7B3F00'
    COLOR_OK     = 'C6EFCE'
    COLOR_WARN   = 'FFEB9C'
    COLOR_ERR    = 'FFC7CE'
    COLOR_INACT  = 'E5E7EB'
    FW  = Font(color='FFFFFF', bold=True, name='Arial', size=10)
    FN  = Font(name='Arial', size=9)
    FB  = Font(name='Arial', size=9, bold=True)
    FI  = Font(name='Arial', size=9, color='9CA3AF', italic=True)
    thin = Side(style='thin', color='CCCCCC')
    brd  = Border(left=thin, right=thin, top=thin, bottom=thin)

    new_cols  = ['COLOR_NOMBRE', 'ACTIVO', 'INV_EFECTIVO',
                 'LBS_ASIGNADO', 'LBS_FALTANTE', 'INV_RESTANTE',
                 'PCT_LINEA', 'PCT_DISPO', 'STATUS_DISPO']
    orig_cols = [c for c in df_result.columns if c not in new_cols]
    df_out    = df_result[orig_cols + [c for c in new_cols if c in df_result.columns]].copy()

    activas    = df_result[df_result['ACTIVO'] == 1]
    dispo_min  = activas.groupby('DISPO')['PCT_LINEA'].min()
    total_d    = dispo_min.count()
    completas  = (dispo_min >= 1).sum()
    parciales  = ((dispo_min > 0) & (dispo_min < 1)).sum()
    sin_inv    = (dispo_min == 0).sum()
    inac_lines = (df_result['ACTIVO'] == 0).sum()
    cob_global = activas['LBS_ASIGNADO'].sum() / activas['LBS_C'].sum() if len(activas) else 0

    cfg_cascada = pd.DataFrame([
        {'Dígito': d, 'Color': CASCADA_COLOR_NOMBRES[d],
         'Estado': '🟢 ACTIVO' if cascada_activo.get(d, 1) == 1 else '🔴 INACTIVO'}
        for d in CASCADA_COLOR_NOMBRES
    ])

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        df_out.to_excel(writer, sheet_name='DETALLE', index=False)

        pd.DataFrame({
            'Métrica': [
                'DISPOs totales (activas)', '✅ DISPOs completas', '⚠️ DISPOs parciales',
                '❌ DISPOs sin inventario', 'Líneas inactivas (excluidas)',
                'LBS necesarias (activas)', 'LBS asignadas', 'LBS faltantes',
                '% Cobertura global (LBS activas)',
                'Inventario usado', 'Modo prioridad', 'Generado'
            ],
            'Valor': [
                total_d, completas, parciales, sin_inv, inac_lines,
                round(activas['LBS_C'].sum(), 1),
                round(activas['LBS_ASIGNADO'].sum(), 1),
                round(activas['LBS_FALTANTE'].sum(), 1),
                f'{cob_global:.1%}',
                'INV + PLAN_INS_DIA1' if usar_plan_ins else 'Solo INV',
                modo,
                datetime.now().strftime('%Y-%m-%d %H:%M')
            ]
        }).to_excel(writer, sheet_name='RESUMEN', index=False)

        cfg_e = pd.DataFrame(list(entrega_pesos.items()), columns=['ENTREGA', 'PESO'])
        cfg_l = pd.DataFrame(list(lotsize_pesos.items()), columns=['LOTSIZE', 'PESO'])
        cfg_e.to_excel(writer, sheet_name='CONFIG', index=False, startrow=1, startcol=0)
        cfg_l.to_excel(writer, sheet_name='CONFIG', index=False, startrow=1, startcol=3)
        cfg_cascada.to_excel(writer, sheet_name='CONFIG', index=False, startrow=1, startcol=6)

        # Hoja REPORTE — placeholder (se llena después con openpyxl)
        pd.DataFrame().to_excel(writer, sheet_name='REPORTE', index=False)

    buf.seek(0)
    wb = load_workbook(buf)

    # ── Formatear DETALLE ─────────────────────────────────────────────────────
    ws     = wb['DETALLE']
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

    s_idx  = df_out.columns.get_loc('STATUS_DISPO') + 1
    pl_idx = df_out.columns.get_loc('PCT_LINEA') + 1
    pd_idx = df_out.columns.get_loc('PCT_DISPO') + 1

    f_ok   = PatternFill('solid', start_color=COLOR_OK)
    f_warn = PatternFill('solid', start_color=COLOR_WARN)
    f_err  = PatternFill('solid', start_color=COLOR_ERR)
    f_inac = PatternFill('solid', start_color=COLOR_INACT)

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        status  = row[s_idx - 1].value
        is_inac = status == '⛔ INACTIVA'
        fill    = (f_inac if is_inac else
                   f_ok   if status == '✅ COMPLETA' else
                   f_warn if status == '⚠️ PARCIAL' else f_err)
        for cell in row:
            cell.font   = FI if is_inac else FN
            cell.border = brd
            if cell.column > n_orig:
                cell.fill      = fill
                cell.alignment = Alignment(horizontal='center')
        row[pl_idx - 1].number_format = '0.0%'
        row[pd_idx - 1].number_format = '0.0%'

    for i, col_name in enumerate(df_out.columns, 1):
        w = (18 if col_name in ('DISPO','ESTILO_EQ','ENTREGA','LOTSIZE','STATUS_DISPO','COLOR_NOMBRE')
             else 14 if col_name in ('LBS_C','LBS_ASIGNADO','LBS_FALTANTE','INV','INV_EFECTIVO','INV_RESTANTE','PLAN_INS_DIA1')
             else 12 if col_name in ('PCT_LINEA','PCT_DISPO')
             else  9 if col_name == 'ACTIVO'
             else 11)
        ws.column_dimensions[get_column_letter(i)].width = w

    # ── Formatear RESUMEN ─────────────────────────────────────────────────────
    ws_r = wb['RESUMEN']
    ws_r.column_dimensions['A'].width = 32
    ws_r.column_dimensions['B'].width = 22
    for row in ws_r.iter_rows(min_row=1, max_row=ws_r.max_row):
        for cell in row:
            cell.border = brd
            cell.font   = FW if cell.row == 1 else FN
            if cell.row == 1:
                cell.fill = PatternFill('solid', start_color=COLOR_H_ORIG)

    # ── Formatear CONFIG ──────────────────────────────────────────────────────
    ws_c = wb['CONFIG']
    titulos = {1: 'PRIORIDAD ENTREGA', 4: 'PRIORIDAD LOTSIZE', 7: 'CASCADA DE COLOR'}
    for col_start, titulo in titulos.items():
        cell = ws_c.cell(row=1, column=col_start)
        cell.value     = titulo
        cell.fill      = PatternFill('solid', start_color=COLOR_H_CFG)
        cell.font      = FW
        cell.alignment = Alignment(horizontal='center')
    for m in ['A1:B1', 'D1:E1', 'G1:I1']:
        ws_c.merge_cells(m)
    for col in ['A','B','D','E','G','H','I']:
        ws_c.column_dimensions[col].width = 16
    for col in ['C','F']:
        ws_c.column_dimensions[col].width = 4
    for row in ws_c.iter_rows(min_row=2, max_row=ws_c.max_row):
        for cell in row:
            cell.font   = FB if cell.row == 2 else FN
            cell.border = brd
            if cell.row == 2:
                cell.fill = PatternFill('solid', start_color='D9D9D9')

    # ── Escribir REPORTE ─────────────────────────────────────────────────────
    ws_rep = wb['REPORTE']
    c1, c2, c3, c4 = generar_reporte(df_result)
    escribir_reporte_excel(ws_rep, c1, c2, c3, c4,
                           FW, FN, FB, brd,
                           COLOR_H_ORIG, COLOR_H_NEW, COLOR_H_CFG)

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out


# ── Layout principal ──────────────────────────────────────────────────────────
col_upload, col_main = st.columns([1, 2.5], gap="large")

with col_upload:
    st.markdown('<p class="section-title">Archivo de entrada</p>', unsafe_allow_html=True)
    uploaded = st.file_uploader("Sube tu archivo Excel", type=['xlsx','xls'],
                                label_visibility="collapsed")

    if uploaded:
        df_raw, header_row = leer_excel(uploaded)
        if df_raw is None:
            st.error("No se encontraron los encabezados. Verifica el archivo.")
        else:
            tiene_color    = 'COLOR_A'   in df_raw.columns
            tiene_plan_ins = 'PLAN_INS_DIA1'  in df_raw.columns

            st.success(f"**{uploaded.name}**")
            st.caption(
                f"{len(df_raw):,} filas · {df_raw['DISPO'].nunique():,} dispos · "
                f"fila encabezado {header_row + 1}"
            )

            # Avisos de columnas opcionales
            if not tiene_color:
                st.markdown(
                    '<div class="warn-box">⚠️ Sin columna COLOR_A — '
                    'todas las líneas serán ACTIVAS.</div>', unsafe_allow_html=True
                )
            if usar_plan_ins and not tiene_plan_ins:
                st.markdown(
                    '<div class="warn-box">⚠️ Activaste PLAN_INS_DIA1 pero el archivo '
                    'no tiene esa columna — se usará solo INV.</div>', unsafe_allow_html=True
                )

            cols_faltantes = [c for c in ['ENTREGA','DISPO','ESTILO_EQ','DTITULAR','LBS_C','INV']
                              if c not in df_raw.columns]
            if cols_faltantes:
                st.warning(f"Columnas faltantes: {cols_faltantes}")
            else:
                preview_cols = [c for c in ['DISPO','ENTREGA','ESTILO_EQ','DTITULAR',
                                            'COLOR_A','LBS_C','INV','PLAN_INS_DIA1']
                                if c in df_raw.columns]
                st.markdown('<p class="section-title">Vista previa</p>', unsafe_allow_html=True)
                st.dataframe(df_raw[preview_cols].head(8),
                             use_container_width=True, hide_index=True)

                if st.button("▶ Procesar", type="primary", use_container_width=True):
                    with st.spinner("Calculando asignación..."):
                        st.session_state['df_result'] = procesar(
                            df_raw, entrega_pesos, lotsize_pesos,
                            modo, cascada_activo, usar_plan_ins
                        )
                        st.session_state['modo']         = modo
                        st.session_state['usar_plan_ins'] = usar_plan_ins

with col_main:
    if 'df_result' in st.session_state:
        df_r    = st.session_state['df_result']
        activas = df_r[df_r['ACTIVO'] == 1]

        dispo_min  = activas.groupby('DISPO')['PCT_LINEA'].min()
        total_d    = dispo_min.count()
        completas  = (dispo_min >= 1).sum()
        parciales  = ((dispo_min > 0) & (dispo_min < 1)).sum()
        sin_inv    = (dispo_min == 0).sum()
        inac_lines = (df_r['ACTIVO'] == 0).sum()
        cob        = activas['LBS_ASIGNADO'].sum() / activas['LBS_C'].sum() if len(activas) else 0
        inv_label  = 'INV + PLAN_INS_DIA1' if st.session_state.get('usar_plan_ins') else 'Solo INV'

        st.markdown('<p class="section-title">Resumen de asignación — por DISPO (líneas activas)</p>',
                    unsafe_allow_html=True)
        st.markdown(f"""
        <div class="kpi-grid">
            <div class="kpi-card green">
                <div class="kpi-label">✅ Completas</div>
                <div class="kpi-value">{completas:,}</div>
                <div class="kpi-sub">{completas/total_d:.1%} de {total_d:,} dispos</div>
            </div>
            <div class="kpi-card yellow">
                <div class="kpi-label">⚠️ Parciales</div>
                <div class="kpi-value">{parciales:,}</div>
                <div class="kpi-sub">{parciales/total_d:.1%} de {total_d:,} dispos</div>
            </div>
            <div class="kpi-card red">
                <div class="kpi-label">❌ Sin inventario</div>
                <div class="kpi-value">{sin_inv:,}</div>
                <div class="kpi-sub">{sin_inv/total_d:.1%} de {total_d:,} dispos</div>
            </div>
            <div class="kpi-card blue">
                <div class="kpi-label">📦 Cobertura LBS</div>
                <div class="kpi-value">{cob:.1%}</div>
                <div class="kpi-sub">{inv_label} · Modo: {st.session_state['modo']}</div>
            </div>
            <div class="kpi-card gray">
                <div class="kpi-label">⛔ Líneas inactivas</div>
                <div class="kpi-value">{inac_lines:,}</div>
                <div class="kpi-sub">excluidas de asignación</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        st.markdown('<p class="section-title">Detalle de resultados</p>', unsafe_allow_html=True)

        fc1, fc2, fc3, fc4 = st.columns(4)
        with fc1:
            filtro_status = st.multiselect(
                "Status",
                ['✅ COMPLETA','⚠️ PARCIAL','❌ SIN INVENTARIO','⛔ INACTIVA'],
                default=['✅ COMPLETA','⚠️ PARCIAL','❌ SIN INVENTARIO','⛔ INACTIVA'],
                label_visibility="collapsed"
            )
        with fc2:
            filtro_entrega = st.multiselect(
                "ENTREGA", sorted(df_r['ENTREGA'].dropna().unique()),
                default=sorted(df_r['ENTREGA'].dropna().unique()),
                label_visibility="collapsed"
            )
        with fc3:
            if 'LOTSIZE' in df_r.columns:
                filtro_lotsize = st.multiselect(
                    "LOTSIZE", sorted(df_r['LOTSIZE'].dropna().unique()),
                    default=sorted(df_r['LOTSIZE'].dropna().unique()),
                    label_visibility="collapsed"
                )
            else:
                filtro_lotsize = None
        with fc4:
            if 'COLOR_NOMBRE' in df_r.columns:
                filtro_color = st.multiselect(
                    "Color", sorted(df_r['COLOR_NOMBRE'].dropna().unique()),
                    default=sorted(df_r['COLOR_NOMBRE'].dropna().unique()),
                    label_visibility="collapsed"
                )
            else:
                filtro_color = None

        df_vis = df_r[
            df_r['STATUS_DISPO'].isin(filtro_status) &
            df_r['ENTREGA'].isin(filtro_entrega)
        ]
        if filtro_lotsize and 'LOTSIZE' in df_r.columns:
            df_vis = df_vis[df_vis['LOTSIZE'].isin(filtro_lotsize)]
        if filtro_color and 'COLOR_NOMBRE' in df_r.columns:
            df_vis = df_vis[df_vis['COLOR_NOMBRE'].isin(filtro_color)]

        cols_show = ['DISPO','ENTREGA','ESTILO_EQ','DTITULAR',
                     'COLOR_NOMBRE','ACTIVO',
                     'LBS_C','INV','INV_EFECTIVO',
                     'LBS_ASIGNADO','LBS_FALTANTE',
                     'PCT_LINEA','PCT_DISPO','STATUS_DISPO']
        if 'LOTSIZE' in df_r.columns:
            cols_show.insert(2, 'LOTSIZE')
        cols_show = [c for c in cols_show if c in df_r.columns]

        st.dataframe(
            df_vis[cols_show].style.format({
                'PCT_LINEA':    '{:.1%}',
                'PCT_DISPO':    '{:.1%}',
                'LBS_C':        '{:,.1f}',
                'LBS_ASIGNADO': '{:,.1f}',
                'LBS_FALTANTE': '{:,.1f}',
                'INV_EFECTIVO': '{:,.0f}',
            }),
            use_container_width=True, hide_index=True, height=420
        )
        st.caption(f"Mostrando {len(df_vis):,} líneas · {df_vis['DISPO'].nunique():,} dispos")

        st.markdown('<p class="section-title">Descargar resultado</p>', unsafe_allow_html=True)
        ts = datetime.now().strftime('%Y%m%d%H%M')
        excel_bytes = generar_excel(
            df_r, entrega_pesos, lotsize_pesos,
            cascada_activo, modo,
            st.session_state.get('usar_plan_ins', False)
        )
        st.download_button(
            label=f"⬇️  Descargar Excel — INVENTARIO_DISPOS_{ts}.xlsx",
            data=excel_bytes,
            file_name=f"INVENTARIO_DISPOS_{ts}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    else:
        st.markdown("""
        <div style="display:flex;align-items:center;justify-content:center;height:400px;
                    border:2px dashed #d1d5db;border-radius:12px;color:#9ca3af;
                    flex-direction:column;gap:0.5rem;">
            <div style="font-size:2.5rem">📂</div>
            <div style="font-weight:600;font-size:1rem;">Sube tu archivo y presiona Procesar</div>
            <div style="font-size:0.82rem;">Los resultados aparecerán aquí</div>
        </div>
        """, unsafe_allow_html=True)
