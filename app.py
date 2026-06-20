import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import FormulaRule
import io
import json

try:
    import plotly.express as px
    import plotly.graph_objects as go
    PLOTLY_OK = True
except Exception:
    PLOTLY_OK = False

st.set_page_config(page_title="Asignación Inventario — ABASTO/CUOTA", page_icon="📦", layout="wide")

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600&family=DM+Mono:wght@400;500&display=swap');

    html, body, [class*="css"], .stMarkdown, .stText, p, div {
        font-family: 'DM Sans', sans-serif !important;
        font-size: 12px !important;
    }
    .block-container { padding: 1.2rem 2rem 2rem; max-width: 1500px; }
    section[data-testid="stSidebar"] { background: #0f172a; }
    section[data-testid="stSidebar"] * { color: #94a3b8 !important; font-size: 11px !important; }
    section[data-testid="stSidebar"] .section-title { color: #475569 !important; border-color: #1e293b !important; }
    section[data-testid="stSidebar"] input, section[data-testid="stSidebar"] select { background: #1e293b !important; color: #e2e8f0 !important; border-color: #334155 !important; font-size: 11px !important; }
    section[data-testid="stSidebar"] strong { color: #cbd5e1 !important; }
    section[data-testid="stSidebar"] .stSelectbox label { color: #64748b !important; }

    .app-header {
        background: #0f172a; border-bottom: 1px solid #1e293b; border-radius: 8px;
        padding: 1.1rem 1.6rem; margin-bottom: 1.2rem; display: flex; align-items: center; gap: 1.2rem;
    }
    .app-header-icon { width: 36px; height: 36px; background: #1d4ed8; border-radius: 8px;
        display: flex; align-items: center; justify-content: center; font-size: 1rem; flex-shrink: 0; }
    .app-header h1 { margin: 0; font-size: 0.95rem; font-weight: 600; color: #f1f5f9; letter-spacing: -0.01em; }
    .app-header p  { margin: 0.15rem 0 0; font-size: 0.68rem; color: #64748b; letter-spacing: 0.02em; }

    .section-title { font-family: 'DM Mono', monospace !important; font-size: 0.6rem !important; font-weight: 500;
        letter-spacing: 0.14em; text-transform: uppercase; color: #94a3b8; margin: 1.2rem 0 0.5rem;
        padding-bottom: 0.3rem; border-bottom: 1px solid #e2e8f0; }

    .sc-card { border-radius: 6px; padding: 0.85rem 1rem; border: 1px solid #e2e8f0; background: #fff; }
    .sc-card-blue   { border-left: 3px solid #1d4ed8; background: #f8faff; }
    .sc-card-green  { border-left: 3px solid #059669; background: #f6fefa; }
    .sc-card-purple { border-left: 3px solid #7c3aed; background: #faf7ff; }
    .sc-label { font-family: 'DM Mono', monospace; font-size: 0.6rem; font-weight: 500;
                letter-spacing: 0.1em; text-transform: uppercase; color: #94a3b8; margin-bottom: 0.6rem; }
    .sc-row { display: flex; justify-content: space-between; align-items: center;
              padding: 0.2rem 0; border-bottom: 1px solid #f1f5f9; }
    .sc-row:last-child { border-bottom: none; padding-top: 0.35rem; margin-top: 0.2rem; }
    .sc-metric { font-size: 0.7rem; color: #64748b; }
    .sc-value  { font-size: 0.75rem; font-weight: 600; color: #0f172a; font-family: 'DM Mono', monospace; }
    .sc-divider { border-top: 1px solid #e2e8f0; margin-top: 0.4rem; padding-top: 0.4rem; }

    div[data-testid="stNumberInput"] label,
    div[data-testid="stSelectbox"] label { font-size: 0.68rem !important; color: #64748b; }
    .stButton > button { font-size: 0.72rem !important; font-weight: 600; border-radius: 5px;
        padding: 0.4rem 1rem; border: none; letter-spacing: 0.02em; }
    .stButton > button[kind="primary"] { background: #1d4ed8; color: white; }
    .stButton > button[kind="primary"]:hover { background: #1e40af; }

    .tip-box  { background: #f8fafc; border-left: 2px solid #3b82f6; border-radius: 0 4px 4px 0;
                padding: 0.5rem 0.75rem; font-size: 0.68rem; color: #475569; margin: 0.5rem 0; }
    .warn-box { background: #fffbeb; border-left: 2px solid #f59e0b; border-radius: 0 4px 4px 0;
                padding: 0.5rem 0.75rem; font-size: 0.68rem; color: #78350f; margin: 0.5rem 0; }
    .evitar-box { background: #fef2f2; border-left: 2px solid #dc2626; border-radius: 0 4px 4px 0;
                padding: 0.4rem 0.6rem; font-size: 0.65rem; color: #7f1d1d; margin: 0.3rem 0; }

    .stDataFrame, .stDataFrame td, .stDataFrame th { font-size: 11px !important; }
    .stTabs [data-baseweb="tab"] { font-size: 0.7rem !important; font-weight: 500; padding: 0.4rem 0.8rem; }
    .stTabs [aria-selected="true"] { font-weight: 600 !important; }
    .stCaption, [data-testid="stCaptionContainer"] { font-size: 0.65rem !important; color: #94a3b8; }
    [data-testid="stFileUploader"] { font-size: 0.7rem !important; }
    [data-testid="stFileUploader"] label { font-size: 0.7rem !important; }
    .stAlert { font-size: 0.7rem !important; padding: 0.4rem 0.75rem !important; }

    .step-header { display: flex; align-items: center; gap: 0.6rem; margin: 1.6rem 0 0.8rem; }
    .step-badge { width: 28px; height: 28px; border-radius: 50%; background: #1d4ed8; color: white;
        display: flex; align-items: center; justify-content: center; font-weight: 700; font-size: 0.85rem;
        font-family: 'DM Mono', monospace; flex-shrink: 0; }
    .step-title { font-size: 1.05rem !important; font-weight: 700; color: #0f172a; margin: 0; }
    .step-sub { font-size: 0.72rem !important; color: #64748b; margin: 0.1rem 0 0 2.2rem; }

    div[data-testid="stExpander"] { border: 1px solid #e2e8f0 !important; border-radius: 8px !important;
        margin-bottom: 0.6rem; }
    div[data-testid="stExpander"] summary { font-size: 0.78rem !important; font-weight: 600 !important; }
    div[data-testid="stDataEditor"] { font-size: 11px !important; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="app-header">
    <div class="app-header-icon">📦</div>
    <div>
        <h1>Asignación de Inventario v2 — ABASTO → CUOTA</h1>
        <p>INV &nbsp;·&nbsp; INV + PLAN_DIA1 &nbsp;·&nbsp; INV + PLAN_SEMANAL &nbsp;·&nbsp; Fuente: hoja WIP de TIM_WIPCR &nbsp;·&nbsp; Llave: PLANTA_WIP + CONSTRUCCION</p>
    </div>
</div>
""", unsafe_allow_html=True)

# ── Constantes ────────────────────────────────────────────────────────────────
DEFAULT_ENTREGA = {
    '01-EXPEDITE': 1, '02-PAST DUE': 2, '03-DUE': 3, '04-AHEAD': 4, '05-AHEAD': 5,
}

# EVITAR=True → solo se habilita con prioridad normal si ENTREGA=01-EXPEDITE o estamos en Fase ABASTO
# LBS_POR_LOTE = capacidad en libras de UNA corrida/lote de ese LOTSIZE (ej. "A-4000" → 4000 lbs/lote).
# La capacidad semanal real = LOTES (lotes/día) × 7 × LBS_POR_LOTE.
CAPACIDAD_LOTSIZE = [
    {'LOTSIZE': 'A-4000', 'MIX': 'DYE',    'LOTES':  4, 'LBS_POR_LOTE': 4000, 'PRIORIDAD': 1, 'ACTIVO': True,  'EVITAR': False},
    {'LOTSIZE': 'B-3300', 'MIX': 'DYE',    'LOTES':  8, 'LBS_POR_LOTE': 3300, 'PRIORIDAD': 2, 'ACTIVO': True,  'EVITAR': False},
    {'LOTSIZE': 'C-2600', 'MIX': 'DYE',    'LOTES': 38, 'LBS_POR_LOTE': 2600, 'PRIORIDAD': 3, 'ACTIVO': True,  'EVITAR': False},
    {'LOTSIZE': 'D-2200', 'MIX': 'DYE',    'LOTES': 29, 'LBS_POR_LOTE': 2200, 'PRIORIDAD': 4, 'ACTIVO': True,  'EVITAR': False},
    {'LOTSIZE': 'D-2200', 'MIX': 'BLEACH', 'LOTES': 14, 'LBS_POR_LOTE': 2200, 'PRIORIDAD': 1, 'ACTIVO': True,  'EVITAR': False},
    {'LOTSIZE': 'F-1000', 'MIX': 'DYE',    'LOTES': 33, 'LBS_POR_LOTE': 1000, 'PRIORIDAD': 6, 'ACTIVO': True,  'EVITAR': False},
    {'LOTSIZE': 'F-1000', 'MIX': 'BLEACH', 'LOTES':  7, 'LBS_POR_LOTE': 1000, 'PRIORIDAD': 2, 'ACTIVO': True,  'EVITAR': False},
]

CASCADA_COLOR_NOMBRES = {
    '0': 'BLANCO', '1': 'AMARILLO', '2': 'NARANJA', '3': 'ROJO',
    '4': 'MORADO', '5': 'ROYAL',    '6': 'VERDE',   '7': 'CAFÉ',
    '8': 'GRIS',   '9': 'NEGRO',
}

ESCENARIOS = [
    {'key': 'INV',      'label': '📦 Solo INV',           'col_extra': None},
    {'key': 'DIA1',     'label': '📅 INV + PLAN_DIA1',    'col_extra': 'PLAN_INS_DIA1'},
    {'key': 'SEMANAL',  'label': '📆 INV + PLAN_SEMANAL', 'col_extra': 'PLAN_INS'},
]

REQ_WIP_COLS = ['DISPO', 'ENTREGA', 'ESTILO_EQ', 'DTITULAR', 'LBS_C', 'INV',
                'LOTSIZE', 'MIX', 'PLANTA_WIP', 'CONSTRUCCION']

# Columnas crudas de WIP que sí queremos en el Excel — WIP trae ~74 columnas y
# la mayoría (RATE, FORMULA, CLAVE2, etc.) no se usan en esta app. Mandarlas
# todas al Excel multiplica el tamaño de la hoja DETALLE y puede tumbar la app
# en Streamlit Cloud (memoria/tiempo) con datasets grandes.
ESSENTIAL_ORIG_COLS = ['DISPO', 'ENTREGA', 'PLANTA_WIP', 'CONSTRUCCION', 'LOTSIZE', 'MIX',
                        'ESTILO_EQ', 'DTITULAR', 'COLOR_A', 'COMPONENTE', 'LBS_C', 'INV',
                        'PLAN_INS_DIA1', 'PLAN_INS']

UMBRAL_PRINCIPAL_DEFAULT  = 0.95   # % mínimo para considerar COMPLETO un componente PRINCIPAL
UMBRAL_SECUNDARIO_DEFAULT = 0.97   # % mínimo para considerar COMPLETO un componente SECUNDARIO-N

# ── Lectura del workbook TIM_WIPCR ──────────────────────────────────────────────
def leer_tim_wipcr(file):
    """Lee WIP (data completa) y CUOTA (metas ABASTO/CUOTA) del workbook TIM_WIPCR."""
    try:
        xls = pd.ExcelFile(file)
    except Exception as e:
        return None, None, f"No se pudo abrir el archivo: {e}"

    faltantes = {'WIP', 'CUOTA'} - set(xls.sheet_names)
    if faltantes:
        return None, None, f"Faltan hojas en el workbook: {sorted(faltantes)}"

    df_wip = pd.read_excel(xls, sheet_name='WIP')
    df_wip.columns = df_wip.columns.str.strip().str.upper()

    df_cuota = pd.read_excel(xls, sheet_name='CUOTA', header=0)
    cols = list(df_cuota.columns)
    if len(cols) < 2:
        return df_wip, None, "La hoja CUOTA no tiene columnas suficientes."
    cols[0], cols[1] = 'PLANTA_WIP', 'CONSTRUCCION'
    df_cuota.columns = [str(c).strip().upper() if isinstance(c, str) else c for c in cols]

    if 'ABASTO' not in df_cuota.columns or 'CUOTA' not in df_cuota.columns:
        return df_wip, None, "La hoja CUOTA no tiene columnas ABASTO/CUOTA reconocibles (revisa encabezados)."

    df_cuota = df_cuota[['PLANTA_WIP', 'CONSTRUCCION', 'ABASTO', 'CUOTA']].dropna(
        subset=['PLANTA_WIP', 'CONSTRUCCION'], how='any'
    )
    df_cuota['ABASTO'] = pd.to_numeric(df_cuota['ABASTO'], errors='coerce').fillna(0)
    df_cuota['CUOTA']  = pd.to_numeric(df_cuota['CUOTA'],  errors='coerce').fillna(0)

    return df_wip, df_cuota, None


def construir_targets(df_cuota):
    targets = {}
    for _, row in df_cuota.iterrows():
        key = (str(row['PLANTA_WIP']).strip(), str(row['CONSTRUCCION']).strip())
        targets[key] = {'ABASTO': float(row['ABASTO']), 'CUOTA': float(row['CUOTA'])}
    return targets


def validar_huerfanos(df_wip, df_cuota):
    """Detecta combinaciones PLANTA_WIP+CONSTRUCCION sin match cruzado."""
    wip_keys   = set(zip(df_wip['PLANTA_WIP'].astype(str).str.strip(), df_wip['CONSTRUCCION'].astype(str).str.strip()))
    cuota_keys = set(zip(df_cuota['PLANTA_WIP'].astype(str).str.strip(), df_cuota['CONSTRUCCION'].astype(str).str.strip()))
    solo_wip   = wip_keys - cuota_keys     # inventario sin meta definida
    solo_cuota = cuota_keys - wip_keys     # metas sin inventario que las cubra
    return solo_wip, solo_cuota


def validar_lotsize_no_configurados(df_wip, capacidad_cfg):
    """Combinaciones LOTSIZE+MIX presentes en WIP pero ausentes de la config
    — esas líneas quedan en 0 (capacidad inactiva), nunca infinita."""
    cfg_keys = {(str(r['LOTSIZE']), str(r['MIX'])) for r in capacidad_cfg}
    wip = df_wip.copy()
    wip.columns = wip.columns.str.strip().str.upper()
    if not {'LOTSIZE', 'MIX'}.issubset(wip.columns):
        return pd.DataFrame()
    wip['_KEY'] = list(zip(wip['LOTSIZE'].astype(str), wip['MIX'].astype(str)))
    faltan = wip[~wip['_KEY'].isin(cfg_keys)]
    if faltan.empty:
        return pd.DataFrame()
    resumen = faltan.groupby(['LOTSIZE', 'MIX']).agg(
        N_LINEAS=('DISPO', 'count'),
        LBS_C_TOTAL=('LBS_C', 'sum'),
    ).reset_index().sort_values('LBS_C_TOTAL', ascending=False)
    return resumen


def extraer_digito3(color_a):
    try:
        s = str(int(float(color_a))) if str(color_a).replace('.', '').isdigit() else str(color_a)
        return s[2] if len(s) >= 3 else None
    except Exception:
        return None


# ── Motor de asignación en 2 fases (ABASTO → CUOTA) ─────────────────────────────
def motor_asignacion(df_wip, targets, entrega_pesos, cascada_activo, capacidad_cfg,
                      col_extra=None, progress_cb=None,
                      umbral_principal=UMBRAL_PRINCIPAL_DEFAULT, umbral_secundario=UMBRAL_SECUNDARIO_DEFAULT):
    """
    Asigna inventario en 2 fases por combinación PLANTA_WIP+CONSTRUCCION:
      Fase 1: cubrir ABASTO (col I de CUOTA)
      Fase 2: con lo que sobra, cubrir CUOTA (col K de CUOTA), sin redistribuir fase 1.
    Respeta capacidad semanal por LOTSIZE+MIX (lotes_dia*7) y el pool físico de
    inventario por ESTILO_EQ+DTITULAR.

    Al final, evalúa cada DISPO completo: si TODAS sus líneas activas con
    demanda alcanzan su umbral (PRINCIPAL ≥ umbral_principal, cualquier
    SECUNDARIO-N ≥ umbral_secundario), se conserva la asignación. Si alguna
    línea se queda corta, se descarta el DISPO entero (todas sus líneas a 0)
    — "si no se completa, no se asigna".
    """
    df = df_wip.copy()
    df.columns = df.columns.str.strip().str.upper()
    req = [c for c in REQ_WIP_COLS if c in df.columns]
    df = df.dropna(subset=req).reset_index(drop=True)

    if col_extra and col_extra in df.columns:
        df['INV_EFECTIVO'] = df['INV'].fillna(0) + df[col_extra].fillna(0)
    else:
        df['INV_EFECTIVO'] = df['INV'].fillna(0)

    if 'COLOR_A' in df.columns:
        df['_DIGITO3']     = df['COLOR_A'].apply(extraer_digito3)
        df['COLOR_NOMBRE'] = df['_DIGITO3'].map(CASCADA_COLOR_NOMBRES).fillna('DESCONOCIDO')
        df['ACTIVO']       = df['_DIGITO3'].map(cascada_activo).fillna(0).astype(int)
    else:
        df['_DIGITO3']     = None
        df['COLOR_NOMBRE'] = 'SIN COLOR_A'
        df['ACTIVO']       = 1

    peso_e       = df['ENTREGA'].map(entrega_pesos).fillna(99).to_numpy(dtype=float)
    es_expedite  = (df['ENTREGA'] == '01-EXPEDITE').to_numpy()
    activo_color = df['ACTIVO'].to_numpy()
    lbs_c        = df['LBS_C'].fillna(0).to_numpy(dtype=float)

    if progress_cb: progress_cb(0.1, "Preparando capacidad por LOTSIZE+MIX (semanal = diaria × 7 × lbs/lote)…")

    # ── Capacidad LOTSIZE+MIX (factorizada) — la capacidad se mide y consume EN LIBRAS ──
    mach_tuple_raw = list(zip(df['LOTSIZE'].astype(str), df['MIX'].astype(str)))
    mach_str = np.array([f"{a}||{b}" for a, b in mach_tuple_raw], dtype=object)
    mach_codes, mach_keys_str = pd.factorize(mach_str)
    mach_keys = [tuple(k.split('||', 1)) for k in mach_keys_str]
    n_mach = len(mach_keys)
    cfg_by_key = {(str(r['LOTSIZE']), str(r['MIX'])): r for r in capacidad_cfg}

    cap_lbs_arr   = np.zeros(n_mach)   # capacidad semanal en LIBRAS = lotes×7×lbs_por_lote
    ls_activo_arr = np.zeros(n_mach, dtype=bool)
    prioridad_arr = np.full(n_mach, 99.0)
    evitar_arr    = np.zeros(n_mach, dtype=bool)
    combos_sin_config = set()
    for c in range(n_mach):
        cfg = cfg_by_key.get(mach_keys[c])
        if cfg is None:
            # LOTSIZE+MIX que no está en la configuración → NO se asigna
            # (capacidad 0 / inactiva), nunca capacidad infinita.
            cap_lbs_arr[c]   = 0.0
            ls_activo_arr[c] = False
            combos_sin_config.add(mach_keys[c])
        else:
            lbs_por_lote     = float(cfg.get('LBS_POR_LOTE', 0)) or 0.0
            cap_lbs_arr[c]   = float(cfg['LOTES']) * 7.0 * lbs_por_lote
            ls_activo_arr[c] = bool(cfg['ACTIVO'])
            prioridad_arr[c] = float(cfg['PRIORIDAD'])
            evitar_arr[c]    = bool(cfg.get('EVITAR', False))
    cap_lbs_usado = np.zeros(n_mach)   # libras ya consumidas de esa capacidad (acumulado fase 1 + fase 2)

    # ── Pool físico de inventario (ESTILO_EQ + DTITULAR) ──
    pool_str = np.array([f"{a}||{b}" for a, b in zip(df['ESTILO_EQ'].astype(str), df['DTITULAR'].astype(str))], dtype=object)
    pool_codes, pool_keys_str = pd.factorize(pool_str)
    n_pools = len(pool_keys_str)
    inv_total = np.zeros(n_pools)
    seen = np.zeros(n_pools, dtype=bool)
    inv_arr = df['INV_EFECTIVO'].to_numpy(dtype=float)
    for i in range(len(df)):
        c = pool_codes[i]
        if not seen[c]:
            inv_total[c] = inv_arr[i]
            seen[c] = True
    inv_used = np.zeros(n_pools)

    # ── Buckets de meta (PLANTA_WIP + CONSTRUCCION) ──
    bucket_tuple_raw = list(zip(df['PLANTA_WIP'].astype(str).str.strip(), df['CONSTRUCCION'].astype(str).str.strip()))
    bucket_str = np.array([f"{a}||{b}" for a, b in bucket_tuple_raw], dtype=object)
    bucket_codes, bucket_keys_str = pd.factorize(bucket_str)
    bucket_keys = [tuple(k.split('||', 1)) for k in bucket_keys_str]
    n_buckets = len(bucket_keys)
    abasto_meta = np.zeros(n_buckets)
    cuota_meta  = np.zeros(n_buckets)
    for c in range(n_buckets):
        t = targets.get(bucket_keys[c], {})
        abasto_meta[c] = t.get('ABASTO', 0.0) or 0.0
        cuota_meta[c]  = t.get('CUOTA', 0.0) or 0.0
    target_asignado = np.zeros(n_buckets)

    ls_activo_row = ls_activo_arr[mach_codes]

    # ── Orden de prioridad: ENTREGA del DISPO → agrupar TODO el DISPO junto →
    # preferencia de máquina dentro del DISPO. Esto hace que el motor complete
    # un pedido (todas sus líneas de color) antes de tocar el siguiente, en vez
    # de repartir el pool en migajas entre muchos DISPOs a la vez.
    dispo_codes, _ = pd.factorize(df['DISPO'].astype(str), sort=True)
    peso_e_dispo = pd.Series(peso_e, index=df.index).groupby(df['DISPO']).transform('min').to_numpy()

    orden_f1 = peso_e_dispo * 1e8 + dispo_codes * 1000 + prioridad_arr[mach_codes]               # Fase ABASTO: evitar SIEMPRE override
    orden_f2 = peso_e_dispo * 1e8 + dispo_codes * 1000 + (prioridad_arr[mach_codes] +
               np.where(evitar_arr[mach_codes] & ~es_expedite, 90, 0))  # Fase CUOTA: solo override si EXPEDITE

    order1 = np.argsort(orden_f1, kind='stable')
    order2 = np.argsort(orden_f2, kind='stable')

    n = len(df)
    asignado = np.zeros(n)          # asignación FINAL — solo de DISPOs ya "cerrados" (completos)
    locked   = np.zeros(n, dtype=bool)   # True = línea de un DISPO ya cerrado (éxito), no se vuelve a tocar
    alguna_vez_intento = np.zeros(n, dtype=bool)

    MAX_ITER = 6
    dispo_series = df['DISPO']

    if 'COMPONENTE' in df.columns:
        componente_arr = df['COMPONENTE'].astype(str).str.upper().str.strip().to_numpy()
    else:
        componente_arr = np.full(len(df), 'PRINCIPAL')
    es_principal = np.char.find(componente_arr.astype(str), 'PRINCIPAL') >= 0
    umbral_arr   = np.where(es_principal, umbral_principal, umbral_secundario)

    for iteracion in range(MAX_ITER):
        candidatos = ~locked
        if not candidatos.any():
            break

        if progress_cb:
            progress_cb(0.3 + 0.5 * iteracion / MAX_ITER,
                        f"Ronda {iteracion + 1}/{MAX_ITER} — completando DISPOs (todo o nada) y liberando lo descartado…")

        intento = np.zeros(n)

        # Fase 1 (ABASTO) — solo sobre líneas candidatas
        for i in order1:
            if not candidatos[i] or activo_color[i] == 0 or not ls_activo_row[i]:
                continue
            b = bucket_codes[i]
            falta_target = abasto_meta[b] - target_asignado[b]
            if falta_target <= 0:
                continue
            m = mach_codes[i]
            cap_rem = cap_lbs_arr[m] - cap_lbs_usado[m]
            if cap_rem <= 0:
                continue
            p = pool_codes[i]
            inv_rem = inv_total[p] - inv_used[p]
            if inv_rem <= 0:
                continue
            cant = min(lbs_c[i], falta_target, inv_rem, cap_rem)
            if cant <= 0:
                continue
            intento[i]          += cant
            inv_used[p]          += cant
            target_asignado[b]   += cant
            cap_lbs_usado[m]     += cant

        # Fase 2 (CUOTA) — solo sobre líneas candidatas
        for i in order2:
            if not candidatos[i] or activo_color[i] == 0 or not ls_activo_row[i]:
                continue
            restante_linea = lbs_c[i] - intento[i]
            if restante_linea <= 0:
                continue
            b = bucket_codes[i]
            falta_target = cuota_meta[b] - target_asignado[b]
            if falta_target <= 0:
                continue
            m = mach_codes[i]
            cap_rem = cap_lbs_arr[m] - cap_lbs_usado[m]
            if cap_rem <= 0:
                continue
            p = pool_codes[i]
            inv_rem = inv_total[p] - inv_used[p]
            if inv_rem <= 0:
                continue
            cant = min(restante_linea, falta_target, inv_rem, cap_rem)
            if cant <= 0:
                continue
            intento[i]          += cant
            inv_used[p]          += cant
            target_asignado[b]   += cant
            cap_lbs_usado[m]     += cant

        alguna_vez_intento |= (intento > 0)

        # ── Evaluar umbral por DISPO (todo o nada) sobre esta ronda ──────────
        pct_tmp = np.divide(intento, lbs_c, out=np.ones_like(intento), where=lbs_c > 0)
        bloquea = candidatos & (lbs_c > 0) & (activo_color == 1) & (pct_tmp < umbral_arr)
        dispo_bloquea    = pd.Series(bloquea, index=df.index).groupby(dispo_series).transform('any').to_numpy()
        dispo_es_cand    = pd.Series(candidatos, index=df.index).groupby(dispo_series).transform('any').to_numpy()
        completa_ronda   = dispo_es_cand & ~dispo_bloquea

        idx_lock   = np.where(completa_ronda)[0]
        idx_revert = np.where(candidatos & ~completa_ronda)[0]

        asignado[idx_lock] = intento[idx_lock]
        locked[idx_lock]   = True

        if len(idx_revert):
            # Devolver lo consumido por los DISPOs que NO se completaron esta
            # ronda, para que la siguiente ronda pueda usarlo en otros DISPOs.
            np.subtract.at(inv_used,         pool_codes[idx_revert],   intento[idx_revert])
            np.subtract.at(cap_lbs_usado,    mach_codes[idx_revert],   intento[idx_revert])
            np.subtract.at(target_asignado,  bucket_codes[idx_revert], intento[idx_revert])

        if len(idx_lock) == 0:
            break   # sin progreso esta ronda — lo que quede candidato se queda en 0

    if progress_cb: progress_cb(0.85, "Umbral de completitud aplicado…")

    dispo_tenia_algo = pd.Series(alguna_vez_intento, index=df.index).groupby(dispo_series).transform('any').to_numpy()
    df['DISPO_DESCARTADA'] = (~locked) & dispo_tenia_algo

    df['LBS_ASIGNADO'] = asignado
    df['LBS_FALTANTE'] = (df['LBS_C'] - df['LBS_ASIGNADO']).clip(lower=0)
    df['PCT_LINEA']    = (df['LBS_ASIGNADO'] / df['LBS_C'].replace(0, np.nan)).fillna(0)

    df['ABASTO_META']      = abasto_meta[bucket_codes]
    df['CUOTA_META']       = cuota_meta[bucket_codes]
    bucket_asig             = target_asignado[bucket_codes]
    df['ABASTO_CUBIERTO']  = np.minimum(bucket_asig, df['ABASTO_META'])
    df['CUOTA_CUBIERTA']   = bucket_asig
    df['PCT_ABASTO']       = np.where(df['ABASTO_META'] > 0, df['ABASTO_CUBIERTO'] / df['ABASTO_META'], 1.0)
    df['PCT_CUOTA']        = np.where(df['CUOTA_META']  > 0, df['CUOTA_CUBIERTA']  / df['CUOTA_META'],  1.0)
    df['LBS_FALTA_ABASTO'] = (df['ABASTO_META'] - df['ABASTO_CUBIERTO']).clip(lower=0)
    df['LBS_FALTA_CUOTA']  = (df['CUOTA_META']  - df['CUOTA_CUBIERTA']).clip(lower=0)

    cond_full = (df['PCT_ABASTO'] >= 1) & (df['PCT_CUOTA'] >= 1)
    cond_aba  = (df['PCT_ABASTO'] >= 1)
    cond_sin_meta = (df['ABASTO_META'] <= 0) & (df['CUOTA_META'] <= 0)
    df['STATUS_AC'] = np.select(
        [cond_sin_meta, cond_full, cond_aba],
        ['⚪ SIN META', '✅ ABASTO+CUOTA', '🟡 SOLO ABASTO'],
        default='🔴 NI ABASTO'
    )

    # Status a nivel línea/DISPO — ahora refleja el umbral relajado (95%/97%)
    # y distingue "se descartó por no llegar al umbral" de "nunca tuvo nada".
    activas_mask = df['ACTIVO'] == 1
    pct_min = df[activas_mask].groupby('DISPO')['PCT_LINEA'].min().rename('_min_pct')
    df = df.merge(pct_min, on='DISPO', how='left')
    df['_min_pct'] = df['_min_pct'].fillna(-1)

    df['STATUS_DISPO'] = np.select(
        [df['ACTIVO'] == 0, df['DISPO_DESCARTADA'], df['_min_pct'] > 0],
        ['⛔ INACTIVA', '🔁 DESCARTADA (no llegó al umbral)', '✅ COMPLETA'],
        default='❌ SIN INVENTARIO'
    )

    pct_dispo = (
        df[activas_mask].groupby('DISPO')['LBS_ASIGNADO'].sum() /
        df[activas_mask].groupby('DISPO')['LBS_C'].sum().replace(0, np.nan)
    ).fillna(0).rename('PCT_DISPO')
    df = df.merge(pct_dispo, on='DISPO', how='left')
    df['PCT_DISPO'] = df['PCT_DISPO'].fillna(0)

    if progress_cb: progress_cb(1.0, "Listo.")

    drop = ['_DIGITO3', '_min_pct']
    return df.drop(columns=[c for c in drop if c in df.columns])


def resumen_buckets(df_r):
    """Tabla resumen 1 fila por PLANTA_WIP+CONSTRUCCION para gráficos y export."""
    g = df_r.groupby(['PLANTA_WIP', 'CONSTRUCCION']).agg(
        ABASTO_META=('ABASTO_META', 'first'),
        CUOTA_META=('CUOTA_META', 'first'),
        ASIGNADO=('CUOTA_CUBIERTA', 'first'),
        LBS_C_TOTAL=('LBS_C', 'sum'),
    ).reset_index()
    g['ABASTO_CUBIERTO'] = np.minimum(g['ASIGNADO'], g['ABASTO_META'])
    g['PCT_ABASTO'] = np.where(g['ABASTO_META'] > 0, g['ABASTO_CUBIERTO'] / g['ABASTO_META'], 1.0)
    g['PCT_CUOTA']  = np.where(g['CUOTA_META']  > 0, g['ASIGNADO']        / g['CUOTA_META'],  1.0)
    g['LBS_FALTA_ABASTO'] = (g['ABASTO_META'] - g['ABASTO_CUBIERTO']).clip(lower=0)
    g['LBS_FALTA_CUOTA']  = (g['CUOTA_META']  - g['ASIGNADO']).clip(lower=0)
    cond_full = (g['PCT_ABASTO'] >= 1) & (g['PCT_CUOTA'] >= 1)
    cond_aba  = (g['PCT_ABASTO'] >= 1)
    cond_sin  = (g['ABASTO_META'] <= 0) & (g['CUOTA_META'] <= 0)
    g['STATUS'] = np.select([cond_sin, cond_full, cond_aba],
                             ['⚪ SIN META', '✅ ABASTO+CUOTA', '🟡 SOLO ABASTO'],
                             default='🔴 NI ABASTO')
    return g.sort_values('CUOTA_META', ascending=False)


def resumen_capacidad(df_r, capacidad_cfg):
    """Consumo de capacidad semanal (en LIBRAS) por LOTSIZE+MIX vs límite."""
    cap_df = pd.DataFrame(capacidad_cfg).copy()
    cap_df['CAPACIDAD_SEMANAL'] = cap_df['LOTES'] * 7 * cap_df.get('LBS_POR_LOTE', 0)
    activos = df_r[(df_r['ACTIVO'] == 1) & (df_r['LBS_ASIGNADO'] > 0)]
    if {'LOTSIZE', 'MIX'}.issubset(activos.columns):
        usados = activos.groupby(['LOTSIZE', 'MIX'])['LBS_ASIGNADO'].sum().rename('LBS_USADOS').reset_index()
        cap_df = cap_df.merge(usados, on=['LOTSIZE', 'MIX'], how='left')
    cap_df['LBS_USADOS'] = cap_df.get('LBS_USADOS', 0)
    cap_df['LBS_USADOS'] = cap_df['LBS_USADOS'].fillna(0)
    cap_df['LOTES_EQUIV_USADOS'] = np.where(cap_df['LBS_POR_LOTE'] > 0,
                                             cap_df['LBS_USADOS'] / cap_df['LBS_POR_LOTE'], 0)
    cap_df['PCT_USO'] = np.where(cap_df['CAPACIDAD_SEMANAL'] > 0,
                                  cap_df['LBS_USADOS'] / cap_df['CAPACIDAD_SEMANAL'], 0)
    return cap_df


def resumen_inventario_libre(df_r):
    """Por ESTILO_EQ + DTITULAR: cuánto del pool físico (INV + plan del
    escenario) quedó SIN asignar — inventario/plan libre que no se usó."""
    g = df_r.groupby(['ESTILO_EQ', 'DTITULAR']).agg(
        INV_EFECTIVO=('INV_EFECTIVO', 'first'),
        LBS_C_TOTAL=('LBS_C', 'sum'),
        LBS_ASIGNADO_TOTAL=('LBS_ASIGNADO', 'sum'),
        N_LINEAS=('DISPO', 'count'),
    ).reset_index()
    g['INV_LIBRE']   = (g['INV_EFECTIVO'] - g['LBS_ASIGNADO_TOTAL']).clip(lower=0)
    g['PCT_USO_POOL'] = np.where(g['INV_EFECTIVO'] > 0, g['LBS_ASIGNADO_TOTAL'] / g['INV_EFECTIVO'], 0)
    return g.sort_values('INV_LIBRE', ascending=False)


def diagnostico_asignacion(df_r, capacidad_cfg):
    """
    Para cada línea con PCT_LINEA < 1, identifica la razón principal por la
    que no se asignó el 100%. Cruza las tablas que ya existen (capacidad,
    inventario libre, metas) en vez de recalcular nada nuevo.

    Devuelve (df_detalle_con_razon, resumen_lbs_por_razon).
    """
    df = df_r.copy()
    df['LBS_NO_ASIGNADO'] = (df['LBS_C'] - df['LBS_ASIGNADO']).clip(lower=0)

    cap_res = resumen_capacidad(df_r, capacidad_cfg)[['LOTSIZE', 'MIX', 'PCT_USO']]
    cap_res = cap_res.groupby(['LOTSIZE', 'MIX'], as_index=False)['PCT_USO'].max()
    df = df.merge(cap_res.rename(columns={'PCT_USO': '_CAP_PCT_USO'}), on=['LOTSIZE', 'MIX'], how='left')
    df['_CAP_PCT_USO'] = df['_CAP_PCT_USO'].fillna(0)

    libre = resumen_inventario_libre(df_r)[['ESTILO_EQ', 'DTITULAR', 'INV_LIBRE']]
    df = df.merge(libre.rename(columns={'INV_LIBRE': '_POOL_LIBRE'}), on=['ESTILO_EQ', 'DTITULAR'], how='left')
    df['_POOL_LIBRE'] = df['_POOL_LIBRE'].fillna(0)

    cfg_act = {(str(r['LOTSIZE']), str(r['MIX'])) for r in capacidad_cfg if r.get('ACTIVO', True)}
    df['_LS_OK'] = list(zip(df['LOTSIZE'].astype(str), df['MIX'].astype(str)))
    df['_LS_OK'] = df['_LS_OK'].isin(cfg_act)

    condiciones = [
        df['PCT_LINEA'] >= 0.999,
        df['ACTIVO'] == 0,
        ~df['_LS_OK'],
        (df['ABASTO_META'] <= 0) & (df['CUOTA_META'] <= 0),
        df['PCT_CUOTA'] >= 1,
        df['_CAP_PCT_USO'] >= 0.999,
        df['_POOL_LIBRE'] <= 0,
        df['STATUS_DISPO'] == '❌ SIN INVENTARIO',
        df['STATUS_DISPO'] == '🔁 DESCARTADA (no llegó al umbral)',
    ]
    etiquetas = [
        '✅ COMPLETA',
        '🎨 COLOR INACTIVO (cascada)',
        '⚙️ LOTSIZE+MIX NO CONFIGURADO/INACTIVO',
        '🚫 SIN META EN CUOTA (huérfano — falta en hoja CUOTA)',
        '🎯 META CUOTA YA CUBIERTA (bucket lleno, no es falla)',
        '🏭 CAPACIDAD LOTSIZE+MIX SATURADA',
        '📦 INVENTARIO/PLAN AGOTADO (pool ESTILO_EQ+DTITULAR)',
        '⏳ NUNCA RECIBIÓ TURNO (DISPOs de mayor prioridad agotaron pool/capacidad antes)',
        '🔁 DISPO DESCARTADA (otra línea/componente no llegó al umbral)',
    ]
    df['RAZON_NO_ASIGNADO'] = np.select(condiciones, etiquetas, default='❓ SIN CLASIFICAR')

    resumen = (df[df['RAZON_NO_ASIGNADO'] != '✅ COMPLETA']
               .groupby('RAZON_NO_ASIGNADO')
               .agg(LBS=('LBS_NO_ASIGNADO', 'sum'), N_LINEAS=('LBS_NO_ASIGNADO', 'count'))
               .sort_values('LBS', ascending=False).reset_index())

    # Extra: cuántos pools distintos necesita cada DISPO (ayuda a explicar
    # la categoría "NUNCA RECIBIÓ TURNO" — entre más pools necesita un DISPO,
    # más difícil que todos coincidan libres al mismo tiempo).
    n_pools_dispo = df_r.groupby('DISPO').apply(
        lambda g: g[['ESTILO_EQ', 'DTITULAR']].drop_duplicates().shape[0]
    ).rename('N_POOLS_DISPO')
    df = df.merge(n_pools_dispo, on='DISPO', how='left')

    drop = [c for c in df.columns if c.startswith('_')]
    return df.drop(columns=drop), resumen


def col_cfg(df, pct_cols=None, num_cols=None, num_fmt='%.1f'):
    """Construye st.column_config solo para columnas presentes — evita el
    Styler de pandas (limitado en celdas y frágil con dataframes grandes)."""
    pct_cols = pct_cols or []
    num_cols = num_cols or []
    cfg = {}
    for c in pct_cols:
        if c in df.columns:
            cfg[c] = st.column_config.NumberColumn(c, format="percent")
    for c in num_cols:
        if c in df.columns:
            cfg[c] = st.column_config.NumberColumn(c, format=num_fmt)
    return cfg


def kpis(df_r):
    activas = df_r[df_r['ACTIVO'] == 1]
    status_por_dispo = activas.drop_duplicates('DISPO')['STATUS_DISPO']
    total_d = status_por_dispo.shape[0]
    buckets = resumen_buckets(df_r)
    return {
        'total':      total_d,
        'completas':  (status_por_dispo == '✅ COMPLETA').sum(),
        'descartadas': (status_por_dispo == '🔁 DESCARTADA (no llegó al umbral)').sum(),
        'sin_inv':    (status_por_dispo == '❌ SIN INVENTARIO').sum(),
        'inactivas':  (df_r['ACTIVO'] == 0).sum(),
        'cob':        activas['LBS_ASIGNADO'].sum() / activas['LBS_C'].sum() if len(activas) else 0,
        'buckets_full':  (buckets['STATUS'] == '✅ ABASTO+CUOTA').sum(),
        'buckets_abasto':(buckets['STATUS'] == '🟡 SOLO ABASTO').sum(),
        'buckets_ni':    (buckets['STATUS'] == '🔴 NI ABASTO').sum(),
        'pct_abasto_global': buckets['ABASTO_CUBIERTO'].sum() / buckets['ABASTO_META'].sum() if buckets['ABASTO_META'].sum() else 1.0,
        'pct_cuota_global':  buckets['ASIGNADO'].sum() / buckets['CUOTA_META'].sum() if buckets['CUOTA_META'].sum() else 1.0,
    }


# ── Excel helpers ─────────────────────────────────────────────────────────────
COLOR_H = {'orig': '2F4F7F', 'new': '1D6A40', 'cfg': '7B3F00',
           'sc1': '1E40AF', 'sc2': '065F46', 'sc3': '6B21A8'}
COLOR_ROW = {'ok': 'C6EFCE', 'warn': 'FFEB9C', 'err': 'FFC7CE', 'inac': 'E5E7EB', 'total': 'D1FAE5'}


def make_styles():
    thin = Side(style='thin', color='CCCCCC')
    brd  = Border(left=thin, right=thin, top=thin, bottom=thin)
    return {
        'FW':  Font(color='FFFFFF', bold=True, name='Arial', size=10),
        'FN':  Font(name='Arial', size=9),
        'FB':  Font(name='Arial', size=9, bold=True),
        'FI':  Font(name='Arial', size=9, color='9CA3AF', italic=True),
        'FT':  Font(name='Arial', size=9, bold=True, color='065F46'),
        'brd': brd,
    }


def formato_detalle(ws, df_out, n_orig, st_):
    """Formatea la hoja DETALLE. Usa formato condicional de Excel (evaluado
    por Excel al abrir, no celda-por-celda en Python) para que funcione bien
    con decenas de miles de filas sin tronar por memoria/tiempo."""
    FW, brd = st_['FW'], st_['brd']
    n_cols = len(df_out.columns)
    last_col_letter = get_column_letter(n_cols)

    for col_idx in range(1, n_cols + 1):
        c = ws.cell(row=1, column=col_idx)
        c.fill = PatternFill('solid', start_color=COLOR_H['new'] if col_idx > n_orig else COLOR_H['orig'])
        c.font = FW
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        c.border = brd
    ws.row_dimensions[1].height = 30
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = ws.dimensions

    n_rows = ws.max_row
    if n_rows >= 2 and 'STATUS_DISPO' in df_out.columns:
        s_idx = df_out.columns.get_loc('STATUS_DISPO') + 1
        s_col = get_column_letter(s_idx)
        data_range = f"A2:{last_col_letter}{n_rows}"
        f_ok, f_warn, f_err, f_inac = (PatternFill('solid', start_color=COLOR_ROW[k])
                                        for k in ('ok', 'warn', 'err', 'inac'))
        reglas = [
            (f'${s_col}2="✅ COMPLETA"', f_ok),
            (f'${s_col}2="🔁 DESCARTADA (no llegó al umbral)"', f_warn),
            (f'${s_col}2="❌ SIN INVENTARIO"', f_err),
            (f'${s_col}2="⛔ INACTIVA"', f_inac),
        ]
        for formula, fill in reglas:
            ws.conditional_formatting.add(data_range, FormulaRule(formula=[formula], fill=fill))

    # Formato de porcentaje — solo en las columnas % (no toda la hoja)
    pct_cols = [c for c in ('PCT_LINEA', 'PCT_DISPO', 'PCT_ABASTO', 'PCT_CUOTA') if c in df_out.columns]
    for pc in pct_cols:
        pi = df_out.columns.get_loc(pc) + 1
        col_letter = get_column_letter(pi)
        for cell in ws[col_letter][1:]:   # salta encabezado
            cell.number_format = '0.0%'

    for i, col_name in enumerate(df_out.columns, 1):
        w = (18 if col_name in ('DISPO', 'ESTILO_EQ', 'ENTREGA', 'LOTSIZE', 'STATUS_DISPO', 'STATUS_AC',
                                 'COLOR_NOMBRE', 'ESCENARIO', 'PLANTA_WIP', 'CONSTRUCCION')
             else 14 if col_name in ('LBS_C', 'LBS_ASIGNADO', 'LBS_FALTANTE', 'INV', 'INV_EFECTIVO',
                                      'PLAN_INS_DIA1', 'PLAN_INS', 'ABASTO_META', 'CUOTA_META',
                                      'ABASTO_CUBIERTO', 'CUOTA_CUBIERTA', 'LBS_FALTA_ABASTO', 'LBS_FALTA_CUOTA')
             else 12 if col_name in ('PCT_LINEA', 'PCT_DISPO', 'PCT_ABASTO', 'PCT_CUOTA')
             else 9 if col_name == 'ACTIVO' else 11)
        ws.column_dimensions[get_column_letter(i)].width = w


def cuadro_comparativo(ws, title, group_col, resultados, start_row, start_col, header_color, st_):
    FW, FN, FB, FT, brd = st_['FW'], st_['FN'], st_['FB'], st_['FT'], st_['brd']
    labels = [r['label'] for r in resultados]
    dfs = []
    for r in resultados:
        completas = r['df'][(r['df']['STATUS_DISPO'] == '✅ COMPLETA') & (r['df']['ACTIVO'] == 1)]
        t = (completas.groupby(group_col)['LBS_C'].sum().rename(r['label'])
             if group_col in completas.columns else pd.Series(dtype=float, name=r['label']))
        dfs.append(t)
    comp = pd.concat(dfs, axis=1).fillna(0).reset_index()
    comp.columns = [group_col] + labels
    comp = comp.sort_values(labels[0], ascending=False)
    total_row = {group_col: 'TOTAL', **{lbl: comp[lbl].sum() for lbl in labels}}
    comp = pd.concat([comp, pd.DataFrame([total_row])], ignore_index=True)

    n_cols = len(comp.columns); end_col = start_col + n_cols - 1
    tc = ws.cell(row=start_row, column=start_col, value=title)
    tc.font = FW; tc.fill = PatternFill('solid', start_color=header_color)
    tc.alignment = Alignment(horizontal='center'); tc.border = brd
    if end_col > start_col:
        ws.merge_cells(start_row=start_row, start_column=start_col, end_row=start_row, end_column=end_col)

    sc_colors = [COLOR_H['sc1'], COLOR_H['sc2'], COLOR_H['sc3']]
    for ci, col_name in enumerate(comp.columns, start=start_col):
        c = ws.cell(row=start_row + 1, column=ci, value=col_name)
        c.font = FW if ci > start_col else FB
        c.fill = PatternFill('solid', start_color=sc_colors[(ci - start_col - 1) % 3] if ci > start_col else 'D9D9D9')
        c.alignment = Alignment(horizontal='center'); c.border = brd

    for ri, (_, row_data) in enumerate(comp.iterrows(), start=start_row + 2):
        is_total = str(row_data.iloc[0]).upper() == 'TOTAL'
        for ci, val in enumerate(row_data.values, start=start_col):
            cell = ws.cell(row=ri, column=ci, value=val)
            cell.border = brd
            if is_total:
                cell.font = FT; cell.fill = PatternFill('solid', start_color=COLOR_ROW['total'])
                cell.alignment = Alignment(horizontal='center')
            else:
                cell.font = FN
                if isinstance(val, (int, float)):
                    cell.number_format = '#,##0.0'; cell.alignment = Alignment(horizontal='right')
    return start_row + len(comp) + 3


def cuadro_abasto_cuota(ws, resultados, start_row, start_col, st_):
    """Una tabla por escenario: PLANTA_WIP+CONSTRUCCION vs ABASTO/CUOTA."""
    FW, FN, FB, FT, brd = st_['FW'], st_['FN'], st_['FB'], st_['FT'], st_['brd']
    sc_colors = [COLOR_H['sc1'], COLOR_H['sc2'], COLOR_H['sc3']]
    row = start_row
    for idx_r, r in enumerate(resultados):
        buckets = resumen_buckets(r['df'])
        cols_out = ['PLANTA_WIP', 'CONSTRUCCION', 'ABASTO_META', 'ABASTO_CUBIERTO', 'PCT_ABASTO',
                    'CUOTA_META', 'ASIGNADO', 'PCT_CUOTA', 'STATUS']
        title = f"📋 ABASTO→CUOTA — {r['label']}"
        n_cols = len(cols_out); end_col = start_col + n_cols - 1
        tc = ws.cell(row=row, column=start_col, value=title)
        tc.font = FW; tc.fill = PatternFill('solid', start_color=sc_colors[idx_r % 3])
        tc.alignment = Alignment(horizontal='center'); tc.border = brd
        ws.merge_cells(start_row=row, start_column=start_col, end_row=row, end_column=end_col)
        row += 1
        for ci, cn in enumerate(cols_out, start=start_col):
            c = ws.cell(row=row, column=ci, value=cn)
            c.font = FB; c.fill = PatternFill('solid', start_color='D9D9D9')
            c.alignment = Alignment(horizontal='center'); c.border = brd
        row += 1
        f_ok, f_warn, f_err = (PatternFill('solid', start_color=COLOR_ROW[k]) for k in ('ok', 'warn', 'err'))
        for _, br in buckets[cols_out].iterrows():
            for ci, cn in enumerate(cols_out, start=start_col):
                val = br[cn]
                cell = ws.cell(row=row, column=ci, value=val)
                cell.border = brd; cell.font = FN
                if cn in ('PCT_ABASTO', 'PCT_CUOTA'):
                    cell.number_format = '0.0%'; cell.alignment = Alignment(horizontal='center')
                    cell.fill = f_ok if val >= 1 else f_warn if val >= 0.5 else f_err
                elif isinstance(val, (int, float)):
                    cell.number_format = '#,##0.0'; cell.alignment = Alignment(horizontal='right')
            row += 1
        row += 2
    return row


def hoja_diagnostico(wb, resultados, capacidad_cfg, st_):
    """Hoja DIAGNOSTICO: resumen por razón de no-asignación, por escenario."""
    FW, FN, FB, FT, brd = st_['FW'], st_['FN'], st_['FB'], st_['FT'], st_['brd']
    ws = wb.create_sheet('DIAGNOSTICO')
    sc_colors = [COLOR_H['sc1'], COLOR_H['sc2'], COLOR_H['sc3']]
    row = 1
    for idx_r, r in enumerate(resultados):
        _, resumen = diagnostico_asignacion(r['df'], capacidad_cfg)
        cols_out = ['RAZON_NO_ASIGNADO', 'LBS', 'N_LINEAS']
        title = f"🔍 Diagnóstico — {r['label']}"
        tc = ws.cell(row=row, column=1, value=title)
        tc.font = FW; tc.fill = PatternFill('solid', start_color=sc_colors[idx_r % 3])
        tc.alignment = Alignment(horizontal='center'); tc.border = brd
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)
        row += 1
        for ci, cn in enumerate(cols_out, start=1):
            c = ws.cell(row=row, column=ci, value=cn)
            c.font = FB; c.fill = PatternFill('solid', start_color='D9D9D9')
            c.alignment = Alignment(horizontal='center'); c.border = brd
        row += 1
        for _, rr in resumen.iterrows():
            for ci, cn in enumerate(cols_out, start=1):
                val = rr[cn]
                cell = ws.cell(row=row, column=ci, value=val)
                cell.border = brd; cell.font = FN
                if isinstance(val, (int, float)):
                    cell.number_format = '#,##0'; cell.alignment = Alignment(horizontal='right')
            row += 1
        row += 2
    for col, w in zip(['A', 'B', 'C'], [55, 16, 14]):
        ws.column_dimensions[col].width = w
    return ws


def generar_excel(resultados, entrega_pesos, cascada_activo, capacidad_cfg, huerfanos_info=None):
    st_ = make_styles()
    FW, FN, FB, brd = st_['FW'], st_['FN'], st_['FB'], st_['brd']

    new_cols = ['COLOR_NOMBRE', 'ACTIVO', 'INV_EFECTIVO', 'LBS_ASIGNADO', 'LBS_FALTANTE',
                'PCT_LINEA', 'PCT_DISPO', 'STATUS_DISPO',
                'ABASTO_META', 'ABASTO_CUBIERTO', 'PCT_ABASTO', 'LBS_FALTA_ABASTO',
                'CUOTA_META', 'CUOTA_CUBIERTA', 'PCT_CUOTA', 'LBS_FALTA_CUOTA', 'STATUS_AC']

    frames = []
    for r in resultados:
        df_sc = r['df'].copy()
        df_sc.insert(0, 'ESCENARIO', r['label'])
        frames.append(df_sc)
    df_all = pd.concat(frames, ignore_index=True)

    orig_cols = [c for c in ESSENTIAL_ORIG_COLS if c in df_all.columns]
    col_order = ['ESCENARIO'] + orig_cols + [c for c in new_cols if c in df_all.columns]
    df_out = df_all[col_order].copy()

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        # DETALLE se construye directo con openpyxl más abajo (mucho más rápido
        # que pandas.to_excel + recarga cuando hay decenas de miles de filas).
        pd.DataFrame().to_excel(writer, sheet_name='REPORTE', index=False)

        resumen_rows = []
        for r in resultados:
            k = kpis(r['df'])
            resumen_rows.append({
                'Escenario': r['label'], 'DISPOs totales': k['total'],
                '✅ Completas': k['completas'], '🔁 Descartadas (no llegó al umbral)': k['descartadas'],
                '❌ Sin inventario': k['sin_inv'], '⛔ Líneas inactivas': k['inactivas'],
                '% Cobertura física': f"{k['cob']:.1%}",
                'Buckets ✅ ABASTO+CUOTA': k['buckets_full'],
                'Buckets 🟡 Solo ABASTO': k['buckets_abasto'],
                'Buckets 🔴 Ni ABASTO': k['buckets_ni'],
                '% ABASTO global': f"{k['pct_abasto_global']:.1%}",
                '% CUOTA global': f"{k['pct_cuota_global']:.1%}",
            })
        pd.DataFrame(resumen_rows).to_excel(writer, sheet_name='RESUMEN', index=False)

        cfg_e  = pd.DataFrame(list(entrega_pesos.items()), columns=['ENTREGA', 'PESO'])
        cfg_ls = pd.DataFrame(capacidad_cfg)
        cfg_ls['CAPACIDAD_SEMANAL_LBS'] = cfg_ls['LOTES'] * 7 * cfg_ls.get('LBS_POR_LOTE', 0)
        cfg_cascada = pd.DataFrame([
            {'Dígito': d, 'Color': CASCADA_COLOR_NOMBRES[d],
             'Estado': '🟢 ACTIVO' if cascada_activo.get(d, 1) == 1 else '🔴 INACTIVO'}
            for d in CASCADA_COLOR_NOMBRES
        ])
        cfg_e.to_excel(writer, sheet_name='CONFIG', index=False, startrow=1, startcol=0)
        cfg_ls.to_excel(writer, sheet_name='CONFIG', index=False, startrow=1, startcol=3)
        cfg_cascada.to_excel(writer, sheet_name='CONFIG', index=False, startrow=1, startcol=12)

        libre_frames = []
        for r in resultados:
            libre_r = resumen_inventario_libre(r['df'])
            libre_r.insert(0, 'ESCENARIO', r['label'])
            libre_frames.append(libre_r)
        df_libre = pd.concat(libre_frames, ignore_index=True).sort_values(
            ['ESCENARIO', 'INV_LIBRE'], ascending=[True, False])
        df_libre.to_excel(writer, sheet_name='INV_LIBRE', index=False)

        if huerfanos_info:
            solo_wip, solo_cuota = huerfanos_info
            pd.DataFrame(sorted(solo_wip), columns=['PLANTA_WIP', 'CONSTRUCCION']).to_excel(
                writer, sheet_name='HUERFANOS', index=False, startrow=1, startcol=0)
            pd.DataFrame(sorted(solo_cuota), columns=['PLANTA_WIP', 'CONSTRUCCION']).to_excel(
                writer, sheet_name='HUERFANOS', index=False, startrow=1, startcol=3)

    buf.seek(0)
    wb = load_workbook(buf)

    ws = wb.create_sheet('DETALLE', 0)
    ws.append(list(df_out.columns))
    df_out_clean = df_out.astype(object).where(df_out.notna(), None)
    for row_vals in df_out_clean.itertuples(index=False, name=None):
        ws.append(row_vals)
    n_orig = 1 + len(orig_cols)
    formato_detalle(ws, df_out, n_orig, st_)
    ws.cell(row=1, column=1).fill = PatternFill('solid', start_color=COLOR_H['cfg'])

    ws_r = wb['RESUMEN']
    sc_fills = [PatternFill('solid', start_color=c) for c in [COLOR_H['sc1'], COLOR_H['sc2'], COLOR_H['sc3']]]
    for col in range(1, 13):
        ws_r.column_dimensions[get_column_letter(col)].width = 20
    for row in ws_r.iter_rows(min_row=1, max_row=ws_r.max_row):
        for cell in row:
            cell.border = brd
            if cell.row == 1:
                cell.font = FW; cell.fill = PatternFill('solid', start_color=COLOR_H['orig'])
                cell.alignment = Alignment(horizontal='center')
            else:
                cell.font = FN
                sc_idx = cell.row - 2
                if cell.column == 1 and 0 <= sc_idx < len(sc_fills):
                    cell.fill = sc_fills[sc_idx]
                    cell.font = Font(color='FFFFFF', bold=True, name='Arial', size=9)

    ws_libre = wb['INV_LIBRE']
    for col_idx in range(1, len(df_libre.columns) + 1):
        c = ws_libre.cell(row=1, column=col_idx)
        c.font = FW; c.fill = PatternFill('solid', start_color=COLOR_H['cfg'])
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        c.border = brd
    ws_libre.freeze_panes = 'A2'
    ws_libre.auto_filter.ref = ws_libre.dimensions
    libre_idx = df_libre.columns.get_loc('INV_LIBRE') + 1
    pct_idx_libre = df_libre.columns.get_loc('PCT_USO_POOL') + 1
    f_libre = PatternFill('solid', start_color=COLOR_ROW['warn'])
    for row in ws_libre.iter_rows(min_row=2, max_row=ws_libre.max_row):
        for cell in row:
            cell.font = FN; cell.border = brd
        row[pct_idx_libre - 1].number_format = '0.0%'
        if isinstance(row[libre_idx - 1].value, (int, float)) and row[libre_idx - 1].value > 0:
            row[libre_idx - 1].fill = f_libre
    for i, col_name in enumerate(df_libre.columns, 1):
        w = 18 if col_name in ('ESCENARIO', 'ESTILO_EQ', 'DTITULAR') else 16
        ws_libre.column_dimensions[get_column_letter(i)].width = w

    ws_c = wb['CONFIG']
    titulos = {1: 'PRIORIDAD ENTREGA', 4: 'CAPACIDAD LOTSIZE+MIX (semanal = diaria×7)', 13: 'CASCADA DE COLOR'}
    for col_start, titulo in titulos.items():
        cell = ws_c.cell(row=1, column=col_start, value=titulo)
        cell.font = FW; cell.fill = PatternFill('solid', start_color=COLOR_H['cfg'])
        cell.alignment = Alignment(horizontal='center')
    for col in range(1, 17):
        ws_c.column_dimensions[get_column_letter(col)].width = 15
    for row in ws_c.iter_rows(min_row=2, max_row=ws_c.max_row):
        for cell in row:
            cell.font = FB if cell.row == 2 else FN
            cell.border = brd
            if cell.row == 2:
                cell.fill = PatternFill('solid', start_color='D9D9D9')

    ws_rep = wb['REPORTE']
    next_row = cuadro_comparativo(ws_rep, '📋 Completas por ENTREGA', 'ENTREGA', resultados, 1, 1, COLOR_H['orig'], st_)
    next_row = cuadro_comparativo(ws_rep, '📋 Completas por LOTSIZE', 'LOTSIZE', resultados, next_row, 1, COLOR_H['orig'], st_)
    next_row = cuadro_comparativo(ws_rep, '📋 Completas por COLOR', 'COLOR_NOMBRE', resultados, next_row, 1, COLOR_H['orig'], st_)
    cuadro_abasto_cuota(ws_rep, resultados, 1, 12, st_)
    for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K']:
        ws_rep.column_dimensions[col].width = 16

    hoja_diagnostico(wb, resultados, capacidad_cfg, st_)

    out = io.BytesIO()
    wb.save(out); out.seek(0)
    return out


# ── Helpers de perfil de configuración ──────────────────────────────────────
def default_cap_df():
    return pd.DataFrame(CAPACIDAD_LOTSIZE)


def default_entrega_df():
    return pd.DataFrame(list(DEFAULT_ENTREGA.items()), columns=['ENTREGA', 'PESO'])


def default_cascada_df():
    return pd.DataFrame([{'DIGITO': d, 'COLOR': n, 'ACTIVO': True} for d, n in CASCADA_COLOR_NOMBRES.items()])


if 'cfg_version' not in st.session_state:
    st.session_state['cfg_version']     = 0
    st.session_state['cap_df_seed']     = default_cap_df()
    st.session_state['entrega_df_seed'] = default_entrega_df()
    st.session_state['cascada_df_seed'] = default_cascada_df()

# ── 1. Cargar datos ──────────────────────────────────────────────────────────
st.markdown("""
<div class="step-header"><div class="step-badge">1</div><div><p class="step-title">Cargar datos</p></div></div>
""", unsafe_allow_html=True)
st.caption("Workbook TIM_WIPCR.xlsx — hojas requeridas: WIP (data completa) y CUOTA (metas ABASTO/CUOTA)")
uploaded = st.file_uploader("Sube TIM_WIPCR.xlsx", type=['xlsx', 'xls'], label_visibility="collapsed")

df_wip, df_cuota, solo_wip, solo_cuota = None, None, set(), set()

if uploaded:
    df_wip, df_cuota, err = leer_tim_wipcr(uploaded)
    if err:
        st.error(err)
        df_wip = None
    else:
        faltan = [c for c in REQ_WIP_COLS if c not in df_wip.columns]
        if faltan:
            st.error(f"Faltan columnas en WIP: {faltan}")
            df_wip = None
        else:
            c1, c2, c3 = st.columns(3)
            c1.success(f"**{uploaded.name}**")
            c2.caption(f"WIP: {len(df_wip):,} filas · {df_wip['DISPO'].nunique():,} dispos")
            c3.caption(f"CUOTA: {len(df_cuota):,} combinaciones PLANTA_WIP+CONSTRUCCION")

            solo_wip, solo_cuota = validar_huerfanos(df_wip, df_cuota)
            if solo_wip:
                st.markdown(f'<div class="warn-box">⚠️ {len(solo_wip)} combinaciones en WIP sin meta en CUOTA (no se les exige ABASTO/CUOTA).</div>', unsafe_allow_html=True)
            if solo_cuota:
                st.markdown(f'<div class="warn-box">⚠️ {len(solo_cuota)} metas en CUOTA sin inventario que las cubra en WIP.</div>', unsafe_allow_html=True)

            tiene_dia1    = 'PLAN_INS_DIA1' in df_wip.columns
            tiene_semanal = 'PLAN_INS' in df_wip.columns
            if not tiene_dia1:
                st.markdown('<div class="warn-box">⚠️ Sin PLAN_INS_DIA1 — Escenario 2 = Solo INV.</div>', unsafe_allow_html=True)
            if not tiene_semanal:
                st.markdown('<div class="warn-box">⚠️ Sin PLAN_INS — Escenario 3 = Solo INV.</div>', unsafe_allow_html=True)

            with st.expander("👁️ Vista previa WIP", expanded=False):
                preview_cols = [c for c in ['DISPO', 'ENTREGA', 'PLANTA_WIP', 'CONSTRUCCION', 'ESTILO_EQ',
                                             'DTITULAR', 'LOTSIZE', 'MIX', 'LBS_C', 'INV'] if c in df_wip.columns]
                st.dataframe(df_wip[preview_cols].head(8), use_container_width=True, hide_index=True)

# ── 2. Configuración ──────────────────────────────────────────────────────────
entrega_pesos, capacidad_cfg, cascada_activo = dict(DEFAULT_ENTREGA), CAPACIDAD_LOTSIZE, {d: 1 for d in CASCADA_COLOR_NOMBRES}

if df_wip is not None:
    st.markdown("""
    <div class="step-header"><div class="step-badge">2</div><div><p class="step-title">Configuración</p></div></div>
    """, unsafe_allow_html=True)

    no_config = validar_lotsize_no_configurados(df_wip, st.session_state['cap_df_seed'].to_dict('records'))
    if not no_config.empty:
        total_lbs_no_cfg = no_config['LBS_C_TOTAL'].sum()
        combos_txt = ', '.join(f"{r.LOTSIZE}·{r.MIX}" for r in no_config.itertuples())
        st.markdown(
            f'<div class="warn-box">⚠️ {len(no_config)} combinación(es) LOTSIZE+MIX en tus datos NO están en la '
            f'configuración ({combos_txt}) — {total_lbs_no_cfg:,.0f} lbs en esas líneas quedarán en 0. '
            f'Agrégalas en la tabla de Capacidad abajo si sí las necesitas.</div>',
            unsafe_allow_html=True
        )

    bcol1, bcol2, bcol3 = st.columns(3)
    with bcol1:
        if st.button("🔄 Restaurar valores por defecto", use_container_width=True):
            st.session_state['cap_df_seed']     = default_cap_df()
            st.session_state['entrega_df_seed'] = default_entrega_df()
            st.session_state['cascada_df_seed'] = default_cascada_df()
            st.session_state['cfg_version']    += 1
            st.rerun()
    with bcol3:
        perfil_file = st.file_uploader("Cargar perfil (JSON)", type=['json'],
                                        label_visibility='collapsed', key='perfil_uploader')
        if perfil_file is not None:
            sig = f"{perfil_file.name}_{perfil_file.size}"
            if st.session_state.get('perfil_aplicado') != sig:
                try:
                    data = json.load(perfil_file)
                    st.session_state['cap_df_seed']     = pd.DataFrame(data['capacidad'])
                    st.session_state['entrega_df_seed'] = pd.DataFrame(data['entrega'])
                    st.session_state['cascada_df_seed'] = pd.DataFrame(data['cascada'])
                    st.session_state['cfg_version']    += 1
                    st.session_state['perfil_aplicado'] = sig
                    st.rerun()
                except Exception as e:
                    st.error(f"Perfil inválido: {e}")

    v = st.session_state['cfg_version']

    with st.expander("📋 Prioridad ENTREGA", expanded=False):
        st.caption("Menor número = mayor prioridad.")
        entrega_edit = st.data_editor(
            st.session_state['entrega_df_seed'], num_rows="fixed", hide_index=True,
            use_container_width=True, key=f"entrega_editor_{v}",
            column_config={'PESO': st.column_config.NumberColumn('PESO', min_value=1, max_value=99)}
        )

    with st.expander("⚙️ Capacidad LOTSIZE + MIX (semanal)", expanded=True):
        st.caption("Capacidad semanal (lbs) = Lotes/día × 7 × Lbs/lote. Puedes agregar o quitar filas. "
                   "EVITAR = solo se usa con prioridad normal si ENTREGA=01-EXPEDITE o en Fase ABASTO.")
        cap_edit = st.data_editor(
            st.session_state['cap_df_seed'], num_rows="dynamic", hide_index=True,
            use_container_width=True, key=f"cap_editor_{v}",
            column_config={
                'LOTSIZE':      st.column_config.TextColumn('LOTSIZE', required=True),
                'MIX':          st.column_config.TextColumn('MIX', required=True),
                'LOTES':        st.column_config.NumberColumn('Lotes/día', min_value=0, step=1),
                'LBS_POR_LOTE': st.column_config.NumberColumn('Lbs/lote', min_value=0, step=100),
                'PRIORIDAD':    st.column_config.NumberColumn('Prioridad', min_value=1, step=1),
                'ACTIVO':       st.column_config.CheckboxColumn('Activo'),
                'EVITAR':       st.column_config.CheckboxColumn('Evitar'),
            }
        )
        cap_edit = cap_edit.dropna(subset=['LOTSIZE', 'MIX']).copy()
        for col, default in [('LOTES', 0), ('LBS_POR_LOTE', 0), ('PRIORIDAD', 99), ('ACTIVO', True), ('EVITAR', False)]:
            if col not in cap_edit.columns:
                cap_edit[col] = default
            cap_edit[col] = cap_edit[col].fillna(default)
        cap_sem_total = (cap_edit['LOTES'] * 7 * cap_edit['LBS_POR_LOTE'] * cap_edit['ACTIVO']).sum()
        st.caption(f"📦 Capacidad total semanal activa: **{cap_sem_total:,.0f} lbs**")

    with st.expander("🎨 Cascada de color (COLOR_A dígito 3)", expanded=False):
        st.caption("🔴 Inactivo = ese color no recibe inventario.")
        cascada_edit = st.data_editor(
            st.session_state['cascada_df_seed'], num_rows="fixed", hide_index=True,
            use_container_width=True, key=f"cascada_editor_{v}",
            column_config={'ACTIVO': st.column_config.CheckboxColumn('Activo')}
        )

    with st.expander("🎯 Umbral de completitud por DISPO (todo o nada)", expanded=False):
        st.caption("Si un DISPO no llega al umbral en TODAS sus líneas activas, se descarta completo (0 lbs). "
                   "Se basa en la columna COMPONENTE de WIP (PRINCIPAL vs SECUNDARIO-N). Si esa columna no existe, "
                   "todo se trata como PRINCIPAL.")
        uc1, uc2 = st.columns(2)
        with uc1:
            umbral_principal_pct = st.number_input(
                "Umbral componente PRINCIPAL (%)", min_value=50, max_value=100,
                value=int(UMBRAL_PRINCIPAL_DEFAULT * 100), step=1, key=f"umbral_p_{v}"
            )
        with uc2:
            umbral_secundario_pct = st.number_input(
                "Umbral componente SECUNDARIO-N (%)", min_value=50, max_value=100,
                value=int(UMBRAL_SECUNDARIO_DEFAULT * 100), step=1, key=f"umbral_s_{v}"
            )
        umbral_principal  = umbral_principal_pct / 100.0
        umbral_secundario = umbral_secundario_pct / 100.0

    with bcol2:
        perfil_export = {
            'capacidad': cap_edit.to_dict('records'),
            'entrega': entrega_edit.to_dict('records'),
            'cascada': cascada_edit.to_dict('records'),
        }
        st.download_button(
            "💾 Exportar perfil (JSON)",
            data=json.dumps(perfil_export, indent=2, ensure_ascii=False),
            file_name="perfil_config.json", mime="application/json", use_container_width=True
        )

    entrega_pesos  = {str(r['ENTREGA']): int(r['PESO']) for _, r in entrega_edit.iterrows()}
    capacidad_cfg  = cap_edit.to_dict('records')
    cascada_activo = {str(r['DIGITO']): (1 if r['ACTIVO'] else 0) for _, r in cascada_edit.iterrows()}

    # ── 3. Ejecutar ───────────────────────────────────────────────────────────
    st.markdown("""
    <div class="step-header"><div class="step-badge">3</div><div><p class="step-title">Ejecutar</p></div></div>
    """, unsafe_allow_html=True)
    st.caption("Elige qué escenario(s) correr — menos escenarios = menos tiempo y un Excel más liviano.")

    label_to_esc = {esc['label']: esc for esc in ESCENARIOS}
    seleccionados = st.multiselect(
        "Escenarios a ejecutar", options=list(label_to_esc.keys()),
        default=[ESCENARIOS[-1]['label']], label_visibility="collapsed"
    )

    if st.button("▶ Ejecutar escenario(s) seleccionados", type="primary",
                 use_container_width=True, disabled=not seleccionados):
        targets = construir_targets(df_cuota)
        progress_bar = st.progress(0, text="Iniciando…")
        resultados = []
        n_esc = len(seleccionados)
        for ei, label in enumerate(seleccionados):
            esc = label_to_esc[label]
            base = ei / n_esc

            def cb(p, msg, base=base, esc=esc):
                progress_bar.progress(min(base + p / n_esc, 1.0), text=f"{esc['label']} — {msg}")

            df_sc = motor_asignacion(df_wip, targets, entrega_pesos, cascada_activo,
                                      capacidad_cfg, col_extra=esc['col_extra'], progress_cb=cb,
                                      umbral_principal=umbral_principal, umbral_secundario=umbral_secundario)
            resultados.append({'key': esc['key'], 'label': esc['label'], 'df': df_sc})
        progress_bar.progress(1.0, text="Completado ✅")
        st.session_state['resultados'] = resultados
        st.session_state['huerfanos'] = (solo_wip, solo_cuota)

col_main = st.container()
with col_main:
    if 'resultados' in st.session_state:
        resultados = st.session_state['resultados']

        st.markdown('<p class="section-title">Comparativo de escenarios</p>', unsafe_allow_html=True)
        kpi_cols = st.columns(len(resultados))
        sc_cls = ['sc-card-blue', 'sc-card-green', 'sc-card-purple']
        for i, r in enumerate(resultados):
            k = kpis(r['df'])
            with kpi_cols[i]:
                st.markdown(f"""
                <div class="sc-card {sc_cls[i % len(sc_cls)]}">
                  <div class="sc-label">{r['label']}</div>
                  <div class="sc-row"><span class="sc-metric">✅ Completas</span><span class="sc-value">{k['completas']:,}</span></div>
                  <div class="sc-row"><span class="sc-metric">🔁 Descartadas (no llegó al umbral)</span><span class="sc-value">{k['descartadas']:,}</span></div>
                  <div class="sc-row"><span class="sc-metric">❌ Sin inventario</span><span class="sc-value">{k['sin_inv']:,}</span></div>
                  <div class="sc-row"><span class="sc-metric">⛔ Inactivas</span><span class="sc-value">{k['inactivas']:,}</span></div>
                  <div class="sc-row sc-divider"><span class="sc-metric">Cobertura física</span><span class="sc-value">{k['cob']:.1%}</span></div>
                  <div class="sc-row"><span class="sc-metric">✅ Buckets ABASTO+CUOTA</span><span class="sc-value">{k['buckets_full']:,}</span></div>
                  <div class="sc-row"><span class="sc-metric">🟡 Solo ABASTO</span><span class="sc-value">{k['buckets_abasto']:,}</span></div>
                  <div class="sc-row"><span class="sc-metric">🔴 Ni ABASTO</span><span class="sc-value">{k['buckets_ni']:,}</span></div>
                  <div class="sc-row"><span class="sc-metric">% ABASTO global</span><span class="sc-value">{k['pct_abasto_global']:.1%}</span></div>
                  <div class="sc-row"><span class="sc-metric">% CUOTA global</span><span class="sc-value">{k['pct_cuota_global']:.1%}</span></div>
                </div>
                """, unsafe_allow_html=True)

        # ── Gráficos comparativos ───────────────────────────────────────────
        st.markdown('<p class="section-title">Gráficos comparativos</p>', unsafe_allow_html=True)
        if PLOTLY_OK:
            esc_sel = st.selectbox("Escenario para gráficos", [r['label'] for r in resultados], key="esc_chart")
            r_sel = next(r for r in resultados if r['label'] == esc_sel)
            buckets = resumen_buckets(r_sel['df'])

            gcol1, gcol2 = st.columns(2)
            with gcol1:
                top_b = buckets.nlargest(20, 'CUOTA_META')
                fig1 = go.Figure()
                fig1.add_bar(name='% ABASTO cubierto', x=top_b['PLANTA_WIP'] + ' · ' + top_b['CONSTRUCCION'],
                             y=(top_b['PCT_ABASTO'] * 100).clip(upper=100), marker_color='#1d4ed8')
                fig1.add_bar(name='% CUOTA cubierta', x=top_b['PLANTA_WIP'] + ' · ' + top_b['CONSTRUCCION'],
                             y=(top_b['PCT_CUOTA'] * 100).clip(upper=100), marker_color='#059669')
                fig1.update_layout(title="% ABASTO vs % CUOTA por PLANTA_WIP+CONSTRUCCION (top 20 por meta)",
                                    barmode='group', height=380, font_size=10, margin=dict(t=40, b=80))
                fig1.update_xaxes(tickangle=-45)
                st.plotly_chart(fig1, use_container_width=True)
            with gcol2:
                cap_res = resumen_capacidad(r_sel['df'], capacidad_cfg)
                cap_res['LOTSIZE_MIX'] = cap_res['LOTSIZE'] + ' · ' + cap_res['MIX']
                fig2 = go.Figure()
                fig2.add_bar(name='Lbs usados (semana)', x=cap_res['LOTSIZE_MIX'], y=cap_res['LBS_USADOS'], marker_color='#7c3aed')
                fig2.add_bar(name='Capacidad semanal', x=cap_res['LOTSIZE_MIX'], y=cap_res['CAPACIDAD_SEMANAL'], marker_color='#e2e8f0')
                fig2.update_layout(title="Consumo de capacidad semanal por LOTSIZE+MIX",
                                    barmode='overlay', height=380, font_size=10, margin=dict(t=40, b=80))
                fig2.update_xaxes(tickangle=-45)
                st.plotly_chart(fig2, use_container_width=True)
        else:
            st.info("Instala `plotly` (`pip install plotly`) para ver los gráficos comparativos.")

        # ── Semáforo de capacidad ────────────────────────────────────────────
        st.markdown('<p class="section-title">Semáforo de capacidad LOTSIZE + MIX</p>', unsafe_allow_html=True)
        esc_sel2 = st.selectbox("Escenario para semáforo", [r['label'] for r in resultados], key="esc_semaforo")
        r_sel2 = next(r for r in resultados if r['label'] == esc_sel2)
        cap_res2 = resumen_capacidad(r_sel2['df'], capacidad_cfg)

        def semaforo(pct):
            return '🟢 Libre' if pct < 0.7 else ('🟡 Casi saturado' if pct < 1.0 else '🔴 Saturado')
        cap_res2['SEMAFORO'] = cap_res2['PCT_USO'].apply(semaforo)
        cap_tab = cap_res2[['LOTSIZE', 'MIX', 'ACTIVO', 'LOTES', 'LBS_POR_LOTE', 'CAPACIDAD_SEMANAL',
                             'LBS_USADOS', 'LOTES_EQUIV_USADOS', 'PCT_USO', 'SEMAFORO']]
        st.dataframe(
            cap_tab, use_container_width=True, hide_index=True,
            column_config=col_cfg(cap_tab, pct_cols=['PCT_USO'],
                                   num_cols=['CAPACIDAD_SEMANAL', 'LBS_USADOS', 'LOTES_EQUIV_USADOS'], num_fmt='%.0f')
        )

        # ── Diagnóstico: por qué no se asignó más ────────────────────────────
        st.markdown('<p class="section-title">🔍 Diagnóstico — por qué no se asignó más</p>', unsafe_allow_html=True)
        st.caption("Cruza capacidad, inventario libre y metas para etiquetar cada línea incompleta con su razón. "
                   "La razón con más LBS es el cuello de botella real — empieza por ahí.")
        tabs_diag = st.tabs([r['label'] for r in resultados])
        for tab, r in zip(tabs_diag, resultados):
            with tab:
                df_diag, resumen_diag = diagnostico_asignacion(r['df'], capacidad_cfg)
                st.dataframe(
                    resumen_diag, use_container_width=True, hide_index=True,
                    column_config=col_cfg(resumen_diag, num_cols=['LBS', 'N_LINEAS'], num_fmt='%.0f')
                )
                with st.expander("Ver detalle línea por línea con la razón asignada"):
                    cols_diag = [c for c in ['DISPO', 'ENTREGA', 'PLANTA_WIP', 'CONSTRUCCION', 'LOTSIZE', 'MIX',
                                              'ESTILO_EQ', 'DTITULAR', 'COMPONENTE', 'LBS_C', 'LBS_ASIGNADO',
                                              'PCT_LINEA', 'N_POOLS_DISPO', 'STATUS_DISPO', 'RAZON_NO_ASIGNADO']
                                 if c in df_diag.columns]
                    st.dataframe(
                        df_diag[df_diag['RAZON_NO_ASIGNADO'] != '✅ COMPLETA'][cols_diag],
                        use_container_width=True, hide_index=True, height=380,
                        column_config=col_cfg(df_diag, pct_cols=['PCT_LINEA'], num_cols=['LBS_C', 'LBS_ASIGNADO'])
                    )

        # ── Resumen ABASTO/CUOTA por bucket ──────────────────────────────────
        st.markdown('<p class="section-title">Resumen ABASTO → CUOTA por PLANTA_WIP + CONSTRUCCION</p>', unsafe_allow_html=True)
        tabs_ac = st.tabs([r['label'] for r in resultados])
        for tab, r in zip(tabs_ac, resultados):
            with tab:
                buckets = resumen_buckets(r['df'])
                buckets_tab = buckets[['PLANTA_WIP', 'CONSTRUCCION', 'ABASTO_META', 'ABASTO_CUBIERTO', 'PCT_ABASTO',
                                        'LBS_FALTA_ABASTO', 'CUOTA_META', 'ASIGNADO', 'PCT_CUOTA', 'LBS_FALTA_CUOTA', 'STATUS']]
                st.dataframe(
                    buckets_tab, use_container_width=True, hide_index=True, height=380,
                    column_config=col_cfg(
                        buckets_tab,
                        pct_cols=['PCT_ABASTO', 'PCT_CUOTA'],
                        num_cols=['ABASTO_META', 'ABASTO_CUBIERTO', 'LBS_FALTA_ABASTO',
                                  'CUOTA_META', 'ASIGNADO', 'LBS_FALTA_CUOTA'],
                        num_fmt='%.0f'
                    )
                )

        # ── Inventario / plan libre por ESTILO_EQ + DTITULAR ──────────────────
        st.markdown('<p class="section-title">Inventario + plan libre (sin asignar) por ESTILO_EQ + DTITULAR</p>', unsafe_allow_html=True)
        tabs_libre = st.tabs([r['label'] for r in resultados])
        for tab, r in zip(tabs_libre, resultados):
            with tab:
                libre = resumen_inventario_libre(r['df'])
                solo_con_libre = st.checkbox("Mostrar solo los que tienen inventario libre (> 0)",
                                              value=True, key=f"libre_chk_{r['key']}")
                libre_vis = libre[libre['INV_LIBRE'] > 0] if solo_con_libre else libre
                st.dataframe(
                    libre_vis, use_container_width=True, hide_index=True, height=380,
                    column_config=col_cfg(
                        libre_vis,
                        pct_cols=['PCT_USO_POOL'],
                        num_cols=['INV_EFECTIVO', 'LBS_C_TOTAL', 'LBS_ASIGNADO_TOTAL', 'INV_LIBRE'],
                        num_fmt='%.0f'
                    )
                )
                st.caption(f"{len(libre_vis):,} combinaciones · {libre_vis['INV_LIBRE'].sum():,.0f} lbs libres en total")

        # ── Tabla detalle por escenario ───────────────────────────────────────
        st.markdown('<p class="section-title">Detalle por línea</p>', unsafe_allow_html=True)
        tabs = st.tabs([r['label'] for r in resultados])
        cols_show_base = ['DISPO', 'ENTREGA', 'PLANTA_WIP', 'CONSTRUCCION', 'LOTSIZE', 'MIX',
                           'ESTILO_EQ', 'DTITULAR', 'COLOR_NOMBRE', 'ACTIVO',
                           'LBS_C', 'INV', 'INV_EFECTIVO', 'LBS_ASIGNADO', 'LBS_FALTANTE', 'PCT_LINEA',
                           'PCT_ABASTO', 'PCT_CUOTA', 'STATUS_DISPO', 'STATUS_AC']

        for tab, r in zip(tabs, resultados):
            with tab:
                df_r = r['df']
                fc1, fc2, fc3 = st.columns(3)
                with fc1:
                    filtro_status = st.multiselect(
                        "Status", ['✅ COMPLETA', '🔁 DESCARTADA (no llegó al umbral)', '❌ SIN INVENTARIO', '⛔ INACTIVA'],
                        default=['✅ COMPLETA', '🔁 DESCARTADA (no llegó al umbral)', '❌ SIN INVENTARIO', '⛔ INACTIVA'],
                        key=f"fs_{r['key']}", label_visibility="collapsed"
                    )
                with fc2:
                    filtro_ac = st.multiselect(
                        "Status AC", ['✅ ABASTO+CUOTA', '🟡 SOLO ABASTO', '🔴 NI ABASTO', '⚪ SIN META'],
                        default=['✅ ABASTO+CUOTA', '🟡 SOLO ABASTO', '🔴 NI ABASTO', '⚪ SIN META'],
                        key=f"fac_{r['key']}", label_visibility="collapsed"
                    )
                with fc3:
                    filtro_entrega = st.multiselect(
                        "ENTREGA", sorted(df_r['ENTREGA'].dropna().unique()),
                        default=sorted(df_r['ENTREGA'].dropna().unique()),
                        key=f"fe_{r['key']}", label_visibility="collapsed"
                    )

                df_vis = df_r[df_r['STATUS_DISPO'].isin(filtro_status) &
                               df_r['STATUS_AC'].isin(filtro_ac) &
                               df_r['ENTREGA'].isin(filtro_entrega)]

                cols_show = [c for c in cols_show_base if c in df_r.columns]
                df_show = df_vis[cols_show]
                st.dataframe(
                    df_show, use_container_width=True, hide_index=True, height=380,
                    column_config=col_cfg(
                        df_show,
                        pct_cols=['PCT_LINEA', 'PCT_ABASTO', 'PCT_CUOTA'],
                        num_cols=['LBS_C', 'LBS_ASIGNADO', 'LBS_FALTANTE'],
                        num_fmt='%.1f'
                    ) | col_cfg(df_show, num_cols=['INV_EFECTIVO'], num_fmt='%.0f')
                )
                st.caption(f"{len(df_vis):,} líneas · {df_vis['DISPO'].nunique():,} dispos")

        # ── Descargar ─────────────────────────────────────────────────────────
        st.markdown('<p class="section-title">Descargar resultado</p>', unsafe_allow_html=True)
        ts = datetime.now().strftime('%Y%m%d%H%M')

        fingerprint = tuple(r['df'].shape for r in resultados)
        if st.session_state.get('excel_fingerprint') != fingerprint:
            st.session_state['excel_bytes'] = None

        if st.button("📦 Preparar archivo Excel", use_container_width=True):
            with st.spinner("Generando Excel… con datasets grandes puede tardar unos segundos."):
                excel_buf = generar_excel(resultados, entrega_pesos, cascada_activo, capacidad_cfg,
                                           huerfanos_info=st.session_state.get('huerfanos'))
                st.session_state['excel_bytes'] = excel_buf.getvalue()
                st.session_state['excel_fingerprint'] = fingerprint

        if st.session_state.get('excel_bytes'):
            st.download_button(
                label=f"⬇️  Descargar Excel — ASIGNACION_ABASTO_CUOTA_{ts}.xlsx",
                data=st.session_state['excel_bytes'],
                file_name=f"ASIGNACION_ABASTO_CUOTA_{ts}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        else:
            st.caption("Da clic en \"Preparar archivo Excel\" para generarlo y habilitar la descarga.")

    else:
        st.markdown("""
        <div style="display:flex;align-items:center;justify-content:center;height:450px;
                    border:2px dashed #d1d5db;border-radius:12px;color:#9ca3af;
                    flex-direction:column;gap:0.5rem;">
            <div style="font-size:2.5rem">📂</div>
            <div style="font-weight:600;font-size:1rem;">Sube TIM_WIPCR.xlsx y presiona Procesar</div>
            <div style="font-size:0.82rem;">Se calcularán los 3 escenarios con asignación ABASTO → CUOTA</div>
        </div>
        """, unsafe_allow_html=True)
