import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import os, warnings, openpyxl, re
warnings.filterwarnings('ignore')

# ── Page Config ──────────────────────────────────────────────
st.set_page_config(page_title="PERU FROST SAC — Dashboard Ejecutivo", layout="wide", page_icon="🦑")

# ── Password Protection ──────────────────────────────────────
def check_password():
    """Devuelve True si el usuario ingresó la contraseña correcta."""
    if "password_correct" not in st.session_state:
        st.session_state["password_correct"] = False

    if st.session_state["password_correct"]:
        return True

    # Login UI
    st.markdown("""
        <style>
        .login-box {
            background-color: #0f1f38;
            padding: 40px;
            border-radius: 15px;
            border: 1px solid #00d4aa;
            max-width: 500px;
            margin: auto;
            text-align: center;
            margin-top: 50px;
        }
        </style>
        <div class="login-box">
            <h1 style="color:white; margin-bottom:10px;">🔒 Acceso Restringido</h1>
            <p style="color:#7a8da6;">Dashboard Ejecutivo - PERU FROST S.A.C.</p>
        </div>
        """, unsafe_allow_html=True)
        
    password = st.text_input("Ingrese la clave de acceso:", type="password")
    
    if st.button("Ingresar"):
        if password == st.secrets["password"]:
            st.session_state["password_correct"] = True
            st.rerun()
        else:
            st.error("😕 Clave incorrecta")
    return False

if not check_password():
    st.stop()

# Acceso Concedido
st.toast("🔑 Acceso Concedido!", icon="✅")

# ── Theme Colors ─────────────────────────────────────────────
C = dict(
    bg="#0a1628", card="#0f1f38", card2="#162a4a", border="#1e3a5f",
    text="#e0e7ef", muted="#7a8da6", white="#ffffff",
    cyan="#00d4aa", green="#22c55e", yellow="#f59e0b", orange="#f97316",
    red="#ef4444", blue="#3b82f6", purple="#a855f7", gray="#7a8da6",
    pf_cyan="#00d4aa", pf_gold="#f59e0b",
    grad1="#0a1628", grad2="#132240",
)
PRODUCT_COLORS = {
    "ALAS CONGELADAS":"#3b82f6","FILETE CONGELADO":"#22c55e","NUCA":"#f59e0b",
    "REPRODUCTOR":"#a855f7","TENTACULO":"#ef4444",
    "ALAS COCIDAS":"#06b6d4","FILETE COCIDO":"#f97316",
}

# ── UT0 Estáticos (Precio de Equilibrio) ─────────────────────
UT0_FIXED = {
    # Nombres EXACTOS según el Excel Clasificado
    "ALAS CONGELADAS": 2046.96,
    "FILETE CONGELADO": 2048.63,
    "NUCA": 2168.63,
    "TENTACULO": 2038.63,  # Rejos
    "REPRODUCTOR": 2008.63, # Rejos OR
    
    # Cocidos
    "ALAS COCIDAS": None,
    "FILETE COCIDO": 2418.63,
}

PROD_FRESCO = ["ALAS CONGELADAS", "FILETE CONGELADO", "NUCA", "REPRODUCTOR", "TENTACULO"]
PROD_COCIDO = ["ALAS COCIDAS", "FILETE COCIDO"]

# ── CSS ──────────────────────────────────────────────────────
st.markdown(f"""<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap');
    
    /* Global fixes for Safari/Webkit dark mode text colors */
    .stApp, .stApp p, .stApp span, .stApp div, .stApp h1, .stApp h2, .stApp h3, .stApp h4, .stApp h5, .stApp h6, .stApp label, .stApp li {{
        color: {C['text']};
    }}
    
    .stApp {{ background: linear-gradient(180deg, {C['grad1']}, {C['grad2']}); font-family: 'Inter', sans-serif; }}
    header[data-testid="stHeader"] {{ background: transparent; }}
    .block-container {{ padding-top: 1rem; max-width: 1400px; }}
    div[data-testid="stSidebar"] {{ background: {C['card']}; border-right: 1px solid {C['border']}; }}
    div[data-testid="stSidebar"] label, div[data-testid="stSidebar"] span {{ color: {C['text']} !important; }}
    .stTabs [data-baseweb="tab-list"] {{ gap: 0; background: {C['card']}; border-radius: 12px; padding: 4px; border: 1px solid {C['border']}; }}
    .stTabs [data-baseweb="tab"] {{ color: {C['muted']}; border-radius: 8px; padding: 8px 20px; font-weight: 600; font-size: 0.85rem; }}
    .stTabs [aria-selected="true"] {{ background: {C['cyan']}; color: {C['bg']} !important; border-radius: 8px; }}
    .kpi-row {{ display: flex; gap: 16px; margin-bottom: 24px; }}
    .kpi-card {{ flex: 1; padding: 20px 24px; border-radius: 12px; border-left: 4px solid; }}
    .kpi-card.c1 {{ background: {C['card']}; border-color: {C['cyan']}; }}
    .kpi-card.c2 {{ background: {C['card']}; border-color: {C['blue']}; }}
    .kpi-card.c3 {{ background: {C['card']}; border-color: {C['green']}; }}
    .kpi-card.c4 {{ background: {C['card']}; border-color: {C['orange']}; }}
    .kpi-label {{ color: {C['muted']}; font-size: 0.72rem; text-transform: uppercase; letter-spacing: 1.5px; font-weight: 600; }}
    .kpi-value {{ color: {C['white']}; font-size: 1.8rem; font-weight: 800; margin-top: 4px; }}
    .kpi-sub {{ color: {C['muted']}; font-size: 0.75rem; margin-top: 2px; }}
    .section-title {{ color: {C['white']}; font-size: 1.5rem; font-weight: 700; margin: 28px 0 16px; }}
    .card-container {{ background: {C['card']}; border: 1px solid {C['border']}; border-radius: 12px; padding: 24px; margin-bottom: 20px; }}
    .info-row {{ display: flex; gap: 16px; margin-bottom: 20px; }}
    .info-card {{ flex: 1; background: {C['card']}; border: 1px solid {C['border']}; border-radius: 12px; padding: 20px; border-left: 4px solid {C['cyan']}; }}
    .info-label {{ color: {C['muted']}; font-size: 0.72rem; text-transform: uppercase; letter-spacing: 1px; }}
    .info-value {{ color: {C['white']}; font-size: 1.6rem; font-weight: 800; margin-top: 4px; }}
    .info-sub {{ color: {C['muted']}; font-size: 0.72rem; }}
    .exec-note {{ background: linear-gradient(90deg, {C['card']}, {C['card2']}); border: 1px solid {C['cyan']}33; border-radius: 12px; padding: 20px; color: {C['text']}; font-size: 0.85rem; line-height: 1.6; margin-top: 20px; }}
    .exec-note b {{ color: {C['white']}; }}
    .header-banner {{ background: linear-gradient(135deg, #0f1f38 0%, #1a3a5c 100%); padding: 28px 32px; border-radius: 16px; margin-bottom: 20px; border: 1px solid {C['border']}; }}
    .header-sup {{ color: {C['cyan']}; font-size: 0.7rem; letter-spacing: 3px; font-weight: 700; text-transform: uppercase; }}
    .header-title {{ color: {C['white']}; font-size: 2.2rem; font-weight: 900; margin: 4px 0; }}
    .header-sub {{ color: {C['muted']}; font-size: 0.9rem; }}
    .pf-highlight {{ background: {C['cyan']}15; border-left: 3px solid {C['cyan']}; }}
    table.styled {{ width: 100%; border-collapse: separate; border-spacing: 0 4px; font-size: 0.85rem; }}
    table.styled th {{ color: {C['muted']}; text-align: left; padding: 8px 12px; border-bottom: 1px solid {C['border']}; font-weight: 600; font-size: 0.7rem; text-transform: uppercase; letter-spacing: 0.5px; }}
    table.styled td {{ color: {C['text']}; padding: 8px 12px; border-bottom: 1px solid {C['border']}11; }}
    table.styled tr.pf {{ background: {C['cyan']}10; border-left: 3px solid {C['cyan']}; }}
    table.styled tr.pf td {{ color: {C['cyan']}; font-weight: 600; }}
    .critical-row {{ background: rgba(239, 68, 68, 0.08) !important; border-left: 3px solid {C['red']} !important; }}
    .badge {{ display: inline-block; padding: 2px 10px; border-radius: 20px; font-size: 0.72rem; font-weight: 600; }}
    .badge-green {{ background: {C['green']}22; color: {C['green']}; }}
    .badge-red {{ background: {C['red']}22; color: {C['red']}; }}
    .badge-gray {{ background: {C['muted']}22; color: {C['muted']}; }}
    .metric-row {{ display: flex; gap: 16px; }}
    .metric-card {{ flex: 1; border-radius: 12px; padding: 16px 20px; }}
    .mc1 {{ background: {C['cyan']}18; border: 1px solid {C['cyan']}33; }}
    .mc2 {{ background: {C['card']}; border: 1px solid {C['border']}; }}
    .mc3 {{ background: {C['orange']}18; border: 1px solid {C['orange']}33; }}
    .mc4 {{ background: {C['card2']}; border: 1px solid {C['border']}; }}
    .mc-label {{ color: {C['muted']}; font-size: 0.7rem; font-weight: 600; }}
    .mc-value {{ font-size: 1.3rem; font-weight: 800; margin-top: 2px; }}
    .mc1 .mc-value {{ color: {C['cyan']}; }}
    .mc3 .mc-value {{ color: {C['orange']}; }}
    .mc2 .mc-value, .mc4 .mc-value {{ color: {C['white']}; }}
    .part-card {{ background: {C['card']}; border: 1px solid {C['border']}; border-radius: 12px; padding: 24px; text-align: center; }}
    .part-pct {{ font-size: 2rem; font-weight: 800; margin: 12px 0 4px; }}
    .part-sub {{ color: {C['muted']}; font-size: 0.78rem; }}
    
    /* Accordion Table CSS */
    .cxc-accordion {{ width: 100%; margin-top: 10px; }}
    .cxc-item {{ margin-bottom: 4px; border-radius: 6px; overflow: hidden; background: rgba(255,255,255,0.02); border: 1px solid {C['border']}44; }}
    .cxc-header {{ display: flex; align-items: center; padding: 12px 16px; cursor: pointer; transition: background 0.3s; user-select: none; }}
    .cxc-header:hover {{ background: rgba(255,255,255,0.05); }}
    .cxc-header .chevron {{ font-size: 0.8rem; color: {C['muted']}; transition: transform 0.3s; margin-right: 15px; }}
    .cxc-toggle {{ display: none; }}
    .cxc-toggle:checked ~ .cxc-header .chevron {{ transform: rotate(90deg); }}
    .cxc-toggle:checked ~ .cxc-content {{ display: block; }}
    .cxc-content {{ display: none; padding: 0 16px 12px 42px; background: rgba(0,0,0,0.1); border-top: 1px solid {C['border']}22; }}
    
    .cxc-col-client {{ flex: 2; font-weight: 600; color: {C['white']}; }}
    .cxc-col-info {{ flex: 1; text-align: right; color: {C['muted']}; font-size: 0.85rem; }}
    .cxc-col-amount {{ flex: 1; text-align: right; color: {C['cyan']}; font-weight: 700; }}
    .cxc-col-days {{ flex: 0.8; text-align: right; font-weight: 600; }}
    
    .cxc-invoice-table {{ width: 100%; border-collapse: collapse; margin-top: 10px; font-size: 0.85rem; }}
    .cxc-invoice-table th {{ text-align: left; color: {C['muted']}; font-size: 0.7rem; text-transform: uppercase; padding: 6px 0; border-bottom: 1px solid {C['border']}33; }}
    .cxc-invoice-table td {{ padding: 8px 0; border-bottom: 1px solid {C['border']}11; }}
</style>""", unsafe_allow_html=True)

# ── Data Loading ─────────────────────────────────────────────
@st.cache_data
def load_data():
    base_dir = os.path.dirname(__file__)
    dirs_to_check = [os.path.join(base_dir, "INPUT"), base_dir]
    v_files = []
    for d in dirs_to_check:
        if os.path.exists(d):
            # Buscar archivos que CONTENGAN "veritrade" en el nombre (más flexible que startswith)
            for f in os.listdir(d):
                if "veritrade" in f.lower() and f.lower().endswith(".xlsx") and not f.startswith("~$"):
                    v_files.append(os.path.join(d, f))
                    
    if not v_files: return pd.DataFrame()
    path = max(v_files, key=os.path.getmtime)
        
    try:
        # Intentar cargar hoja 'Veritrade'
        df = pd.read_excel(path, sheet_name='Veritrade', header=5)
    except Exception:
        # Fallback a la primera hoja si 'Veritrade' no existe
        df = pd.read_excel(path, sheet_name=0, header=5)
        
    # ── Limpieza y Filtrado de Filas Fantasmas ──
    df.columns = [str(c).strip() for c in df.columns]
    if 'Fecha' in df.columns:
        df = df[df['Fecha'].notna()]
    
    df['Fecha'] = pd.to_datetime(df['Fecha'], errors='coerce')
    df = df[df['Fecha'].notna()]
    
    # Asegurar tipos numéricos para evitar errores de cálculo
    for col in ['U$ FOB Tot', 'Kg Neto', 'Partida Aduanera']:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
            
    df['FOB_KG'] = df['U$ FOB Tot'] / (df['Kg Neto'].replace(0, 1))
    df['USD_TM'] = df['FOB_KG'] * 1000
    df['TM'] = df['Kg Neto'] / 1000
    df['MES'] = df['Fecha'].dt.to_period('M')
    df['PARTIDA_TIPO'] = df['Partida Aduanera'].map({307430000:'FRESCO', 1605540000:'COCIDO'})
    
    return df
        
@st.cache_data
def load_rentabilidad():
    base_dir = os.path.dirname(__file__)
    dirs_to_check = [os.path.join(base_dir, "INPUT"), base_dir]
    r_files = []
    for d in dirs_to_check:
        if os.path.exists(d):
            r_files.extend([os.path.join(d, f) for f in os.listdir(d) if f.lower().startswith("rentabilidad") and f.lower().endswith(".xlsx")])
    if not r_files: return pd.DataFrame(), f"No files found in {dirs_to_check}"
    path = max(r_files, key=os.path.getmtime)
        
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
        sheet_name_list = [s for s in wb.sheetnames if 'resumen' in s.lower()]
        ws = wb[sheet_name_list[0]] if sheet_name_list else (wb['Resumen'] if 'Resumen' in wb.sheetnames else None)
        if not ws: return pd.DataFrame(), f"Sheet 'Resumen' missing. Found: {wb.sheetnames}"
        
        # 1. Mapeo Dinámico de Columnas (Fila 3) - Normalizado
        headers = {}
        import re, unicodedata
        for c in range(1, ws.max_column + 1):
            v = ws.cell(3, c).value
            if v:
                clean_v = re.sub(r'\s+', ' ', str(v).upper()).strip()
                clean_v = "".join(ch for ch in unicodedata.normalize('NFKD', clean_v) if not unicodedata.combining(ch))
                headers[c] = clean_v
        
        # Mapa de búsqueda (Lo que buscamos en el Excel -> El nombre que usaremos en el DataFrame)
        target_map = {
            'FECHA EMBARQUE': 'Fecha Embarque',
            'CANTIDAD': 'Cantidad',
            'PRODUCTO': 'Producto',
            'CLIENTE (RAZON SOCIAL)': 'Cliente',
            'VALOR CFR': 'VALOR CFR',
            'UTILIDAD NETA': 'UTILIDAD NETA',
            'UTILIDAD BRUTA': 'UTILIDAD BRUTA',
            'PRECIO TM': 'Precio TM',
            'VALOR FOB': 'VALOR FOB',
            'COSTO UNITARIO': 'COSTO UNITARIO'
        }
        
        final_col_map = {} # {Nombre_DF: c_idx}
        for target_excel, target_df in target_map.items():
            for c_idx, h_text in headers.items():
                if target_excel == h_text: # Prioridad Exacto
                    final_col_map[target_df] = c_idx
                    break
                elif target_excel in h_text and target_df not in final_col_map: # Parcial (solo si no hay exacto)
                    final_col_map[target_df] = c_idx
        
        # Validar mínimos
        if 'Producto' not in final_col_map or 'Cliente' not in final_col_map:
            return pd.DataFrame(), f"Columnas críticas (Producto/Cliente) no encontradas en {path}. Mapa: {final_col_map}"

        rows = []
        for r in range(4, ws.max_row + 1):
            p_val = ws.cell(r, final_col_map['Producto']).value
            if not p_val or str(p_val).strip() == '' or 'TOTAL' in str(p_val).upper():
                continue
            
            # Solo si hay algo de cantidad o utilidad para evitar filas fantasmas
            qty_val = ws.cell(r, final_col_map.get('Cantidad', 1)).value
            if qty_val is None: continue

            row_data = {}
            for target_df, c_idx in final_col_map.items():
                row_data[target_df] = ws.cell(r, c_idx).value
            rows.append(row_data)
            
        wb.close()
        df = pd.DataFrame(rows)
        
        if not df.empty:
            # Tipado Numérico
            num_cols = ['Cantidad', 'VALOR CFR', 'UTILIDAD NETA', 'UTILIDAD BRUTA', 'Precio TM', 'VALOR FOB', 'COSTO UNITARIO']
            for col in num_cols:
                if col in df.columns:
                    df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
            
            if 'Fecha Embarque' in df.columns:
                df['Fecha Embarque'] = pd.to_datetime(df['Fecha Embarque'], errors='coerce')
            
        return df, f"OK: {len(df)} filas útiles cargadas de Resumen"
    except Exception as e:
        import traceback
        traceback.print_exc()
        return pd.DataFrame(), f"Error processing {path}: {str(e)}"

    return df, "Filtered columns applied"

@st.cache_data
def load_inventario():
    import glob
    input_dir = os.path.join(os.path.dirname(__file__), "INPUT")
    # Auto-detect any inventory excel file in the directory
    inv_files = glob.glob(os.path.join(input_dir, "inventario*.xlsx"))
    if not inv_files:
        return {}, pd.DataFrame()
    # If multiple, pick the most recently modified
    path = max(inv_files, key=os.path.getmtime)
    
    wb = openpyxl.load_workbook(path, data_only=True)
    
    all_months = {}  # {month_name: DataFrame}
    
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        
        # Find MOVIMIENTO INVENTARIO column dynamically (row 1)
        mov_col = None
        for c in range(1, ws.max_column + 1):
            v = ws.cell(1, c).value
            if v and 'MOVIMIENTO' in str(v).upper():
                mov_col = c
                break
        
        if mov_col is None:
            continue
        
        # Row 2 has sub-headers: Stock Inicial, INGRESOS, SALIDAS, STOCK KG
        # These are at mov_col, mov_col+1, mov_col+2, mov_col+3
        stock_ini_col = mov_col      # Stock Inicial
        ingresos_col = mov_col + 1   # INGRESOS
        salidas_col = mov_col + 2    # SALIDAS
        stock_kg_col = mov_col + 3   # STOCK KG (final)
        
        rows = []
        for r in range(3, ws.max_row + 1):
            mat = ws.cell(r, 3).value  # C = MATERIAL
            mat_str = str(mat).strip().upper()
            if not mat or mat_str == '' or 'TOTAL' in mat_str or mat_str.startswith('STOCK') or 'STOCK A ' in mat_str:
                continue
            rows.append({
                'CODIGO_SAP': ws.cell(r, 2).value,  # B
                'MATERIAL': str(mat).strip(),
                'STOCK_INICIO_DIA1': ws.cell(r, 4).value,  # D = stock inicio del mes
                'STOCK_INICIAL': ws.cell(r, stock_ini_col).value,
                'INGRESOS': ws.cell(r, ingresos_col).value,
                'SALIDAS': ws.cell(r, salidas_col).value,
                'STOCK_KG': ws.cell(r, stock_kg_col).value,
            })
        
        df_month = pd.DataFrame(rows)
        for col in ['STOCK_INICIO_DIA1', 'STOCK_INICIAL', 'INGRESOS', 'SALIDAS', 'STOCK_KG']:
            if col in df_month.columns:
                df_month[col] = pd.to_numeric(df_month[col], errors='coerce').fillna(0)
        
        all_months[sheet_name] = df_month
    
        all_months[sheet_name] = df_month
    
    # Target common sheet names en orden de prioridad (mes más reciente primero)
    target_sheets = [
        'abril2026', 'abril', 
        'marzo2026', 'marzo', 
        'febrero2026', 'febrero', 
        'enero2026', 'enero',
        'mayo2026', 'junio2026'
    ]
    latest_key = None
    all_months_clean = {k.lower().strip().replace(" ", ""): k for k in all_months.keys()}
    
    for ts in target_sheets:
        if ts in all_months_clean:
            latest_key = all_months_clean[ts]
            break
            
    if not latest_key and all_months:
        latest_key = list(all_months.keys())[0]
            
    df_latest = pd.DataFrame()
    if latest_key and latest_key in all_months:
        df_latest = all_months[latest_key].copy()
        # Keep only rows with valid material and any movement/stock
        mask = (df_latest['STOCK_INICIAL'] > 0) | (df_latest['INGRESOS'] > 0) | (df_latest['SALIDAS'] > 0) | (df_latest['STOCK_KG'] > 0)
        df_latest = df_latest[mask]
        df_latest['TM'] = df_latest['STOCK_KG'] / 1000
    
    wb.close()
    return all_months, df_latest

@st.cache_data
def load_cxc():
    """Load CxC independent rows from Sheet1: Cliente, Nº documento, Deuda (USD HOMOL), Días Atrasados"""
    base_dir = os.path.dirname(__file__)
    dirs_to_check = [os.path.join(base_dir, "INPUT"), base_dir]
    cxc_files = []
    for d in dirs_to_check:
        if os.path.exists(d):
            cxc_files.extend([os.path.join(d, f) for f in os.listdir(d) if f.lower().startswith("cxc") and f.lower().endswith(".xlsx")])
    if not cxc_files: return pd.DataFrame()
    path = max(cxc_files, key=os.path.getmtime)
    
    wb = openpyxl.load_workbook(path, data_only=True)
    # Use 'Data' sheet if it exists, otherwise fallback to the first active sheet
    ws = wb['Data'] if 'Data' in wb.sheetnames else wb.active

    # Dynamically find column indices from Row 1
    cols = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(1, c).value
        if v:
            v_up = str(v).upper().strip()
            if 'UNIDAD DE NEGOCIO' in v_up: cols['unidad'] = c
            elif 'USD HOMOL' in v_up: cols['usd'] = c
            elif 'NOMBRE DE CLIENTE' in v_up: cols['nombre'] = c
            elif 'DOCUMENTO' in v_up and 'CLASE' not in v_up and 'FECHA' not in v_up and 'MONEDA' not in v_up: cols['ndoc'] = c
            elif 'FECHA IDEAL' in v_up: cols['fecha_ideal'] = c
            elif 'ATRASADOS' in v_up: cols['dias'] = c

    rows = []
    for r in range(2, ws.max_row + 1):
        unidad_val = ws.cell(r, cols.get('unidad', 11)).value
        unidad_str = str(unidad_val or '').upper()
        if 'EXPORT' not in unidad_str:
            continue
            
        usd_homol = ws.cell(r, cols.get('usd', 16)).value
        if not usd_homol:
            continue
        
        nombre = ws.cell(r, cols.get('nombre', 2)).value
        n_doc = ws.cell(r, cols.get('ndoc', 3)).value
        fecha_ideal = ws.cell(r, cols.get('fecha_ideal', 12)).value
        dias_q = ws.cell(r, cols.get('dias', 17)).value
        
        # Calculate dias dynamically if possible
        dias = float(dias_q) if dias_q is not None else 0
        try:
            import datetime
            if fecha_ideal and isinstance(fecha_ideal, datetime.datetime):
                dias = (datetime.datetime.now() - fecha_ideal).days
            elif fecha_ideal and isinstance(fecha_ideal, datetime.date):
                dias = (datetime.date.today() - fecha_ideal).days
        except Exception:
            pass
            
        if not nombre: continue
        rows.append({
            'Cliente': str(nombre).strip(),
            'N_Doc': str(n_doc).strip() if n_doc else '—',
            'Deuda_Pendiente': float(usd_homol),
            'Dias_Atrasados': max(0, dias),
        })
    wb.close()
    if not rows: return pd.DataFrame()
    df = pd.DataFrame(rows)
    return df.sort_values(['Cliente', 'Deuda_Pendiente'], ascending=[True, False])

@st.cache_data
def get_manual_summary(df, group_col):
    if df.empty: 
        return pd.DataFrame()
    
    # 1. Agrupar totales (Cantidad, Valor CFR, Utilidad Total)
    grouped = df.groupby(group_col).agg({
        'Cantidad': 'sum',
        'VALOR CFR': 'sum',
        'UTILIDAD NETA': 'sum'
    }).reset_index()
    
    # 2. Calcular Utilidad Mensual (Ene, Feb, Mar)
    if 'Fecha Embarque' in df.columns:
        df_copy = df.copy()
        df_copy['Mes'] = df_copy['Fecha Embarque'].dt.month
        
        for m_idx, m_name in [(1, 'Util_Ene'), (2, 'Util_Feb'), (3, 'Util_Mar')]:
            # Sumar UTILIDAD NETA por mes
            m_sum = df_copy[df_copy['Mes'] == m_idx].groupby(group_col)['UTILIDAD NETA'].sum().reset_index()
            m_sum.columns = [group_col, m_name]
            grouped = grouped.merge(m_sum, on=group_col, how='left')
    
    grouped.fillna(0, inplace=True)
    
    # Asegurar que existan columnas mensuales aunque no haya datos
    for m in ['Util_Ene', 'Util_Feb', 'Util_Mar']:
        if m not in grouped.columns:
            grouped[m] = 0.0

    # 3. Calcular Margen Neto Ponderado (Total Utilidad / Total CFR)
    grouped['MG_Neto'] = (grouped['UTILIDAD NETA'] / (grouped['VALOR CFR'].replace(0, 1)) * 100).fillna(0)
    
    # 4. Renombrar columnas para compatibilidad con el resto del dashboard
    grouped.rename(columns={
        'UTILIDAD NETA': 'Util_Neta',
        'Cantidad': 'TM_Vendidas',
        'VALOR CFR': 'Venta_CFR'
    }, inplace=True)
    # Crear alias para evitar errores de referencia
    grouped['Util_Total'] = grouped['Util_Neta']
    
    if group_col == 'Producto':
        grouped.rename(columns={'Producto': 'Producto_TD'}, inplace=True)
    elif group_col == 'Cliente':
        grouped.rename(columns={'Cliente': 'Cliente_TD'}, inplace=True)
        
    return grouped

df_raw = load_data()
df_rent, dbg_rent = load_rentabilidad()
inv_months, df_inv = load_inventario()
df_cxc = load_cxc()

# Generar tablas de rentabilidad de manera manual desde el Excel Resumen
df_td_prod = get_manual_summary(df_rent, 'Producto')
df_td_cli = get_manual_summary(df_rent, 'Cliente')

# Debug info on sidebar to diagnose Streamlit Cloud
with st.sidebar:
    with st.expander("🛠️ Debug System Info"):
        st.write(f"Rentabilidad load: {dbg_rent}")
        st.write(f"Len Rentabilidad: {len(df_rent)}")
        st.write(f"Len Inventario: {len(df_inv)}")
        st.write(f"Len TD_Prod: {len(df_td_prod)}")


# Default FOB/KG ranges per product
DEFAULT_RANGES = {
    'FILETE CONGELADO': (0.7, 2.4),
    'ALAS CONGELADAS': (0.7, 2.4),
    'NUCA': (0.5, 2.0),
    'TENTACULO': (1.0, 3.3),
    'REPRODUCTOR': (1.0, 3.3),
    'ALAS COCIDAS': (1.0, 5.0),
    'FILETE COCIDO': (1.0, 5.0),
}

def apply_fob_filter(df, ranges=None):
    """Filter rows by configurable FOB/KG ranges per product"""
    if ranges is None:
        ranges = DEFAULT_RANGES
    mask = (df['PRODUCTO'].isna()) | (df['PRODUCTO']=='')
    for prod, (lo, hi) in ranges.items():
        # Se comenta el filtro superior (<=hi) a pedido del usuario, dejando solo el inferior
        mask = mask | ((df['PRODUCTO']==prod) & (df['FOB_KG']>=lo)) # & (df['FOB_KG']<=hi))
    return df[mask]

def is_pf(name):
    return 'PERU FROST' in str(name).upper()

def fmt_usd(v): return f"${v:,.0f}" if pd.notna(v) else "—"
def fmt_tm(v): return f"{v:,.1f} TM" if pd.notna(v) else "—"
def fmt_pct(v): return f"{v:.2f}%" if pd.notna(v) else "—"

# ── Sidebar ──────────────────────────────────────────────────
with st.sidebar:
    st.markdown(f"<div style='color:{C['cyan']};font-weight:800;font-size:1.1rem;margin-bottom:16px;'>⚙️ FILTROS</div>", unsafe_allow_html=True)
    if df_raw.empty:
        base_dir = os.path.dirname(__file__)
        dirs_scanned = [os.path.join(base_dir, "INPUT"), base_dir]
        st.error(f"❌ No se encontraron archivos que contengan 'Veritrade' en: {dirs_scanned}")
        st.stop()

    min_date = df_raw['Fecha'].min().date()
    max_date = df_raw['Fecha'].max().date()

    date_range = st.date_input("📅 Rango de Fechas", value=(min_date, max_date), min_value=min_date, max_value=max_date, key="date_filter")
    if isinstance(date_range, tuple) and len(date_range)==2:
        d_start, d_end = date_range
    else:
        d_start, d_end = min_date, max_date
    st.markdown(f"<div style='color:{C['muted']};font-size:0.75rem;margin-top:8px;'>Registros totales: <b style=\"color:{C['white']}\">{len(df_raw):,}</b></div>", unsafe_allow_html=True)
    st.markdown("---")

    # ── Registro: configuración de rangos FOB/KG por producto ──
    # with st.expander("📋 Registro — Rangos FOB/KG", expanded=False):
    #     st.markdown(f"<div style='color:{C['muted']};font-size:0.78rem;margin-bottom:8px;'>Configura el intervalo de precio FOB/KG válido para cada producto. Modifica los valores y presiona <b style=\"color:{C['cyan']}\">Aplicar</b>.</div>", unsafe_allow_html=True)
    #     
    #     # Initialize widget keys in session_state ONLY if they don't exist yet
    #     for prod, (def_lo, def_hi) in DEFAULT_RANGES.items():
    #         if f"rng_lo_{prod}" not in st.session_state:
    #             st.session_state[f"rng_lo_{prod}"] = def_lo
    #         if f"rng_hi_{prod}" not in st.session_state:
    #             st.session_state[f"rng_hi_{prod}"] = def_hi
    #     
    #     for prod in DEFAULT_RANGES:
    #         st.markdown(f"<div style='color:{C['cyan']};font-size:0.82rem;font-weight:600;margin-top:6px;'>{prod}</div>", unsafe_allow_html=True)
    #         col_lo, col_hi = st.columns(2)
    #         with col_lo:
    #             st.number_input("Mín", min_value=0.0, max_value=20.0, step=0.1, key=f"rng_lo_{prod}", label_visibility="collapsed")
    #             st.caption("Mínimo")
    #         with col_hi:
    #             st.number_input("Máx", min_value=0.0, max_value=20.0, step=0.1, key=f"rng_hi_{prod}", label_visibility="collapsed")
    #             st.caption("Máximo")
    #     
    #     col_btn1, col_btn2 = st.columns(2)
    #     with col_btn1:
    #         st.button("✅ Aplicar", use_container_width=True, key="btn_apply_ranges")
    #     with col_btn2:
    #         if st.button("🔄 Restaurar", use_container_width=True, key="btn_reset_ranges"):
    #             for prod, (def_lo, def_hi) in DEFAULT_RANGES.items():
    #                 st.session_state[f"rng_lo_{prod}"] = def_lo
    #                 st.session_state[f"rng_hi_{prod}"] = def_hi
    #             st.rerun()
    # 
    # Build user_ranges directly from widget session_state keys (always current)
    user_ranges = DEFAULT_RANGES # Use defaults since filters are disabled
    # for prod in DEFAULT_RANGES:
    #     lo = st.session_state.get(f"rng_lo_{prod}", DEFAULT_RANGES[prod][0])
    #     hi = st.session_state.get(f"rng_hi_{prod}", DEFAULT_RANGES[prod][1])
    #     user_ranges[prod] = (lo, hi)
    st.markdown("---")

# Apply date filter
df_dated = df_raw[(df_raw['Fecha'].dt.date >= d_start) & (df_raw['Fecha'].dt.date <= d_end)]
df = apply_fob_filter(df_dated, user_ranges)

# Subsets
df_pf = df[df['Exportador'].apply(is_pf)]
df_fresco = df[df['PRODUCTO'].isin(PROD_FRESCO)]
df_cocido = df[df['PRODUCTO'].isin(PROD_COCIDO)]
df_classified = df[df['PRODUCTO'].notna() & (df['PRODUCTO']!='')]
df_classified = df[df['PRODUCTO'].notna() & (df['PRODUCTO']!='')]

# ── Header Banner ────────────────────────────────────────────
period_str = f"{d_start.strftime('%b %Y')} — {d_end.strftime('%b %Y')}" if d_start != d_end else d_start.strftime('%B %Y')
st.markdown(f"""<div class="header-banner">
    <div class="header-sup">🦑 DASHBOARD EJECUTIVO INTEGRAL</div>
    <div class="header-title">PERU FROST SAC</div>
    <div class="header-sub">Análisis de Exportaciones — {period_str}<br>
    <span style="font-size:0.78rem;">Información confidencial | Gerencia General & Directorio</span></div>
</div>""", unsafe_allow_html=True)

# ── KPI Row ──────────────────────────────────────────────────
df_pf_clean = df_pf[df_pf['PRODUCTO'].isin(PROD_FRESCO + PROD_COCIDO)]
fob_total_pf = df_pf_clean['U$ FOB Tot'].sum()
peso_neto_pf = df_pf_clean['Kg Neto'].sum()
fob_fresco_pf = df_pf_clean[df_pf_clean['PRODUCTO'].isin(PROD_FRESCO)]['U$ FOB Tot'].sum()
fob_cocido_pf = df_pf_clean[df_pf_clean['PRODUCTO'].isin(PROD_COCIDO)]['U$ FOB Tot'].sum()

st.markdown(f"""<div class="kpi-row">
    <div class="kpi-card c1"><div class="kpi-label">FOB TOTAL</div><div class="kpi-value">{fmt_usd(fob_total_pf)}</div><div class="kpi-sub">{period_str}</div></div>
    <div class="kpi-card c2"><div class="kpi-label">PESO NETO</div><div class="kpi-value">{peso_neto_pf/1000:,.1f} TM</div><div class="kpi-sub">{peso_neto_pf:,.0f} kg</div></div>
    <div class="kpi-card c3"><div class="kpi-label">FOB FRESCO</div><div class="kpi-value">{fmt_usd(fob_fresco_pf)}</div><div class="kpi-sub">Total Fresco PF</div></div>
    <div class="kpi-card c4"><div class="kpi-label">FOB COCIDO</div><div class="kpi-value">{fmt_usd(fob_cocido_pf)}</div><div class="kpi-sub">Total Cocido PF</div></div>
</div>""", unsafe_allow_html=True)

# ── TABS ─────────────────────────────────────────────────────
tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8, tab9, tab10, tab11 = st.tabs(["📊 Resumen", "🧩 Mix & Participación", "🌍 Mercados", "🏆 Rankings", "💲 Precios & UT0", "📈 Histórico 12M", "🔝 Top 5", "👥 Clientes", "📦 Inventario", "💰 Rentabilidad", "🗄️ Base de Datos"])

# ═══════════════ TAB 1: RESUMEN EJECUTIVO ═══════════════════
with tab1:
    st.markdown('<div class="section-title">Resumen Ejecutivo</div>', unsafe_allow_html=True)
    # Rankings
    rank_fresco = df_fresco.groupby('Exportador')['U$ FOB Tot'].sum().sort_values(ascending=False)
    rank_cocido = df_cocido.groupby('Exportador')['U$ FOB Tot'].sum().sort_values(ascending=False)
    pf_rank_f = next((i+1 for i,e in enumerate(rank_fresco.index) if is_pf(e)), "—")
    pf_rank_c = next((i+1 for i,e in enumerate(rank_cocido.index) if is_pf(e)), "—")
    # Top market
    pf_markets = df_pf_clean.groupby('Pais de Destino')['U$ FOB Tot'].sum().sort_values(ascending=False)
    top_market = pf_markets.index[0] if len(pf_markets)>0 else "—"
    top_market_pct = (pf_markets.iloc[0]/fob_total_pf*100) if len(pf_markets)>0 and fob_total_pf>0 else 0
    # Active products
    prods_activos = df_pf_clean['PRODUCTO'].nunique()
    total_prods = df_classified['PRODUCTO'].nunique()

    st.markdown(f"""<div class="info-row">
        <div class="info-card"><div class="info-label">Ranking Fresco</div><div class="info-value">#{pf_rank_f}</div><div class="info-sub">de {len(rank_fresco)} exportadores</div></div>
        <div class="info-card"><div class="info-label">Ranking Cocido</div><div class="info-value">#{pf_rank_c}</div><div class="info-sub">de {len(rank_cocido)} exportadores</div></div>
        <div class="info-card"><div class="info-label">Mercado Principal</div><div class="info-value">{top_market}</div><div class="info-sub">{top_market_pct:.1f}% del FOB</div></div>
        <div class="info-card"><div class="info-label">Productos Activos</div><div class="info-value">{prods_activos}/{total_prods}</div><div class="info-sub">con exportación en periodo</div></div>
    </div>""", unsafe_allow_html=True)

    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown('<div class="card-container"><b style="color:white;">FOB por Mercado de Destino PF</b>', unsafe_allow_html=True)
        if len(pf_markets) > 0:
            top5m = pf_markets.head(5).reset_index()
            top5m.columns = ['Pais','FOB']
            fig_m = px.bar(top5m, y='Pais', x='FOB', orientation='h', color_discrete_sequence=[C['cyan']])
            fig_m.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font_color=C['text'],
                showlegend=False, margin=dict(l=0,r=0,t=10,b=0), height=250, xaxis=dict(gridcolor='rgba(30,58,95,0.27)', tickformat='$,.0f'), yaxis=dict(autorange='reversed'))
            st.plotly_chart(fig_m, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)
    with col_b:
        st.markdown('<div class="card-container"><b style="color:white;">Mix: Fresco vs Cocido (FOB) PF</b>', unsafe_allow_html=True)
        if fob_fresco_pf > 0 or fob_cocido_pf > 0:
            fig_pie = go.Figure(go.Pie(labels=['Fresco','Cocido'], values=[fob_fresco_pf, fob_cocido_pf],
                marker_colors=[C['green'], C['orange']], hole=0.55, textinfo='label+percent', textfont_size=13))
            fig_pie.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font_color=C['text'],
                showlegend=False, margin=dict(l=0,r=0,t=10,b=0), height=250)
            st.plotly_chart(fig_pie, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)

    # Prices vs UT0 table
    st.markdown('<div class="card-container">', unsafe_allow_html=True)
    st.markdown(f'<b style="color:{C["white"]}; font-size:1.1rem;">Resumen de Precios vs Equilibrio (UT0)</b>', unsafe_allow_html=True)
    all_products = sorted(df_classified['PRODUCTO'].unique())
    rows_html = ""
    for prod in all_products:
        prod_df = df_classified[df_classified['PRODUCTO']==prod]
        pf_prod = prod_df[prod_df['Exportador'].apply(is_pf)]
        pf_utm = (pf_prod['U$ FOB Tot'].sum() / pf_prod['Kg Neto'].sum() * 1000) if pf_prod['Kg Neto'].sum()>0 else None
        
        # UT0 lookup
        ut0 = UT0_FIXED.get(prod)
        if prod == "REPRODUCTOR": ut0 = UT0_FIXED.get("REPRODUCTOR")
        
        mkt_utm = (prod_df['U$ FOB Tot'].sum() / prod_df['Kg Neto'].sum() * 1000) if prod_df['Kg Neto'].sum()>0 else None
        vs_mkt = (pf_utm - mkt_utm) if pf_utm and mkt_utm else None
        vs_ut0 = (pf_utm - ut0) if pf_utm and ut0 else None
        vs_mkt_icon = f'<span style="color:{C["green"]}">↗ +{vs_mkt:.0f}</span>' if vs_mkt and vs_mkt>=0 else (f'<span style="color:{C["red"]}">↘ {vs_mkt:.0f}</span>' if vs_mkt else "—")
        vs_ut0_icon = f'<span style="color:{C["green"]}">↗ +{vs_ut0:.0f}</span>' if vs_ut0 and vs_ut0>=0 else (f'<span style="color:{C["red"]}">↘ {vs_ut0:.0f}</span>' if vs_ut0 else "—")
        estado = '<span class="badge badge-green">Sobre UT0</span>' if vs_ut0 and vs_ut0>=0 else ('<span class="badge badge-red">Bajo UT0</span>' if vs_ut0 and vs_ut0<0 else '<span class="badge badge-gray">Sin referencia</span>')
        rows_html += f"<tr><td>{prod}</td><td style='color:{C['cyan']};font-weight:700;'>{fmt_usd(pf_utm) if pf_utm else '—'}</td><td>{fmt_usd(mkt_utm) if mkt_utm else '—'}</td><td style='color:{C['orange']};font-weight:700;'>{fmt_usd(ut0) if ut0 else '—'}</td><td>{vs_mkt_icon}</td><td>{vs_ut0_icon}</td><td>{estado}</td></tr>"
    st.markdown(f"""<table class="styled"><thead><tr><th>Producto</th><th style="color:{C['cyan']}">PERU FROST</th><th>$TM Mercado</th><th style="color:{C['orange']}">UT0</th><th>vs Mercado</th><th>vs UT0</th><th>Estado</th></tr></thead><tbody>{rows_html}</tbody></table>""", unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

    # Executive note
    sobre_count = sum(1 for prod in all_products if (lambda p: (lambda pf_p, mkt_np: ((pf_p['U$ FOB Tot'].sum()/pf_p['Kg Neto'].sum()*1000) - (mkt_np['U$ FOB Tot'].sum()/mkt_np['Kg Neto'].sum()*1000)) if pf_p['Kg Neto'].sum()>0 and mkt_np['Kg Neto'].sum()>0 else None)(df_classified[(df_classified['PRODUCTO']==p) & df_classified['Exportador'].apply(is_pf)], df_classified[(df_classified['PRODUCTO']==p) & ~df_classified['Exportador'].apply(is_pf)]))(prod) is not None and (lambda p: (lambda pf_p, mkt_np: ((pf_p['U$ FOB Tot'].sum()/pf_p['Kg Neto'].sum()*1000) - (mkt_np['U$ FOB Tot'].sum()/mkt_np['Kg Neto'].sum()*1000)) if pf_p['Kg Neto'].sum()>0 and mkt_np['Kg Neto'].sum()>0 else None)(df_classified[(df_classified['PRODUCTO']==p) & df_classified['Exportador'].apply(is_pf)], df_classified[(df_classified['PRODUCTO']==p) & ~df_classified['Exportador'].apply(is_pf)]))(prod) >= 0)
    st.markdown(f"""<div class="exec-note">
    <b>Lectura ejecutiva:</b> PERU FROST inicia el periodo con <b>{fmt_usd(fob_total_pf)}</b> en exportaciones y <b>{peso_neto_pf/1000:,.1f} TM</b> de volumen. 
    Se posiciona <b>#{pf_rank_f} en fresco</b> y <b>#{pf_rank_c} en cocido</b> a nivel nacional. 
    {sobre_count} de {prods_activos} productos activos operan por encima del precio de equilibrio (UT0).
    </div>""", unsafe_allow_html=True)

# ═══════════════ TAB 2: MIX & PARTICIPACIÓN ═════════════════
with tab2:
    st.markdown('<div class="section-title">Mix de Productos & Participación de Mercado</div>', unsafe_allow_html=True)
    col1, col2 = st.columns(2)
    pf_by_prod = df_pf[df_pf['PRODUCTO'].notna() & (df_pf['PRODUCTO']!='')].groupby('PRODUCTO').agg({'U$ FOB Tot':'sum','Kg Neto':'sum'}).reset_index()
    pf_by_prod['TM'] = pf_by_prod['Kg Neto']/1000
    colors_list = [PRODUCT_COLORS.get(p, C['muted']) for p in pf_by_prod['PRODUCTO']]
    with col1:
        st.markdown('<div class="card-container"><b style="color:white;">Mix por FOB (USD)</b>', unsafe_allow_html=True)
        if len(pf_by_prod)>0:
            fig_pf = go.Figure(go.Pie(labels=pf_by_prod['PRODUCTO'], values=pf_by_prod['U$ FOB Tot'], marker_colors=colors_list, hole=0.5, textinfo='label+percent', textfont_size=11))
            fig_pf.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font_color=C['text'], showlegend=False, margin=dict(l=0,r=0,t=10,b=0), height=350)
            st.plotly_chart(fig_pf, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)
    with col2:
        st.markdown('<div class="card-container"><b style="color:white;">Mix por Volumen (TM)</b>', unsafe_allow_html=True)
        if len(pf_by_prod)>0:
            fig_pv = go.Figure(go.Pie(labels=pf_by_prod['PRODUCTO'], values=pf_by_prod['TM'], marker_colors=colors_list, hole=0.5, textinfo='label+percent', textfont_size=11))
            fig_pv.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font_color=C['text'], showlegend=False, margin=dict(l=0,r=0,t=10,b=0), height=350)
            st.plotly_chart(fig_pv, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)

    st.markdown(f'<div class="section-title">Participación de Mercado — PERU FROST vs Industria</div>', unsafe_allow_html=True)
    total_fresco_fob = df[df['PRODUCTO'].isin(PROD_FRESCO)]['U$ FOB Tot'].sum()
    total_cocido_fob = df[df['PRODUCTO'].isin(PROD_COCIDO)]['U$ FOB Tot'].sum()
    total_fresco_tm = df[df['PRODUCTO'].isin(PROD_FRESCO)]['Kg Neto'].sum()/1000
    total_cocido_tm = df[df['PRODUCTO'].isin(PROD_COCIDO)]['Kg Neto'].sum()/1000
    pf_f_fob = df_pf[df_pf['PRODUCTO'].isin(PROD_FRESCO)]['U$ FOB Tot'].sum()
    pf_c_fob = df_pf[df_pf['PRODUCTO'].isin(PROD_COCIDO)]['U$ FOB Tot'].sum()
    pf_f_tm = df_pf[df_pf['PRODUCTO'].isin(PROD_FRESCO)]['Kg Neto'].sum()/1000
    pf_c_tm = df_pf[df_pf['PRODUCTO'].isin(PROD_COCIDO)]['Kg Neto'].sum()/1000
    parts = [
        ("Fresco FOB", pf_f_fob/total_fresco_fob*100 if total_fresco_fob else 0, fmt_usd(pf_f_fob), C['cyan']),
        ("Cocido FOB", pf_c_fob/total_cocido_fob*100 if total_cocido_fob else 0, fmt_usd(pf_c_fob), C['green']),
        ("Fresco Vol.", pf_f_tm/total_fresco_tm*100 if total_fresco_tm else 0, fmt_tm(pf_f_tm), C['blue']),
        ("Cocido Vol.", pf_c_tm/total_cocido_tm*100 if total_cocido_tm else 0, fmt_tm(pf_c_tm), C['orange']),
    ]
    cols = st.columns(4)
    for i, (label, pct, sub, color) in enumerate(parts):
        cols[i].markdown(f"""<div class="part-card"><div style="color:{C['muted']};font-weight:600;font-size:0.85rem;">{label}</div>
        <div class="part-pct" style="color:{color};">{pct:.2f}%</div><div class="part-sub">PF: {sub}</div></div>""", unsafe_allow_html=True)

# ═══════════════ TAB 3: MERCADOS ════════════════════════════
with tab3:
    st.markdown('<div class="section-title">Mercados de Destino</div>', unsafe_allow_html=True)
    st.markdown('<div class="card-container"><b style="color:white;">PERU FROST — Distribución por País y Participación de Mercado</b>', unsafe_allow_html=True)
    
    # 1. Market totals per country
    mkt_country_total = df.groupby('Pais de Destino').agg({'U$ FOB Tot':'sum', 'Kg Neto':'sum'}).rename(columns={'U$ FOB Tot':'MKT_FOB', 'Kg Neto':'MKT_KG'})
    
    # 2. Peru Frost totals per country
    pf_country_all = df_pf.groupby('Pais de Destino').agg({'U$ FOB Tot':'sum','Kg Neto':'sum'}).sort_values('U$ FOB Tot', ascending=False)
    pf_country_all['TM'] = pf_country_all['Kg Neto']/1000
    pf_country_all['USD_TM'] = (pf_country_all['U$ FOB Tot']/pf_country_all['Kg Neto']*1000).fillna(0)
    
    # 3. Calculate Participation (%) = (PF_FOB / MKT_FOB) * 100
    pf_country_all = pf_country_all.merge(mkt_country_total[['MKT_FOB']], left_index=True, right_index=True, how='left')
    pf_country_all['%PARTICIPACION'] = (pf_country_all['U$ FOB Tot'] / pf_country_all['MKT_FOB'] * 100).fillna(0)
    # Also calculate internal distribution for reference in a separate column if needed, but we'll use %PARTICIPACION as requested
    
    col_ch, col_tb = st.columns([1,1])
    with col_ch:
        top_c = pf_country_all.head(8).reset_index()
        fig_c = px.bar(top_c, y='Pais de Destino', x='U$ FOB Tot', orientation='h', color_discrete_sequence=[C['cyan']])
        fig_c.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font_color=C['text'],
            showlegend=False, margin=dict(l=0,r=0,t=10,b=0), height=300, xaxis=dict(gridcolor='rgba(30,58,95,0.27)', tickformat='$,.0f'), yaxis=dict(autorange='reversed'))
        st.plotly_chart(fig_c, use_container_width=True)
    with col_tb:
        rows_c = ""
        for pais, row in pf_country_all.iterrows():
            rows_c += f"<tr><td>{pais}</td><td style='color:{C['cyan']};font-weight:700;'>{fmt_usd(row['U$ FOB Tot'])}</td><td>{row['TM']:,.1f} TM</td><td style='font-weight:700;'>${row['USD_TM']:,.0f}</td><td style='color:{C['green']};font-weight:600;'>{row['%PARTICIPACION']:.2f}%</td></tr>"
        st.markdown(f'<table class="styled"><tr><th>País</th><th style="color:{C["cyan"]}">FOB PF</th><th>TM PF</th><th>USD/TM PF</th><th>% Participación</th></tr>{rows_c}</table>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

    # Bloques dinámicos al final
    pf_country_blocks = pf_country_all[pf_country_all['%PARTICIPACION'] > 0].copy()
    n_blocks = len(pf_country_blocks)
    if n_blocks > 0:
        st.markdown('<div style="margin-top:30px;"></div>', unsafe_allow_html=True)
        st.markdown('<div class="section-title">Participación de Mercado por País (Bloques)</div>', unsafe_allow_html=True)
        # Todos en la misma fila como solicitado
        cols_b = st.columns(n_blocks)
        for i, (pais, row) in enumerate(pf_country_blocks.iterrows()):
            with cols_b[i]:
                st.markdown(f"""
                    <div class="part-card" style="margin-bottom:16px; border-top: 3px solid {C['cyan']}; padding: 15px 10px; height: 100%; display: flex; flex-direction: column; justify-content: space-between;">
                        <div style="color:{C['white']}; font-weight:700; font-size:0.75rem; min-height:40px; display:flex; align-items:center; justify-content:center; text-transform:uppercase;">{pais}</div>
                        <div class="part-pct" style="color:{C['cyan']}; font-size:1.4rem; margin: 8px 0 2px;">{row['%PARTICIPACION']:.1f}%</div>
                        <div style="color:{C['text']}; font-weight:600; font-size:0.8rem;">{fmt_usd(row['U$ FOB Tot'])}</div>
                        <div class="part-sub" style="font-size:0.65rem; margin-top:4px;">PF en MTP Total</div>
                    </div>
                """, unsafe_allow_html=True)

    # Market comparison
    st.markdown('<div class="card-container"><b style="color:white;">Perú Total vs PERU FROST — Destinos de Exportación</b>', unsafe_allow_html=True)
    mkt_country = df.groupby('Pais de Destino')['U$ FOB Tot'].sum().sort_values(ascending=False).head(10).reset_index()
    pf_country_comp = df_pf.groupby('Pais de Destino')['U$ FOB Tot'].sum().reset_index()
    merged = mkt_country.merge(pf_country_comp, on='Pais de Destino', how='left', suffixes=('_total','_pf')).fillna(0)
    fig_comp = go.Figure()
    fig_comp.add_trace(go.Bar(name='Perú Total', x=merged['Pais de Destino'], y=merged['U$ FOB Tot_total'], marker_color='rgba(122,141,166,0.53)'))
    fig_comp.add_trace(go.Bar(name='PERU FROST', x=merged['Pais de Destino'], y=merged['U$ FOB Tot_pf'], marker_color=C['cyan']))
    fig_comp.update_layout(barmode='group', paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font_color=C['text'],
        legend=dict(orientation='h', y=-0.15), margin=dict(l=0,r=0,t=10,b=40), height=350, yaxis=dict(gridcolor='rgba(30,58,95,0.27)', tickformat='$,.0f'))
    st.plotly_chart(fig_comp, use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

# ═══════════════ TAB 4: RANKINGS TOP 15 ═════════════════════
with tab4:
    st.markdown('<div class="section-title">Rankings Top 15 Exportadores</div>', unsafe_allow_html=True)

    # Definir listas exactas de productos para los rankings
    list_congelado = ["ALAS CONGELADAS", "FILETE CONGELADO", "NUCA", "TENTACULO", "REPRODUCTOR"]
    list_cocido = ["ALAS COCIDAS", "FILETE COCIDO"]
    
    df_rank_congelado = df_classified[df_classified['PRODUCTO'].isin(list_congelado)]
    df_rank_cocido = df_classified[df_classified['PRODUCTO'].isin(list_cocido)]

    def build_ranking_table(subset_df, title):
        st.markdown(f'<div class="card-container"><b style="color:{C["white"]};font-size:1.05rem;">{title}</b>', unsafe_allow_html=True)
        # Agrupar por exportador sumando FOB (Col O) y Volumen (Col J)
        grp = subset_df.groupby('Exportador').agg({'U$ FOB Tot':'sum','Kg Neto':'sum'}).sort_values('U$ FOB Tot', ascending=False)
        total_fob = grp['U$ FOB Tot'].sum() # Base total para el % FOB
        grp = grp.head(15) # Tomar el top 15
        
        # Calcular TM y USD/TM (FOB/Volumen*1000)
        grp['TM'] = grp['Kg Neto'] / 1000
        grp['USD_TM'] = (grp['U$ FOB Tot'] / grp['Kg Neto'] * 1000).fillna(0)
        
        # Calcular % FOB = FOB del Exportador / FOB Total del grupo * 100
        grp['%FOB'] = (grp['U$ FOB Tot'] / total_fob) * 100 if total_fob > 0 else 0
        
        rows = ""
        for i, (exp, r) in enumerate(grp.iterrows()):
            is_pf_row = is_pf(exp)
            tr_class = ' class="pf"' if is_pf_row else ''
            name_display = exp[:40] + '...' if len(exp) > 40 else exp
            rank_display = f"{'🏆' if i==0 else '🥈' if i==1 else '🥉' if i==2 else f'#{i+1}'}"
            rows += f'<tr{tr_class}><td>{rank_display}</td><td>{name_display}</td><td style="font-weight:700;">{fmt_usd(r["U$ FOB Tot"])}</td><td>{r["TM"]:,.1f} TM</td><td>${r["USD_TM"]:,.0f}</td><td style="color:{C["green"]};font-weight:600;">{r["%FOB"]:.2f}%</td></tr>'
        st.markdown(f'<table class="styled"><tr><th>#</th><th>Exportador</th><th>FOB</th><th>Volumen</th><th>USD/TM</th><th>% FOB</th></tr>{rows}</table>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

    build_ranking_table(df_rank_congelado, "🧊 Top 15 Exportadores — Productos Congelados (Alas, Filete, Nuca, Tentáculo, Reproductor)")
    build_ranking_table(df_rank_cocido, "🔥 Top 15 Exportadores — Productos Cocidos (Alas Cocidas, Filete Cocido)")

# ═══════════════ TAB 5: PRECIOS & UT0 ═══════════════════════
with tab5:
    st.markdown('<div class="section-title">Análisis de Precios & Precio de Equilibrio (UT0)</div>', unsafe_allow_html=True)

    all_prods = sorted(df_classified['PRODUCTO'].unique())
    selected_prod = st.selectbox("Seleccionar Producto", all_prods, key="prod_select")

    if selected_prod:
        prod_data = df_classified[df_classified['PRODUCTO']==selected_prod]
        pf_prod = prod_data[prod_data['Exportador'].apply(is_pf)]
        mkt_no_pf = prod_data[~prod_data['Exportador'].apply(is_pf)]

        pf_utm_val = (pf_prod['U$ FOB Tot'].sum()/pf_prod['Kg Neto'].sum()*1000) if pf_prod['Kg Neto'].sum()>0 else 0
        mkt_utm_val = (prod_data['U$ FOB Tot'].sum()/prod_data['Kg Neto'].sum()*1000) if prod_data['Kg Neto'].sum()>0 else 0
        ut0_val = UT0_FIXED.get(selected_prod) or 0

        # Rankings for this product
        prod_rank = prod_data.groupby('Exportador')['U$ FOB Tot'].sum().sort_values(ascending=False)
        pf_pos = next((i+1 for i,e in enumerate(prod_rank.index) if is_pf(e)), "—")

        st.markdown(f"""<div class="metric-row">
            <div class="metric-card mc1"><div class="mc-label">PERU FROST $/TM</div><div class="mc-value">${pf_utm_val:,.0f}</div></div>
            <div class="metric-card mc2"><div class="mc-label">$TM Mercado</div><div class="mc-value">${mkt_utm_val:,.0f}</div></div>
            <div class="metric-card mc3"><div class="mc-label">UT0 (Equilibrio)</div><div class="mc-value">${ut0_val:,.0f}</div></div>
            <div class="metric-card mc4"><div class="mc-label">Posición PF</div><div class="mc-value">#{pf_pos}</div></div>
        </div>""", unsafe_allow_html=True)

        # Bar chart by exporter — use Top 15 from Rankings (Fresco or Cocido)
        is_cocido_prod = selected_prod in PROD_COCIDO
        if is_cocido_prod:
            rank_subset = df_classified[df_classified['PRODUCTO'].isin(PROD_COCIDO)]
            rank_label = "Cocido"
        else:
            rank_subset = df_classified[df_classified['PRODUCTO'].isin(PROD_FRESCO)]
            rank_label = "Fresco"
        
        # Get Top 15 exporters by FOB from the ranking category
        rank_top15 = rank_subset.groupby('Exportador')['U$ FOB Tot'].sum().sort_values(ascending=False).head(15).index.tolist()
        
        # Always include PF
        pf_in_data = [e for e in prod_data['Exportador'].unique() if is_pf(e)]
        for pf_e in pf_in_data:
            if pf_e not in rank_top15:
                rank_top15.append(pf_e)
        
        # Filter product data to only these top 15 exporters
        prod_top15 = prod_data[prod_data['Exportador'].isin(rank_top15)]
        exp_grp_all = prod_top15.groupby('Exportador').agg({'U$ FOB Tot':'sum','Kg Neto':'sum'}).sort_values('U$ FOB Tot', ascending=False)
        exp_grp_all['USD_TM'] = (exp_grp_all['U$ FOB Tot']/exp_grp_all['Kg Neto']*1000).fillna(0)
        # Remove exporters with 0 volume for this specific product
        exp_grp_all = exp_grp_all[exp_grp_all['Kg Neto'] > 0]
        
        st.markdown('<div class="card-container" style="margin-top:20px;">', unsafe_allow_html=True)
        st.markdown(f'<b style="color:{C["white"]};">USD/TM por Exportador — {selected_prod} (Top 15 {rank_label})</b>', unsafe_allow_html=True)
        
        exp_grp = exp_grp_all.reset_index()
        exp_grp['short_name'] = exp_grp['Exportador'].apply(lambda x: x[:25]+'...' if len(x)>25 else x)
        colors = [C['cyan'] if is_pf(e) else 'rgba(122,141,166,0.67)' for e in exp_grp['Exportador']]
        fig_exp = go.Figure(go.Bar(x=exp_grp['short_name'], y=exp_grp['USD_TM'], marker_color=colors))
        fig_exp.add_hline(y=mkt_utm_val, line_dash="dash", line_color=C['yellow'], annotation_text=f"Mercado ${mkt_utm_val:,.0f}", annotation_font_color=C['yellow'])
        fig_exp.add_hline(y=ut0_val, line_dash="dot", line_color=C['orange'], annotation_text=f"UT0 ${ut0_val:,.0f}", annotation_font_color=C['orange'])
        fig_exp.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font_color=C['text'],
            margin=dict(l=0,r=0,t=10,b=80), height=400, yaxis=dict(gridcolor='rgba(30,58,95,0.27)', tickformat='$,.0f'),
            xaxis=dict(tickangle=-45))
        st.plotly_chart(fig_exp, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)

    # Comparative all products
    st.markdown('<div class="card-container">', unsafe_allow_html=True)
    st.markdown(f'<b style="color:{C["white"]};">Comparativo UT0 vs PERU FROST vs Mercado — Todos los Productos</b>', unsafe_allow_html=True)
    comp_data = []
    for p in all_prods:
        pd_p = df_classified[df_classified['PRODUCTO']==p]
        pf_p = pd_p[pd_p['Exportador'].apply(is_pf)]
        pf_v = (pf_p['U$ FOB Tot'].sum()/pf_p['Kg Neto'].sum()*1000) if pf_p['Kg Neto'].sum()>0 else 0
        mk_v = (pd_p['U$ FOB Tot'].sum()/pd_p['Kg Neto'].sum()*1000) if pd_p['Kg Neto'].sum()>0 else 0
        
        # Get specialized UT0
        ut_v = UT0_FIXED.get(p)
        if ut_v is None: ut_v = 0
            
        comp_data.append({'Producto':p, 'PERU FROST':pf_v, '$TM Mercado':mk_v, 'UT0':ut_v})
    comp_df = pd.DataFrame(comp_data)
    fig_all = go.Figure()
    fig_all.add_trace(go.Bar(name='PERU FROST', x=comp_df['Producto'], y=comp_df['PERU FROST'], marker_color=C['cyan']))
    fig_all.add_trace(go.Bar(name='$TM Mercado', x=comp_df['Producto'], y=comp_df['$TM Mercado'], marker_color='rgba(122,141,166,0.53)'))
    fig_all.add_trace(go.Bar(name='UT0', x=comp_df['Producto'], y=comp_df['UT0'], marker_color=C['orange']))
    fig_all.update_layout(barmode='group', paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font_color=C['text'],
        legend=dict(orientation='h', y=-0.2), margin=dict(l=0,r=0,t=10,b=50), height=400,
        yaxis=dict(gridcolor='rgba(30,58,95,0.27)', tickformat='$,.0f'))
    st.plotly_chart(fig_all, use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

# ═══════════════ TAB 6: HISTÓRICO 12M ═══════════════════════
with tab6:
    st.markdown('<div class="section-title">Histórico Últimos 12 Meses por Producto</div>', unsafe_allow_html=True)
    st.markdown(f'<div style="color:{C["muted"]};font-size:0.85rem;margin-bottom:16px;">📌 Esta vista siempre muestra los últimos 12 meses del dataset, sin afectar el filtro de fechas.</div>', unsafe_allow_html=True)

    # Use raw data WITHOUT current fob filters to keep historical peaks intact
    df_hist = df_raw.copy()
    max_month = df_hist['Fecha'].max().to_period('M')
    min_month_12 = max_month - 11
    df_hist = df_hist[df_hist['MES'] >= min_month_12]
    
    # Filter only classified products
    df_hist_cl = df_hist[df_hist['PRODUCTO'].notna() & (df_hist['PRODUCTO']!='')]
    all_prods_hist = sorted(df_hist_cl['PRODUCTO'].unique())
    selected_prod_t6 = st.selectbox("Seleccionar Producto para Histórico", all_prods_hist, key="hist_prod")
    
    if selected_prod_t6:
        # Filter by product
        df_hist_prod = df_hist_cl[df_hist_cl['PRODUCTO'] == selected_prod_t6]
        df_hist_pf = df_hist_prod[df_hist_prod['Exportador'].apply(is_pf)]
        df_hist_mkt = df_hist_prod[~df_hist_prod['Exportador'].apply(is_pf)]
    
        # Monthly aggregations PF
        monthly_pf = df_hist_pf.groupby('MES').agg({'U$ FOB Tot':'sum','Kg Neto':'sum'}).reset_index()
        monthly_pf['TM'] = monthly_pf['Kg Neto']/1000
        monthly_pf['USD_TM'] = (monthly_pf['U$ FOB Tot']/monthly_pf['Kg Neto']*1000).fillna(0)
        monthly_pf['mes_str'] = monthly_pf['MES'].astype(str)
        
        # Monthly aggregations Market
        monthly_mkt = df_hist_mkt.groupby('MES').agg({'U$ FOB Tot':'sum','Kg Neto':'sum'}).reset_index()
        monthly_mkt['USD_TM'] = (monthly_mkt['U$ FOB Tot']/monthly_mkt['Kg Neto']*1000).fillna(0)
        monthly_mkt['mes_str'] = monthly_mkt['MES'].astype(str)
    
        # 4 KPI cards (PF) based on product
        precio_prom = monthly_pf['USD_TM'].mean() if len(monthly_pf) > 0 else 0
        precio_max = monthly_pf['USD_TM'].max() if len(monthly_pf) > 0 else 0
        precio_min = monthly_pf['USD_TM'].min() if len(monthly_pf) > 0 else 0
        vol_total = monthly_pf['TM'].sum()
        meses_activos = len(monthly_pf[monthly_pf['U$ FOB Tot'] > 0])
    
        st.markdown(f"""<div class="kpi-row">
            <div class="kpi-card c1"><div class="kpi-label">PRECIO PROMEDIO</div><div class="kpi-value">${precio_prom:,.0f}</div><div class="kpi-sub">USD/TM</div></div>
            <div class="kpi-card c2"><div class="kpi-label">PRECIO MÁXIMO</div><div class="kpi-value">${precio_max:,.0f}</div><div class="kpi-sub">USD/TM</div></div>
            <div class="kpi-card c3"><div class="kpi-label">PRECIO MÍNIMO</div><div class="kpi-value">${precio_min:,.0f}</div><div class="kpi-sub">USD/TM</div></div>
            <div class="kpi-card c4"><div class="kpi-label">VOLUMEN TOTAL</div><div class="kpi-value">{vol_total:,.1f} TM</div><div class="kpi-sub">{meses_activos}/12 meses activos</div></div>
        </div>""", unsafe_allow_html=True)
    
        # Two side-by-side charts: Price line + Volume bars
        col_h1, col_h2 = st.columns(2)
        with col_h1:
            st.markdown('<div class="card-container">', unsafe_allow_html=True)
            st.markdown(f'<b style="color:{C["white"]};">Evolución Precio USD/TM — PERU FROST vs Industria Mercado</b>', unsafe_allow_html=True)
            
            fig_price = go.Figure()
            if len(monthly_pf) > 0:
                fig_price.add_trace(go.Scatter(x=monthly_pf['mes_str'], y=monthly_pf['USD_TM'], name='PERU FROST', line=dict(color=C['cyan'], width=3), mode='lines+markers', marker=dict(size=8)))
            if len(monthly_mkt) > 0:
                fig_price.add_trace(go.Scatter(x=monthly_mkt['mes_str'], y=monthly_mkt['USD_TM'], name='Promedio Industria', line=dict(color='rgba(122,141,166,0.8)', width=2, dash='dash'), mode='lines+markers', marker=dict(size=4)))
                
            fig_price.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font_color=C['text'],
                legend=dict(orientation='h', y=-0.15), margin=dict(l=0,r=0,t=10,b=40), height=350,
                yaxis=dict(gridcolor='rgba(30,58,95,0.27)', tickformat='$,.0f'), xaxis=dict(gridcolor='rgba(30,58,95,0.13)'))
            st.plotly_chart(fig_price, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)
            
        with col_h2:
            st.markdown('<div class="card-container">', unsafe_allow_html=True)
            st.markdown(f'<b style="color:{C["white"]};">Volumen Mensual (TM) — PERU FROST</b>', unsafe_allow_html=True)
            if len(monthly_pf) > 0:
                fig_vol = go.Figure(go.Bar(x=monthly_pf['mes_str'], y=monthly_pf['TM'], marker_color=C['cyan'], marker=dict(opacity=0.8)))
                fig_vol.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font_color=C['text'],
                    margin=dict(l=0,r=0,t=10,b=40), height=350,
                    yaxis=dict(gridcolor='rgba(30,58,95,0.27)', ticksuffix=' TM'), xaxis=dict(gridcolor='rgba(30,58,95,0.13)'))
                st.plotly_chart(fig_vol, use_container_width=True)
            else:
                st.caption("No hay volumen en los últimos 12 meses.")
            st.markdown('</div>', unsafe_allow_html=True)

# ═══════════════ TAB 7: TOP 5 COMPETIDORES ══════════════════
with tab7:
    st.markdown('<div class="section-title">Top 5 Competidores — Análisis de Precios</div>', unsafe_allow_html=True)

    # Use raw un-filtered historical data to keep historical peaks intact for 12 month charts
    df_comp = df_raw.copy()
    max_m = df_comp['Fecha'].max().to_period('M')
    min_m = max_m - 11
    df_comp = df_comp[df_comp['MES'] >= min_m]
    df_comp_cl = df_comp[df_comp['PRODUCTO'].notna() & (df_comp['PRODUCTO']!='')]

    # Product selector
    prods_avail = sorted(df_comp_cl['PRODUCTO'].unique())
    selected_prod_t5 = st.selectbox("Seleccionar Producto", prods_avail, key="top5_prod")

    df_prod = df_comp_cl[df_comp_cl['PRODUCTO']==selected_prod_t5]
    # PF data
    df_prod_pf = df_prod[df_prod['Exportador'].apply(is_pf)]
    # Top 5 competitors (by Volume/Tonnes, excluding PF)
    df_prod_comp = df_prod[~df_prod['Exportador'].apply(is_pf)]
    top5_exp = df_prod_comp.groupby('Exportador')['Kg Neto'].sum().nlargest(5).index.tolist()

    # Multi-line chart: price evolution PF vs Top 5
    st.markdown('<div class="card-container">', unsafe_allow_html=True)
    st.markdown(f'<b style="color:{C["white"]};font-size:1.05rem;">{selected_prod_t5} — Evolución Precio USD/TM (12 Meses)</b>', unsafe_allow_html=True)
    fig_t5 = go.Figure()
    # PF line (thick cyan)
    pf_monthly = df_prod_pf.groupby('MES').agg({'U$ FOB Tot':'sum','Kg Neto':'sum'}).reset_index()
    pf_monthly['USD_TM'] = (pf_monthly['U$ FOB Tot']/pf_monthly['Kg Neto']*1000).fillna(0)
    pf_monthly['mes_str'] = pf_monthly['MES'].astype(str)
    if len(pf_monthly) > 0:
        fig_t5.add_trace(go.Scatter(x=pf_monthly['mes_str'], y=pf_monthly['USD_TM'], name='PERU FROST',
            line=dict(color=C['cyan'], width=4), mode='lines+markers', marker=dict(size=9)))

    comp_colors = [C['orange'], C['red'], '#a855f7', C['yellow'], C['green']]
    for i, exp in enumerate(top5_exp):
        exp_data = df_prod_comp[df_prod_comp['Exportador']==exp].groupby('MES').agg({'U$ FOB Tot':'sum','Kg Neto':'sum'}).reset_index()
        exp_data['USD_TM'] = (exp_data['U$ FOB Tot']/exp_data['Kg Neto']*1000).fillna(0)
        exp_data['mes_str'] = exp_data['MES'].astype(str)
        short_name = str(exp)[:25]
        fig_t5.add_trace(go.Scatter(x=exp_data['mes_str'], y=exp_data['USD_TM'], name=short_name,
            line=dict(color=comp_colors[i % len(comp_colors)], width=2, dash='dot'), mode='lines+markers', marker=dict(size=5)))

    fig_t5.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font_color=C['text'],
        legend=dict(orientation='h', y=-0.25, font=dict(size=10)), margin=dict(l=0,r=0,t=10,b=70), height=450,
        yaxis=dict(gridcolor='rgba(30,58,95,0.27)', tickformat='$,.0f', title='USD/TM'),
        xaxis=dict(gridcolor='rgba(30,58,95,0.13)'))
    st.plotly_chart(fig_t5, use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

    # Current filter price comparison bar chart
    df_prod_filter = df[df['PRODUCTO']==selected_prod_t5]
    last_by_exp = df_prod_filter.groupby('Exportador').agg({'U$ FOB Tot':'sum','Kg Neto':'sum'}).reset_index()
    last_by_exp['USD_TM'] = (last_by_exp['U$ FOB Tot']/last_by_exp['Kg Neto']*1000).fillna(0)
    
    # Only keep Top 5 + PF for this chart
    last_by_exp = last_by_exp[last_by_exp['Exportador'].isin(top5_exp) | last_by_exp['Exportador'].apply(is_pf)]
    last_by_exp = last_by_exp.sort_values('USD_TM', ascending=True)

    # Color: PF = cyan, others = gradient
    colors_bar = []
    for _, r in last_by_exp.iterrows():
        colors_bar.append(C['cyan'] if is_pf(r['Exportador']) else 'rgba(122,141,166,0.53)')

    st.markdown('<div class="card-container">', unsafe_allow_html=True)
    st.markdown(f'<b style="color:{C["white"]};">{selected_prod_t5} — Precio {period_str} (USD/TM)</b>', unsafe_allow_html=True)
    fig_bar = go.Figure(go.Bar(
        y=last_by_exp['Exportador'].apply(lambda x: str(x)[:30]),
        x=last_by_exp['USD_TM'], orientation='h', marker_color=colors_bar,
        text=last_by_exp['USD_TM'].apply(lambda v: f'${v:,.0f}' if v>0 else ''), textposition='outside'))
    fig_bar.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font_color=C['text'],
        margin=dict(l=0,r=60,t=10,b=0), height=max(250, len(last_by_exp)*35),
        xaxis=dict(gridcolor='rgba(30,58,95,0.27)', tickformat='$,.0f'))
    st.plotly_chart(fig_bar, use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

    # Summary table
    st.markdown('<div class="card-container">', unsafe_allow_html=True)
    st.markdown(f'<b style="color:{C["white"]};">Comparativo Top 5 vs PERU FROST — {selected_prod_t5}</b>', unsafe_allow_html=True)
    
    # PF Metrics
    pf_df_filter = df_prod_filter[df_prod_filter['Exportador'].apply(is_pf)]
    pf_usd_filter = (pf_df_filter['U$ FOB Tot'].sum() / pf_df_filter['Kg Neto'].sum() * 1000) if pf_df_filter['Kg Neto'].sum()>0 else 0
    pf_usd_12m = (pf_monthly['U$ FOB Tot'].sum() / pf_monthly['Kg Neto'].sum() * 1000) if pf_monthly['Kg Neto'].sum()>0 else 0
    pf_active = len(pf_monthly[pf_monthly['U$ FOB Tot']>0])
    
    rows_t5 = f'<tr style="background:rgba(0,255,204,0.1);"><td style="color:{C["cyan"]};font-weight:800;">PERU FROST</td><td style="color:{C["cyan"]};font-weight:800;">${pf_usd_filter:,.0f}</td><td style="color:{C["muted"]}">${pf_usd_12m:,.0f}</td><td style="color:{C["muted"]}">{pf_active}/12</td><td>—</td></tr>'
    
    for exp in top5_exp:
        # Filtered data
        exp_filter = df_prod_filter[df_prod_filter['Exportador']==exp]
        exp_usd_filter = (exp_filter['U$ FOB Tot'].sum() / exp_filter['Kg Neto'].sum() * 1000) if exp_filter['Kg Neto'].sum()>0 else 0
        
        # 12M data
        exp_12m = df_prod_comp[df_prod_comp['Exportador']==exp]
        exp_12m_grp = exp_12m.groupby('MES').agg({'U$ FOB Tot':'sum','Kg Neto':'sum'})
        meses_act = len(exp_12m_grp[exp_12m_grp['U$ FOB Tot'] > 0])
        exp_usd_12m = (exp_12m['U$ FOB Tot'].sum() / exp_12m['Kg Neto'].sum() * 1000) if exp_12m['Kg Neto'].sum()>0 else 0
        
        # Difference vs PF in the period
        diff = exp_usd_filter - pf_usd_filter 
        if exp_usd_filter > 0 and pf_usd_filter > 0:
            diff_badge = f'<span style="color:{C["green"]}">' + (f'+${diff:,.0f}' if diff>0 else f'${diff:,.0f}') + '</span>' if diff >= 0 else f'<span style="color:{C["red"]}">${diff:,.0f}</span>'
        else:
            diff_badge = '<span style="color:gray">Sin Mov.</span>'
            
        usd_str = f"${exp_usd_filter:,.0f}" if exp_usd_filter > 0 else "—"
        rows_t5 += f'<tr><td>{str(exp)[:35]}</td><td style="font-weight:700;">{usd_str}</td><td style="color:{C["muted"]}">${exp_usd_12m:,.0f}</td><td style="color:{C["muted"]}">{meses_act}/12</td><td>{diff_badge}</td></tr>'

    st.markdown(f'<table class="styled"><tr><th>Exportador</th><th>USD/TM ({period_str})</th><th>Prom. 12M</th><th>Meses Activos</th><th>vs PF (Periodo)</th></tr>{rows_t5}</table>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

# ═══════════════ TAB 8: CLIENTES ════════════════════════════
with tab8:
    st.markdown('<div class="section-title">Análisis de Clientes — PERU FROST</div>', unsafe_allow_html=True)

    # ── Rentabilidad por Cliente (Manual from Resumen) ──
    if len(df_td_cli) > 0:
        # Usamos df_td_cli que ya tiene todo el agregado manual
        df_final_cli = df_td_cli.copy()
        df_final_cli.rename(columns={'Cliente_TD': 'Cliente'}, inplace=True)
        
        total_cfr_cli = df_final_cli['Venta_CFR'].sum()
        total_util_cli = df_final_cli['Util_Total'].sum()
        total_mg_cli = (total_util_cli / total_cfr_cli * 100) if total_cfr_cli > 0 else 0
        n_cli = len(df_final_cli)
        n_rentable = len(df_final_cli[df_final_cli['Util_Total'] > 0])
        n_cli = len(df_final_cli)
        n_rentable = len(df_final_cli[df_final_cli['Util_Neta'] > 0])

        # KPI cards
        st.markdown(f"""<div class="kpi-row">
            <div class="kpi-card c1"><div class="kpi-label">TM VENDIDAS</div><div class="kpi-value">{df_final_cli['TM_Vendidas'].sum():,.1f}</div><div class="kpi-sub">Total Periodo</div></div>
            <div class="kpi-card c2"><div class="kpi-label">VENTA CFR TOTAL</div><div class="kpi-value">{fmt_usd(total_cfr_cli)}</div><div class="kpi-sub">Total periodo</div></div>
            <div class="kpi-card c3"><div class="kpi-label">UTILIDAD NETA</div><div class="kpi-value">{fmt_usd(total_util_cli)}</div><div class="kpi-sub">{'🟢' if total_util_cli > 0 else '🔴'} Total periodo</div></div>
            <div class="kpi-card c4"><div class="kpi-label">MARGEN NETO</div><div class="kpi-value">{total_mg_cli:.1f}%</div><div class="kpi-sub">{'🟢' if total_mg_cli > 0 else '🔴'} Ponderado</div></div>
        </div>""", unsafe_allow_html=True)

        # Rentabilidad table
        st.markdown('<div class="card-container">', unsafe_allow_html=True)
        st.markdown(f'<b style="color:{C["white"]};font-size:1.05rem;">Rentabilidad por Cliente</b>', unsafe_allow_html=True)
        rows_rc = ""
        for _, r in df_final_cli.sort_values('Venta_CFR', ascending=False).iterrows():
            cli = str(r['Cliente'])[:35]
            mg = r['MG_Neto']
            mg_color = C['green'] if mg > 0 else C['red']
            u_color = C['green'] if r['Util_Neta'] > 0 else C['red']
            u_ene = f'${r["Util_Ene"]:,.0f}' if r['Util_Ene'] != 0 else '—'
            u_feb = f'${r["Util_Feb"]:,.0f}' if r['Util_Feb'] != 0 else '—'
            u_mar = f'${r["Util_Mar"]:,.0f}' if r['Util_Mar'] != 0 else '—'
            rows_rc += f'<tr><td>{cli}</td><td style="color:{C["cyan"]};font-weight:700;">{r["TM_Vendidas"]:,.1f}</td><td>{fmt_usd(r["Venta_CFR"])}</td><td>{u_ene}</td><td>{u_feb}</td><td>{u_mar}</td><td style="color:{u_color};font-weight:700;">{fmt_usd(r["Util_Neta"])}</td><td style="color:{mg_color};font-weight:700;">{mg:.1f}%</td></tr>'
        st.markdown(f'<table class="styled"><tr><th>Cliente</th><th style="color:{C["cyan"]}">TM</th><th>Venta CFR</th><th>Util. Ene</th><th>Util. Feb</th><th>Util. Mar</th><th>Utilidad Neta</th><th>Margen</th></tr>{rows_rc}</table>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

        # ── Horizontal bar chart: Client profitability ──
        st.markdown('<div class="card-container">', unsafe_allow_html=True)
        st.markdown(f'<b style="color:{C["white"]};font-size:1.05rem;">¿Qué cliente es más rentable?</b>', unsafe_allow_html=True)
        cli_chart = df_final_cli[df_final_cli['Util_Neta'] != 0].sort_values('Util_Neta', ascending=True)
        colors_cli = [C['green'] if v > 0 else C['red'] for v in cli_chart['Util_Neta']]
        fig_cli_bar = go.Figure(go.Bar(
            y=cli_chart['Cliente'].apply(lambda x: str(x)[:30]),
            x=cli_chart['Util_Neta'],
            orientation='h',
            marker_color=colors_cli,
            text=[f'${v:,.0f}' for v in cli_chart['Util_Neta']],
            textposition='outside',
            textfont=dict(size=10)
        ))
        fig_cli_bar.add_vline(x=0, line_color=C['text'], line_width=1)
        fig_cli_bar.update_layout(
            paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font_color=C['text'],
            margin=dict(l=220, r=80, t=10, b=30), height=max(400, len(cli_chart)*38),
            xaxis=dict(gridcolor='rgba(30,58,95,0.27)', tickformat='$,.0f', title='Utilidad Neta (USD)', zeroline=False),
            yaxis=dict(side='left', automargin=True),
        )
        st.plotly_chart(fig_cli_bar, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)

    # ── CxC section (independent invoices) ──
    if len(df_cxc) > 0:
        st.markdown(f'<div class="section-title">Cuentas por Cobrar — Detalle Facturas</div>', unsafe_allow_html=True)
        cxc_positivo = df_cxc[df_cxc['Deuda_Pendiente'] > 0]
        cxc_total = cxc_positivo['Deuda_Pendiente'].sum()
        avg_dias = cxc_positivo['Dias_Atrasados'].mean() if len(cxc_positivo) > 0 else 0
        st.markdown(f"""<div class="info-row">
            <div class="info-card"><div class="info-label">DEUDA TOTAL PENDIENTE (USD)</div><div class="info-value">{fmt_usd(cxc_total)}</div></div>
            <div class="info-card"><div class="info-label">FACTURAS PENDIENTES</div><div class="info-value">{len(cxc_positivo)}</div></div>
            <div class="info-card"><div class="info-label">DÍAS ATRASO PROM.</div><div class="info-value" style="color:{C['red'] if avg_dias > 30 else C['yellow']}">{avg_dias:.0f} días</div></div>
        </div>""", unsafe_allow_html=True)

        # ── Visualización Ranking de Deuda ──
        cxc_agg = cxc_positivo.groupby('Cliente')['Deuda_Pendiente'].sum().reset_index().sort_values('Deuda_Pendiente', ascending=False).head(10)
        if not cxc_agg.empty:
            st.markdown('<div class="card-container">', unsafe_allow_html=True)
            st.markdown(f'<div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:10px;"><b style="color:{C["white"]}; font-size:1rem;">Ranking Top 10 — Deuda por Cliente</b><span style="color:{C["gray"]}; font-size:0.8rem;">USD</span></div>', unsafe_allow_html=True)
            
            # En Plotly horizontal, para que el mayor esté ARRIBA, el dataframe debe estar en orden ASCENDENTE (el último es el más grande y queda arriba)
            fig_cxc = px.bar(
                cxc_agg.sort_values('Deuda_Pendiente', ascending=True), 
                y='Cliente', 
                x='Deuda_Pendiente', 
                orientation='h',
                color='Deuda_Pendiente',
                color_continuous_scale='Reds',
                text_auto='.2s'
            )
            fig_cxc.update_layout(
                paper_bgcolor='rgba(0,0,0,0)', 
                plot_bgcolor='rgba(0,0,0,0)', 
                font_color=C['text'],
                showlegend=False,
                coloraxis_showscale=False,
                margin=dict(l=0,r=50,t=10,b=10), 
                height=380,
                xaxis=dict(gridcolor='rgba(255,255,255,0.05)', tickformat='$,.0f', title=None, zeroline=False),
                yaxis=dict(title=None, tickfont=dict(size=12, color=C['white']))
            )
            fig_cxc.update_traces(
                hovertemplate="<b>%{y}</b><br>Deuda: $%{x:,.2f}<extra></extra>",
                textposition='outside',
                textfont=dict(color=C['white'], size=11),
                marker_line_color='rgba(0,0,0,0)',
                marker_line_width=0
            )
            st.plotly_chart(fig_cxc, use_container_width=True, key="cxc_ranking_chart")
            st.markdown('</div>', unsafe_allow_html=True)

        # ── Detalle de Facturas (Accordion Table Premium) ──
        st.markdown(f'<div style="margin: 30px 0 15px 5px; border-left: 4px solid {C["cyan"]}; padding-left:15px;"><b style="color:{C["white"]}; font-size:1.2rem;">Detalle de Facturas por Cliente</b><br><small style="color:{C["gray"]};">Haga clic sobre un cliente para expandir sus facturas</small></div>', unsafe_allow_html=True)
        
        # Agrupamos por cliente
        facturas_por_cliente = cxc_positivo.groupby('Cliente')
        clientes_ordenados = cxc_positivo.groupby('Cliente')['Deuda_Pendiente'].sum().sort_values(ascending=False).index

        accordion_html = '<div class="cxc-accordion">'
        
        # Encabezado de la "pseudo-tabla"
        accordion_html += f'<div style="display:flex; padding: 10px 16px; font-size: 0.7rem; color:{C["muted"]}; text-transform:uppercase; font-weight:700; letter-spacing:1px;">'
        accordion_html += '<div style="flex:2; padding-left:25px;">Cliente</div>'
        accordion_html += '<div style="flex:1; text-align:right;">Nro Facturas</div>'
        accordion_html += '<div style="flex:1; text-align:right;">Deuda Total</div>'
        accordion_html += '<div style="flex:0.8; text-align:right;">Máx. Atraso</div>'
        accordion_html += '</div>'

        for i, cliente in enumerate(clientes_ordenados):
            df_cli = facturas_por_cliente.get_group(cliente).sort_values('Dias_Atrasados', ascending=False)
            deuda_total = df_cli['Deuda_Pendiente'].sum()
            max_atraso = df_cli['Dias_Atrasados'].max()
            n_facturas = len(df_cli)
            
            # Formateo de alerta de color
            atraso_color = C['red'] if max_atraso > 30 else (C['yellow'] if max_atraso > 15 else C['text'])
            item_id = f"cxc_item_{i}"
            
            # Filas de facturas dentro del acordeón
            invoice_rows = ""
            for _, r in df_cli.iterrows():
                d_atraso = r['Dias_Atrasados']
                row_style = f"background:rgba(239,68,68,0.05);" if d_atraso > 45 else ""
                invoice_rows += f'<tr style="{row_style}">'
                invoice_rows += f'<td style="color:{C["gray"]};">#{str(r["N_Doc"]).strip()}</td>'
                invoice_rows += f'<td style="text-align:right; font-weight:600;">{fmt_usd(r["Deuda_Pendiente"])}</td>'
                invoice_rows += f'<td style="text-align:right; color:{C["red"] if d_atraso > 30 else C["text"]};">{d_atraso:.0f} días</td>'
                invoice_rows += '</tr>'

            accordion_html += f'<div class="cxc-item">'
            accordion_html += f'<input type="checkbox" id="{item_id}" class="cxc-toggle">'
            accordion_html += f'<label class="cxc-header" for="{item_id}">'
            accordion_html += f'<span class="chevron">▶</span>'
            accordion_html += f'<div class="cxc-col-client">{cliente[:35]}</div>'
            accordion_html += f'<div class="cxc-col-info">{n_facturas} fact.</div>'
            accordion_html += f'<div class="cxc-col-amount">{fmt_usd(deuda_total)}</div>'
            accordion_html += f'<div class="cxc-col-days" style="color:{atraso_color};">{max_atraso:.0f} días</div>'
            accordion_html += f'</label>'
            accordion_html += f'<div class="cxc-content">'
            accordion_html += f'<table class="cxc-invoice-table">'
            accordion_html += f'<thead><tr><th>Documento</th><th style="text-align:right;">Monto</th><th style="text-align:right;">Atraso</th></tr></thead>'
            accordion_html += f'<tbody>{invoice_rows}</tbody>'
            accordion_html += f'</table></div></div>'
            
        accordion_html += '</div>'
        st.markdown(accordion_html, unsafe_allow_html=True)



# ═══════════════ TAB 9: INVENTARIO ══════════════════════════
with tab9:
    st.markdown('<div class="section-title">Inventario de Producto Terminado</div>', unsafe_allow_html=True)

    if len(df_inv) > 0:
        total_stock_kg = df_inv['STOCK_KG'].sum()
        total_stock_tm = total_stock_kg / 1000
        n_products = len(df_inv)
        latest_sheet = list(inv_months.keys())[0] if inv_months else 'Último mes'

        st.markdown(f"""<div class="kpi-row">
            <div class="kpi-card c1"><div class="kpi-label">STOCK ACTUAL</div><div class="kpi-value">{total_stock_tm:,.1f} TM</div><div class="kpi-sub">{total_stock_kg:,.0f} kg</div></div>
            <div class="kpi-card c2"><div class="kpi-label">PRODUCTOS EN STOCK</div><div class="kpi-value">{n_products}</div><div class="kpi-sub">Con stock > 0</div></div>
            <div class="kpi-card c3"><div class="kpi-label">INGRESOS</div><div class="kpi-value">{df_inv['INGRESOS'].sum()/1000:,.1f} TM</div><div class="kpi-sub">{df_inv['INGRESOS'].sum():,.0f} kg</div></div>
            <div class="kpi-card c4"><div class="kpi-label">SALIDAS</div><div class="kpi-value">{df_inv['SALIDAS'].sum()/1000:,.1f} TM</div><div class="kpi-sub">{df_inv['SALIDAS'].sum():,.0f} kg</div></div>
        </div>""", unsafe_allow_html=True)

        # Table: show data exactly as-is from latest month, same order as Excel
        st.markdown('<div class="card-container">', unsafe_allow_html=True)
        st.markdown(f'<b style="color:{C["white"]};font-size:1.05rem;">Movimiento de Inventario — {latest_sheet}</b>', unsafe_allow_html=True)
        rows_inv = ''
        for _, r in df_inv.iterrows():
            sap = str(r['CODIGO_SAP']).strip() if r['CODIGO_SAP'] else '—'
            mat = str(r['MATERIAL'])[:45]
            # Convert all to TM
            ini = (r['STOCK_INICIAL'] or 0) / 1000
            ing = (r['INGRESOS'] or 0) / 1000
            sal = (r['SALIDAS'] or 0) / 1000
            fin = (r['STOCK_KG'] or 0) / 1000
            ini_str = f'{ini:,.2f}' if ini > 0 else '—'
            ing_str = f'{ing:,.2f}' if ing > 0 else '—'
            sal_str = f'{sal:,.2f}' if sal > 0 else '—'
            fin_str = f'{fin:,.2f}'
            rows_inv += f'<tr><td style="font-size:0.8rem;color:{C["muted"]}">{sap}</td><td>{mat}</td><td style="text-align:right;">{ini_str}</td><td style="text-align:right;color:{C["green"]};">{ing_str}</td><td style="text-align:right;color:{C["red"]};">{sal_str}</td><td style="text-align:right;color:{C["cyan"]};font-weight:700;">{fin_str}</td></tr>'
        st.markdown(f'<table class="styled"><tr><th>Cód. SAP</th><th>Material</th><th>Stock Inicial (TM)</th><th style="color:{C["green"]}">Ingresos (TM)</th><th style="color:{C["red"]}">Salidas (TM)</th><th style="color:{C["cyan"]}">Stock Final (TM)</th></tr>{rows_inv}</table>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)
    else:
        st.warning("⚠️ **Inventario vacío o fórmulas no calculadas:** No se encontraron productos con stock en el mes actual. Si acabas de subir el Excel desde Drive o un software externo, **ábrelo en tu Excel de escritorio y presiona 'Guardar'**. Esto obliga a Excel a calcular las fórmulas para que el Dashboard pueda leer los valores finales.")

# ═══════════════ TAB 10: RENTABILIDAD ═══════════════════════
with tab10:
    st.markdown('<div class="section-title">Análisis de Rentabilidad por Producto</div>', unsafe_allow_html=True)

    if len(df_td_prod) > 0 and len(df_rent) > 0:
        rent_filter = df_rent  # No se aplica filtro de fechas global por petición del usuario para respetar los consolidados

        # ── Static KPIs from Resumen totals ──
        total_tm_vendidas = rent_filter['Cantidad'].sum()
        total_venta_cfr = rent_filter['VALOR CFR'].sum()
        total_util_neta = df_td_prod['Util_Total'].sum()
        total_mg_neto = (total_util_neta / total_venta_cfr * 100) if total_venta_cfr > 0 else 0

        st.markdown(f"""<div class="kpi-row">
            <div class="kpi-card c1"><div class="kpi-label">TM VENDIDAS</div><div class="kpi-value">{total_tm_vendidas:,.1f}</div><div class="kpi-sub">Total periodo</div></div>
            <div class="kpi-card c2"><div class="kpi-label">VENTA CFR</div><div class="kpi-value">{fmt_usd(total_venta_cfr)}</div><div class="kpi-sub">Total periodo</div></div>
            <div class="kpi-card c3"><div class="kpi-label">UTILIDAD NETA</div><div class="kpi-value">{fmt_usd(total_util_neta)}</div><div class="kpi-sub">{'🟢' if total_util_neta > 0 else '🔴'} Total periodo</div></div>
            <div class="kpi-card c4"><div class="kpi-label">MARGEN NETO</div><div class="kpi-value">{total_mg_neto:.1f}%</div><div class="kpi-sub">{'🟢' if total_mg_neto > 0 else '🔴'} Ponderado</div></div>
        </div>""", unsafe_allow_html=True)

        # ── Usamos el agregado manual por producto ──
        df_final_prod = df_td_prod.copy()
        df_final_prod.rename(columns={'Producto_TD': 'Producto'}, inplace=True)

        # Table
        st.markdown('<div class="card-container">', unsafe_allow_html=True)
        st.markdown(f'<b style="color:{C["white"]};font-size:1.05rem;">Rentabilidad por Producto</b>', unsafe_allow_html=True)
        rows_rent = ""
        for _, r in df_final_prod.sort_values('Venta_CFR', ascending=False).iterrows():
            prod = str(r['Producto'])[:30]
            mg = r['MG_Neto']
            mg_color = C['green'] if mg > 0 else C['red']
            u_color = C['green'] if r['Util_Neta'] > 0 else C['red']
            if mg > 15:
                clase = f'<span class="badge badge-green">Estrella</span>'
            elif mg > 0:
                clase = f'<span style="color:{C["orange"]}">Moderado</span>'
            else:
                clase = f'<span class="badge badge-red">Crítico</span>'
            u_ene = f'${r["Util_Ene"]:,.0f}' if r['Util_Ene'] != 0 else '—'
            u_feb = f'${r["Util_Feb"]:,.0f}' if r['Util_Feb'] != 0 else '—'
            u_mar = f'${r["Util_Mar"]:,.0f}' if r['Util_Mar'] != 0 else '—'
            rows_rent += f'<tr><td>{prod}</td><td style="color:{C["cyan"]};font-weight:700;">{r["TM_Vendidas"]:,.1f}</td><td>{fmt_usd(r["Venta_CFR"])}</td><td>{u_ene}</td><td>{u_feb}</td><td>{u_mar}</td><td style="color:{u_color};font-weight:700;">{fmt_usd(r["Util_Neta"])}</td><td style="color:{mg_color};font-weight:700;">{mg:.1f}%</td><td>{clase}</td></tr>'
        st.markdown(f'<table class="styled"><tr><th>Producto</th><th style="color:{C["cyan"]}">TM</th><th>Venta CFR</th><th>Util. Ene</th><th>Util. Feb</th><th>Util. Mar</th><th>Utilidad Neta</th><th>Margen</th><th>Tipo</th></tr>{rows_rent}</table>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

        # ── Horizontal bar chart: Product profitability ──
        st.markdown('<div class="card-container">', unsafe_allow_html=True)
        st.markdown(f'<b style="color:{C["white"]};font-size:1.05rem;">¿Qué producto tiene mayor rentabilidad?</b>', unsafe_allow_html=True)
        prod_chart = df_final_prod[df_final_prod['Util_Neta'] != 0].sort_values('Util_Neta', ascending=True)
        colors_prod = [C['green'] if v > 0 else C['red'] for v in prod_chart['Util_Neta']]
        fig_prod_bar = go.Figure(go.Bar(
            y=prod_chart['Producto'].apply(lambda x: str(x)[:25]),
            x=prod_chart['Util_Neta'],
            orientation='h',
            marker_color=colors_prod,
            text=[f'${v:,.0f}' for v in prod_chart['Util_Neta']],
            textposition='outside',
            textfont=dict(size=10)
        ))
        fig_prod_bar.add_vline(x=0, line_color=C['text'], line_width=1)
        fig_prod_bar.update_layout(
            paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font_color=C['text'],
            margin=dict(l=180, r=80, t=10, b=30), height=max(350, len(prod_chart)*40),
            xaxis=dict(gridcolor='rgba(30,58,95,0.27)', tickformat='$,.0f', title='Utilidad Neta (USD)', zeroline=False),
            yaxis=dict(side='left', automargin=True),
        )
        st.plotly_chart(fig_prod_bar, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)

        # Stacked cost chart
        st.markdown('<div class="card-container">', unsafe_allow_html=True)
        st.markdown(f'<b style="color:{C["white"]};">Costo Producción vs Margen Bruto (USD/TM)</b>', unsafe_allow_html=True)
        # Use original aggregation for cost/price data
        rent_grp = df_rent.groupby('Producto').agg({
            'Cantidad':'sum', 'VALOR FOB':'sum', 'VALOR CFR':'sum',
            'COSTO UNITARIO':'mean', 'Precio TM':'mean',
            'UTILIDAD BRUTA':'sum'
        }).reset_index()
        rent_sorted = rent_grp.sort_values('VALOR FOB', ascending=False).head(10)
        costos = []
        margenes = []
        for _, r in rent_sorted.iterrows():
            costo = r['COSTO UNITARIO'] if pd.notna(r['COSTO UNITARIO']) else 0
            precio = r['Precio TM'] if pd.notna(r['Precio TM']) else 0
            margen = max(0, precio - costo)
            costos.append(costo)
            margenes.append(margen)
        fig_r = go.Figure()
        fig_r.add_trace(go.Bar(x=rent_sorted['Producto'].apply(lambda x: str(x)[:20]), y=costos, name='Costo Producción', marker_color='rgba(122,141,166,0.6)'))
        fig_r.add_trace(go.Bar(x=rent_sorted['Producto'].apply(lambda x: str(x)[:20]), y=margenes, name='Margen Bruto', marker_color=C['green']))
        fig_r.update_layout(barmode='stack', paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font_color=C['text'],
            legend=dict(orientation='h', y=-0.25), margin=dict(l=0,r=0,t=10,b=80), height=380,
            yaxis=dict(gridcolor='rgba(30,58,95,0.27)', tickformat='$,.0f'), xaxis=dict(tickangle=-45))
        st.plotly_chart(fig_r, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)

        # Executive insight
        criticos_list = [str(r['Producto'])[:20] for _, r in df_final_prod.iterrows() if r['Util_Neta'] < 0]
        st.markdown(f"""<div class="exec-note">
            <b>Lectura ejecutiva — Rentabilidad:</b> TM vendidas: <b>{total_tm_vendidas:,.1f}</b> | Venta CFR: <b>{fmt_usd(total_venta_cfr)}</b> | Utilidad Neta: <b>{fmt_usd(total_util_neta)}</b> | Margen Neto: <b>{total_mg_neto:.1f}%</b>.
            {('<span style="color:'+C['red']+'">⚠️ Productos con margen negativo: '+', '.join(criticos_list[:4])+'</span>') if criticos_list else '✅ Todos los productos operan con margen positivo.'}
        </div>""", unsafe_allow_html=True)
    else:
        st.info("No se encontró el archivo de rentabilidad en INPUT/")


# ═══════════════ TAB 11: BASE DE DATOS ══════════════════════
with tab11:
    st.markdown('<div class="section-title">Base de Datos Explorable</div>', unsafe_allow_html=True)
    st.markdown(f'<div style="color:{C["muted"]};font-size:0.85rem;margin-bottom:16px;">Tabla con datos filtrados y estandarizados. Usa los filtros de la sidebar para ajustar el rango.</div>', unsafe_allow_html=True)

    # Build standardized table
    db = df_classified.copy()
    db['Año'] = db['Fecha'].dt.year
    db['Mes'] = db['Fecha'].dt.month_name()
    db['Es_PF'] = db['Exportador'].apply(is_pf)

    # Select columns for display
    db_display = db[['Fecha','Año','Mes','Exportador','Importador','Pais de Destino','PRODUCTO','PARTIDA_TIPO','TM','Kg Neto','U$ FOB Tot','USD_TM']].copy()
    db_display.columns = ['Fecha','Año','Mes','Exportador','Cliente','País Destino','Producto','Tipo','TM','KG Neto','FOB USD','USD/TM']

    # KPIs
    st.markdown(f"""<div class="info-row">
        <div class="info-card"><div class="info-label">REGISTROS</div><div class="info-value">{len(db_display):,}</div></div>
        <div class="info-card"><div class="info-label">EXPORTADORES</div><div class="info-value">{db_display['Exportador'].nunique()}</div></div>
        <div class="info-card"><div class="info-label">PRODUCTOS</div><div class="info-value">{db_display['Producto'].nunique()}</div></div>
        <div class="info-card"><div class="info-label">PAÍSES</div><div class="info-value">{db_display['País Destino'].nunique()}</div></div>
    </div>""", unsafe_allow_html=True)

    # Filters for the table
    col_f1, col_f2, col_f3 = st.columns(3)
    with col_f1:
        tipo_filter = st.multiselect("Tipo", db_display['Tipo'].dropna().unique(), default=db_display['Tipo'].dropna().unique(), key="db_tipo")
    with col_f2:
        prod_filter = st.multiselect("Producto", db_display['Producto'].dropna().unique(), default=db_display['Producto'].dropna().unique(), key="db_prod")
    with col_f3:
        pais_filter = st.multiselect("País", db_display['País Destino'].dropna().unique().tolist()[:20], default=db_display['País Destino'].dropna().unique().tolist()[:20], key="db_pais")

    filtered_db = db_display[db_display['Tipo'].isin(tipo_filter) & db_display['Producto'].isin(prod_filter) & db_display['País Destino'].isin(pais_filter)]

    st.dataframe(filtered_db.sort_values('Fecha', ascending=False), use_container_width=True, height=400)

    # ── Resumen Ejecutivo Consolidado ──
    st.markdown('<div class="card-container">', unsafe_allow_html=True)
    st.markdown(f'<b style="color:{C["white"]};font-size:1.1rem;">Resumen Ejecutivo Consolidado</b>', unsafe_allow_html=True)
    st.markdown(f'<div style="color:{C["muted"]};font-size:0.8rem;margin-bottom:12px;">Cruce integral: Precios, UT0, Márgenes y Rankings por producto</div>', unsafe_allow_html=True)

    # Build consolidated table
    pf_data = df_classified[df_classified['Exportador'].apply(is_pf)]
    all_data = df_classified
    prods_consol = sorted(pf_data[pf_data['PRODUCTO'].notna() & (pf_data['PRODUCTO']!='')]['PRODUCTO'].unique())
    rows_consol = ""
    for prod in prods_consol:
        pf_p = pf_data[pf_data['PRODUCTO']==prod]
        all_p = all_data[all_data['PRODUCTO']==prod]
        pf_utm = pf_p['U$ FOB Tot'].sum()/pf_p['Kg Neto'].sum()*1000 if pf_p['Kg Neto'].sum()>0 else 0
        mkt_utm = all_p['U$ FOB Tot'].sum()/all_p['Kg Neto'].sum()*1000 if all_p['Kg Neto'].sum()>0 else 0
        
        # UT0 ESTÁTICO (CORRECCIÓN)
        ut0 = UT0_FIXED.get(prod)
        vs_mkt = pf_utm - mkt_utm
        
        # Ranking
        rank_list = all_p.groupby('Exportador')['U$ FOB Tot'].sum().sort_values(ascending=False)
        pf_names = [e for e in rank_list.index if is_pf(e)]
        rank_pos = list(rank_list.index).index(pf_names[0])+1 if pf_names else 0
        rank_total = len(rank_list)
        
        # Margin from rentabilidad
        mg_str = "—"
        if len(df_rent) > 0:
            rent_p = df_rent[df_rent['Producto'].astype(str).str.contains(prod[:8], case=False, na=False)]
            if len(rent_p) > 0:
                util_sum = rent_p['UTILIDAD NETA'].sum()
                cfr_sum = rent_p['VALOR CFR'].sum()
                if cfr_sum > 0:
                    mg_pct = (util_sum / cfr_sum) * 100
                else:
                    mg_pct = 0
                mg_color = C['green'] if mg_pct > 0 else C['red']
                mg_str = f'<span style="color:{mg_color};font-weight:700;">{mg_pct:.1f}%</span>'
                
        vs_mkt_badge = f'<span style="color:{C["green"]}">+{vs_mkt:,.0f}</span>' if vs_mkt >= 0 else f'<span style="color:{C["red"]}">{vs_mkt:,.0f}</span>'
        
        if ut0 is not None:
            vs_ut0 = pf_utm - ut0
            vs_ut0_badge = f'<span style="color:{C["green"]}">+{vs_ut0:,.0f}</span>' if vs_ut0 >= 0 else f'<span style="color:{C["red"]}">{vs_ut0:,.0f}</span>'
            ut0_str = f'${ut0:,.0f}'
        else:
            vs_ut0_badge = "—"
            ut0_str = "—"
            
        rank_badge = f'<span style="color:{C["cyan"]};font-weight:700;">#{rank_pos}/{rank_total}</span>'
        rows_consol += f'<tr><td style="font-weight:700;">{prod}</td><td style="color:{C["cyan"]};font-weight:700;">${pf_utm:,.0f}</td><td>${mkt_utm:,.0f}</td><td style="color:{C["orange"]};font-weight:700;">{ut0_str}</td><td>{vs_mkt_badge}</td><td>{vs_ut0_badge}</td><td>{mg_str}</td><td>{rank_badge}</td></tr>'
    st.markdown(f'<table class="styled"><tr><th>Producto</th><th style="color:{C["cyan"]}">PF USD/TM</th><th>$TM Mercado</th><th style="color:{C["orange"]}">UT0</th><th>vs Mercado</th><th>vs UT0</th><th>Margen %</th><th>Ranking</th></tr>{rows_consol}</table>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

    # ── Download buttons ──
    st.markdown('<div class="section-title">Archivos Fuente del Dashboard</div>', unsafe_allow_html=True)
    st.markdown(f'<div style="color:{C["muted"]};font-size:0.85rem;margin-bottom:16px;">Descarga la tabla procesada o los archivos Excel originales utilizados para auditar los datos.</div>', unsafe_allow_html=True)
    
    dl_c1, dl_c2, dl_c3, dl_c4, dl_c5 = st.columns(5)
    
    csv = filtered_db.to_csv(index=False).encode('utf-8')
    dl_c1.download_button("📥 Descargar Tabla (CSV)", csv, "peru_frost_base_datos.csv", "text/csv", use_container_width=True)
    
    input_dir = os.path.join(os.path.dirname(__file__), "INPUT")
    
    # 1. Veritrade Excel
    v_path = os.path.join(input_dir, "Veritrade_MARCOS@PERUFROST.COM_PE_E_20260327145809_CLASIFICADO.xlsx")
    if os.path.exists(v_path):
        with open(v_path, "rb") as f: v_bytes = f.read()
        dl_c2.download_button("📊 Excel Veritrade", v_bytes, "Veritrade_Clasificado.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        
    # 2. Rentabilidad
    r_path = os.path.join(input_dir, "Rentabilidad por FP 2026.xlsx")
    if os.path.exists(r_path):
        with open(r_path, "rb") as f: r_bytes = f.read()
        dl_c3.download_button("💰 Excel Rentabilidad", r_bytes, "Rentabilidad_2026.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        
    # 3. Inventario
    i_path = os.path.join(input_dir, "inventario PT al 23-03-2026 (2).xlsx")
    if os.path.exists(i_path):
        with open(i_path, "rb") as f: i_bytes = f.read()
        dl_c4.download_button("📦 Excel Inventario", i_bytes, "Inventario_Kardex.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        
    # 4. CxC
    c_path = os.path.join(input_dir, "CxC al 25-3-2026.xlsx")
    if os.path.exists(c_path):
        with open(c_path, "rb") as f: c_bytes = f.read()
        dl_c5.download_button("👥 Excel CxC", c_bytes, "Cartera_CxC.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

# ── Footer ───────────────────────────────────────────────────
st.markdown(f"""<div style="text-align:center;margin-top:40px;padding:20px;border-top:1px solid {C['border']};">
    <div style="color:{C['cyan']};font-weight:700;">PERU FROST S.A.C.</div>
    <div style="color:{C['muted']};font-size:0.75rem;">Dashboard Ejecutivo Integral — Análisis de Exportaciones {period_str}<br>
    Información confidencial para uso exclusivo de Gerencia General y Directorio</div>
</div>""", unsafe_allow_html=True)
