"""Helpers for locating sheets in the inventory workbook."""


def find_sheet(sheetnames, target):
    """Return the sheet name matching target, ignoring case and surrounding spaces.

    Returns None when no sheet matches.
    """
    wanted = target.strip().lower()
    for name in sheetnames:
        if name.strip().lower() == wanted:
            return name
    return None


def parse_resumen_inventario(wb):
    """Extract Stock Libre breakdown for Planta and China from the RESUMEN sheet.

    Reads Column C (Producto) and Column D (TM en Stock Libre).
    Returns a dict with 'planta', 'china', 'total' DataFrames and scalar totals,
    or None if no RESUMEN sheet is found.
    """
    import pandas as pd

    resumen_sheet = None
    for name in wb.sheetnames:
        if 'resumen' in name.strip().lower():
            resumen_sheet = name
            break

    if not resumen_sheet:
        return None

    ws = wb[resumen_sheet]
    planta_rows = []
    china_rows = []
    total_rows = []

    current_section = None

    for r in range(1, ws.max_row + 1):
        c_val = ws.cell(r, 3).value  # Column C: Producto
        d_val = ws.cell(r, 4).value  # Column D: TM en Stock Libre

        # Check for section headers across first few columns
        row_text = ' '.join(str(ws.cell(r, c).value or '') for c in range(1, 12)).upper()
        if 'RENTABILIDAD DEL INVENTARIO' in row_text:
            current_section = 'PLANTA'
            continue
        elif 'YANTAI' in row_text or 'CHINA' in row_text or 'ALMACÉN DE' in row_text or 'ALMACEN DE' in row_text:
            current_section = 'CHINA'
            continue
        elif 'INVENTARIO TOTAL VALORIZADO' in row_text:
            current_section = 'TOTAL'
            continue

        c_str = str(c_val).strip() if c_val is not None else ''
        if not c_str or c_str.upper() in ('PRODUCTO', 'PERU FROST') or 'AVANCE COMERCIAL' in c_str.upper():
            continue
        if c_str.upper() == 'TOTAL':
            if current_section == 'PLANTA':
                current_section = None
            elif current_section == 'CHINA':
                current_section = None
            elif current_section == 'TOTAL':
                break
            continue

        tm_val = pd.to_numeric(d_val, errors='coerce')
        tm_val = 0.0 if pd.isna(tm_val) else float(tm_val)

        item = {'PRODUCTO': c_str, 'STOCK_LIBRE_TM': tm_val}
        if current_section == 'PLANTA':
            planta_rows.append(item)
        elif current_section == 'CHINA':
            china_rows.append(item)
        elif current_section == 'TOTAL':
            total_rows.append(item)

    df_planta = pd.DataFrame(planta_rows)
    df_china = pd.DataFrame(china_rows)
    df_total = pd.DataFrame(total_rows)

    tot_planta = float(df_planta['STOCK_LIBRE_TM'].sum()) if not df_planta.empty else 0.0
    tot_china = float(df_china['STOCK_LIBRE_TM'].sum()) if not df_china.empty else 0.0
    tot_total = float(df_total['STOCK_LIBRE_TM'].sum()) if not df_total.empty else (tot_planta + tot_china)

    return {
        'planta': df_planta,
        'china': df_china,
        'total': df_total,
        'tot_planta': tot_planta,
        'tot_china': tot_china,
        'tot_total': tot_total,
    }

