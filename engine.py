"""
explosion_insumos.py
--------------------
Replica la macro VBA GenerarExplosionInsumos() en Python puro.
Lee el archivo .xlsm, procesa las hojas PPTO + "item N",
y genera la hoja "EXPLOSION INSUMOS" con totales por categoría.

Uso:
    python explosion_insumos.py <ruta_archivo.xlsm>
"""

import sys
import re
from collections import defaultdict
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────

def rgb_to_hex(r, g, b):
    return f"{r:02X}{g:02X}{b:02X}"


def safe_float(val):
    try:
        return float(val) if val not in (None, "", " ") else 0.0
    except (ValueError, TypeError):
        return 0.0


def thin_border():
    s = Side(style="thin")
    return Border(left=s, right=s, top=s, bottom=s)


def apply_header_style(cell, hex_color="00467F", font_color="FFFFFF"):
    cell.font = Font(bold=True, color=font_color, name="Arial")
    cell.fill = PatternFill("solid", fgColor=hex_color)
    cell.alignment = Alignment(horizontal="center", vertical="center")


def apply_cat_style(cell, hex_color):
    cell.fill = PatternFill("solid", fgColor=hex_color)
    cell.font = Font(bold=True, name="Arial")


def money_fmt(cell):
    cell.number_format = '$ #,##0'


def qty_fmt(cell):
    cell.number_format = '#,##0.00'


def pct_fmt(cell):
    cell.number_format = '0.00%'


# ─────────────────────────────────────────────
# 1. LEER TABLA DE CATEGORÍAS (hoja AUX)
# ─────────────────────────────────────────────

def leer_categorias(ws_aux):
    """
    Espera columnas:
      A=TEXTO_APU | B=CATEGORIA | C=NUM | D=COLOR_RGB | E=TIPO
    """
    cats = []
    for row in ws_aux.iter_rows(min_row=2, values_only=True):
        texto_apu = str(row[0] or "").strip()
        if not texto_apu:
            continue
        categoria = str(row[1] or "").strip()
        num       = str(row[2] or "").strip()
        color_str = str(row[3] or "220,220,220").strip()
        tipo      = str(row[4] or "ESTANDAR").strip().upper()

        parts = [p.strip() for p in color_str.split(",")]
        try:
            r, g, b = int(parts[0]), int(parts[1]), int(parts[2])
        except Exception:
            r, g, b = 220, 220, 220

        cats.append({
            "texto_apu": texto_apu,
            "categoria": categoria,
            "num": num,
            "hex": rgb_to_hex(r, g, b),
            "tipo": tipo,
        })
    return cats


# ─────────────────────────────────────────────
# 2. LEER CANTIDADES DEL PRESUPUESTO (hoja PPTO)
# ─────────────────────────────────────────────

def leer_cant_ppto(ws_ppto):
    """Devuelve {item_num: cantidad}"""
    cant_map = {}
    for row in ws_ppto.iter_rows(min_row=3, max_row=200, values_only=True):
        item_val = row[0]
        if item_val is None:
            continue
        try:
            item_num = int(item_val)
            cant     = safe_float(row[3])
            cant_map[item_num] = cant
        except (ValueError, TypeError):
            pass
    return cant_map


# ─────────────────────────────────────────────
# 3. DETECTAR CATEGORÍA de una celda A
# ─────────────────────────────────────────────

def detectar_categoria(cell_a, cats):
    for cat in cats:
        if cat["texto_apu"].upper() in cell_a.upper():
            return cat["categoria"], cat["tipo"]
    return None, None


# ─────────────────────────────────────────────
# 4. PROCESAR UNA HOJA "item N"
# ─────────────────────────────────────────────

def procesar_item(ws_item, item_num, cant_ppto, cats, diccionarios):
    """Acumula insumos al diccionario global. Devuelve base del ítem."""
    if cant_ppto == 0:
        return 0.0

    last_row = ws_item.max_row
    categoria_actual = None
    tipo_actual      = None
    base_item        = 0.0

    d_insumos    = diccionarios["insumos"]
    d_unidades   = diccionarios["unidades"]
    d_cantidades = diccionarios["cantidades"]
    d_valores    = diccionarios["valores"]
    d_categorias = diccionarios["categorias"]
    d_vunit      = diccionarios["vunit"]
    d_claves     = diccionarios["claves"]

    for row in ws_item.iter_rows(min_row=11, max_row=last_row, values_only=True):
        cell_a = str(row[0] or "").strip()
        cell_b = str(row[1] or "").strip()
        cell_f = str(row[5] or "").strip()   # columna F (índice 5)

        # Detectar sección / categoría
        cat, tipo = detectar_categoria(cell_a, cats)
        if cat:
            categoria_actual = cat
            tipo_actual      = tipo

        # Salir de categoría si encontramos TOTAL COSTO
        if "TOTAL COSTO" in cell_a.upper():
            categoria_actual = None
            tipo_actual      = None
            continue

        if categoria_actual is None:
            continue

        # Filtrar filas de cabecera o subtotales
        if not cell_b:
            continue
        if "DESCRIPCI" in cell_b.upper():
            continue
        if "CLAVE" in cell_a.upper():
            continue
        if "SUB-TOTAL" in cell_a.upper() or "SUB-TOTAL" in cell_f.upper():
            continue

        # Leer valores según tipo
        clave_insumo = cell_a
        descripcion  = cell_b
        unidad       = str(row[2] or "").strip() or "UND"
        cant_insumo  = 0.0
        v_unit       = 0.0

        if tipo_actual == "RENDIMIENTO":
            cant_d  = safe_float(row[3])
            rend_e  = safe_float(row[4]) or 1.0
            cant_insumo = cant_d * rend_e
            v_unit  = safe_float(row[5])
        elif tipo_actual == "PORCENTAJE":
            cant_insumo = safe_float(row[4])
            v_unit      = safe_float(row[5])
            unidad      = str(row[2] or "").strip() or "GL"
        else:  # ESTANDAR
            cant_insumo = safe_float(row[3])
            v_unit      = safe_float(row[5])

        if v_unit <= 0:
            continue

        clave = descripcion.upper().strip()
        valor_total = cant_insumo * v_unit

        # Acumular base (solo no-PORCENTAJE)
        if tipo_actual != "PORCENTAJE":
            base_item += valor_total * cant_ppto

        # Acumular en diccionarios globales
        if clave in d_insumos:
            d_cantidades[clave] += cant_insumo * cant_ppto
            d_valores[clave]    += valor_total * cant_ppto
            if d_cantidades[clave] > 0:
                d_vunit[clave] = d_valores[clave] / d_cantidades[clave]
            if not d_claves[clave] and clave_insumo:
                d_claves[clave] = clave_insumo
        else:
            d_insumos[clave]    = descripcion
            d_unidades[clave]   = unidad
            d_cantidades[clave] = cant_insumo * cant_ppto
            d_valores[clave]    = valor_total * cant_ppto
            d_categorias[clave] = categoria_actual
            d_vunit[clave]      = v_unit
            d_claves[clave]     = clave_insumo

    return base_item


# ─────────────────────────────────────────────
# 5. ESCRIBIR HOJA EXPLOSION INSUMOS
# ─────────────────────────────────────────────

def escribir_explosion(wb, cats, diccionarios, base_total):
    sheet_name = "EXPLOSION INSUMOS"

    # Eliminar si existe
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]

    ws = wb.create_sheet(sheet_name)

    d_insumos    = diccionarios["insumos"]
    d_unidades   = diccionarios["unidades"]
    d_cantidades = diccionarios["cantidades"]
    d_valores    = diccionarios["valores"]
    d_categorias = diccionarios["categorias"]
    d_vunit      = diccionarios["vunit"]
    d_claves     = diccionarios["claves"]

    # Título
    ws.merge_cells("A1:G1")
    ws["A1"] = "EXPLOSIÓN DE INSUMOS - CONSOLIDADO"
    ws["A1"].font      = Font(bold=True, size=14, name="Arial")
    ws["A1"].alignment = Alignment(horizontal="center")

    # Cabeceras
    headers = ["CATEGORÍA", "CLAVE", "DESCRIPCIÓN DEL INSUMO",
               "UNIDAD", "CANTIDAD TOTAL", "COSTO", "VALOR TOTAL"]
    for col_idx, h in enumerate(headers, start=1):
        cell = ws.cell(row=3, column=col_idx, value=h)
        apply_header_style(cell)
        cell.border = thin_border()

    # Anchos de columna
    col_widths = [18, 12, 55, 10, 18, 18, 18]
    for i, w in enumerate(col_widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    fila        = 4
    total_gral  = 0.0

    for cat in cats:
        cat_nombre = cat["categoria"]
        cat_hex    = cat["hex"]
        cat_tipo   = cat["tipo"]
        cat_label  = f"{cat['num']} {cat_nombre}"

        # Insumos de esta categoría
        insumos_cat = [(k, v) for k, v in d_insumos.items()
                       if d_categorias.get(k) == cat_nombre]
        if not insumos_cat:
            continue

        # Fila de categoría
        ws.merge_cells(start_row=fila, start_column=1, end_row=fila, end_column=7)
        cell = ws.cell(row=fila, column=1, value=cat_label)
        cell.font = Font(bold=True, name="Arial")
        cell.fill = PatternFill("solid", fgColor=cat_hex)
        fila += 1

        subtotal = 0.0
        for clave, descripcion in insumos_cat:
            cant_total  = d_cantidades[clave]
            valor_total = d_valores[clave]

            if cat_tipo == "PORCENTAJE":
                v_unit_mostrar = base_total
                if base_total > 0:
                    cant_total = valor_total / base_total
                else:
                    cant_total = 0.0
            else:
                if cant_total > 0:
                    v_unit_mostrar = valor_total / cant_total
                else:
                    v_unit_mostrar = d_vunit[clave]

            row_data = [
                cat_nombre,
                d_claves[clave],
                descripcion,
                d_unidades[clave],
                cant_total,
                v_unit_mostrar,
                valor_total,
            ]
            for col_idx, val in enumerate(row_data, start=1):
                cell = ws.cell(row=fila, column=col_idx, value=val)
                cell.border  = thin_border()
                cell.font    = Font(name="Arial")

            # Formatos numéricos
            if cat_tipo == "PORCENTAJE":
                pct_fmt(ws.cell(row=fila, column=5))
            else:
                qty_fmt(ws.cell(row=fila, column=5))
            money_fmt(ws.cell(row=fila, column=6))
            money_fmt(ws.cell(row=fila, column=7))

            subtotal += valor_total
            fila += 1

        # Fila subtotal de categoría
        sub_cell_label = ws.cell(row=fila, column=6, value=f"SUBTOTAL {cat_nombre}:")
        sub_cell_val   = ws.cell(row=fila, column=7, value=subtotal)
        sub_cell_label.font = Font(bold=True, name="Arial")
        sub_cell_val.font   = Font(bold=True, name="Arial")
        money_fmt(sub_cell_val)
        for col_idx in range(1, 8):
            c = ws.cell(row=fila, column=col_idx)
            c.fill   = PatternFill("solid", fgColor=cat_hex)
            c.border = thin_border()

        total_gral += subtotal
        fila += 2   # línea en blanco entre categorías

    # TOTAL GENERAL
    ws.cell(row=fila, column=6, value="TOTAL GENERAL:").font = Font(bold=True, color="FFFFFF", name="Arial")
    total_cell = ws.cell(row=fila, column=7, value=total_gral)
    total_cell.font = Font(bold=True, color="FFFFFF", name="Arial")
    money_fmt(total_cell)
    for col_idx in range(1, 8):
        c = ws.cell(row=fila, column=col_idx)
        c.fill   = PatternFill("solid", fgColor="00467F")
        c.border = thin_border()

    print(f"  → TOTAL GENERAL: $ {total_gral:,.0f}")


# ─────────────────────────────────────────────
# 6. MAIN
# ─────────────────────────────────────────────

def generar_explosion(ruta_archivo):
    print(f"Cargando: {ruta_archivo}")
    wb = load_workbook(ruta_archivo, data_only=True)

    # Hoja AUX
    ws_aux = wb["AUX"] if "AUX" in wb.sheetnames else (wb["Aux"] if "Aux" in wb.sheetnames else None)
    if ws_aux is None:
        raise ValueError("No se encontró la hoja AUX.")
    cats = leer_categorias(ws_aux)
    print(f"  Categorías leídas: {len(cats)}")

    # Hoja PPTO
    ws_ppto = wb["PPTO"] if "PPTO" in wb.sheetnames else None
    if ws_ppto is None:
        raise ValueError("No se encontró la hoja PPTO.")
    cant_ppto_map = leer_cant_ppto(ws_ppto)
    print(f"  Ítems en PPTO: {list(cant_ppto_map.keys())}")

    # Diccionarios globales
    diccionarios = {
        "insumos":    {},
        "unidades":   {},
        "cantidades": defaultdict(float),
        "valores":    defaultdict(float),
        "categorias": {},
        "vunit":      {},
        "claves":     defaultdict(str),
    }

    # Recorrer hojas "item N"
    item_pattern = re.compile(r"^item\s+(\d+)$", re.IGNORECASE)
    base_total = 0.0

    for sheet_name in wb.sheetnames:
        m = item_pattern.match(sheet_name.strip())
        if not m:
            continue
        item_num  = int(m.group(1))
        cant_ppto = cant_ppto_map.get(item_num, 0.0)
        ws_item   = wb[sheet_name]
        base_item = procesar_item(ws_item, item_num, cant_ppto, cats, diccionarios)
        base_total += base_item
        print(f"  Ítem {item_num}: cant={cant_ppto}, base=${base_item:,.0f}")

    print(f"  Base total consolidada: $ {base_total:,.0f}")

    # Escribir hoja resultado
    escribir_explosion(wb, cats, diccionarios, base_total)

    # Guardar
    import os; nombre = os.path.splitext(os.path.basename(ruta_archivo))[0]; ruta_salida = f"/mnt/user-data/outputs/{nombre}_explosion.xlsx"
    wb.save(ruta_salida)
    print(f"\n✓ Archivo guardado: {ruta_salida}")
    return ruta_salida


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Uso: python explosion_insumos.py <ruta_archivo.xlsm>")
        sys.exit(1)
    generar_explosion(sys.argv[1])


# ─────────────────────────────────────────────
# FUNCIÓN PARA STREAMLIT: devuelve datos + bytes
# ─────────────────────────────────────────────

import io

def procesar_archivo(file_bytes: bytes):
    """
    Recibe los bytes del .xlsm/.xlsx subido.
    Devuelve (tabla_dict, total_general, excel_bytes, cats).
    tabla_dict = {categoria: [{"clave", "descripcion", "unidad", "cantidad", "costo", "valor_total"}, ...]}
    """
    wb = load_workbook(io.BytesIO(file_bytes), data_only=True)

    ws_aux = wb["AUX"] if "AUX" in wb.sheetnames else (wb["Aux"] if "Aux" in wb.sheetnames else None)
    if ws_aux is None:
        raise ValueError("No se encontró la hoja AUX.")
    cats = leer_categorias(ws_aux)

    ws_ppto = wb["PPTO"] if "PPTO" in wb.sheetnames else None
    if ws_ppto is None:
        raise ValueError("No se encontró la hoja PPTO.")
    cant_ppto_map = leer_cant_ppto(ws_ppto)

    diccionarios = {
        "insumos":    {},
        "unidades":   {},
        "cantidades": defaultdict(float),
        "valores":    defaultdict(float),
        "categorias": {},
        "vunit":      {},
        "claves":     defaultdict(str),
    }

    item_pattern = re.compile(r"^item\s+(\d+)$", re.IGNORECASE)
    base_total = 0.0

    items_procesados = []
    for sheet_name in wb.sheetnames:
        m = item_pattern.match(sheet_name.strip())
        if not m:
            continue
        item_num  = int(m.group(1))
        cant_ppto = cant_ppto_map.get(item_num, 0.0)
        ws_item   = wb[sheet_name]
        base_item = procesar_item(ws_item, item_num, cant_ppto, cats, diccionarios)
        base_total += base_item
        items_procesados.append({"item": item_num, "cantidad": cant_ppto, "base": base_item})

    # Construir tabla para UI
    d_insumos    = diccionarios["insumos"]
    d_unidades   = diccionarios["unidades"]
    d_cantidades = diccionarios["cantidades"]
    d_valores    = diccionarios["valores"]
    d_categorias = diccionarios["categorias"]
    d_vunit      = diccionarios["vunit"]
    d_claves     = diccionarios["claves"]

    tabla = {}
    subtotales = {}
    total_general = 0.0

    for cat in cats:
        cat_nombre = cat["categoria"]
        cat_tipo   = cat["tipo"]
        insumos_cat = [(k, v) for k, v in d_insumos.items()
                       if d_categorias.get(k) == cat_nombre]
        if not insumos_cat:
            continue

        filas = []
        subtotal = 0.0
        for clave, descripcion in insumos_cat:
            cant_total  = d_cantidades[clave]
            valor_total = d_valores[clave]

            if cat_tipo == "PORCENTAJE":
                v_unit_mostrar = base_total
                cant_total = valor_total / base_total if base_total > 0 else 0.0
            else:
                v_unit_mostrar = valor_total / cant_total if cant_total > 0 else d_vunit[clave]

            filas.append({
                "clave":       d_claves[clave],
                "descripcion": descripcion,
                "unidad":      d_unidades[clave],
                "cantidad":    cant_total,
                "costo":       v_unit_mostrar,
                "valor_total": valor_total,
                "es_pct":      cat_tipo == "PORCENTAJE",
            })
            subtotal += valor_total

        tabla[cat_nombre]      = filas
        subtotales[cat_nombre] = subtotal
        total_general         += subtotal

    # Generar Excel en memoria
    escribir_explosion(wb, cats, diccionarios, base_total)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    excel_bytes = buf.read()

    return tabla, subtotales, total_general, excel_bytes, items_procesados, cats
