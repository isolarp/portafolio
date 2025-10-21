#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Script lineal que abre SE16N vía SAP GUI Scripting, ejecuta consulta de una tabla y vuelca resultados a un pandas.DataFrame.
# Requisitos: Windows + SAP GUI con Scripting habilitado, Python, pywin32, pandas.
# No incluye credenciales: usa la sesión SAP GUI abierta.
import time
import sys
import re
try:
    from win32com.client import GetObject
except Exception as e:
    print("Necesitas pywin32. Instálalo con: pip install pywin32", file=sys.stderr)
    raise
try:
    import pandas as pd
except Exception as e:
    print("Necesitas pandas y openpyxl. Instálalos con: pip install pandas openpyxl", file=sys.stderr)
    raise

TABLE_NAME = "T001"
CONNECTION_INDEX = 0
SESSION_INDEX = 0
# WAIT_AFTER_EXEC: segundos a esperar después de ejecutar (ajusta si tu sistema es lento)
WAIT_AFTER_EXEC = 1.0

print("Conectando a SAP GUI...", file=sys.stderr)
try:
    sapgui = GetObject("SAPGUI")
    application = sapgui.GetScriptingEngine()
    connection = application.Children(CONNECTION_INDEX)
    session = connection.Children(SESSION_INDEX)
except Exception as e:
    print(f"Error conectando a SAP GUI/Scripting engine: {e}", file=sys.stderr)
    raise

try:
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nSE16N"
    session.findById("wnd[0]/tbar[0]/btn[0]").press()
    time.sleep(0.3)
    session.findById("wnd[0]/usr/ctxtGD-TAB").text = TABLE_NAME
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    time.sleep(WAIT_AFTER_EXEC)
except Exception as e:
    print(f"Error navegando a SE16N o ejecutando: {e}", file=sys.stderr)
    raise

candidates = [
    "wnd[0]/usr/cntlRESULT_LIST/shellcont/shell",
    "wnd[0]/usr/cntlGRID1/shellcont/shell",
    "wnd[0]/usr/cntlGRID/shellcont/shell",
    "wnd[0]/usr/cntlCONTAINER/shellcont/shell",
]
grid = None
grid_id = None
for cid in candidates:
    try:
        grid = session.findById(cid)
        grid_id = cid
        break
    except Exception:
        continue
if grid is None:
    try:
        wnd = session.findById("wnd[0]")
        for child in wnd.Children:
            try:
                if hasattr(child, "Id") and "shell" in str(child.Id).lower():
                    grid = child
                    grid_id = getattr(child, "Id", "<dynamic>")
                    break
            except Exception:
                continue
    except Exception:
        pass
if grid is None:
    raise RuntimeError("No se encontró control grid en la pantalla. Revisa la ventana SE16N y el ID del control.")

print(f"Grid encontrado en: {grid_id}", file=sys.stderr)

df = None
try:
    grid.selectAll()
    grid.copy()
    df = pd.read_clipboard(sep='\t')
except Exception:
    row_count = None
    col_count = None
    try:
        row_count = int(getattr(grid, "RowCount"))
    except Exception:
        try:
            row_count = int(grid.RowCount)
        except Exception:
            row_count = None
    try:
        col_count = int(getattr(grid, "ColumnCount"))
    except Exception:
        try:
            col_count = int(grid.ColumnCount)
        except Exception:
            try:
                col_count = int(grid.Columns.Count)
            except Exception:
                col_count = None
    if row_count is None or col_count is None:
        raise RuntimeError("No se pudo determinar RowCount/ColumnCount del grid y la copia al portapapeles falló.")
    columns = []
    for ci in range(col_count):
        col_name = f"col_{ci}"
        try:
            col_name = grid.GetColumnTitle(ci)
        except Exception:
            try:
                col_name = grid.Columns.Item(ci).Title
            except Exception:
                try:
                    col_name = grid.Columns.Item(ci).Name
                except Exception:
                    try:
                        col_name = grid.Columns.Item(ci).Text
                    except Exception:
                        col_name = f"col_{ci}"
        columns.append(str(col_name))
    data = []
    for ri in range(row_count):
        row_vals = []
        for ci in range(col_count):
            val = ""
            tried = False
            for method in ("GetCellValue", "getCellValue", "GetCellText", "getCellText", "GetValue", "getValue"):
                try:
                    func = getattr(grid, method)
                    try:
                        val = func(ri, ci)
                    except Exception:
                        try:
                            val = func(ri, columns[ci])
                        except Exception:
                            val = func(ri)
                    tried = True
                    break
                except Exception:
                    continue
            if not tried:
                try:
                    cell = grid.Cells(ri, ci)
                    val = str(cell.Text)
                except Exception:
                    val = ""
            row_vals.append(val)
        data.append(row_vals)
    df = pd.DataFrame(data, columns=columns)

print(f"Extrajimos {len(df)} filas y {len(df.columns)} columnas.", file=sys.stderr)
out = f"se16n_{TABLE_NAME}.xlsx"
try:
    df.to_excel(out, index=False, engine='openpyxl')
    print(f"Guardado a: {out}", file=sys.stderr)
except Exception as e:
    csv_fallback = f"se16n_{TABLE_NAME}.csv"
    df.to_csv(csv_fallback, index=False, encoding='utf-8-sig')
    print(f"No se pudo guardar .xlsx ({e}). Se guardó CSV en {csv_fallback}", file=sys.stderr)