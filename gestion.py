"""
Lógica de gestión (sin Streamlit): el Excel es la base de datos.

Hojas:
  Instrucciones · Resumen · Config · Productos · Clientes · Pedidos · Líneas · Movimientos · Facturas
"""
from __future__ import annotations

import datetime as dt
import io
from pathlib import Path

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

# --------------------------------------------------------------------------- #
# Estructura del Excel
# --------------------------------------------------------------------------- #
COLUMNAS = {
    "Config": ["Campo", "Valor"],
    "Productos": ["SKU", "Nombre", "Precio", "IVA %", "Coste unitario", "Stock mínimo", "Stock actual"],
    "Clientes": ["ID", "Nombre", "NIF", "Dirección", "CP", "Ciudad", "Email", "Teléfono"],
    "Pedidos": ["Nº pedido", "Fecha", "ID cliente", "Cliente", "Estado", "Nº factura", "Notas"],
    "Líneas": ["Nº pedido", "SKU", "Descripción", "Cantidad", "Precio", "Descuento %", "IVA %",
               "Base", "Cuota IVA", "Total"],
    "Movimientos": ["Fecha", "SKU", "Tipo", "Cantidad", "Coste unitario", "Referencia", "Notas"],
    "Facturas": ["Nº factura", "Fecha", "Nº pedido", "ID cliente", "Cliente", "NIF",
                 "Base", "IVA", "Total", "Cobro", "Fecha cobro"],
    "Gastos": ["Fecha", "Concepto", "Categoría", "Tipo", "Proveedor", "Base", "IVA %", "IVA", "Total", "Notas"],
    "Proveedores": ["ID", "Nombre", "Qué suministra", "Contacto", "Email", "Teléfono", "Web", "País", "Notas"],
    "Tareas": ["ID", "Tarea", "Estado", "Prioridad", "Fecha límite", "Notas"],
}

CATEGORIAS_GASTO = ["Impresión", "Envío", "Aduanas", "Impuestos", "Embalaje", "Envíos a clientes", "Web y dominio",
                    "Marca y registros", "Publicidad", "Material", "Gestoría", "Otros"]
TIPOS_GASTO = ["Gasto general", "Compra de stock"]
ESTADOS_TAREA = ["Pendiente", "En curso", "Hecha"]
PRIORIDADES = ["Alta", "Media", "Baja"]

EUR = '#,##0.00 "€"'
PCT = "0%"
FECHA = "dd/mm/yyyy"
FORMATOS = {
    "Productos": {"Precio": EUR, "IVA %": PCT, "Coste unitario": EUR},
    "Pedidos": {"Fecha": FECHA},
    "Líneas": {"Precio": EUR, "Descuento %": PCT, "IVA %": PCT, "Base": EUR, "Cuota IVA": EUR, "Total": EUR},
    "Movimientos": {"Fecha": FECHA, "Coste unitario": EUR},
    "Facturas": {"Fecha": FECHA, "Base": EUR, "IVA": EUR, "Total": EUR, "Fecha cobro": FECHA},
    "Gastos": {"Fecha": FECHA, "Base": EUR, "IVA %": PCT, "IVA": EUR, "Total": EUR},
    "Tareas": {"Fecha límite": FECHA},
}
ANCHOS = {"Nombre": 30, "Descripción": 32, "Dirección": 30, "Email": 26, "Cliente": 26,
          "Notas": 30, "Referencia": 18, "Campo": 30, "Valor": 45, "Estado": 20, "Tarea": 45,
          "Concepto": 32, "Qué suministra": 26, "Web": 28, "Categoría": 18, "Tipo": 16, "Proveedor": 24}

CONFIG_POR_DEFECTO = [
    ("Nombre o razón social", "Tu nombre y apellidos"),
    ("NIF", "00000000X"),
    ("Dirección", "Calle Ejemplo 1"),
    ("CP y ciudad", "08000 Barcelona"),
    ("Email", "xingonplay@gmail.com"),
    ("Teléfono", ""),
    ("IBAN", ""),
    ("Serie factura", "X"),
    ("Precios con IVA incluido", "Sí"),
    ("IVA por defecto", 0.21),
    ("Texto al pie", "Gracias por jugar a XINGON!"),
]

TIPOS_MOV = ["Entrada", "Venta", "Ajuste"]
ESTADO_PENDIENTE = "Pendiente de facturar"
ESTADO_FACTURADO = "Facturado"
ESTADO_ANULADO = "Anulado"

AZUL = Font(name="Arial", color="0000FF")
NEGRO = Font(name="Arial")
NEGRITA = Font(name="Arial", bold=True, color="FFFFFF")
CABECERA = PatternFill("solid", start_color="1F3A5F")
AMARILLO = PatternFill("solid", start_color="FFF2CC")


# --------------------------------------------------------------------------- #
# Utilidades de formato
# --------------------------------------------------------------------------- #
def eur(x: float) -> str:
    """12345.6 -> '12.345,60 €'"""
    s = f"{float(x):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"{s} €"


def pct(x: float) -> str:
    return f"{float(x) * 100:.0f} %".replace(".", ",")


def _a_fecha(v) -> dt.date | None:
    if v is None or (isinstance(v, float) and pd.isna(v)) or v is pd.NaT:
        return None
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    try:
        return pd.to_datetime(v, dayfirst=True).date()
    except Exception:
        return None


def _vacio(v) -> bool:
    if v is None:
        return True
    try:
        if pd.isna(v):
            return True
    except (TypeError, ValueError):
        pass
    return str(v).strip() == ""


# --------------------------------------------------------------------------- #
# Crear la plantilla
# --------------------------------------------------------------------------- #
def _cabecera(ws, cols):
    ws.append(cols)
    for i, c in enumerate(cols, start=1):
        cell = ws.cell(row=1, column=i)
        cell.font = NEGRITA
        cell.fill = CABECERA
        cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.column_dimensions[cell.column_letter].width = ANCHOS.get(c, 15)
    ws.freeze_panes = "A2"


def _formula_stock(fila: int) -> str:
    return f"=SUMIFS(Movimientos!$D:$D,Movimientos!$B:$B,A{fila})"


def crear_plantilla(ruta: str | Path) -> Path:
    ruta = Path(ruta)
    wb = Workbook()

    # Instrucciones
    ws = wb.active
    ws.title = "Instrucciones"
    ws.column_dimensions["A"].width = 110
    lineas = [
        ("GESTIÓN DE PEDIDOS, STOCK Y FACTURAS", True),
        ("", False),
        ("Cómo funciona", True),
        ("• Este Excel es la base de datos de la app de Streamlit. Puedes usarlo solo, pero lo más cómodo es "
         "registrar pedidos y facturas desde la app.", False),
        ("• Celdas en AZUL = datos que escribes tú. Celdas en NEGRO = fórmulas (no las toques).", False),
        ("• Config: rellena tus datos fiscales (salen en la factura). Celdas amarillas = obligatorias.", False),
        ("• Productos: catálogo. 'Precio' es el PVP con IVA si en Config pones 'Precios con IVA incluido = Sí'. "
         "'Stock actual' se calcula solo sumando la hoja Movimientos.", False),
        ("• Movimientos: cada entrada (llegada de un pedido a la imprenta) suma; cada venta resta (cantidad "
         "negativa). Para corregir un recuento usa Tipo = Ajuste con cantidad + o −.", False),
        ("• Pedidos + Líneas: un pedido tiene una fila en Pedidos y una fila por producto en Líneas.", False),
        ("• Facturas: numeración correlativa por serie y año (p. ej. X2026-0001). No borres facturas: si hay "
         "un error se hace una factura rectificativa.", False),
        ("• Gastos: gastos de la actividad. Tipo 'Gasto general' (web, marca, publicidad…) o 'Compra de stock' "
         "(impresión, envío y aduanas de un lote). El margen usa los gastos generales.", False),
        ("• Proveedores y Tareas: tu agenda de imprentas y transportistas, y tu lista de cosas por hacer.", False),
        ("• Resumen: totales facturados, IVA repercutido, pendiente de cobro, gastos y unidades.", False),
        ("", False),
        ("Flujo típico", True),
        ("1. Config → tus datos.   2. Productos → precio y coste.   3. Movimientos → Entrada con las unidades "
         "que te llegan.   4. App → Nuevo pedido.   5. App → Facturar → descargar PDF.   6. Marcar como cobrada.",
         False),
        ("", False),
        ("Las filas de ejemplo (cliente C-0001) puedes borrarlas.", False),
    ]
    for texto, negrita in lineas:
        ws.append([texto])
        ws.cell(row=ws.max_row, column=1).font = Font(name="Arial", bold=negrita, size=13 if negrita else 11)
        ws.cell(row=ws.max_row, column=1).alignment = Alignment(wrap_text=True, vertical="top")

    # Resumen (fórmulas)
    ws = wb.create_sheet("Resumen")
    ws.column_dimensions["A"].width = 38
    ws.column_dimensions["B"].width = 18
    ws.append(["Indicador", "Valor"])
    for c in ws[1]:
        c.font, c.fill = NEGRITA, CABECERA
    resumen = [
        ("Total facturado (IVA incl.)", "=SUM(Facturas!I:I)", EUR),
        ("Base imponible facturada", "=SUM(Facturas!G:G)", EUR),
        ("IVA repercutido", "=SUM(Facturas!H:H)", EUR),
        ("Pendiente de cobro", '=SUMIFS(Facturas!I:I,Facturas!J:J,"Pendiente")', EUR),
        ("Nº facturas emitidas", "=COUNTA(Facturas!A:A)-1", "0"),
        ("Pedidos pendientes de facturar", f'=COUNTIFS(Pedidos!E:E,"{ESTADO_PENDIENTE}")', "0"),
        ("Unidades compradas", '=SUMIFS(Movimientos!D:D,Movimientos!C:C,"Entrada")', "0"),
        ("Unidades vendidas", '=-SUMIFS(Movimientos!D:D,Movimientos!C:C,"Venta")', "0"),
        ("Unidades en stock (todas)", "=SUM(Movimientos!D:D)", "0"),
        ("Gastos generales (IVA incl.)", '=SUMIFS(Gastos!I:I,Gastos!D:D,"Gasto general")', EUR),
        ("Gastos de compra de stock (IVA incl.)", '=SUMIFS(Gastos!I:I,Gastos!D:D,"Compra de stock")', EUR),
        ("Tareas pendientes", '=COUNTIFS(Tareas!C:C,"<>Hecha",Tareas!B:B,"<>")-1', "0"),
        ("Inversión en stock comprado", "=SUMPRODUCT((Movimientos!C2:C5000=\"Entrada\")*"
                                        "Movimientos!D2:D5000*Movimientos!E2:E5000)", EUR),
    ]
    for etiqueta, formula, fmt in resumen:
        ws.append([etiqueta, formula])
        ws.cell(row=ws.max_row, column=1).font = NEGRO
        c = ws.cell(row=ws.max_row, column=2)
        c.font, c.number_format = NEGRO, fmt

    # Config
    ws = wb.create_sheet("Config")
    _cabecera(ws, COLUMNAS["Config"])
    for campo, valor in CONFIG_POR_DEFECTO:
        ws.append([campo, valor])
        ws.cell(row=ws.max_row, column=1).font = NEGRO
        v = ws.cell(row=ws.max_row, column=2)
        v.font = AZUL
        if campo in ("Nombre o razón social", "NIF", "Dirección", "CP y ciudad"):
            v.fill = AMARILLO
        if campo == "IVA por defecto":
            v.number_format = PCT
    ws["B10"].comment = Comment("21 % es el tipo general en España (juegos de mesa/cartas). "
                                "Pon 0 si aún no repercutes IVA.", "App")
    ws["B9"].comment = Comment("Sí = el Precio de Productos es el PVP final (p. ej. 14,99 €) y la app "
                               "calcula la base. No = el Precio es sin IVA.", "App")

    # Tablas de datos
    for hoja in ["Productos", "Clientes", "Pedidos", "Líneas", "Movimientos", "Facturas",
                 "Gastos", "Proveedores", "Tareas"]:
        ws = wb.create_sheet(hoja)
        _cabecera(ws, COLUMNAS[hoja])

    # Fila de ejemplo de producto (PVP dado por la usuaria) y cliente de ejemplo
    ws = wb["Productos"]
    ws.append(["XINGON-001", "XINGON! – juego de cartas", 14.99, 0.21, None, 10, _formula_stock(2)])
    _formatear_fila(ws, "Productos", 2)
    ws["E2"].comment = Comment("Rellena lo que te cuesta cada unidad (impresión + envío + aduanas). "
                               "Sirve para calcular el margen.", "App")
    ws["C2"].comment = Comment("PVP con IVA incluido indicado por ti: 14,99 €.", "App")

    ws = wb["Clientes"]
    ws.append(["C-0001", "Cliente de ejemplo", "", "Calle Falsa 123", "08830", "Sant Boi de Llobregat",
               "", ""])
    _formatear_fila(ws, "Clientes", 2)

    ws = wb["Proveedores"]
    ws.append(["PR-0001", "MakePlayingCards (MPC)", "Impresión de barajas (papel M31)", "", "", "",
               "makeplayingcards.com", "China", "Envío con aduanas a pagar"])
    _formatear_fila(ws, "Proveedores", 2)
    ws = wb["Tareas"]
    ws.append(["T-0001", "Seguir el registro de la marca XINGON!", "En curso", "Media", None, ""])
    _formatear_fila(ws, "Tareas", 2)

    for w in wb.worksheets:
        if w.title in COLUMNAS:
            w.auto_filter.ref = w.dimensions
    ruta.parent.mkdir(parents=True, exist_ok=True)
    wb.save(ruta)
    return ruta


def _formatear_fila(ws, hoja: str, fila: int):
    cols = COLUMNAS[hoja]
    fmts = FORMATOS.get(hoja, {})
    for i, nombre in enumerate(cols, start=1):
        c = ws.cell(row=fila, column=i)
        es_formula = isinstance(c.value, str) and c.value.startswith("=")
        c.font = NEGRO if es_formula else AZUL
        if nombre in fmts:
            c.number_format = fmts[nombre]


# --------------------------------------------------------------------------- #
# Lectura
# --------------------------------------------------------------------------- #
def cargar(ruta: str | Path) -> dict[str, pd.DataFrame]:
    todas = pd.read_excel(ruta, sheet_name=None, engine="openpyxl")
    dfs = {h: todas.get(h, pd.DataFrame(columns=COLUMNAS[h])) for h in COLUMNAS}
    # Clientes y Proveedores como texto (para no perder el 0 inicial de CP o teléfono)
    for h in ("Clientes", "Proveedores"):
        if h in todas:
            dfs[h] = pd.read_excel(ruta, sheet_name=h, engine="openpyxl", dtype=str)
    for hoja, cols in COLUMNAS.items():
        df = dfs[hoja]
        for c in cols:
            if c not in df.columns:
                df[c] = None
        df = df[cols].dropna(how="all")
        dfs[hoja] = df.reset_index(drop=True)
    # tipos
    for c in ["SKU", "Nombre"]:
        dfs["Productos"][c] = dfs["Productos"][c].astype(str)
    for h in ("Clientes", "Proveedores"):
        dfs[h] = dfs[h].astype(object).where(dfs[h].notna(), "").astype(str)
    return dfs


def config(dfs) -> dict:
    cfg = {campo: valor for campo, valor in CONFIG_POR_DEFECTO}
    for _, r in dfs["Config"].iterrows():
        if not _vacio(r["Campo"]):
            cfg[str(r["Campo"]).strip()] = "" if _vacio(r["Valor"]) else r["Valor"]
    v = cfg.get("IVA por defecto", 0.21)
    try:
        v = float(str(v).replace("%", "").replace(",", "."))
        cfg["IVA por defecto"] = v / 100 if v > 1 else v
    except ValueError:
        cfg["IVA por defecto"] = 0.21
    cfg["iva_incluido"] = str(cfg.get("Precios con IVA incluido", "Sí")).strip().lower() in ("sí", "si", "s", "yes", "1", "true")
    return cfg


def stock_actual(dfs) -> pd.DataFrame:
    prod = dfs["Productos"].copy()
    mov = dfs["Movimientos"]
    suma = pd.to_numeric(mov["Cantidad"], errors="coerce").fillna(0).groupby(mov["SKU"].astype(str)).sum()
    prod["Stock actual"] = prod["SKU"].map(suma).fillna(0).astype(int)
    prod["Stock mínimo"] = pd.to_numeric(prod["Stock mínimo"], errors="coerce").fillna(0).astype(int)
    prod["Alerta"] = prod["Stock actual"] <= prod["Stock mínimo"]
    return prod


def _siguiente(prefijo: str, existentes, ancho: int = 4) -> str:
    n = 0
    for e in existentes:
        e = str(e)
        if e.startswith(prefijo):
            try:
                n = max(n, int(e[len(prefijo):]))
            except ValueError:
                pass
    return f"{prefijo}{n + 1:0{ancho}d}"


def calcular_linea(cantidad, precio, descuento, iva, iva_incluido: bool):
    """Devuelve (base, cuota, total) redondeados a céntimos."""
    importe = round(float(cantidad) * float(precio) * (1 - float(descuento or 0)), 2)
    iva = float(iva or 0)
    if iva_incluido:
        total = importe
        base = round(total / (1 + iva), 2)
        cuota = round(total - base, 2)
    else:
        base = importe
        cuota = round(base * iva, 2)
        total = round(base + cuota, 2)
    return base, cuota, total


# --------------------------------------------------------------------------- #
# Escritura
# --------------------------------------------------------------------------- #
def _abrir(ruta):
    wb = load_workbook(ruta)
    for hoja, cols in COLUMNAS.items():   # Excel creado con una versión anterior
        if hoja not in wb.sheetnames:
            _cabecera(wb.create_sheet(hoja), cols)
    return wb


def _guardar(wb, ruta):
    for hoja in COLUMNAS:
        ws = wb[hoja]
        ws.auto_filter.ref = ws.dimensions
    wb.save(ruta)


def _append(wb, hoja: str, fila: list):
    ws = wb[hoja]
    # primera fila libre (ignorando filas vacías al final)
    r = ws.max_row
    while r > 1 and all(_vacio(c.value) for c in ws[r]):
        r -= 1
    r += 1
    for i, v in enumerate(fila, start=1):
        ws.cell(row=r, column=i, value=v)
    if hoja == "Productos":
        ws.cell(row=r, column=7, value=_formula_stock(r))
    _formatear_fila(ws, hoja, r)
    return r


def _buscar_fila(ws, clave, col: int = 1) -> int | None:
    for r in range(2, ws.max_row + 1):
        if str(ws.cell(row=r, column=col).value) == str(clave):
            return r
    return None


def reemplazar_hoja(ruta, hoja: str, df: pd.DataFrame):
    """Sobrescribe Productos, Clientes o Config con lo editado en la app."""
    wb = _abrir(ruta)
    ws = wb[hoja]
    if ws.max_row > 1:
        ws.delete_rows(2, ws.max_row - 1)
    cols = COLUMNAS[hoja]
    for _, r in df.iterrows():
        if all(_vacio(r.get(c)) for c in cols[:2]):
            continue
        fila = [None if _vacio(r.get(c)) else r.get(c) for c in cols]
        fila = [v.date() if isinstance(v, (pd.Timestamp, dt.datetime)) else v for v in fila]
        if hoja == "Gastos":
            total = float(r.get("Total") or 0) if not _vacio(r.get("Total")) else 0.0
            iva = float(r.get("IVA %")) if not _vacio(r.get("IVA %")) else 0.0
            base = round(total / (1 + iva), 2)
            fila[5], fila[6], fila[7], fila[8] = base, iva, round(total - base, 2), round(total, 2)
        if hoja == "Productos":
            fila = fila[:6]
        _append(wb, hoja, fila)
    _guardar(wb, ruta)


def nuevo_cliente(ruta, nombre, nif="", direccion="", cp="", ciudad="", email="", telefono="") -> str:
    dfs = cargar(ruta)
    cid = _siguiente("C-", dfs["Clientes"]["ID"])
    wb = _abrir(ruta)
    _append(wb, "Clientes", [cid, nombre, nif, direccion, cp, ciudad, email, telefono])
    _guardar(wb, ruta)
    return cid


def registrar_movimiento(ruta, sku, cantidad: int, tipo="Entrada", coste=None, referencia="", notas="",
                         fecha: dt.date | None = None):
    wb = _abrir(ruta)
    _append(wb, "Movimientos", [fecha or dt.date.today(), sku, tipo, int(cantidad), coste, referencia, notas])
    _guardar(wb, ruta)


def crear_pedido(ruta, cliente_id: str, lineas: list[dict], fecha: dt.date | None = None, notas="") -> str:
    """
    lineas: [{"SKU":..., "Cantidad":..., "Descuento %": 0.1}, ...]
    Descuenta stock (movimiento 'Venta') al crear el pedido.
    """
    dfs = cargar(ruta)
    cfg = config(dfs)
    fecha = fecha or dt.date.today()
    prod = dfs["Productos"].set_index("SKU")
    cli = dfs["Clientes"].set_index("ID")
    if cliente_id not in cli.index:
        raise ValueError(f"Cliente {cliente_id} no existe")
    lineas = [l for l in lineas if not _vacio(l.get("SKU")) and float(l.get("Cantidad") or 0) > 0]
    if not lineas:
        raise ValueError("El pedido no tiene líneas")

    pid = _siguiente(f"P{fecha.year}-", dfs["Pedidos"]["Nº pedido"])
    wb = _abrir(ruta)
    _append(wb, "Pedidos", [pid, fecha, cliente_id, cli.loc[cliente_id, "Nombre"], ESTADO_PENDIENTE, None, notas])
    for l in lineas:
        sku = str(l["SKU"])
        if sku not in prod.index:
            raise ValueError(f"Producto {sku} no existe")
        p = prod.loc[sku]
        precio = float(l.get("Precio") if not _vacio(l.get("Precio")) else p["Precio"])
        iva = float(p["IVA %"]) if not _vacio(p["IVA %"]) else cfg["IVA por defecto"]
        desc = float(l.get("Descuento %") or 0)
        cant = int(l["Cantidad"])
        base, cuota, total = calcular_linea(cant, precio, desc, iva, cfg["iva_incluido"])
        _append(wb, "Líneas", [pid, sku, p["Nombre"], cant, precio, desc, iva, base, cuota, total])
        _append(wb, "Movimientos", [fecha, sku, "Venta", -cant, None, pid, ""])
    _guardar(wb, ruta)
    return pid


def anular_pedido(ruta, pid: str):
    """Solo pedidos sin facturar: devuelve el stock y marca el pedido como Anulado."""
    dfs = cargar(ruta)
    ped = dfs["Pedidos"].set_index("Nº pedido")
    if ped.loc[pid, "Estado"] != ESTADO_PENDIENTE:
        raise ValueError("Solo se pueden anular pedidos pendientes de facturar")
    wb = _abrir(ruta)
    ws = wb["Pedidos"]
    ws.cell(row=_buscar_fila(ws, pid), column=5, value=ESTADO_ANULADO)
    for _, l in dfs["Líneas"][dfs["Líneas"]["Nº pedido"] == pid].iterrows():
        _append(wb, "Movimientos", [dt.date.today(), l["SKU"], "Venta", int(l["Cantidad"]), None, pid,
                                    "Anulación de pedido"])
    _guardar(wb, ruta)


def facturar_pedido(ruta, pid: str, fecha: dt.date | None = None) -> str:
    dfs = cargar(ruta)
    cfg = config(dfs)
    fecha = fecha or dt.date.today()
    ped = dfs["Pedidos"].set_index("Nº pedido")
    if pid not in ped.index:
        raise ValueError(f"Pedido {pid} no existe")
    if ped.loc[pid, "Estado"] != ESTADO_PENDIENTE:
        raise ValueError(f"El pedido {pid} ya está {ped.loc[pid, 'Estado'].lower()}")
    lin = dfs["Líneas"][dfs["Líneas"]["Nº pedido"] == pid]
    base = round(pd.to_numeric(lin["Base"]).sum(), 2)
    cuota = round(pd.to_numeric(lin["Cuota IVA"]).sum(), 2)
    total = round(pd.to_numeric(lin["Total"]).sum(), 2)
    cid = ped.loc[pid, "ID cliente"]
    cli = dfs["Clientes"].set_index("ID").loc[cid]

    serie = str(cfg.get("Serie factura") or "")
    fid = _siguiente(f"{serie}{fecha.year}-", dfs["Facturas"]["Nº factura"])
    # Fecha no anterior a la última factura de la serie (correlatividad)
    fac = dfs["Facturas"]
    if not fac.empty:
        ult = max((_a_fecha(f) for f in fac["Fecha"] if _a_fecha(f)), default=None)
        if ult and fecha < ult:
            raise ValueError(f"La fecha no puede ser anterior a la última factura ({ult:%d/%m/%Y})")

    wb = _abrir(ruta)
    _append(wb, "Facturas", [fid, fecha, pid, cid, cli["Nombre"], cli["NIF"], base, cuota, total, "Pendiente", None])
    ws = wb["Pedidos"]
    r = _buscar_fila(ws, pid)
    ws.cell(row=r, column=5, value=ESTADO_FACTURADO)
    ws.cell(row=r, column=6, value=fid)
    _guardar(wb, ruta)
    return fid


def marcar_cobro(ruta, fid: str, cobrada: bool = True, fecha: dt.date | None = None):
    wb = _abrir(ruta)
    ws = wb["Facturas"]
    r = _buscar_fila(ws, fid)
    ws.cell(row=r, column=10, value="Cobrada" if cobrada else "Pendiente")
    c = ws.cell(row=r, column=11, value=(fecha or dt.date.today()) if cobrada else None)
    c.number_format = FECHA
    _guardar(wb, ruta)


def nuevo_gasto(ruta, fecha, concepto, categoria, tipo, total, iva=0.21, proveedor="", notas=""):
    base = round(float(total) / (1 + float(iva or 0)), 2)
    wb = _abrir(ruta)
    _append(wb, "Gastos", [fecha, concepto, categoria, tipo, proveedor, base, float(iva or 0),
                           round(float(total) - base, 2), round(float(total), 2), notas])
    _guardar(wb, ruta)


def calcular_margen(*, unidades, impresion=0.0, envio=0.0, aranceles=0.0, despacho=0.0, otros_lote=0.0,
                    iva_importacion=0.0, iva_deducible=False, pvp=14.99, iva_venta=0.21,
                    repercute_iva=True, embalaje=0.0, envio_cliente=0.0, comision_pct=0.0,
                    gastos_generales=0.0, unidades_previstas=None) -> dict:
    """
    Margen por unidad de un lote.
    - Si el IVA es deducible (estás dada de alta en régimen general), el IVA de importación no es coste.
    - gastos_generales: gastos fijos a repartir entre las unidades previstas (web, marca, publicidad…).
    """
    unidades = max(int(unidades), 1)
    prev = max(int(unidades_previstas or unidades), 1)
    coste_lote = impresion + envio + aranceles + despacho + otros_lote + (0 if iva_deducible else iva_importacion)
    coste_ud = coste_lote / unidades
    ingreso_ud = pvp / (1 + iva_venta) if repercute_iva else pvp
    iva_ud = pvp - ingreso_ud
    comision_ud = pvp * comision_pct
    variable_ud = embalaje + envio_cliente + comision_ud
    generales_ud = gastos_generales / prev
    coste_total_ud = coste_ud + variable_ud + generales_ud
    margen_ud = ingreso_ud - coste_total_ud
    contrib_ud = ingreso_ud - variable_ud           # lo que deja cada venta para cubrir lote + fijos
    equilibrio = None if contrib_ud <= 0 else -(-(coste_lote + gastos_generales) // contrib_ud)
    fijo_ud = coste_ud + embalaje + envio_cliente + generales_ud
    factor = (1 / (1 + iva_venta) if repercute_iva else 1) - comision_pct
    pvp_minimo = fijo_ud / factor if factor > 0 else None
    return {
        "coste_lote": coste_lote, "coste_ud": coste_ud, "ingreso_ud": ingreso_ud, "iva_ud": iva_ud,
        "comision_ud": comision_ud, "variable_ud": variable_ud, "generales_ud": generales_ud,
        "coste_total_ud": coste_total_ud, "margen_ud": margen_ud,
        "margen_pct": margen_ud / ingreso_ud if ingreso_ud else 0.0,
        "beneficio_previsto": margen_ud * prev, "equilibrio_uds": int(equilibrio) if equilibrio else None,
        "pvp_minimo": pvp_minimo, "unidades_previstas": prev,
    }


def registrar_lote(ruta, sku, unidades, coste_ud, componentes: dict, fecha=None, referencia="", proveedor=""):
    """Entrada de stock + una fila de gasto 'Compra de stock' por cada componente con importe."""
    fecha = fecha or dt.date.today()
    wb = _abrir(ruta)
    _append(wb, "Movimientos", [fecha, sku, "Entrada", int(unidades), round(coste_ud, 4), referencia, "Lote"])
    cat = {"Impresión": "Impresión", "Envío": "Envío", "Aranceles": "Aduanas", "Gestión aduanera": "Aduanas",
           "IVA de importación": "Impuestos", "Otros del lote": "Otros"}
    for nombre, importe in componentes.items():
        if importe:
            _append(wb, "Gastos", [fecha, f"{nombre} – {referencia or sku}", cat.get(nombre, "Otros"),
                                   "Compra de stock", proveedor, round(importe, 2), 0, 0, round(importe, 2), ""])
    _guardar(wb, ruta)


# --------------------------------------------------------------------------- #
# PDF de la factura
# --------------------------------------------------------------------------- #
def _fuentes_pdf():
    """Usa Nunito Sans / ADLaM Display si están en la carpeta fuentes/; si no, Helvetica."""
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont

    carpeta = Path(__file__).parent / "fuentes"
    ttf = list(carpeta.glob("*.ttf"))

    def buscar(nombre, estilo):
        for f in ttf:
            n = f.stem.lower().replace(" ", "")
            if nombre in n and estilo in n and not any(x in n for x in ("italic", "semi", "extra", "variable")):
                return f
        return None

    normal, negrita = "Helvetica", "Helvetica-Bold"
    try:
        r, b = buscar("nunitosans", "regular"), buscar("nunitosans", "-bold")
        if r:
            pdfmetrics.registerFont(TTFont("Nunito", str(r)))
            normal = negrita = "Nunito"
        if b:
            pdfmetrics.registerFont(TTFont("Nunito-Bold", str(b)))
            negrita = "Nunito-Bold"
        if r:
            from reportlab.lib.fonts import addMapping
            addMapping("Nunito", 0, 0, "Nunito"); addMapping("Nunito", 1, 0, negrita)
            addMapping("Nunito", 0, 1, "Nunito"); addMapping("Nunito", 1, 1, negrita)
        t = buscar("adlam", "")
        titulo = negrita
        if t:
            pdfmetrics.registerFont(TTFont("ADLaM", str(t)))
            titulo = "ADLaM"
    except Exception:
        normal, negrita, titulo = "Helvetica", "Helvetica-Bold", "Helvetica-Bold"
    return normal, negrita, titulo


def pdf_factura(dfs, fid: str) -> bytes:
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_RIGHT
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.lib.units import mm
    from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

    cfg = config(dfs)
    F, FB, FT = _fuentes_pdf()
    fac = dfs["Facturas"].set_index("Nº factura").loc[fid]
    pid = fac["Nº pedido"]
    lin = dfs["Líneas"][dfs["Líneas"]["Nº pedido"] == pid]
    cli_df = dfs["Clientes"].set_index("ID")
    cli = cli_df.loc[fac["ID cliente"]] if fac["ID cliente"] in cli_df.index else None
    fecha = _a_fecha(fac["Fecha"])

    azul = colors.HexColor("#041539")   # azul noche XINGON!
    amarillo = colors.HexColor("#EAAD02")
    gris = colors.HexColor("#EEF1F8")
    ss = getSampleStyleSheet()
    normal = ParagraphStyle("n", parent=ss["Normal"], fontName=F, fontSize=9.5, leading=13)
    peq = ParagraphStyle("p", parent=normal, fontSize=8.5, textColor=colors.HexColor("#555555"))
    titulo = ParagraphStyle("t", parent=normal, fontName=FT, fontSize=24, textColor=azul,
                            alignment=TA_RIGHT, leading=30)
    der = ParagraphStyle("d", parent=normal, alignment=TA_RIGHT)
    
    def txt(v):
        return "" if _vacio(v) else str(v)

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=18 * mm, rightMargin=18 * mm,
                            topMargin=16 * mm, bottomMargin=16 * mm, title=f"Factura {fid}")
    W = A4[0] - 36 * mm
    el = []

    emisor = (f"<b>{txt(cfg['Nombre o razón social'])}</b><br/>NIF: {txt(cfg['NIF'])}<br/>"
              f"{txt(cfg['Dirección'])}<br/>{txt(cfg['CP y ciudad'])}")
    if txt(cfg.get("Email")):
        emisor += f"<br/>{txt(cfg['Email'])}"
    if txt(cfg.get("Teléfono")):
        emisor += f"<br/>Tel. {txt(cfg['Teléfono'])}"
    cab = Table([[Paragraph(emisor, normal),
                  [Paragraph("FACTURA", titulo), Spacer(1, 3),
                   Paragraph(f"<b>Nº:</b> {fid}<br/><b>Fecha:</b> {fecha:%d/%m/%Y}<br/>"
                             f"<b>Pedido:</b> {pid}", der)]]],
                colWidths=[W * 0.55, W * 0.45])
    cab.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"),
                             ("LINEBELOW", (0, 0), (-1, 0), 3, amarillo), ("BOTTOMPADDING", (0, 0), (-1, 0), 10)]))
    el += [cab, Spacer(1, 10 * mm)]

    if cli is not None:
        ctxt = f"<b>{txt(cli['Nombre'])}</b>"
        if txt(cli["NIF"]):
            ctxt += f"<br/>NIF: {txt(cli['NIF'])}"
        if txt(cli["Dirección"]):
            ctxt += f"<br/>{txt(cli['Dirección'])}"
        if txt(cli["CP"]) or txt(cli["Ciudad"]):
            ctxt += f"<br/>{txt(cli['CP'])} {txt(cli['Ciudad'])}"
    else:
        ctxt = f"<b>{txt(fac['Cliente'])}</b>"
    caja = Table([[Paragraph("FACTURAR A", ParagraphStyle("x", parent=peq, fontName=FB))],
                  [Paragraph(ctxt, normal)]], colWidths=[W * 0.55], hAlign="LEFT")
    caja.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), gris), ("LEFTPADDING", (0, 0), (-1, -1), 8),
                              ("TOPPADDING", (0, 0), (-1, -1), 5), ("BOTTOMPADDING", (0, -1), (-1, -1), 8)]))
    el += [caja, Spacer(1, 8 * mm)]

    filas = [["Descripción", "Cant.", "Precio unit.\n(sin IVA)", "Dto.", "IVA", "Base"]]
    por_iva: dict[float, list[float]] = {}
    for _, l in lin.iterrows():
        iva = float(l["IVA %"] or 0)
        precio = float(l["Precio"])
        unit_sin = precio / (1 + iva) if cfg["iva_incluido"] else precio
        filas.append([Paragraph(txt(l["Descripción"]), normal), f"{int(l['Cantidad'])}", eur(unit_sin),
                      pct(l["Descuento %"] or 0) if float(l["Descuento %"] or 0) else "—",
                      pct(iva), eur(l["Base"])])
        acc = por_iva.setdefault(iva, [0.0, 0.0])
        acc[0] += float(l["Base"])
        acc[1] += float(l["Cuota IVA"])
    t = Table(filas, colWidths=[W * 0.40, W * 0.08, W * 0.16, W * 0.09, W * 0.10, W * 0.17], repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), azul), ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), FB), ("FONTNAME", (1, 1), (-1, -1), F), ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("ALIGN", (1, 0), (-1, -1), "RIGHT"), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, gris]),
        ("LINEBELOW", (0, -1), (-1, -1), 0.5, azul),
        ("TOPPADDING", (0, 0), (-1, -1), 5), ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    el += [t, Spacer(1, 6 * mm)]

    tot = []
    for iva, (b, c) in sorted(por_iva.items()):
        tot.append([f"Base imponible ({pct(iva)})", eur(b)])
        tot.append([f"IVA {pct(iva)}", eur(c)])
    tot.append(["TOTAL", eur(fac["Total"])])
    tt = Table(tot, colWidths=[W * 0.28, W * 0.17], hAlign="RIGHT")
    tt.setStyle(TableStyle([
        ("ALIGN", (1, 0), (1, -1), "RIGHT"), ("FONTNAME", (0, 0), (-1, -1), F), ("FONTSIZE", (0, 0), (-1, -1), 9.5),
        ("FONTNAME", (0, -1), (-1, -1), FB), ("FONTSIZE", (0, -1), (-1, -1), 12),
        ("LINEBELOW", (0, -1), (-1, -1), 3, amarillo),
        ("TEXTCOLOR", (0, -1), (-1, -1), colors.white), ("BACKGROUND", (0, -1), (-1, -1), azul),
        ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    el += [tt, Spacer(1, 12 * mm)]

    pie = []
    if txt(cfg.get("IBAN")):
        pie.append(f"<b>Forma de pago:</b> transferencia a {txt(cfg['IBAN'])}")
    if txt(fac["Cobro"]) == "Cobrada":
        pie.append(f"<b>Estado:</b> cobrada el {_a_fecha(fac['Fecha cobro']):%d/%m/%Y}"
                   if _a_fecha(fac["Fecha cobro"]) else "<b>Estado:</b> cobrada")
    if txt(cfg.get("Texto al pie")):
        pie.append(txt(cfg["Texto al pie"]))
    for p in pie:
        el.append(Paragraph(p, peq))

    doc.build(el)
    return buf.getvalue()
