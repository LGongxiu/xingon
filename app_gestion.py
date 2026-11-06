"""
App de Streamlit: pedidos, stock y facturas con un Excel como base de datos.

Ejecutar en local:   streamlit run app.py
El Excel por defecto es gestion_xingon.xlsx (en la misma carpeta). Si no existe, se crea.
"""
import datetime as dt
import hashlib
import os
from pathlib import Path

import pandas as pd
import streamlit as st

import estilo
import gestion as g

st.set_page_config(page_title="XINGON! · Pedidos y facturas", page_icon="🃏", layout="wide")
estilo.aplicar()

RUTA = Path(os.environ.get("EXCEL_GESTION", Path(__file__).parent / "gestion_xingon.xlsx"))
if not RUTA.exists():
    g.crear_plantilla(RUTA)


def recargar():
    st.rerun()


def aviso(msg):
    st.session_state["_aviso"] = msg


dfs = g.cargar(RUTA)
cfg = g.config(dfs)

# --------------------------------------------------------------------------- #
# Barra lateral
# --------------------------------------------------------------------------- #
with st.sidebar:
    st.title("XINGON!")
    st.caption("Pedidos · stock · facturas")
    pagina = st.radio("Sección", ["📊 Panel", "🛒 Nuevo pedido", "📦 Pedidos y facturas", "🏷️ Stock",
                                  "🧮 Márgenes", "💸 Gastos", "🚚 Proveedores", "✅ Tareas",
                                  "🎴 Productos", "👥 Clientes", "⚙️ Mis datos"], label_visibility="collapsed")
    st.divider()
    st.caption(f"Excel: `{RUTA.name}`")
    st.download_button("⬇️ Descargar Excel", RUTA.read_bytes(), file_name=RUTA.name,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       use_container_width=True)
    subido = st.file_uploader("Cargar mi Excel", type=["xlsx"],
                              help="Sustituye el Excel con el que trabaja la app por el tuyo.")
    if subido is not None:
        h = hashlib.md5(subido.getvalue()).hexdigest()
        if st.session_state.get("_subido") != h:
            RUTA.write_bytes(subido.getvalue())
            st.session_state["_subido"] = h
            aviso("Excel cargado.")
            recargar()
    st.caption("En Streamlit Cloud los cambios se pierden al reiniciarse la app: "
               "descarga el Excel después de trabajar.")

if msg := st.session_state.pop("_aviso", None):
    st.success(msg)

if str(cfg.get("NIF", "")) in ("", "00000000X"):
    st.warning("Aún no has puesto tus datos fiscales. Ve a **⚙️ Mis datos** antes de emitir facturas.")

stock = g.stock_actual(dfs)
productos = dfs["Productos"]
clientes = dfs["Clientes"]
opc_cli = {f"{r.ID} — {r.Nombre}": r.ID for r in clientes.itertuples()}
opc_prod = {f"{r.SKU} — {r.Nombre}": r.SKU for r in productos.itertuples()}
gastos = dfs["Gastos"].copy()
gastos["Fecha"] = pd.to_datetime(gastos["Fecha"], errors="coerce")
proveedores = dfs["Proveedores"]
tareas = dfs["Tareas"].copy()
tareas["Fecha límite"] = pd.to_datetime(tareas["Fecha límite"], errors="coerce")


def gastos_generales(anio=None, solo_base=False):
    gg = gastos[gastos["Tipo"] == "Gasto general"]
    if anio:
        gg = gg[gg["Fecha"].dt.year == anio]
    return float(pd.to_numeric(gg["Base" if solo_base else "Total"], errors="coerce").sum())


def editor_guardar(df, hoja, column_config, texto_boton, prep=None, **kw):
    ed = st.data_editor(df, num_rows="dynamic", use_container_width=True, hide_index=True,
                        column_config=column_config, key=f"ed_{hoja}", **kw)
    if st.button(f"💾 {texto_boton}", type="primary", key=f"save_{hoja}"):
        if prep:
            ed = prep(ed)
            if ed is None:
                return
        g.reemplazar_hoja(RUTA, hoja, ed)
        aviso("Cambios guardados.")
        recargar()


# --------------------------------------------------------------------------- #
# Panel
# --------------------------------------------------------------------------- #
if pagina == "📊 Panel":
    st.header("Panel")
    fac = dfs["Facturas"].copy()
    fac["Fecha"] = pd.to_datetime(fac["Fecha"], errors="coerce")
    anio = st.selectbox("Año", sorted({dt.date.today().year, *fac["Fecha"].dt.year.dropna().astype(int)},
                                      reverse=True))
    f = fac[fac["Fecha"].dt.year == anio]
    mov = dfs["Movimientos"]
    vendidas = -pd.to_numeric(mov.loc[mov["Tipo"] == "Venta", "Cantidad"], errors="coerce").sum()
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Facturado (IVA incl.)", g.eur(f["Total"].sum()))
    c2.metric("Base imponible", g.eur(f["Base"].sum()), help="Lo que es tuyo antes de IVA")
    c3.metric("IVA repercutido", g.eur(f["IVA"].sum()), help="Lo que tendrías que ingresar a Hacienda")
    c4.metric("Pendiente de cobro", g.eur(f.loc[f["Cobro"] == "Pendiente", "Total"].sum()))

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Unidades vendidas (total)", int(vendidas))
    c2.metric("Unidades en stock", int(stock["Stock actual"].sum()))
    c3.metric("Pedidos por facturar", int((dfs["Pedidos"]["Estado"] == g.ESTADO_PENDIENTE).sum()))
    # Margen estimado con el coste medio de las entradas
    ent = mov[mov["Tipo"] == "Entrada"].copy()
    ent["Cantidad"] = pd.to_numeric(ent["Cantidad"], errors="coerce")
    ent["Coste unitario"] = pd.to_numeric(ent["Coste unitario"], errors="coerce")
    inversion = (ent["Cantidad"] * ent["Coste unitario"]).sum()
    base_total = pd.to_numeric(dfs["Facturas"]["Base"], errors="coerce").sum()
    c4.metric("Base facturada − inversión en stock", g.eur(base_total - inversion),
              help="Todas las facturas menos lo que has pagado por todo el stock comprado "
                   "(si es positivo, ya has recuperado la inversión).")

    # Beneficio estimado: base facturada - coste de lo vendido - gastos generales (todo del año)
    coste_medio = inversion / ent["Cantidad"].sum() if ent["Cantidad"].sum() else 0
    mov_anio = mov[pd.to_datetime(mov["Fecha"], errors="coerce").dt.year == anio]
    vend_anio = -pd.to_numeric(mov_anio.loc[mov_anio["Tipo"] == "Venta", "Cantidad"], errors="coerce").sum()
    gg = gastos_generales(anio)
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Gastos generales", g.eur(gg), help="Hoja Gastos, tipo 'Gasto general', este año (IVA incl.)")
    c2.metric("Coste de lo vendido", g.eur(vend_anio * coste_medio),
              help="Unidades vendidas este año × coste medio por unidad de tus lotes")
    c3.metric("Beneficio estimado", g.eur(f["Base"].sum() - vend_anio * coste_medio - gg),
              help="Base facturada − coste de lo vendido − gastos generales")
    pend_t = tareas[tareas["Estado"] != "Hecha"]
    vencidas = int((pend_t["Fecha límite"] < pd.Timestamp(dt.date.today())).sum())
    c4.metric("Tareas pendientes", len(pend_t), delta=f"{vencidas} vencidas" if vencidas else None,
              delta_color="inverse")

    bajos = stock[stock["Alerta"]]
    if not bajos.empty:
        st.error("Stock bajo: " + ", ".join(f"{r['Nombre']} ({r['Stock actual']} uds)" for _, r in bajos.iterrows()))

    if not f.empty:
        st.subheader("Facturación por mes")
        mes = f.groupby(f["Fecha"].dt.strftime("%Y-%m"))["Total"].sum()
        st.bar_chart(mes, y_label="€", color=estilo.COLORES["amarillo"])


# --------------------------------------------------------------------------- #
# Nuevo pedido
# --------------------------------------------------------------------------- #
elif pagina == "🛒 Nuevo pedido":
    st.header("Nuevo pedido")
    if productos.empty:
        st.info("Primero añade algún producto en 🎴 Productos.")
        st.stop()

    with st.expander("➕ Cliente nuevo"):
        with st.form("cli_nuevo", clear_on_submit=True):
            a, b = st.columns(2)
            nombre = a.text_input("Nombre *")
            nif = b.text_input("NIF / DNI", help="Necesario en factura completa si el cliente lo pide "
                                                  "(empresas, profesionales).")
            direccion = a.text_input("Dirección")
            cp = b.text_input("CP")
            ciudad = a.text_input("Ciudad")
            email = b.text_input("Email")
            tel = a.text_input("Teléfono")
            if st.form_submit_button("Guardar cliente"):
                if not nombre.strip():
                    st.error("El nombre es obligatorio.")
                else:
                    cid = g.nuevo_cliente(RUTA, nombre.strip(), nif.strip(), direccion, cp, ciudad, email, tel)
                    st.session_state["cli_sel"] = f"{cid} — {nombre.strip()}"
                    aviso(f"Cliente {cid} creado.")
                    recargar()

    a, b = st.columns([2, 1])
    claves = list(opc_cli)
    idx = claves.index(st.session_state["cli_sel"]) if st.session_state.get("cli_sel") in claves else 0
    cli_txt = a.selectbox("Cliente", claves, index=idx)
    fecha = b.date_input("Fecha", dt.date.today(), format="DD/MM/YYYY")

    st.caption("Añade una fila por producto. Precios " +
               ("con IVA incluido." if cfg["iva_incluido"] else "sin IVA."))
    base_lineas = pd.DataFrame({"Producto": [next(iter(opc_prod))], "Cantidad": [1], "Descuento %": [0.0]})
    ed = st.data_editor(
        base_lineas, num_rows="dynamic", use_container_width=True, key="lineas",
        column_config={
            "Producto": st.column_config.SelectboxColumn(options=list(opc_prod), required=True, width="large"),
            "Cantidad": st.column_config.NumberColumn(min_value=1, step=1, required=True),
            "Descuento %": st.column_config.NumberColumn(min_value=0.0, max_value=100.0, step=5.0, format="%.0f %%"),
        },
    )
    ed = ed.dropna(subset=["Producto"])
    lineas, filas_prev, avisos_stock = [], [], []
    disp = stock.set_index("SKU")["Stock actual"].to_dict()
    pinfo = productos.set_index("SKU")
    for _, r in ed.iterrows():
        sku = opc_prod[r["Producto"]]
        cant = 0 if pd.isna(r["Cantidad"]) else int(r["Cantidad"])
        desc = 0.0 if pd.isna(r["Descuento %"]) else float(r["Descuento %"]) / 100
        p = pinfo.loc[sku]
        iva = float(p["IVA %"]) if not pd.isna(p["IVA %"]) else cfg["IVA por defecto"]
        base, cuota, total = g.calcular_linea(cant, p["Precio"], desc, iva, cfg["iva_incluido"])
        lineas.append({"SKU": sku, "Cantidad": cant, "Descuento %": desc})
        filas_prev.append({"Producto": p["Nombre"], "Cant.": cant, "Base": base, "IVA": cuota, "Total": total})
        if cant > disp.get(sku, 0):
            avisos_stock.append(f"{p['Nombre']}: pides {cant} y hay {disp.get(sku, 0)} en stock")

    if filas_prev:
        prev = pd.DataFrame(filas_prev)
        st.dataframe(prev.style.format({"Base": g.eur, "IVA": g.eur, "Total": g.eur}),
                     hide_index=True, use_container_width=True)
        a, b, c = st.columns(3)
        a.metric("Base", g.eur(prev["Base"].sum()))
        b.metric("IVA", g.eur(prev["IVA"].sum()))
        c.metric("Total", g.eur(prev["Total"].sum()))
    for w in avisos_stock:
        st.warning("⚠️ " + w)
    notas = st.text_input("Notas (opcional)", placeholder="p. ej. entregado en mano, envío por correo…")
    facturar_ya = st.checkbox("Facturar ahora", value=True)

    if st.button("✅ Crear pedido", type="primary", disabled=not lineas):
        try:
            pid = g.crear_pedido(RUTA, opc_cli[cli_txt], lineas, fecha, notas)
            texto = f"Pedido {pid} creado."
            if facturar_ya:
                fid = g.facturar_pedido(RUTA, pid, fecha)
                st.session_state["ultima_factura"] = fid
                texto += f" Factura {fid} emitida: descárgala en 📦 Pedidos y facturas."
            st.session_state.pop("lineas", None)
            aviso(texto)
            recargar()
        except ValueError as e:
            st.error(str(e))


# --------------------------------------------------------------------------- #
# Pedidos y facturas
# --------------------------------------------------------------------------- #
elif pagina == "📦 Pedidos y facturas":
    st.header("Pedidos y facturas")
    t1, t2 = st.tabs(["Facturas", "Pedidos"])

    with t1:
        fac = dfs["Facturas"]
        if fac.empty:
            st.info("Todavía no hay facturas.")
        else:
            vista = fac.copy()
            vista["Fecha"] = pd.to_datetime(vista["Fecha"]).dt.strftime("%d/%m/%Y")
            st.dataframe(vista[["Nº factura", "Fecha", "Cliente", "Base", "IVA", "Total", "Cobro"]]
                         .style.format({"Base": g.eur, "IVA": g.eur, "Total": g.eur}),
                         hide_index=True, use_container_width=True)
            ids = list(fac["Nº factura"])[::-1]
            ult = st.session_state.get("ultima_factura")
            fid = st.selectbox("Factura", ids, index=ids.index(ult) if ult in ids else 0)
            fila = fac.set_index("Nº factura").loc[fid]
            a, b, c = st.columns(3)
            a.download_button("⬇️ Descargar PDF", g.pdf_factura(dfs, fid), file_name=f"Factura_{fid}.pdf",
                              mime="application/pdf", type="primary", use_container_width=True)
            if fila["Cobro"] == "Pendiente":
                fc = b.date_input("Fecha de cobro", dt.date.today(), format="DD/MM/YYYY")
                if c.button("💶 Marcar como cobrada", use_container_width=True):
                    g.marcar_cobro(RUTA, fid, True, fc)
                    aviso(f"{fid} marcada como cobrada.")
                    recargar()
            else:
                b.success(f"Cobrada el {pd.to_datetime(fila['Fecha cobro']):%d/%m/%Y}"
                          if not pd.isna(fila["Fecha cobro"]) else "Cobrada")
                if c.button("↩️ Volver a pendiente", use_container_width=True):
                    g.marcar_cobro(RUTA, fid, False)
                    recargar()

    with t2:
        ped = dfs["Pedidos"]
        if ped.empty:
            st.info("Todavía no hay pedidos.")
        else:
            vista = ped.copy()
            vista["Fecha"] = pd.to_datetime(vista["Fecha"]).dt.strftime("%d/%m/%Y")
            tot = dfs["Líneas"].groupby("Nº pedido")["Total"].sum()
            vista["Total"] = vista["Nº pedido"].map(tot)
            st.dataframe(vista[["Nº pedido", "Fecha", "Cliente", "Total", "Estado", "Nº factura", "Notas"]]
                         .style.format({"Total": g.eur}, na_rep=""),
                         hide_index=True, use_container_width=True)
            pend = list(ped.loc[ped["Estado"] == g.ESTADO_PENDIENTE, "Nº pedido"])
            if pend:
                st.subheader("Pedidos pendientes de facturar")
                a, b, c, d = st.columns([2, 2, 1.3, 1.3])
                pid = a.selectbox("Pedido", pend)
                ff = b.date_input("Fecha de factura", dt.date.today(), format="DD/MM/YYYY", key="ff")
                if c.button("🧾 Facturar", type="primary", use_container_width=True):
                    try:
                        fid = g.facturar_pedido(RUTA, pid, ff)
                        st.session_state["ultima_factura"] = fid
                        aviso(f"Factura {fid} emitida. Descárgala en la pestaña Facturas.")
                        recargar()
                    except ValueError as e:
                        st.error(str(e))
                if d.button("🗑️ Anular", use_container_width=True,
                            help="Marca el pedido como anulado y devuelve las unidades al stock."):
                    g.anular_pedido(RUTA, pid)
                    aviso(f"Pedido {pid} anulado; stock devuelto.")
                    recargar()


# --------------------------------------------------------------------------- #
# Stock
# --------------------------------------------------------------------------- #
elif pagina == "🏷️ Stock":
    st.header("Stock")
    if productos.empty:
        st.info("Primero añade algún producto en 🎴 Productos.")
        st.stop()
    st.dataframe(stock[["SKU", "Nombre", "Stock actual", "Stock mínimo", "Alerta"]],
                 hide_index=True, use_container_width=True,
                 column_config={"Alerta": st.column_config.CheckboxColumn("¿Reponer?")})

    st.subheader("Registrar entrada o ajuste")
    with st.form("mov", clear_on_submit=True):
        a, b, c = st.columns(3)
        prod_txt = a.selectbox("Producto", list(opc_prod))
        tipo = b.selectbox("Tipo", ["Entrada", "Ajuste"],
                           help="Entrada = te llegan unidades (p. ej. de la imprenta). "
                                "Ajuste = corregir un recuento (+ o −), regalos, roturas…")
        cant = c.number_input("Cantidad", value=50, step=1,
                              help="En un ajuste, pon negativo para restar (p. ej. -2 por 2 regaladas).")
        coste = a.number_input("Coste por unidad (€)", min_value=0.0, step=0.1, format="%.2f",
                               help="Impresión + envío + aduanas ÷ unidades. Solo en entradas.")
        ref = b.text_input("Referencia", placeholder="p. ej. pedido MPC nº…")
        fecha = c.date_input("Fecha", dt.date.today(), format="DD/MM/YYYY")
        notas = st.text_input("Notas")
        if st.form_submit_button("Guardar movimiento", type="primary"):
            if cant == 0:
                st.error("La cantidad no puede ser 0.")
            else:
                g.registrar_movimiento(RUTA, opc_prod[prod_txt], int(cant), tipo,
                                       coste if tipo == "Entrada" and coste else None, ref, notas, fecha)
                aviso("Movimiento guardado.")
                recargar()

    st.subheader("Historial de movimientos")
    mov = dfs["Movimientos"].copy()
    if not mov.empty:
        mov["Fecha"] = pd.to_datetime(mov["Fecha"]).dt.strftime("%d/%m/%Y")
        st.dataframe(mov.iloc[::-1], hide_index=True, use_container_width=True)


# --------------------------------------------------------------------------- #
# Márgenes
# --------------------------------------------------------------------------- #
elif pagina == "🧮 Márgenes":
    st.header("Calculadora de márgenes")
    st.caption("Calcula cuánto te cuesta cada baraja de un lote y cuánto ganas al venderla. "
               "Pon los importes tal como aparecen en las facturas de la imprenta y del transportista.")

    if productos.empty:
        st.info("Primero añade algún producto en 🎴 Productos.")
        st.stop()
    prod_txt = st.selectbox("Producto", list(opc_prod))
    sku = opc_prod[prod_txt]
    p = productos.set_index("SKU").loc[sku]
    iva_prod = float(p["IVA %"]) if not pd.isna(p["IVA %"]) else cfg["IVA por defecto"]

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("1 · Coste del lote")
        uds = st.number_input("Unidades del lote", min_value=1, value=50, step=10)
        impresion = st.number_input("Impresión (total €)", min_value=0.0, value=0.0, step=10.0, format="%.2f")
        envio = st.number_input("Envío hasta España (€)", min_value=0.0, value=0.0, step=5.0, format="%.2f")
        aranceles = st.number_input("Aranceles / derechos de aduana (€)", min_value=0.0, value=0.0, step=1.0,
                                    format="%.2f")
        despacho = st.number_input("Gestión aduanera del transportista (€)", min_value=0.0, value=0.0,
                                   step=1.0, format="%.2f", help="Lo que cobra DHL/UPS/FedEx por tramitar la aduana.")
        auto_iva = st.toggle("Calcular IVA de importación (21 %)", value=True,
                             help="El IVA de importación se paga sobre impresión + envío + aranceles. "
                                  "Desactívalo para poner el importe exacto de tu factura de aduanas.")
        if auto_iva:
            iva_imp = round((impresion + envio + aranceles) * 0.21, 2)
            st.caption(f"IVA de importación: **{g.eur(iva_imp)}**")
        else:
            iva_imp = st.number_input("IVA de importación pagado (€)", min_value=0.0, value=0.0, format="%.2f")
        otros_lote = st.number_input("Otros gastos del lote (€)", min_value=0.0, value=0.0, format="%.2f",
                                     help="Muestras, cajas, cambio de divisa…")
        deducible = st.toggle("Puedo deducirme el IVA", value=False,
                              help="Actívalo solo si estás dada de alta y en régimen general de IVA. "
                                   "Si no, el IVA que pagas es un coste más.")

    with col2:
        st.subheader("2 · Venta")
        pvp = st.number_input("PVP al cliente (€)", min_value=0.0, value=float(p["Precio"]), step=0.5, format="%.2f")
        repercute = st.toggle("El PVP incluye IVA que ingreso a Hacienda", value=cfg["iva_incluido"],
                              help="Si estás dada de alta, del PVP de 14,99 € una parte es IVA y no es tuya.")
        iva_v = st.number_input("IVA de la venta (%)", 0.0, 100.0, iva_prod * 100, step=1.0) / 100
        embalaje = st.number_input("Embalaje por unidad (€)", min_value=0.0, value=0.0, step=0.1, format="%.2f")
        envio_cli = st.number_input("Envío al cliente por unidad, si lo pagas tú (€)", min_value=0.0, value=0.0,
                                    step=0.5, format="%.2f")
        comision = st.number_input("Comisión de cobro (%)", 0.0, 30.0, 0.0, step=0.5,
                                   help="Bizum o efectivo: 0 %. Tiendas online o TPV: suele ser 1–4 %.") / 100

        st.subheader("3 · Gastos generales")
        anio = dt.date.today().year
        gg_def = gastos_generales(anio, solo_base=deducible)
        gg = st.number_input("Gastos generales a repartir (€)", min_value=0.0, value=round(gg_def, 2), step=10.0,
                             format="%.2f", help=f"Por defecto, los 'Gasto general' de {anio} de la pantalla 💸 Gastos "
                                                 "(web, marca, publicidad…).")
        prev = st.number_input("Unidades entre las que repartirlos", min_value=1, value=int(uds), step=10)

    r = g.calcular_margen(unidades=uds, impresion=impresion, envio=envio, aranceles=aranceles, despacho=despacho,
                          otros_lote=otros_lote, iva_importacion=iva_imp, iva_deducible=deducible, pvp=pvp,
                          iva_venta=iva_v, repercute_iva=repercute, embalaje=embalaje, envio_cliente=envio_cli,
                          comision_pct=comision, gastos_generales=gg, unidades_previstas=prev)
    st.divider()
    st.subheader("Resultado por unidad")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Coste por baraja (lote)", g.eur(r["coste_ud"]))
    c2.metric("Coste total por venta", g.eur(r["coste_total_ud"]), help="Lote + embalaje + envío + comisión + "
                                                                       "parte de gastos generales")
    c3.metric("Te queda de cada venta", g.eur(r["ingreso_ud"]), help="PVP sin el IVA que ingresas a Hacienda")
    c4.metric("Margen por unidad", g.eur(r["margen_ud"]), delta=f"{r['margen_pct'] * 100:.0f} %".replace(".", ","))
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Inversión del lote", g.eur(r["coste_lote"]))
    c2.metric(f"Beneficio si vendes {r['unidades_previstas']}", g.eur(r["beneficio_previsto"]))
    c3.metric("Unidades para no perder", r["equilibrio_uds"] if r["equilibrio_uds"] else "—",
              help="Ventas necesarias para recuperar el lote y los gastos generales")
    c4.metric("PVP mínimo sin pérdidas", g.eur(r["pvp_minimo"]) if r["pvp_minimo"] else "—")
    if r["margen_ud"] < 0:
        st.error("Con estos números pierdes dinero en cada venta. Sube el PVP o reduce costes.")
    elif r["equilibrio_uds"] and r["equilibrio_uds"] > uds:
        st.warning(f"Con un lote de {uds} no llegas a cubrir los gastos generales: necesitas vender "
                   f"{r['equilibrio_uds']} unidades.")

    desglose = pd.DataFrame({
        "Concepto": ["Lote (impresión, envío, aduanas" + ("" if deducible else ", IVA") + ")", "Embalaje",
                     "Envío al cliente", "Comisión de cobro", "Gastos generales", "IVA que ingresas",
                     "Margen"],
        "€ por unidad": [r["coste_ud"], embalaje, envio_cli, r["comision_ud"], r["generales_ud"],
                         r["iva_ud"] if repercute else 0.0, r["margen_ud"]],
    })
    st.caption(f"En qué se va cada PVP de {g.eur(pvp)}:")
    st.dataframe(desglose.style.format({"€ por unidad": g.eur}), hide_index=True, use_container_width=True)

    with st.expander("📥 Registrar este lote en el stock"):
        st.caption("Añade una entrada de stock con el coste por unidad calculado y apunta cada importe "
                   "en 💸 Gastos como 'Compra de stock'.")
        a, b, c = st.columns(3)
        f_lote = a.date_input("Fecha de llegada", dt.date.today(), format="DD/MM/YYYY")
        ref = b.text_input("Referencia", placeholder="pedido MPC nº…")
        nombres_prov = [""] + list(proveedores["Nombre"])
        prov = c.selectbox("Proveedor", nombres_prov, index=1 if len(nombres_prov) > 1 else 0)
        if st.button("Registrar lote", type="primary"):
            g.registrar_lote(RUTA, sku, uds, r["coste_ud"], {
                "Impresión": impresion, "Envío": envio, "Aranceles": aranceles, "Gestión aduanera": despacho,
                "IVA de importación": iva_imp, "Otros del lote": otros_lote}, f_lote, ref, prov)
            aviso(f"Lote registrado: +{uds} uds a {g.eur(r['coste_ud'])} por unidad.")
            recargar()


# --------------------------------------------------------------------------- #
# Gastos
# --------------------------------------------------------------------------- #
elif pagina == "💸 Gastos":
    st.header("Gastos")
    anio = dt.date.today().year
    ga = gastos[gastos["Fecha"].dt.year == anio]
    c1, c2, c3 = st.columns(3)
    c1.metric(f"Gastos generales {anio}", g.eur(ga.loc[ga["Tipo"] == "Gasto general", "Total"].sum()))
    c2.metric(f"Compras de stock {anio}", g.eur(ga.loc[ga["Tipo"] == "Compra de stock", "Total"].sum()))
    c3.metric("IVA soportado", g.eur(pd.to_numeric(ga["IVA"], errors="coerce").sum()),
              help="Solo lo puedes deducir si estás dada de alta en régimen general de IVA.")

    with st.form("gasto", clear_on_submit=True):
        st.subheader("Añadir gasto")
        a, b, c = st.columns(3)
        f_g = a.date_input("Fecha", dt.date.today(), format="DD/MM/YYYY")
        concepto = b.text_input("Concepto *", placeholder="Dominio xingon.es")
        cat = c.selectbox("Categoría", g.CATEGORIAS_GASTO, index=g.CATEGORIAS_GASTO.index("Web y dominio"))
        tipo = a.selectbox("Tipo", g.TIPOS_GASTO,
                           help="Gasto general = de la actividad. Compra de stock = forma parte del coste de un lote.")
        total = b.number_input("Importe total pagado, IVA incl. (€)", min_value=0.0, step=1.0, format="%.2f")
        iva_g = c.number_input("IVA incluido (%)", 0.0, 100.0, 21.0, step=1.0)
        prov = a.selectbox("Proveedor", [""] + list(proveedores["Nombre"]))
        notas = b.text_input("Notas")
        if st.form_submit_button("Guardar gasto", type="primary"):
            if not concepto.strip() or total <= 0:
                st.error("Pon un concepto y un importe mayor que 0.")
            else:
                g.nuevo_gasto(RUTA, f_g, concepto.strip(), cat, tipo, total, iva_g / 100, prov, notas)
                aviso("Gasto guardado.")
                recargar()

    st.subheader("Todos los gastos")
    st.caption("Puedes corregir o borrar filas en la tabla. La base y el IVA se recalculan al guardar.")
    vista = gastos.copy()
    vista["IVA %"] = pd.to_numeric(vista["IVA %"], errors="coerce") * 100

    def prep_gastos(ed):
        ed = ed.copy()
        ed["IVA %"] = pd.to_numeric(ed["IVA %"], errors="coerce").fillna(0) / 100
        return ed

    editor_guardar(vista, "Gastos", {
        "Fecha": st.column_config.DateColumn(format="DD/MM/YYYY"),
        "Concepto": st.column_config.TextColumn(width="large", required=True),
        "Categoría": st.column_config.SelectboxColumn(options=g.CATEGORIAS_GASTO),
        "Tipo": st.column_config.SelectboxColumn(options=g.TIPOS_GASTO, required=True),
        "Proveedor": st.column_config.SelectboxColumn(options=[""] + list(proveedores["Nombre"])),
        "Base": st.column_config.NumberColumn(format="%.2f €", disabled=True),
        "IVA %": st.column_config.NumberColumn(format="%.0f %%", min_value=0, max_value=100),
        "IVA": st.column_config.NumberColumn(format="%.2f €", disabled=True),
        "Total": st.column_config.NumberColumn("Total (IVA incl.)", format="%.2f €", min_value=0.0),
    }, "Guardar cambios", prep=prep_gastos)

    if not ga.empty:
        st.subheader(f"Gastos {anio} por categoría")
        st.bar_chart(ga.groupby("Categoría")["Total"].sum(), y_label="€", color=estilo.COLORES["amarillo"])


# --------------------------------------------------------------------------- #
# Proveedores
# --------------------------------------------------------------------------- #
elif pagina == "🚚 Proveedores":
    st.header("Proveedores")
    st.caption("Imprentas, transportistas, gestoría, tiendas de material… Deja el ID vacío en filas nuevas.")

    def prep_prov(ed):
        ed = ed.fillna("")
        ed = ed[ed["Nombre"].astype(str).str.strip() != ""].copy()
        usados = list(ed["ID"])
        for i in ed.index:
            if not str(ed.at[i, "ID"]).strip():
                ed.at[i, "ID"] = g._siguiente("PR-", usados)
                usados.append(ed.at[i, "ID"])
        return ed

    editor_guardar(proveedores, "Proveedores", {
        "ID": st.column_config.TextColumn(width="small"),
        "Nombre": st.column_config.TextColumn(required=True, width="medium"),
        "Web": st.column_config.LinkColumn(),
    }, "Guardar proveedores", prep=prep_prov)

    if not gastos.empty and not proveedores.empty:
        st.subheader("Gastado con cada proveedor")
        por_prov = gastos[gastos["Proveedor"].astype(str).str.strip() != ""].groupby("Proveedor")["Total"].sum()
        if not por_prov.empty:
            st.dataframe(por_prov.rename("Total").to_frame().style.format(g.eur), use_container_width=True)


# --------------------------------------------------------------------------- #
# Tareas
# --------------------------------------------------------------------------- #
elif pagina == "✅ Tareas":
    st.header("Tareas")
    hoy = pd.Timestamp(dt.date.today())
    pend = tareas[tareas["Estado"] != "Hecha"]
    c1, c2, c3 = st.columns(3)
    c1.metric("Pendientes", int((tareas["Estado"] == "Pendiente").sum()))
    c2.metric("En curso", int((tareas["Estado"] == "En curso").sum()))
    c3.metric("Vencidas", int((pend["Fecha límite"] < hoy).sum()))

    with st.form("tarea", clear_on_submit=True):
        a, b, c = st.columns([3, 1, 1])
        t_txt = a.text_input("Nueva tarea", placeholder="Pedir presupuesto de 100 barajas a MPC")
        prio = b.selectbox("Prioridad", g.PRIORIDADES, index=1)
        limite = c.date_input("Fecha límite", value=None, format="DD/MM/YYYY")
        if st.form_submit_button("Añadir tarea", type="primary") and t_txt.strip():
            nueva = pd.DataFrame([{"ID": g._siguiente("T-", tareas["ID"]), "Tarea": t_txt.strip(),
                                   "Estado": "Pendiente", "Prioridad": prio, "Fecha límite": limite, "Notas": ""}])
            g.reemplazar_hoja(RUTA, "Tareas", pd.concat([dfs["Tareas"], nueva], ignore_index=True))
            aviso("Tarea añadida.")
            recargar()

    ver_hechas = st.toggle("Mostrar tareas hechas", value=False)
    orden = {"Alta": 0, "Media": 1, "Baja": 2}
    vista = tareas.sort_values(["Fecha límite", "Prioridad"], key=lambda s: s.map(orden) if s.name == "Prioridad" else s,
                               na_position="last")
    if not ver_hechas:
        ocultas = vista[vista["Estado"] == "Hecha"]
        vista = vista[vista["Estado"] != "Hecha"]

    def prep_tareas(ed):
        ed = ed.copy()
        if not ver_hechas:
            ed = pd.concat([ed, ocultas], ignore_index=True)
        ed = ed[ed["Tarea"].astype(str).str.strip().replace("nan", "") != ""]
        usados = list(ed["ID"].dropna())
        for i in ed.index:
            if g._vacio(ed.at[i, "ID"]):
                ed.at[i, "ID"] = g._siguiente("T-", usados)
                usados.append(ed.at[i, "ID"])
            if g._vacio(ed.at[i, "Estado"]):
                ed.at[i, "Estado"] = "Pendiente"
        return ed

    st.caption("Cambia el estado en la tabla y pulsa Guardar. Para borrar una tarea, selecciona la fila y pulsa la papelera.")
    editor_guardar(vista, "Tareas", {
        "ID": st.column_config.TextColumn(width="small", disabled=True),
        "Tarea": st.column_config.TextColumn(width="large", required=True),
        "Estado": st.column_config.SelectboxColumn(options=g.ESTADOS_TAREA, required=True),
        "Prioridad": st.column_config.SelectboxColumn(options=g.PRIORIDADES),
        "Fecha límite": st.column_config.DateColumn(format="DD/MM/YYYY"),
    }, "Guardar tareas", prep=prep_tareas)


# --------------------------------------------------------------------------- #
# Productos
# --------------------------------------------------------------------------- #
elif pagina == "🎴 Productos":
    st.header("Productos")
    st.caption("Edita directamente en la tabla y pulsa Guardar. El stock no se edita aquí: "
               "se calcula con los movimientos (🏷️ Stock).")
    df = productos.drop(columns=["Stock actual"]).copy()
    df["IVA %"] = pd.to_numeric(df["IVA %"], errors="coerce") * 100
    ed = st.data_editor(
        df, num_rows="dynamic", use_container_width=True, hide_index=True,
        column_config={
            "SKU": st.column_config.TextColumn(required=True, help="Código único, p. ej. XINGON-001"),
            "Nombre": st.column_config.TextColumn(required=True, width="large"),
            "Precio": st.column_config.NumberColumn(
                "PVP (IVA incl.) €" if cfg["iva_incluido"] else "Precio sin IVA €", min_value=0.0, format="%.2f"),
            "IVA %": st.column_config.NumberColumn(min_value=0, max_value=100, step=1, format="%.0f %%"),
            "Coste unitario": st.column_config.NumberColumn("Coste unitario €", min_value=0.0, format="%.2f"),
            "Stock mínimo": st.column_config.NumberColumn(min_value=0, step=1),
        },
    )
    if st.button("💾 Guardar productos", type="primary"):
        ed = ed.dropna(subset=["SKU"])
        if ed["SKU"].duplicated().any():
            st.error("Hay SKUs repetidos.")
        else:
            ed["IVA %"] = pd.to_numeric(ed["IVA %"], errors="coerce") / 100
            g.reemplazar_hoja(RUTA, "Productos", ed)
            aviso("Productos guardados.")
            recargar()


# --------------------------------------------------------------------------- #
# Clientes
# --------------------------------------------------------------------------- #
elif pagina == "👥 Clientes":
    st.header("Clientes")
    st.caption("Edita en la tabla y pulsa Guardar. Para clientes nuevos deja el ID vacío y se asigna solo.")
    ed = st.data_editor(clientes, num_rows="dynamic", use_container_width=True, hide_index=True,
                        column_config={"ID": st.column_config.TextColumn(disabled=False),
                                       "Nombre": st.column_config.TextColumn(required=True, width="medium")})
    if st.button("💾 Guardar clientes", type="primary"):
        ed = ed.fillna("")
        ed = ed[ed["Nombre"].astype(str).str.strip() != ""].copy()
        usados = list(ed["ID"])
        for i in ed.index:
            if not str(ed.at[i, "ID"]).strip():
                ed.at[i, "ID"] = g._siguiente("C-", usados)
                usados.append(ed.at[i, "ID"])
        if ed["ID"].duplicated().any():
            st.error("Hay IDs repetidos.")
        else:
            g.reemplazar_hoja(RUTA, "Clientes", ed)
            aviso("Clientes guardados.")
            recargar()


# --------------------------------------------------------------------------- #
# Mis datos
# --------------------------------------------------------------------------- #
elif pagina == "⚙️ Mis datos":
    st.header("Mis datos (salen en la factura)")
    with st.form("cfg"):
        a, b = st.columns(2)
        nuevo = {}
        nuevo["Nombre o razón social"] = a.text_input("Nombre o razón social *", str(cfg["Nombre o razón social"]))
        nuevo["NIF"] = b.text_input("NIF *", str(cfg["NIF"]))
        nuevo["Dirección"] = a.text_input("Dirección *", str(cfg["Dirección"]))
        nuevo["CP y ciudad"] = b.text_input("CP y ciudad *", str(cfg["CP y ciudad"]))
        nuevo["Email"] = a.text_input("Email", str(cfg["Email"]))
        nuevo["Teléfono"] = b.text_input("Teléfono", str(cfg["Teléfono"]))
        nuevo["IBAN"] = a.text_input("IBAN (para cobrar por transferencia)", str(cfg["IBAN"]))
        nuevo["Serie factura"] = b.text_input("Serie de factura", cfg["Serie factura"],
                                              help="Prefijo del número: X → X2026-0001. No lo cambies a mitad de año.")
        nuevo["Precios con IVA incluido"] = "Sí" if a.toggle("Mis precios llevan el IVA incluido",
                                                             cfg["iva_incluido"]) else "No"
        nuevo["IVA por defecto"] = b.number_input("IVA por defecto (%)", 0.0, 100.0,
                                                  float(cfg["IVA por defecto"]) * 100, step=1.0) / 100
        nuevo["Texto al pie"] = st.text_area("Texto al pie de la factura", str(cfg["Texto al pie"]))
        if st.form_submit_button("💾 Guardar", type="primary"):
            g.reemplazar_hoja(RUTA, "Config", pd.DataFrame(list(nuevo.items()), columns=["Campo", "Valor"]))
            aviso("Datos guardados.")
            recargar()
