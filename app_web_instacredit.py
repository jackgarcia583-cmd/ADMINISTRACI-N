#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SISTEMA WEB DE CONTROL DE COMPRAS - INSTACREDIT S.A. (NIC)
Versión Corporativa Completa:
1. Dashboard Ejecutivo Desplegable (Mes, Proveedor, Clasificación, Año, Moneda)
2. Flujo Ordenado en 5 Pasos:
   Paso 1: Elaborar Requisición (REQ + PED)
   Paso 2: Elaborar Orden de Compra (OC)
   Paso 3: Colocar Precios y Descuentos (SUM)
   Paso 4: Validar y Aprobar Información (APR / REG)
   Paso 5: Imprimir / Descargar PDF Oficial (CON)
3. Gestión Completa y Editable de Proveedores (PRO), Inventario (INV) y Personal (DAT)
4. Recepción de Facturas (ING) y Pendientes (PEN / ACT)
5. Migración Blindada Macro Excel (.xlsm) <-> Nube (Supabase / SQLite)
"""

import io
import os
from datetime import datetime
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from sqlalchemy import create_engine, text

st.set_page_config(
    page_title="Control de Compras — Instacredit S.A.",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilos visuales corporativos
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(90deg, #002b5c 0%, #004b93 100%);
        color: white;
        padding: 16px 22px;
        border-radius: 10px;
        margin-bottom: 16px;
        box-shadow: 0 2px 6px rgba(0,0,0,0.12);
    }
    .main-header h2 { color: white !important; margin: 0; font-size: 22px; }
    .main-header p { color: #d9e6f2; margin: 4px 0 0 0; font-size: 13px; }
    div[data-testid="stMetric"] {
        background-color: #f8fafd;
        border: 1px solid #dce4ec;
        border-left: 5px solid #004b93;
        padding: 12px 16px;
        border-radius: 8px;
    }
</style>
""", unsafe_allow_html=True)

MESES_ES = {
    1: "ENE", 2: "FEB", 3: "MAR", 4: "ABR", 5: "MAY", 6: "JUN",
    7: "JUL", 8: "AGO", 9: "SEP", 10: "OCT", 11: "NOV", 12: "DIC"
}
ORDEN_MESES = ["ENE", "FEB", "MAR", "ABR", "MAY", "JUN", "JUL", "AGO", "SEP", "OCT", "NOV", "DIC"]

TIPOS_GASTO_DEFAULT = [
    "SUMINISTROS", "ACTIVOS", "MERCADEO", "MANTENIMIENTO", "ACTIVIDADES",
    "ALQUILER DE VEHICULO", "ATENCION FUNCIONARIOS", "CAPACITACIONES",
    "DONACIONES", "FLOTILLA", "FUMIGACION", "HERRAMIENTAS", "TI", "PUBLICIDAD"
]

CONDICIONES_DEFAULT = [
    "CONTADO", "8 DIAS CREDITO", "15 DIAS CREDITO",
    "20 DIAS CREDITO", "30 DIAS CREDITO", "AVANCE OBRA"
]

UNIDADES_DEFAULT = ["UND", "GLOB", "SERV", "CJA", "GAL", "KG", "M", "L", "PAQ", "CUB", "IMP", "KM"]
MONEDAS_DEFAULT = ["CORDOBAS", "DOLARES", "QUETZALES", "COLONES"]
ESTATUS_OC_LIST = ["PENDIENTE", "APROBADO", "CERRADA", "NULO"]


def clean_val(val, default=""):
    if val is None or pd.isna(val):
        return default
    s = str(val).strip()
    if s in ("nan", "None", "NaT", "#¡REF!", "#N/A", "Error 2023", "Err:504"):
        return default
    if s.endswith(".0") and s[:-2].isdigit():
        return s[:-2]
    return s


def clean_float(val, default=0.0):
    if val is None or pd.isna(val):
        return default
    try:
        s = str(val).replace(",", "").replace("$", "").replace("C$", "").strip()
        if s in ("", "nan", "None", "NULA", "NULO", "PENDIENTE", "Error 2023"):
            return default
        return float(s)
    except Exception:
        return default


def format_date_str(val):
    if val is None or pd.isna(val):
        return datetime.now().strftime("%d/%m/%Y")
    if isinstance(val, datetime):
        return val.strftime("%d/%m/%Y")
    s = str(val).strip()
    if " " in s:
        s = s.split(" ")[0]
    if "-" in s and len(s) == 10 and s[:4].isdigit():
        p = s.split("-")
        return f"{p[2]}/{p[1]}/{p[0]}"
    return s


# =====================================================================
# CONEXIÓN A BASE DE DATOS (SUPABASE NUBE O SQLITE LOCAL)
# =====================================================================
def get_db_url():
    if "db_url_override" in st.session_state and st.session_state["db_url_override"]:
        return st.session_state["db_url_override"]
    try:
        if "DATABASE_URL" in st.secrets:
            return st.secrets["DATABASE_URL"]
    except Exception:
        pass
    if os.environ.get("DATABASE_URL"):
        return os.environ.get("DATABASE_URL")
    return "sqlite:///instacredit_compras_nic.db"


@st.cache_resource
def get_engine(db_url: str):
    if db_url.startswith("postgres://"):
        db_url = db_url.replace("postgres://", "postgresql+psycopg2://", 1)
    elif db_url.startswith("postgresql://") and "+psycopg2" not in db_url:
        db_url = db_url.replace("postgresql://", "postgresql+psycopg2://", 1)
    return create_engine(db_url, pool_pre_ping=True)


def init_database(engine):
    is_sqlite = engine.url.drivername.startswith("sqlite")
    pk_type = "INTEGER PRIMARY KEY AUTOINCREMENT" if is_sqlite else "SERIAL PRIMARY KEY"

    queries = [
        f"""CREATE TABLE IF NOT EXISTS proveedores (
            id {pk_type}, nombre_comercial TEXT NOT NULL, razon_social TEXT,
            cedula TEXT, direccion TEXT, numero TEXT, contacto TEXT,
            tel_cont TEXT, condicion TEXT DEFAULT 'CONTADO', email TEXT, segmento TEXT
        )""",
        f"""CREATE TABLE IF NOT EXISTS inventario (
            id {pk_type}, descripcion TEXT NOT NULL, medida TEXT DEFAULT 'UND'
        )""",
        f"""CREATE TABLE IF NOT EXISTS solicitantes (
            id {pk_type}, nombre TEXT NOT NULL, departamento TEXT
        )""",
        f"""CREATE TABLE IF NOT EXISTS req (
            id {pk_type}, rq TEXT NOT NULL, solicitante TEXT, departamento TEXT,
            clasificacion_gasto TEXT, fecha_solic TEXT, fecha_req TEXT, detalle TEXT,
            oc_asociada TEXT DEFAULT 'SIN OC', factura TEXT DEFAULT 'SIN FACTURA',
            est TEXT DEFAULT 'SIN FACTURA', estatus TEXT DEFAULT 'ABIERTA'
        )""",
        f"""CREATE TABLE IF NOT EXISTS ped (
            id {pk_type}, rq TEXT NOT NULL, item INTEGER, descripcion TEXT NOT NULL,
            cant REAL DEFAULT 1, und TEXT DEFAULT 'UND'
        )""",
        f"""CREATE TABLE IF NOT EXISTS reg (
            id {pk_type}, rq TEXT, oc TEXT NOT NULL, fecha_oc TEXT, proveedor TEXT,
            clasificacion TEXT, moneda TEXT DEFAULT 'CORDOBAS', subtotal REAL DEFAULT 0.0,
            iva REAL DEFAULT 0.0, descuento REAL DEFAULT 0.0, total REAL DEFAULT 0.0,
            condicion TEXT DEFAULT 'CONTADO', mes TEXT, anio TEXT, destino TEXT,
            descripcion TEXT, comp TEXT DEFAULT 'INSTACREDIT', pais TEXT DEFAULT 'NIC',
            estatus TEXT DEFAULT 'PENDIENTE', solicita TEXT, departamento TEXT,
            revisa TEXT DEFAULT 'JACKSON GARCÍA RAYO',
            cargo_revisa TEXT DEFAULT 'COORDINADOR ADMINISTRATIVO',
            autoriza TEXT DEFAULT 'CECILIA BONILLA',
            cargo_autoriza TEXT DEFAULT 'GERENCIA PAIS'
        )""",
        f"""CREATE TABLE IF NOT EXISTS sum_detalles (
            id {pk_type}, oc TEXT NOT NULL, item INTEGER, articulo TEXT NOT NULL,
            cant REAL DEFAULT 1.0, und TEXT DEFAULT 'UND', moneda TEXT DEFAULT 'CORDOBAS',
            precio_unit REAL DEFAULT 0.0, descuento REAL DEFAULT 0.0, iva REAL DEFAULT 0.0,
            total_unit REAL DEFAULT 0.0, proveedor TEXT, total_linea REAL DEFAULT 0.0, fecha TEXT
        )""",
        f"""CREATE TABLE IF NOT EXISTS ing (
            id {pk_type}, rq TEXT, oc TEXT NOT NULL, proveedor TEXT, clasificacion TEXT,
            moneda TEXT, fecha_ingreso TEXT, factura TEXT, monto REAL DEFAULT 0.0,
            estatus TEXT DEFAULT 'COMPLETO', descripcion TEXT, stat TEXT DEFAULT 'CERRADA',
            monto_oc REAL DEFAULT 0.0
        )"""
    ]

    with engine.begin() as conn:
        for q in queries:
            conn.execute(text(q))

        cnt = conn.execute(text("SELECT COUNT(*) FROM solicitantes")).scalar()
        if cnt == 0:
            sol_base = [
                ("JACKSON GARCÍA RAYO", "ADMINISTRATIVO"),
                ("JULIO SOSA", "GERENTE FINANCIERO"),
                ("JORDY CHAVARRIA", "INFRAESTRUCTURA"),
                ("SHIRLEY MEDINA", "COMERCIAL"),
                ("MARCELL SANABRIA", "INFRAESTRUCTURA"),
                ("DARLING AVALOS", "CALL CENTER NIC"),
                ("EDUARDO FITORIA", "MERCADEO"),
                ("MAYTE RIGUERO", "COMERCIAL"),
                ("SAUL RODRIGUEZ", "MERCADEO"),
                ("MARTIN TRUJILLO", "HIGIENE Y SEGURIDAD"),
                ("LUIS CARLOS SANCHEZ", "ADMINISTRATIVO"),
                ("MAURICIO ULATE", "DEPARTAMENTO TI")
            ]
            for nom, dpto in sol_base:
                conn.execute(
                    text("INSERT INTO solicitantes (nombre, departamento) VALUES (:n, :d)"),
                    {"n": nom, "d": dpto}
                )


def get_next_number(engine, table_name: str, col_name: str, base_min: int) -> str:
    with engine.connect() as conn:
        rows = conn.execute(text(f"SELECT {col_name} FROM {table_name}")).fetchall()
    max_val = base_min
    for r in rows:
        s = str(r[0]).strip()
        if s.isdigit() and int(s) > max_val:
            max_val = int(s)
    return str(max_val + 1)


def generate_printable_html(engine, oc_num: str) -> str:
    with engine.connect() as conn:
        reg_df = pd.read_sql(text("SELECT * FROM reg WHERE oc = :oc"), conn, params={"oc": str(oc_num)})
        items_df = pd.read_sql(text("SELECT * FROM sum_detalles WHERE oc = :oc ORDER BY item"), conn, params={"oc": str(oc_num)})
        prov_df = pd.DataFrame()
        if not reg_df.empty and reg_df.iloc[0]["proveedor"]:
            prov_df = pd.read_sql(
                text("SELECT * FROM proveedores WHERE nombre_comercial = :p"),
                conn,
                params={"p": reg_df.iloc[0]["proveedor"]}
            )

    if reg_df.empty:
        return "<h3>Orden de Compra no encontrada</h3>"

    r = reg_df.iloc[0]
    p = prov_df.iloc[0] if not prov_df.empty else {}

    rows_html = ""
    for _, it in items_df.iterrows():
        rows_html += f"""
        <tr>
            <td style="text-align:center;">{it['item']}</td>
            <td>{it['articulo']}</td>
            <td style="text-align:center;">{clean_float(it['cant']):,.2f}</td>
            <td style="text-align:center;">{it['und']}</td>
            <td style="text-align:right;">{clean_float(it['precio_unit']):,.2f}</td>
            <td style="text-align:right;">{clean_float(it['descuento']):,.2f}</td>
            <td style="text-align:right;">{clean_float(it['iva']):,.2f}</td>
            <td style="text-align:right;font-weight:bold;">{clean_float(it['total_linea']):,.2f}</td>
        </tr>
        """

    return f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<title>OC_{r['oc']}_INSTACREDIT</title>
<style>
    body {{ font-family: 'Segoe UI', Arial, sans-serif; margin: 20px; color: #111; font-size: 12px; background: white; }}
    .top {{ display:flex; justify-content:space-between; border-bottom:3px solid #003366; padding-bottom:10px; }}
    .title {{ font-size:20px; font-weight:bold; color:#003366; }}
    .box {{ border:2px solid #003366; padding:8px 14px; background:#f4f7fb; text-align:right; border-radius:4px; }}
    .sec {{ background:#003366; color:#fff; font-weight:bold; padding:5px 8px; margin-top:12px; }}
    table {{ width:100%; border-collapse:collapse; }}
    td, th {{ border:1px solid #bbb; padding:6px; }}
    th {{ background:#003366; color:#fff; }}
    .tot {{ width:320px; margin-left:auto; margin-top:12px; }}
    .sigs {{ display:flex; justify-content:space-between; margin-top:55px; text-align:center; }}
    .sig {{ width:30%; border-top:1px solid #111; padding-top:6px; }}
    @media print {{ .no-print {{ display:none; }} }}
</style>
</head>
<body>
    <div class="no-print" style="margin-bottom:12px;">
        <button onclick="window.print()" style="padding:8px 16px;font-weight:bold;background:#003366;color:white;border:none;border-radius:4px;cursor:pointer;">
            🖨️ Imprimir / Guardar como PDF
        </button>
    </div>
    <div class="top">
        <div>
            <div class="title">INSTACREDIT S.A.</div>
            <div>Cédula RUC: J0310000115168</div>
            <div>Mansión Teodolinda 3 cuadras al sur, 75 vrs abajo Bolonia, Managua</div>
            <div>Tel: +(505) 22527676 | comprasnic@instacredit.com</div>
        </div>
        <div class="box">
            <div style="font-size:16px;font-weight:bold;color:#003366;">ORDEN DE COMPRA N° {r['oc']}</div>
            <div><strong>REQUISICIÓN N°:</strong> {r['rq'] or 'N/A'}</div>
            <div><strong>FECHA:</strong> {r['fecha_oc']}</div>
            <div><strong>ESTATUS:</strong> {r['estatus']}</div>
        </div>
    </div>

    <div class="sec">INFORMACIÓN DEL PROVEEDOR</div>
    <table>
        <tr>
            <td><strong>NOMBRE COMERCIAL:</strong> {r['proveedor']}</td>
            <td><strong>CÉDULA / RUC:</strong> {p.get('cedula', 'N/A')}</td>
        </tr>
        <tr>
            <td><strong>RAZÓN SOCIAL:</strong> {p.get('razon_social', r['proveedor'])}</td>
            <td><strong>CONTACTO / TEL:</strong> {p.get('contacto', '')} ({p.get('tel_cont', '')})</td>
        </tr>
        <tr>
            <td><strong>DIRECCIÓN:</strong> {p.get('direccion', 'Managua, Nicaragua')}</td>
            <td><strong>EMAIL:</strong> {p.get('email', '')}</td>
        </tr>
    </table>

    <div class="sec">INFORMACIÓN DE LA COMPRA</div>
    <table>
        <tr>
            <td><strong>SOLICITA:</strong> {r.get('solicita') or 'ADMINISTRATIVO'}</td>
            <td><strong>CONDICIÓN:</strong> {r['condicion']}</td>
            <td><strong>MONEDA:</strong> {r['moneda']}</td>
        </tr>
        <tr>
            <td><strong>DEPARTAMENTO:</strong> {r.get('departamento') or 'ADMINISTRATIVO'}</td>
            <td><strong>CLASIFICACIÓN:</strong> {r['clasificacion']}</td>
            <td><strong>DESTINO:</strong> {r['destino'] or 'NIC'}</td>
        </tr>
        <tr>
            <td colspan="3"><strong>CONCEPTO / DESCRIPCIÓN:</strong> {r['descripcion'] or ''}</td>
        </tr>
    </table>

    <table style="margin-top:10px;">
        <thead>
            <tr>
                <th>ITEM</th><th>ARTÍCULO / DESCRIPCIÓN</th><th>CANT</th><th>UND</th>
                <th>PRECIO UNIT</th><th>DESC</th><th>IVA</th><th>TOTAL</th>
            </tr>
        </thead>
        <tbody>{rows_html}</tbody>
    </table>

    <table class="tot">
        <tr><td>SUBTOTAL:</td><td style="text-align:right;">{r['moneda']} {clean_float(r['subtotal']):,.2f}</td></tr>
        <tr><td>DESCUENTO:</td><td style="text-align:right;">{r['moneda']} {clean_float(r['descuento']):,.2f}</td></tr>
        <tr><td>IVA (15%):</td><td style="text-align:right;">{r['moneda']} {clean_float(r['iva']):,.2f}</td></tr>
        <tr style="background:#e8eef5;font-weight:bold;font-size:14px;color:#003366;">
            <td>TOTAL ORDEN:</td><td style="text-align:right;">{r['moneda']} {clean_float(r['total']):,.2f}</td>
        </tr>
    </table>

    <div class="sigs">
        <div class="sig"><strong>SOLICITA: {r.get('solicita') or 'SOLICITANTE'}</strong><br>{r.get('departamento') or 'DEPARTAMENTO'}</div>
        <div class="sig"><strong>REVISA: {r.get('revisa') or 'JACKSON GARCÍA RAYO'}</strong><br>{r.get('cargo_revisa') or 'COORDINADOR ADMINISTRATIVO'}</div>
        <div class="sig"><strong>AUTORIZA: {r.get('autoriza') or 'CECILIA BONILLA'}</strong><br>{r.get('cargo_autoriza') or 'GERENCIA PAIS'}</div>
    </div>
</body>
</html>"""


# Inicializar base de datos con respaldo automático a SQLite si hay error de red
db_url = get_db_url()
try:
    engine = get_engine(db_url)
    init_database(engine)
except Exception as err_db:
    st.sidebar.error(f"Error de red con Nube: {err_db}")
    db_url = "sqlite:///instacredit_compras_nic.db"
    engine = get_engine(db_url)
    init_database(engine)

# Encabezado Corporativo
st.markdown("""
<div class="main-header">
    <h2>🏢 INSTACREDIT S.A. — SISTEMA CORPORATIVO DE COMPRAS Y SUMINISTROS</h2>
    <p>RUC: J0310000115168 | Mansión Teodolinda 3c al sur, 75 vrs abajo Bolonia, Managua | comprasnic@instacredit.com</p>
</div>
""", unsafe_allow_html=True)

# Menú Lateral
st.sidebar.title("🏢 Menú Principal")
modulo = st.sidebar.radio(
    "Selecciona el Módulo:",
    [
        "📊 1. Resumen y Dashboard Dinámico",
        "🛒 2. Proceso de Compras (5 Pasos)",
        "🏢 3. Proveedores y Catálogos (Editable)",
        "🧾 4. Recepción de Facturas y Pendientes",
        "⚙️ 5. Migración Excel y Respaldos"
    ]
)

st.sidebar.divider()
modo_actual = "☁️ Nube (Supabase)" if "postgresql" in db_url else "💻 Local (SQLite)"
st.sidebar.info(f"**Base de Datos:** {modo_actual}")

with st.sidebar.expander("🔗 Conectar a Supabase (Nube)"):
    url_input = st.text_input(
        "URL de Supabase (con ':' entre usuario y clave):",
        value=st.session_state.get("db_url_override", ""),
        type="password"
    )
    c_b1, c_b2 = st.columns(2)
    if c_b1.button("Conectar Nube"):
        st.session_state["db_url_override"] = url_input.strip()
        st.rerun()
    if c_b2.button("Usar Local"):
        st.session_state["db_url_override"] = "sqlite:///instacredit_compras_nic.db"
        st.rerun()


# =====================================================================
# MÓDULO 1: RESUMEN Y DASHBOARD DINÁMICO DESPLEGABLE
# =====================================================================
if modulo == "📊 1. Resumen y Dashboard Dinámico":
    with engine.connect() as conn:
        df_reg = pd.read_sql("SELECT * FROM reg", conn)
        df_req = pd.read_sql("SELECT * FROM req", conn)
        df_ing = pd.read_sql("SELECT * FROM ing", conn)

    if df_reg.empty:
        st.warning("⚠️ Aún no hay datos cargados. Ve al **Módulo 5 (Migración Excel)** para subir tu archivo `.xlsm` en 1 clic.")
    else:
        df_reg["total_num"] = df_reg["total"].apply(clean_float)
        df_reg["subtotal_num"] = df_reg["subtotal"].apply(clean_float)
        df_reg["iva_num"] = df_reg["iva"].apply(clean_float)
        df_reg["anio_clean"] = df_reg["anio"].replace("", "S/A").fillna("S/A")
        df_reg["mes_clean"] = df_reg["mes"].replace("", "S/M").fillna("S/M")
        df_reg["prov_clean"] = df_reg["proveedor"].replace("", "SIN PROVEEDOR").fillna("SIN PROVEEDOR")
        df_reg["clas_clean"] = df_reg["clasificacion"].replace("", "SIN CLASIFICAR").fillna("SIN CLASIFICAR")

        # Barra de Filtros Interactivos
        with st.container(border=True):
            st.markdown("#### 🔍 Filtros Interactivos del Dashboard")
            f1, f2, f3, f4, f5 = st.columns(5)

            anios_list = ["TODOS"] + sorted(df_reg["anio_clean"].unique().tolist(), reverse=True)
            meses_list = ["TODOS"] + [m for m in ORDEN_MESES if m in df_reg["mes_clean"].unique()] + [m for m in df_reg["mes_clean"].unique() if m not in ORDEN_MESES]
            provs_list = ["TODOS"] + sorted(df_reg["prov_clean"].unique().tolist())
            clas_list = ["TODOS"] + sorted(df_reg["clas_clean"].unique().tolist())
            est_list = ["TODOS (Sin Nulos)", "PENDIENTE", "APROBADO", "CERRADA", "NULO", " absolutamente TODOS"]

            sel_anio = f1.selectbox("📆 Año:", anios_list)
            sel_mes = f2.selectbox("📅 Mes:", meses_list)
            sel_prov = f3.selectbox("🏢 Proveedor:", provs_list)
            sel_clas = f4.selectbox("🏷️ Clasificación:", clas_list)
            sel_est = f5.selectbox("📌 Estatus:", est_list)

        df_f = df_reg.copy()
        if sel_est == "TODOS (Sin Nulos)":
            df_f = df_f[~df_f["estatus"].str.upper().isin(["NULO", "NULA"])]
        elif sel_est in ("PENDIENTE", "APROBADO", "CERRADA", "NULO"):
            df_f = df_f[df_f["estatus"].str.upper() == sel_est]

        if sel_anio != "TODOS":
            df_f = df_f[df_f["anio_clean"] == sel_anio]
        if sel_mes != "TODOS":
            df_f = df_f[df_f["mes_clean"] == sel_mes]
        if sel_prov != "TODOS":
            df_f = df_f[df_f["prov_clean"] == sel_prov]
        if sel_clas != "TODOS":
            df_f = df_f[df_f["clas_clean"] == sel_clas]

        # KPIs principales
        tot_nio = df_f[df_f["moneda"].str.upper() == "CORDOBAS"]["total_num"].sum()
        tot_usd = df_f[df_f["moneda"].str.upper() == "DOLARES"]["total_num"].sum()
        tot_gtq = df_f[df_f["moneda"].str.upper() == "QUETZALES"]["total_num"].sum()
        cant_ocs = len(df_f)
        cant_pend = len(df_f[df_f["estatus"].str.upper() == "PENDIENTE"])

        k1, k2, k3, k4, k5 = st.columns(5)
        k1.metric("Compras en CÓRDOBAS", f"C$ {tot_nio:,.2f}")
        k2.metric("Compras en DÓLARES", f"$ {tot_usd:,.2f}")
        k3.metric("Compras en QUETZALES", f"Q {tot_gtq:,.2f}")
        k4.metric("Órdenes Filtradas", f"{cant_ocs:,}")
        k5.metric("Órdenes Pendientes", f"{cant_pend:,}")

        # Gráficos visuales
        col_g1, col_g2 = st.columns(2)
        with col_g1:
            st.markdown("##### 📈 Top 10 Proveedores (según filtro actual)")
            g_prov = df_f.groupby("prov_clean", as_index=False)["total_num"].sum().sort_values("total_num", ascending=False).head(10)
            if not g_prov.empty:
                st.bar_chart(g_prov.set_index("prov_clean")["total_num"])

        with col_g2:
            st.markdown("##### 📊 Gasto por Clasificación (según filtro actual)")
            g_clas = df_f.groupby("clas_clean", as_index=False)["total_num"].sum().sort_values("total_num", ascending=False).head(10)
            if not g_clas.empty:
                st.bar_chart(g_clas.set_index("clas_clean")["total_num"])

        # Secciones desplegables detalladas
        st.markdown("#### 📂 Desglose Desplegable por Mes, Proveedor y Clasificación")

        with st.expander("📅 1. Desplegar Resumen por MES y AÑO", expanded=True):
            res_mes = (
                df_f.groupby(["anio_clean", "mes_clean", "moneda"], as_index=False)
                .agg(Cant_Ordenes=("oc", "count"), Subtotal=("subtotal_num", "sum"), IVA=("iva_num", "sum"), Total=("total_num", "sum"))
                .sort_values(["anio_clean", "Total"], ascending=[False, False])
            )
            st.dataframe(res_mes.rename(columns={"anio_clean": "AÑO", "mes_clean": "MES", "moneda": "MONEDA", "Cant_Ordenes": "N° ÓRDENES"}), use_container_width=True, hide_index=True)

        with st.expander("🏢 2. Desplegar Resumen por PROVEEDOR", expanded=False):
            res_prv = (
                df_f.groupby(["prov_clean", "moneda"], as_index=False)
                .agg(Cant_Ordenes=("oc", "count"), Subtotal=("subtotal_num", "sum"), IVA=("iva_num", "sum"), Total=("total_num", "sum"))
                .sort_values("Total", ascending=False)
            )
            st.dataframe(res_prv.rename(columns={"prov_clean": "PROVEEDOR", "moneda": "MONEDA", "Cant_Ordenes": "N° ÓRDENES"}), use_container_width=True, hide_index=True)

        with st.expander("🏷️ 3. Desplegar Resumen por CLASIFICACIÓN DE GASTO", expanded=False):
            res_cls = (
                df_f.groupby(["clas_clean", "moneda"], as_index=False)
                .agg(Cant_Ordenes=("oc", "count"), Subtotal=("subtotal_num", "sum"), IVA=("iva_num", "sum"), Total=("total_num", "sum"))
                .sort_values("Total", ascending=False)
            )
            st.dataframe(res_cls.rename(columns={"clas_clean": "CLASIFICACIÓN", "moneda": "MONEDA", "Cant_Ordenes": "N° ÓRDENES"}), use_container_width=True, hide_index=True)

        with st.expander("📋 4. Ver Detalle de Órdenes de Compra Filtradas", expanded=False):
            st.dataframe(
                df_f[["rq", "oc", "fecha_oc", "proveedor", "clasificacion", "moneda", "subtotal", "iva", "descuento", "total", "condicion", "mes", "anio", "descripcion", "estatus"]],
                use_container_width=True,
                hide_index=True
            )


# =====================================================================
# MÓDULO 2: PROCESO DE COMPRAS ORDENADO EN 5 PASOS
# =====================================================================
elif modulo == "🛒 2. Proceso de Compras (5 Pasos)":
    st.subheader("🛒 Flujo Operativo de Compras — Paso a Paso")

    with engine.connect() as conn:
        provs_df = pd.read_sql("SELECT * FROM proveedores ORDER BY nombre_comercial", conn)
        sols_df = pd.read_sql("SELECT * FROM solicitantes ORDER BY nombre", conn)
        inv_df = pd.read_sql("SELECT * FROM inventario ORDER BY descripcion", conn)
        rqs_df = pd.read_sql("SELECT * FROM req ORDER BY CAST(rq AS INTEGER) DESC LIMIT 400", conn)

    lista_provs = provs_df["nombre_comercial"].tolist() if not provs_df.empty else ["PROVEEDOR GENERAL"]
    lista_sols = sols_df["nombre"].tolist() if not sols_df.empty else ["JACKSON GARCÍA RAYO"]
    dict_sols = dict(zip(sols_df["nombre"], sols_df["departamento"])) if not sols_df.empty else {}
    lista_arts = inv_df["descripcion"].tolist() if not inv_df.empty else []

    paso = st.radio(
        "Selecciona la etapa del proceso:",
        [
            "1️⃣ Elaborar Requisición",
            "2️⃣ Elaborar Orden de Compra",
            "3️⃣ Colocar Precios y Descuentos",
            "4️⃣ Validar y Aprobar Información",
            "5️⃣ Imprimir / Descargar OC"
        ],
        horizontal=True
    )
    st.divider()

    # -----------------------------------------------------------------
    # PASO 1: ELABORAR REQUISICIÓN (REQ + PED)
    # -----------------------------------------------------------------
    if paso == "1️⃣ Elaborar Requisición":
        st.markdown("#### 1️⃣ Paso 1: Elaborar Solicitud / Requisición Interna (`REQ` y `PED`)")
        prox_rq = get_next_number(engine, "req", "rq", 2814)

        col1, col2, col3, col4 = st.columns(4)
        rq_num = col1.text_input("N° de Requisición (RQ):", value=prox_rq)
        sol_rq = col2.selectbox("Solicitante:", lista_sols)
        depto_rq = col3.text_input("Departamento:", value=dict_sols.get(sol_rq, "ADMINISTRATIVO"))
        clasif_rq = col4.selectbox("Clasificación de Gasto:", TIPOS_GASTO_DEFAULT)

        col5, col6, col7 = st.columns([1, 1, 2])
        f_sol = col5.text_input("Fecha Solicitud:", value=datetime.now().strftime("%d/%m/%Y"))
        f_req = col6.text_input("Fecha Requerida:", value=datetime.now().strftime("%d/%m/%Y"))
        det_rq = col7.text_input("Detalle / Justificación General:")

        if not lista_arts:
            st.caption("💡 Tip: Escribe los artículos requeridos en la tabla inferior.")
        else:
            art_sugerido = st.selectbox("🔍 Buscar artículo en Catálogo INV (opcional para copiar):", ["(Escribir manualmente)"] + lista_arts)
            if art_sugerido != "(Escribir manualmente)":
                st.info(f"Artículo seleccionado: **{art_sugerido}** — cópialo o escríbelo abajo.")

        df_ped_base = pd.DataFrame([
            {"DESCRIPCION": "", "CANT": 1.0, "UND": "UND"}
        ])
        edit_ped = st.data_editor(
            df_ped_base,
            num_rows="dynamic",
            use_container_width=True,
            column_config={
                "DESCRIPCION": st.column_config.TextColumn("DESCRIPCIÓN DEL ARTÍCULO / SERVICIO", required=True, width="large"),
                "CANT": st.column_config.NumberColumn("CANTIDAD", min_value=0.01, default=1.0, format="%.2f"),
                "UND": st.column_config.SelectboxColumn("UNIDAD", options=UNIDADES_DEFAULT, default="UND")
            }
        )

        if st.button("💾 Guardar Requisición y Pasar al Paso 2", type="primary"):
            items_validos = []
            for _, r in edit_ped.iterrows():
                d_txt = clean_val(r.get("DESCRIPCION"))
                if d_txt:
                    items_validos.append({
                        "item": len(items_validos) + 1,
                        "desc": d_txt,
                        "cant": clean_float(r.get("CANT"), 1.0),
                        "und": clean_val(r.get("UND"), "UND")
                    })
            if not items_validos:
                st.error("⚠️ Debes agregar al menos un artículo en la tabla de pedidos.")
            else:
                det_final = det_rq.strip() or items_validos[0]["desc"]
                with engine.begin() as conn:
                    conn.execute(text("DELETE FROM req WHERE rq = :rq"), {"rq": rq_num.strip()})
                    conn.execute(text("""
                        INSERT INTO req (rq, solicitante, departamento, clasificacion_gasto, fecha_solic, fecha_req, detalle, oc_asociada, factura, est, estatus)
                        VALUES (:rq, :sol, :dep, :cla, :fs, :fr, :det, 'SIN OC', 'SIN FACTURA', 'SIN FACTURA', 'ABIERTA')
                    """), {
                        "rq": rq_num.strip(), "sol": sol_rq, "dep": depto_rq,
                        "cla": clasif_rq, "fs": f_sol, "fr": f_req, "det": det_final
                    })
                    conn.execute(text("DELETE FROM ped WHERE rq = :rq"), {"rq": rq_num.strip()})
                    for it in items_validos:
                        conn.execute(text("""
                            INSERT INTO ped (rq, item, descripcion, cant, und)
                            VALUES (:rq, :it, :d, :c, :u)
                        """), {"rq": rq_num.strip(), "it": it["item"], "d": it["desc"], "c": it["cant"], "u": it["und"]})
                st.session_state["rq_activa_proceso"] = rq_num.strip()
                st.success(f"✅ Requisición **RQ N° {rq_num}** guardada como ABIERTA. Ahora selecciona el **Paso 2️⃣ Elaborar Orden de Compra** arriba.")

    # -----------------------------------------------------------------
    # PASO 2: ELABORAR ORDEN DE COMPRA (ENCABEZADO + VINCULACIÓN RQ)
    # -----------------------------------------------------------------
    elif paso == "2️⃣ Elaborar Orden de Compra":
        st.markdown("#### 2️⃣ Paso 2: Elaborar Encabezado de Orden de Compra (`OC`)")
        prox_oc = get_next_number(engine, "reg", "oc", 16573)

        rqs_abiertas = rqs_df[rqs_df["estatus"].str.upper() == "ABIERTA"] if not rqs_df.empty else pd.DataFrame()
        opciones_rq = ["(Sin RQ / Elaborar Orden Directa)"] + [
            f"RQ {r['rq']} — {r['solicitante']} | {r['detalle']}"
            for _, r in rqs_abiertas.iterrows()
        ]

        idx_rq_default = 0
        rq_mem = st.session_state.get("rq_activa_proceso", "")
        if rq_mem:
            for i, opt in enumerate(opciones_rq):
                if opt.startswith(f"RQ {rq_mem} "):
                    idx_rq_default = i
                    break

        sel_rq = st.selectbox("🔗 Selecciona la Requisición (RQ) para convertirla en Orden de Compra:", opciones_rq, index=idx_rq_default)

        rq_sel_num = "0"
        sol_def = "JACKSON GARCÍA RAYO"
        depto_def = "ADMINISTRATIVO"
        clas_def = "SUMINISTROS"
        desc_def = ""

        if sel_rq != "(Sin RQ / Elaborar Orden Directa)":
            rq_sel_num = sel_rq.split("—")[0].replace("RQ", "").strip()
            row_rq = rqs_df[rqs_df["rq"] == rq_sel_num]
            if not row_rq.empty:
                sol_def = row_rq.iloc[0]["solicitante"] or sol_def
                depto_def = row_rq.iloc[0]["departamento"] or depto_def
                clas_def = row_rq.iloc[0]["clasificacion_gasto"] or clas_def
                desc_def = row_rq.iloc[0]["detalle"] or ""

        c1, c2, c3, c4 = st.columns(4)
        oc_num = c1.text_input("ORDEN DE COMPRA N°:", value=st.session_state.get("oc_borrador_num", prox_oc))
        rq_val = c2.text_input("REQUISICIÓN ASOCIADA (RQ):", value=rq_sel_num)
        fecha_oc = c3.text_input("FECHA DE EMISIÓN:", value=datetime.now().strftime("%d/%m/%Y"))
        moneda_oc = c4.selectbox("MONEDA:", MONEDAS_DEFAULT)

        c5, c6, c7, c8 = st.columns(4)
        prov_oc = c5.selectbox("PROVEEDOR:", lista_provs)

        ruc_p, cond_p, contacto_p = "", "CONTADO", ""
        if not provs_df.empty and prov_oc in provs_df["nombre_comercial"].values:
            p_row = provs_df[provs_df["nombre_comercial"] == prov_oc].iloc[0]
            ruc_p = p_row["cedula"] or ""
            cond_p = p_row["condicion"] if p_row["condicion"] in CONDICIONES_DEFAULT else "CONTADO"
            contacto_p = p_row["contacto"] or ""

        c6.text_input("CÉDULA / RUC:", value=ruc_p, disabled=True)
        cond_oc = c7.selectbox("CONDICIÓN DE PAGO:", CONDICIONES_DEFAULT, index=CONDICIONES_DEFAULT.index(cond_p) if cond_p in CONDICIONES_DEFAULT else 0)
        clas_oc = c8.selectbox("CLASIFICACIÓN DE GASTO:", TIPOS_GASTO_DEFAULT, index=TIPOS_GASTO_DEFAULT.index(clas_def) if clas_def in TIPOS_GASTO_DEFAULT else 0)

        c9, c10, c11, c12 = st.columns(4)
        sol_oc = c9.selectbox("SOLICITA:", lista_sols, index=lista_sols.index(sol_def) if sol_def in lista_sols else 0)
        dep_oc = c10.text_input("DEPARTAMENTO:", value=dict_sols.get(sol_oc, depto_def))
        dest_oc = c11.text_input("DESTINO / SUCURSAL:", value="CORPORATIVO")
        rev_oc = c12.selectbox("REVISA:", ["JACKSON GARCÍA RAYO", "JULIO SOSA", "LUIS CARLOS SANCHEZ"])

        desc_oc = st.text_input("CONCEPTO / DESCRIPCIÓN GENERAL DE LA COMPRA:", value=desc_def)

        if st.button("➡️ Confirmar Encabezado y Pasar al Paso 3 (Colocar Precios)", type="primary"):
            st.session_state["oc_borrador"] = {
                "oc": oc_num.strip(),
                "rq": rq_val.strip(),
                "fecha_oc": fecha_oc.strip(),
                "moneda": moneda_oc,
                "proveedor": prov_oc,
                "ruc": ruc_p,
                "condicion": cond_oc,
                "clasificacion": clas_oc,
                "solicita": sol_oc,
                "departamento": dep_oc,
                "destino": dest_oc,
                "revisa": rev_oc,
                "descripcion": desc_oc.strip()
            }
            # Precargar ítems desde PED si viene de una RQ
            if rq_val.strip() and rq_val.strip() != "0":
                with engine.connect() as conn:
                    ped_df = pd.read_sql(text("SELECT descripcion AS \"ARTICULO\", cant AS \"CANT\", und AS \"UND\" FROM ped WHERE rq = :rq ORDER BY item"), conn, params={"rq": rq_val.strip()})
                if not ped_df.empty:
                    ped_df["PRECIO_UNIT"] = 0.0
                    ped_df["DESCUENTO"] = 0.0
                    ped_df["APLICA_IVA"] = True
                    st.session_state["oc_items_borrador"] = ped_df
            st.success(f"✅ Encabezado de **OC N° {oc_num}** listo. Ahora haz clic arriba en **3️⃣ Colocar Precios y Descuentos**.")

    # -----------------------------------------------------------------
    # PASO 3: COLOCAR PRECIOS Y DESCUENTOS (HOJA SUM)
    # -----------------------------------------------------------------
    elif paso == "3️⃣ Colocar Precios y Descuentos":
        st.markdown("#### 3️⃣ Paso 3: Colocar Precios Unitarios, Descuentos e IVA (`SUM`)")

        borrador = st.session_state.get("oc_borrador", None)
        if not borrador:
            st.info("ℹ️ Puedes cargar una OC existente para editar sus precios o configurar primero el Paso 2.")

        col_sel1, col_sel2 = st.columns([1, 2])
        oc_target = col_sel1.text_input("N° de OC a cotizar / editar precios:", value=borrador["oc"] if borrador else get_next_number(engine, "reg", "oc", 16573))
        if col_sel2.button("🔄 Cargar artículos existentes de esta OC (si ya fue guardada)"):
            with engine.connect() as conn:
                ex_reg = pd.read_sql(text("SELECT * FROM reg WHERE oc = :oc"), conn, params={"oc": oc_target.strip()})
                ex_sum = pd.read_sql(text("SELECT articulo AS \"ARTICULO\", cant AS \"CANT\", und AS \"UND\", precio_unit AS \"PRECIO_UNIT\", descuento AS \"DESCUENTO\", iva FROM sum_detalles WHERE oc = :oc ORDER BY item"), conn, params={"oc": oc_target.strip()})
            if not ex_reg.empty:
                r0 = ex_reg.iloc[0]
                st.session_state["oc_borrador"] = {
                    "oc": r0["oc"], "rq": r0["rq"] or "0", "fecha_oc": r0["fecha_oc"] or datetime.now().strftime("%d/%m/%Y"),
                    "moneda": r0["moneda"] or "CORDOBAS", "proveedor": r0["proveedor"] or lista_provs[0],
                    "condicion": r0["condicion"] or "CONTADO", "clasificacion": r0["clasificacion"] or "SUMINISTROS",
                    "solicita": r0.get("solicita") or "JACKSON GARCÍA RAYO", "departamento": r0.get("departamento") or "ADMINISTRATIVO",
                    "destino": r0["destino"] or "CORPORATIVO", "revisa": r0.get("revisa") or "JACKSON GARCÍA RAYO",
                    "descripcion": r0["descripcion"] or ""
                }
            if not ex_sum.empty:
                ex_sum["APLICA_IVA"] = ex_sum["iva"].apply(lambda x: clean_float(x) > 0)
                st.session_state["oc_items_borrador"] = ex_sum[["ARTICULO", "CANT", "UND", "PRECIO_UNIT", "DESCUENTO", "APLICA_IVA"]]

        borrador = st.session_state.get("oc_borrador", {
            "oc": oc_target.strip(), "rq": "0", "fecha_oc": datetime.now().strftime("%d/%m/%Y"),
            "moneda": "CORDOBAS", "proveedor": lista_provs[0], "condicion": "CONTADO",
            "clasificacion": "SUMINISTROS", "solicita": "JACKSON GARCÍA RAYO",
            "departamento": "ADMINISTRATIVO", "destino": "CORPORATIVO",
            "revisa": "JACKSON GARCÍA RAYO", "descripcion": ""
        })

        st.markdown(f"**Orden activa:** `OC N° {borrador['oc']}` | **RQ:** `{borrador['rq']}` | **Proveedor:** `{borrador['proveedor']}` | **Moneda:** `{borrador['moneda']}`")

        df_init_prices = st.session_state.get("oc_items_borrador", pd.DataFrame([
            {"ARTICULO": "", "CANT": 1.0, "UND": "UND", "PRECIO_UNIT": 0.0, "DESCUENTO": 0.0, "APLICA_IVA": True}
        ]))

        edited_prices = st.data_editor(
            df_init_prices,
            num_rows="dynamic",
            use_container_width=True,
            column_config={
                "ARTICULO": st.column_config.TextColumn("ARTÍCULO / DESCRIPCIÓN", required=True, width="large"),
                "CANT": st.column_config.NumberColumn("CANTIDAD", min_value=0.01, default=1.0, format="%.2f"),
                "UND": st.column_config.SelectboxColumn("UND", options=UNIDADES_DEFAULT, default="UND"),
                "PRECIO_UNIT": st.column_config.NumberColumn("PRECIO UNITARIO", min_value=0.0, default=0.0, format="%.4f"),
                "DESCUENTO": st.column_config.NumberColumn("DESCUENTO MONTO", min_value=0.0, default=0.0, format="%.2f"),
                "APLICA_IVA": st.column_config.CheckboxColumn("IVA 15%", default=True)
            }
        )

        # Cálculo automático de líneas
        lineas_calculadas = []
        sub_acum, desc_acum, iva_acum = 0.0, 0.0, 0.0
        for _, r in edited_prices.iterrows():
            art_t = clean_val(r.get("ARTICULO"))
            if not art_t:
                continue
            c_v = clean_float(r.get("CANT"), 1.0)
            pu_v = clean_float(r.get("PRECIO_UNIT"), 0.0)
            de_v = clean_float(r.get("DESCUENTO"), 0.0)
            iv_flag = bool(r.get("APLICA_IVA", True))

            bruto = round(c_v * pu_v, 4)
            base = max(0.0, bruto - de_v)
            iva_m = round(base * 0.15, 4) if iv_flag else 0.0
            tot_l = round(base + iva_m, 4)

            sub_acum += bruto
            desc_acum += de_v
            iva_acum += iva_m
            lineas_calculadas.append({
                "item": len(lineas_calculadas) + 1,
                "articulo": art_t,
                "cant": c_v,
                "und": clean_val(r.get("UND"), "UND"),
                "precio_unit": pu_v,
                "descuento": de_v,
                "iva": iva_m,
                "total_linea": tot_l
            })

        tot_final = round(sub_acum - desc_acum + iva_acum, 2)
        mon_simb = borrador["moneda"]

        p_m1, p_m2, p_m3, p_m4 = st.columns(4)
        p_m1.metric("SUBTOTAL", f"{mon_simb} {sub_acum:,.2f}")
        p_m2.metric("DESCUENTO", f"{mon_simb} {desc_acum:,.2f}")
        p_m3.metric("IVA (15%)", f"{mon_simb} {iva_acum:,.2f}")
        p_m4.metric("TOTAL CALCULADO", f"{mon_simb} {tot_final:,.2f}")

        if st.button("💾 Guardar Precios y Pasar al Paso 4 (Validar Información)", type="primary"):
            if not lineas_calculadas:
                st.error("⚠️ Ingresa al menos un artículo con su descripción y precio.")
            else:
                try:
                    dt_oc = datetime.strptime(borrador["fecha_oc"], "%d/%m/%Y")
                except Exception:
                    dt_oc = datetime.now()
                mes_oc = MESES_ES.get(dt_oc.month, "OCT")
                anio_oc = str(dt_oc.year)
                cargo_rev = "GERENTE FINANCIERO" if "JULIO" in borrador["revisa"].upper() else "COORDINADOR ADMINISTRATIVO"
                desc_final = borrador["descripcion"] or lineas_calculadas[0]["articulo"]

                with engine.begin() as conn:
                    conn.execute(text("DELETE FROM reg WHERE oc = :oc"), {"oc": borrador["oc"]})
                    conn.execute(text("""
                        INSERT INTO reg (
                            rq, oc, fecha_oc, proveedor, clasificacion, moneda,
                            subtotal, iva, descuento, total, condicion, mes, anio,
                            destino, descripcion, comp, pais, estatus,
                            solicita, departamento, revisa, cargo_revisa, autoriza, cargo_autoriza
                        ) VALUES (
                            :rq, :oc, :f_oc, :prov, :clasif, :mon,
                            :sub, :iva, :desc, :tot, :cond, :mes, :anio,
                            :dest, :d_gen, 'INSTACREDIT', 'NIC', 'PENDIENTE',
                            :sol, :depto, :rev, :c_rev, 'CECILIA BONILLA', 'GERENCIA PAIS'
                        )
                    """), {
                        "rq": borrador["rq"], "oc": borrador["oc"], "f_oc": borrador["fecha_oc"],
                        "prov": borrador["proveedor"], "clasif": borrador["clasificacion"], "mon": borrador["moneda"],
                        "sub": round(sub_acum, 2), "iva": round(iva_acum, 2), "desc": round(desc_acum, 2),
                        "tot": tot_final, "cond": borrador["condicion"], "mes": mes_oc, "anio": anio_oc,
                        "dest": borrador["destino"], "d_gen": desc_final, "sol": borrador["solicita"],
                        "depto": borrador["departamento"], "rev": borrador["revisa"], "c_rev": cargo_rev
                    })

                    conn.execute(text("DELETE FROM sum_detalles WHERE oc = :oc"), {"oc": borrador["oc"]})
                    for it in lineas_calculadas:
                        conn.execute(text("""
                            INSERT INTO sum_detalles (
                                oc, item, articulo, cant, und, moneda,
                                precio_unit, descuento, iva, total_unit, proveedor, total_linea, fecha
                            ) VALUES (
                                :oc, :item, :art, :cant, :und, :mon,
                                :pu, :desc, :iva, :pu, :prov, :tlin, :f_oc
                            )
                        """), {
                            "oc": borrador["oc"], "item": it["item"], "art": it["articulo"],
                            "cant": it["cant"], "und": it["und"], "mon": borrador["moneda"],
                            "pu": it["precio_unit"], "desc": it["descuento"], "iva": it["iva"],
                            "prov": borrador["proveedor"], "tlin": it["total_linea"], "f_oc": borrador["fecha_oc"]
                        })

                    if borrador["rq"] and borrador["rq"] != "0":
                        conn.execute(text("UPDATE req SET oc_asociada = :oc WHERE rq = :rq"), {"oc": borrador["oc"], "rq": borrador["rq"]})

                st.session_state["oc_validar_num"] = borrador["oc"]
                st.success(f"✅ Precios guardados en **OC N° {borrador['oc']}**. Ahora pasa al **Paso 4️⃣ Validar y Aprobar Información**.")

    # -----------------------------------------------------------------
    # PASO 4: VALIDAR Y APROBAR INFORMACIÓN (APR / REG)
    # -----------------------------------------------------------------
    elif paso == "4️⃣ Validar y Aprobar Información":
        st.markdown("#### 4️⃣ Paso 4: Validar Datos, Firmas y Aprobar Orden (`APR` / `REG`)")

        with engine.connect() as conn:
            ultimas_ocs = pd.read_sql("SELECT oc, rq, fecha_oc, proveedor, moneda, total, estatus FROM reg ORDER BY CAST(oc AS INTEGER) DESC LIMIT 200", conn)

        oc_def_val = st.session_state.get("oc_validar_num", ultimas_ocs.iloc[0]["oc"] if not ultimas_ocs.empty else "")
        v_col1, v_col2 = st.columns([1, 2])
        oc_val_input = v_col1.text_input("🔎 Ingresa el N° de OC a Validar:", value=oc_def_val)

        if oc_val_input.strip():
            with engine.connect() as conn:
                reg_v = pd.read_sql(text("SELECT * FROM reg WHERE oc = :oc"), conn, params={"oc": oc_val_input.strip()})
                sum_v = pd.read_sql(text("SELECT item AS ITEM, articulo AS ARTICULO, cant AS CANT, und AS UND, precio_unit AS PRECIO_UNIT, descuento AS DESC, iva AS IVA, total_linea AS TOTAL FROM sum_detalles WHERE oc = :oc ORDER BY item"), conn, params={"oc": oc_val_input.strip()})

            if reg_v.empty:
                st.warning(f"No se encontró la OC N° {oc_val_input}.")
            else:
                rv = reg_v.iloc[0]
                with st.container(border=True):
                    st.markdown(f"##### 📋 Resumen de Validación — OC N° {rv['oc']} (RQ: {rv['rq'] or 'N/A'})")
                    vc1, vc2, vc3, vc4 = st.columns(4)
                    nuevo_prov = vc1.selectbox("Proveedor:", lista_provs, index=lista_provs.index(rv["proveedor"]) if rv["proveedor"] in lista_provs else 0)
                    nueva_cond = vc2.selectbox("Condición:", CONDICIONES_DEFAULT, index=CONDICIONES_DEFAULT.index(rv["condicion"]) if rv["condicion"] in CONDICIONES_DEFAULT else 0)
                    nueva_clas = vc3.selectbox("Clasificación:", TIPOS_GASTO_DEFAULT, index=TIPOS_GASTO_DEFAULT.index(rv["clasificacion"]) if rv["clasificacion"] in TIPOS_GASTO_DEFAULT else 0)
                    nuevo_est = vc4.selectbox("Estatus de Aprobación:", ESTATUS_OC_LIST, index=ESTATUS_OC_LIST.index(rv["estatus"]) if rv["estatus"] in ESTATUS_OC_LIST else 0)

                    vc5, vc6, vc7, vc8 = st.columns(4)
                    nuevo_sol = vc5.text_input("Solicita:", value=rv.get("solicita") or "JACKSON GARCÍA RAYO")
                    nuevo_rev = vc6.text_input("Revisa:", value=rv.get("revisa") or "JACKSON GARCÍA RAYO")
                    nuevo_aut = vc7.text_input("Autoriza:", value=rv.get("autoriza") or "CECILIA BONILLA")
                    nueva_desc = vc8.text_input("Descripción:", value=rv["descripcion"] or "")

                    st.dataframe(sum_v, use_container_width=True, hide_index=True)

                    tot_c1, tot_c2, tot_c3 = st.columns(3)
                    tot_c1.info(f"**Subtotal:** {rv['moneda']} {clean_float(rv['subtotal']):,.2f}")
                    tot_c2.info(f"**IVA:** {rv['moneda']} {clean_float(rv['iva']):,.2f}")
                    tot_c3.success(f"**TOTAL VALIDADO:** {rv['moneda']} {clean_float(rv['total']):,.2f}")

                    if st.button("✅ Confirmar Validación y Guardar Cambios", type="primary"):
                        with engine.begin() as conn:
                            conn.execute(text("""
                                UPDATE reg
                                SET proveedor = :p, condicion = :c, clasificacion = :cl, estatus = :es,
                                    solicita = :so, revisa = :re, autoriza = :au, descripcion = :de
                                WHERE oc = :oc
                            """), {
                                "p": nuevo_prov, "c": nueva_cond, "cl": nueva_clas, "es": nuevo_est,
                                "so": nuevo_sol, "re": nuevo_rev, "au": nuevo_aut, "de": nueva_desc,
                                "oc": oc_val_input.strip()
                            })
                        st.session_state["oc_imprimir_num"] = oc_val_input.strip()
                        st.success(f"✅ OC N° {oc_val_input} validada con estatus **{nuevo_est}**. Pasa al **Paso 5️⃣ Imprimir / Descargar OC**.")

    # -----------------------------------------------------------------
    # PASO 5: IMPRIMIR / DESCARGAR PDF OFICIAL (CON / OC)
    # -----------------------------------------------------------------
    elif paso == "5️⃣ Imprimir / Descargar OC":
        st.markdown("#### 5️⃣ Paso 5: Vista Previa e Impresión de Orden de Compra (`OC` / `CON`)")
        oc_imp_default = st.session_state.get("oc_imprimir_num", st.session_state.get("oc_validar_num", "16568"))
        oc_print = st.text_input("🖨️ Ingresa el N° de Orden de Compra (OC) para visualizar e imprimir:", value=oc_imp_default)

        if oc_print.strip():
            html_vista = generate_printable_html(engine, oc_print.strip())
            st.download_button(
                label=f"📥 Descargar Orden de Compra N° {oc_print.strip()} (Formato Oficial Imprimible)",
                data=html_vista,
                file_name=f"OC_{oc_print.strip()}_INSTACREDIT.html",
                mime="text/html",
                type="primary"
            )
            st.markdown("##### Vista Previa del Documento Oficial (Haz clic en el botón azul dentro del documento para imprimir directo):")
            components.html(html_vista, height=780, scrolling=True)


# =====================================================================
# MÓDULO 3: GESTIÓN COMPLETA Y EDITABLE DE PROVEEDORES Y CATÁLOGOS
# =====================================================================
elif modulo == "🏢 3. Proveedores y Catálogos (Editable)":
    st.subheader("🏢 Gestión Completa y Editable de Proveedores, Artículos y Personal")

    tab_pro, tab_inv, tab_dat = st.tabs([
        "🏢 1. Proveedores (Hoja PRO - Editable)",
        "📦 2. Artículos de Inventario (Hoja INV - Editable)",
        "👥 3. Solicitantes y Departamentos (Hoja DAT)"
    ])

    with tab_pro:
        with st.expander("➕ Agregar Nuevo Proveedor al Catálogo", expanded=False):
            with st.form("form_add_prov"):
                p1, p2, p3 = st.columns(3)
                n_com = p1.text_input("Nombre Comercial *")
                r_soc = p2.text_input("Razón Social")
                ced_p = p3.text_input("Cédula / RUC")

                p4, p5, p6, p7 = st.columns(4)
                con_p = p4.text_input("Nombre de Contacto")
                tel_p = p5.text_input("Teléfono Contacto")
                cnd_p = p6.selectbox("Condición de Pago", CONDICIONES_DEFAULT)
                seg_p = p7.text_input("Segmento / Rubro")

                p8, p9 = st.columns(2)
                ema_p = p8.text_input("Correo Electrónico (Email)")
                dir_p = p9.text_input("Dirección Física")

                if st.form_submit_button("💾 Guardar Nuevo Proveedor", type="primary"):
                    if not n_com.strip():
                        st.error("El Nombre Comercial es obligatorio.")
                    else:
                        with engine.begin() as conn:
                            conn.execute(text("""
                                INSERT INTO proveedores (nombre_comercial, razon_social, cedula, direccion, contacto, tel_cont, condicion, email, segmento)
                                VALUES (:n, :r, :c, :d, :co, :t, :cn, :e, :s)
                            """), {
                                "n": n_com.strip(), "r": r_soc.strip() or n_com.strip(), "c": ced_p.strip(),
                                "d": dir_p.strip(), "co": con_p.strip(), "t": tel_p.strip(),
                                "cn": cnd_p, "e": ema_p.strip(), "s": seg_p.strip()
                            })
                        st.success(f"✅ Proveedor '{n_com}' agregado exitosamente.")
                        st.rerun()

        st.markdown("##### ✏️ Tabla Editable de Proveedores (Haz doble clic en cualquier celda para editar, agregar filas al final o borrar)")
        busq_prv = st.text_input("🔎 Filtrar proveedores por nombre, RUC o segmento:")

        with engine.connect() as conn:
            df_pro_edit = pd.read_sql("SELECT id, nombre_comercial, razon_social, cedula, direccion, contacto, tel_cont, condicion, email, segmento FROM proveedores ORDER BY nombre_comercial", conn)

        if busq_prv.strip():
            q_p = busq_prv.strip().lower()
            df_pro_ filtered = df_pro_edit[
                df_pro_edit["nombre_comercial"].str.lower().str.contains(q_p, na=False) |
                df_pro_edit["cedula"].str.lower().str.contains(q_p, na=False) |
                df_pro_edit["segmento"].str.lower().str.contains(q_p, na=False)
            ]
        else:
            df_pro_filtered = df_pro_edit

        edited_pro_table = st.data_editor(
            df_pro_filtered,
            num_rows="dynamic",
            use_container_width=True,
            hide_index=True,
            disabled=["id"],
            column_config={
                "id": st.column_config.NumberColumn("ID", width="small"),
                "nombre_comercial": st.column_config.TextColumn("NOMBRE COMERCIAL", required=True),
                "razon_social": st.column_config.TextColumn("RAZÓN SOCIAL"),
                "cedula": st.column_config.TextColumn("CÉDULA / RUC"),
                "direccion": st.column_config.TextColumn("DIRECCIÓN"),
                "contacto": st.column_config.TextColumn("CONTACTO"),
                "tel_cont": st.column_config.TextColumn("TELÉFONO"),
                "condicion": st.column_config.SelectboxColumn("CONDICIÓN", options=CONDICIONES_DEFAULT),
                "email": st.column_config.TextColumn("EMAIL"),
                "segmento": st.column_config.TextColumn("SEGMENTO")
            }
        )

        if st.button("💾 Guardar Cambios Realizados en la Tabla de Proveedores", type="primary"):
            with engine.begin() as conn:
                if not busq_prv.strip():
                    conn.execute(text("DELETE FROM proveedores"))
                    rows_save = []
                    for _, r in edited_pro_table.iterrows():
                        nc = clean_val(r.get("nombre_comercial"))
                        if nc:
                            rows_save.append({
                                "nombre_comercial": nc,
                                "razon_social": clean_val(r.get("razon_social"), nc),
                                "cedula": clean_val(r.get("cedula")),
                                "direccion": clean_val(r.get("direccion")),
                                "contacto": clean_val(r.get("contacto")),
                                "tel_cont": clean_val(r.get("tel_cont")),
                                "condicion": clean_val(r.get("condicion"), "CONTADO"),
                                "email": clean_val(r.get("email")),
                                "segmento": clean_val(r.get("segmento"))
                            })
                    if rows_save:
                        pd.DataFrame(rows_save).to_sql("proveedores", conn, if_exists="append", index=False)
                else:
                    for _, r in edited_pro_table.iterrows():
                        if pd.notna(r.get("id")):
                            conn.execute(text("""
                                UPDATE proveedores
                                SET nombre_comercial=:nc, razon_social=:rs, cedula=:ce, direccion=:di,
                                    contacto=:co, tel_cont=:te, condicion=:cn, email=:em, segmento=:se
                                WHERE id=:id
                            """), {
                                "nc": clean_val(r.get("nombre_comercial")), "rs": clean_val(r.get("razon_social")),
                                "ce": clean_val(r.get("cedula")), "di": clean_val(r.get("direccion")),
                                "co": clean_val(r.get("contacto")), "te": clean_val(r.get("tel_cont")),
                                "cn": clean_val(r.get("condicion"), "CONTADO"), "em": clean_val(r.get("email")),
                                "se": clean_val(r.get("segmento")), "id": int(r["id"])
                            })
            st.success("✅ Catálogo de proveedores actualizado correctamente.")
            st.rerun()

    with tab_inv:
        st.markdown("##### 📦 Catálogo Editable de Artículos y Servicios (`INV`)")
        with engine.connect() as conn:
            df_inv_e = pd.read_sql("SELECT descripcion, medida FROM inventario ORDER BY descripcion", conn)
        edit_inv = st.data_editor(df_inv_e, num_rows="dynamic", use_container_width=True, hide_index=True)
        if st.button("💾 Guardar Cambios en Inventario (INV)", type="primary"):
            with engine.begin() as conn:
                conn.execute(text("DELETE FROM inventario"))
                r_inv = [{"descripcion": clean_val(r.get("descripcion")), "medida": clean_val(r.get("medida"), "UND")} for _, r in edit_inv.iterrows() if clean_val(r.get("descripcion"))]
                if r_inv:
                    pd.DataFrame(r_inv).to_sql("inventario", conn, if_exists="append", index=False)
            st.success("✅ Inventario actualizado.")
            st.rerun()

    with tab_dat:
        st.markdown("##### 👥 Catálogo Editable de Solicitantes y Departamentos (`DAT`)")
        with engine.connect() as conn:
            df_sol_e = pd.read_sql("SELECT nombre, departamento FROM solicitantes ORDER BY nombre", conn)
        edit_sol = st.data_editor(df_sol_e, num_rows="dynamic", use_container_width=True, hide_index=True)
        if st.button("💾 Guardar Solicitantes y Departamentos", type="primary"):
            with engine.begin() as conn:
                conn.execute(text("DELETE FROM solicitantes"))
                r_sol = [{"nombre": clean_val(r.get("nombre")), "departamento": clean_val(r.get("departamento"), "ADMINISTRATIVO")} for _, r in edit_sol.iterrows() if clean_val(r.get("nombre"))]
                if r_sol:
                    pd.DataFrame(r_sol).to_sql("solicitantes", conn, if_exists="append", index=False)
            st.success("✅ Solicitantes actualizados.")
            st.rerun()


# =====================================================================
# MÓDULO 4: RECEPCIÓN DE FACTURAS Y PENDIENTES (ING / PEN / ACT)
# =====================================================================
elif modulo == "🧾 4. Recepción de Facturas y Pendientes":
    st.subheader("🧾 Control de Facturas Recibidas (`ING`), OC Pendientes (`PEN`) y RQ Abiertas (`ACT`)")

    t_ing, t_pen, t_act = st.tabs([
        "🧾 1. Registrar Factura y Cerrar OC (Hoja ING)",
        "⏳ 2. Órdenes de Compra Pendientes sin Factura (Hoja PEN)",
        "📌 3. Requisiciones Abiertas (Hoja ACT)"
    ])

    with t_ing:
        c1, c2, c3, c4 = st.columns(4)
        oc_f_in = c1.text_input("N° Orden de Compra (OC):")
        fac_in = c2.text_input("N° de Factura Proveedor:")
        fec_in = c3.text_input("Fecha Recepción:", value=datetime.now().strftime("%d/%m/%Y"))
        mon_in = c4.number_input("Monto Facturado:", min_value=0.0, value=0.0, step=100.0)

        if st.button("✅ Registrar Factura y Cerrar Orden de Compra", type="primary"):
            if not oc_f_in.strip() or not fac_in.strip():
                st.error("Ingresa el N° de OC y el N° de Factura.")
            else:
                with engine.begin() as conn:
                    r_oc = conn.execute(text("SELECT rq, proveedor, clasificacion, moneda, total, descripcion FROM reg WHERE oc = :oc"), {"oc": oc_f_in.strip()}).fetchone()
                    rq_v = r_oc[0] if r_oc else ""
                    prov_v = r_oc[1] if r_oc else "PROVEEDOR"
                    cla_v = r_oc[2] if r_oc else "SUMINISTROS"
                    mon_v = r_oc[3] if r_oc else "CORDOBAS"
                    tot_oc = clean_float(r_oc[4], mon_in) if r_oc else mon_in
                    des_v = r_oc[5] if r_oc else ""

                    conn.execute(text("""
                        INSERT INTO ing (rq, oc, proveedor, clasificacion, moneda, fecha_ingreso, factura, monto, estatus, descripcion, stat, monto_oc)
                        VALUES (:rq, :oc, :pr, :cl, :mo, :fe, :fa, :mf, 'COMPLETO', :de, 'CERRADA', :moc)
                    """), {
                        "rq": rq_v, "oc": oc_f_in.strip(), "pr": prov_v, "cl": cla_v, "mo": mon_v,
                        "fe": fec_in, "fa": fac_in.strip(), "mf": mon_in if mon_in > 0 else tot_oc,
                        "de": des_v, "moc": tot_oc
                    })
                    conn.execute(text("UPDATE reg SET estatus = 'CERRADA' WHERE oc = :oc"), {"oc": oc_f_in.strip()})
                    if rq_v:
                        conn.execute(text("UPDATE req SET factura = :fa, est = 'CON FACTURA', estatus = 'CERRADA' WHERE rq = :rq"), {"fa": fac_in.strip(), "rq": rq_v})
                st.success(f"✅ Factura {fac_in} ingresada y OC {oc_f_in} marcada como CERRADA.")

        with engine.connect() as conn:
            st.dataframe(pd.read_sql("SELECT rq, oc, proveedor, moneda, fecha_ingreso, factura, monto, monto_oc, estatus, stat FROM ing ORDER BY id DESC LIMIT 400", conn), use_container_width=True, hide_index=True)

    with t_pen:
        with engine.connect() as conn:
            df_pen_v = pd.read_sql("SELECT rq, oc, fecha_oc, proveedor, clasificacion, moneda, total, condicion, descripcion, estatus FROM reg WHERE UPPER(estatus) = 'PENDIENTE' ORDER BY CAST(oc AS INTEGER) DESC", conn)
        st.metric("Total de Órdenes Pendientes de Factura", f"{len(df_pen_v):,}")
        st.dataframe(df_pen_v, use_container_width=True, hide_index=True)

    with t_act:
        with engine.connect() as conn:
            df_act_v = pd.read_sql("SELECT rq, solicitante, departamento, clasificacion_gasto, fecha_solic, detalle, oc_asociada, estatus FROM req WHERE UPPER(estatus) = 'ABIERTA' ORDER BY CAST(rq AS INTEGER) DESC", conn)
        st.metric("Total de Requisiciones Abiertas", f"{len(df_act_v):,}")
        st.dataframe(df_act_v, use_container_width=True, hide_index=True)


# =====================================================================
# MÓDULO 5: MIGRACIÓN BLINDADA MACRO EXCEL (.XLSM) Y EXPORTACIÓN
# =====================================================================
elif modulo == "⚙️ 5. Migración Excel y Respaldos":
    st.subheader("⚙️ Migración Completa desde Macro Excel (`.xlsm`) y Descarga de Respaldo (`.xlsx`)")

    archivo_xlsm = st.file_uploader(
        "1️⃣ Sube aquí tu archivo M01_CONTROL DE COMPRAS NI ACTUALIZADO.xlsm (o PRUEBA.xlsm):",
        type=["xlsm", "xlsx"]
    )

    if archivo_xlsm is not None and st.button("🚀 Ejecutar Migración Completa de las 8 Tablas (Sin Errores de Duplicados)", type="primary"):
        with st.spinner("Importando PRO, INV, DAT, REG, REQ, SUM, PED e ING y reparando 'Error 2023'..."):
            xls = pd.ExcelFile(archivo_xlsm, engine="openpyxl")
            reg_lookup = {}
            resumen_migracion = []

            # Reiniciar tablas limpiamente para evitar cualquier bloqueo de UNIQUE antiguo
            with engine.begin() as conn:
                for tbl in ["proveedores", "inventario", "reg", "req", "sum_detalles", "ped", "ing"]:
                    conn.execute(text(f"DROP TABLE IF EXISTS {tbl}"))
            init_database(engine)

            with engine.begin() as conn:
                # 1. PRO
                if "PRO" in xls.sheet_names:
                    df_p = pd.read_excel(xls, "PRO", header=1)
                    rows_p, vistos_p = [], set()
                    for _, r in df_p.iterrows():
                        nom = clean_val(r.get("NOMBRE COMERCIAL"))
                        clave = nom.upper().strip()
                        if nom and clave not in vistos_p:
                            vistos_p.add(clave)
                            rows_p.append({
                                "nombre_comercial": nom,
                                "razon_social": clean_val(r.get("RAZON SOCIAL"), nom),
                                "cedula": clean_val(r.get("CEDULA")),
                                "direccion": clean_val(r.get("DIRECCION")),
                                "numero": clean_val(r.get("NUMERO")),
                                "contacto": clean_val(r.get("CONTACT")),
                                "tel_cont": clean_val(r.get("TEL CONT")),
                                "condicion": clean_val(r.get("CONDICION"), "CONTADO"),
                                "email": clean_val(r.get("EMAIL")),
                                "segmento": clean_val(r.get("SEGMENTO"))
                            })
                    if rows_p:
                        pd.DataFrame(rows_p).to_sql("proveedores", conn, if_exists="append", index=False, chunksize=400)
                    resumen_migracion.append(f"Proveedores (PRO): {len(rows_p)}")

                # 2. INV
                if "INV" in xls.sheet_names:
                    df_i = pd.read_excel(xls, "INV", header=0)
                    rows_i, vistos_i = [], set()
                    for _, r in df_i.iterrows():
                        desc = clean_val(r.get("DESCRIPCION"))
                        if desc and desc.upper() not in vistos_i:
                            vistos_i.add(desc.upper())
                            rows_i.append({"descripcion": desc, "medida": clean_val(r.get("MEDIDA"), "UND")})
                    if rows_i:
                        pd.DataFrame(rows_i).to_sql("inventario", conn, if_exists="append", index=False, chunksize=400)
                    resumen_migracion.append(f"Inventario (INV): {len(rows_i)}")

                # 3. REG
                if "REG" in xls.sheet_names:
                    df_r = pd.read_excel(xls, "REG", header=2)
                    rows_r, vistos_oc = [], set()
                    for _, r in df_r.iterrows():
                        oc_v = clean_val(r.get("OC"))
                        if not oc_v or oc_v in vistos_oc:
                            continue
                        vistos_oc.add(oc_v)
                        f_oc = format_date_str(r.get("FECHA OC"))
                        prov = clean_val(r.get("PROVEEDOR"), "NULA")
                        cla = clean_val(r.get("CLASIFICACION"))
                        mon = clean_val(r.get("MONEDA"), "CORDOBAS")
                        tot = clean_float(r.get("TOTAL"))
                        d_gen = clean_val(r.get("DESCRIPCION"))
                        reg_lookup[oc_v] = {"moneda": mon, "proveedor": prov, "fecha": f_oc, "total": tot, "clasificacion": cla, "descripcion": d_gen}

                        rows_r.append({
                            "rq": clean_val(r.get("RQ")), "oc": oc_v, "fecha_oc": f_oc, "proveedor": prov,
                            "clasificacion": cla, "moneda": mon, "subtotal": clean_float(r.get("SUBTOTAL")),
                            "iva": clean_float(r.get("IVA")), "descuento": clean_float(r.get("DESCUENTO")),
                            "total": tot, "condicion": clean_val(r.get("CONDICION"), "CONTADO"),
                            "mes": clean_val(r.get("MES")), "anio": clean_val(r.get("AÑO")),
                            "destino": clean_val(r.get("DESTINO")), "descripcion": d_gen,
                            "comp": clean_val(r.get("COMP"), "INSTACREDIT"), "pais": clean_val(r.get("PAIS"), "NIC"),
                            "estatus": clean_val(r.get("ESTATUS"), "PENDIENTE")
                        })
                    if rows_r:
                        pd.DataFrame(rows_r).to_sql("reg", conn, if_exists="append", index=False, chunksize=400)
                    resumen_migracion.append(f"Órdenes (REG): {len(rows_r)}")

                # 4. REQ
                if "REQ" in xls.sheet_names:
                    df_rq = pd.read_excel(xls, "REQ", header=2)
                    rows_rq, vistos_rq = [], set()
                    for _, r in df_rq.iterrows():
                        rq_v = clean_val(r.get("RQ"))
                        if not rq_v or rq_v in vistos_rq:
                            continue
                        vistos_rq.add(rq_v)
                        rows_rq.append({
                            "rq": rq_v, "solicitante": clean_val(r.get("SOLICITANTE")),
                            "departamento": clean_val(r.get("DEPARTAMENTO")),
                            "clasificacion_gasto": clean_val(r.get("CLASIFICACION GASTO")),
                            "fecha_solic": format_date_str(r.get("FECHA SOLIC")),
                            "fecha_req": format_date_str(r.get("FECHA REQ")),
                            "detalle": clean_val(r.get("DETALLE")),
                            "oc_asociada": clean_val(r.get("OC ASOCIADA"), "SIN OC"),
                            "factura": clean_val(r.get("FACTURA"), "SIN FACTURA"),
                            "est": clean_val(r.get("EST"), "SIN FACTURA"),
                            "estatus": clean_val(r.get("ESTATUS"), "ABIERTA")
                        })
                    if rows_rq:
                        pd.DataFrame(rows_rq).to_sql("req", conn, if_exists="append", index=False, chunksize=400)
                    resumen_migracion.append(f"Requisiciones (REQ): {len(rows_rq)}")

                # 5. SUM
                if "SUM" in xls.sheet_names:
                    df_s = pd.read_excel(xls, "SUM", header=1)
                    rows_s = []
                    for _, r in df_s.iterrows():
                        oc_v = clean_val(r.get("OC"))
                        art = clean_val(r.get("ARTICULO"))
                        if not oc_v or not art:
                            continue
                        inf = reg_lookup.get(oc_v, {})
                        cant = clean_float(r.get("CANT"), 1.0)
                        pu = clean_float(r.get("PRECIO UNIT"), 0.0)
                        tot_u = clean_float(r.get("TOTAL"), pu)
                        rows_s.append({
                            "oc": oc_v, "item": int(clean_float(r.get("ITEM"), 1)), "articulo": art,
                            "cant": cant, "und": clean_val(r.get("UND"), "UND"),
                            "moneda": clean_val(r.get("MONEDA")) or inf.get("moneda", "CORDOBAS"),
                            "precio_unit": pu, "descuento": clean_float(r.get("DESC"), 0.0),
                            "iva": clean_float(r.get("IVA"), 0.0), "total_unit": tot_u,
                            "proveedor": clean_val(r.get("PROVEEDOR")) or inf.get("proveedor", ""),
                            "total_linea": clean_float(r.get("Columna1"), cant * tot_u),
                            "fecha": clean_val(r.get("FECHA")) or inf.get("fecha", "")
                        })
                    if rows_s:
                        pd.DataFrame(rows_s).to_sql("sum_detalles", conn, if_exists="append", index=False, chunksize=600)
                    resumen_migracion.append(f"Detalle Ítems OC (SUM): {len(rows_s)}")

                # 6. PED
                if "PED" in xls.sheet_names:
                    df_pd = pd.read_excel(xls, "PED", header=23)
                    rows_pd = []
                    for _, r in df_pd.iterrows():
                        rq_v = clean_val(r.get("RQ"))
                        desc = clean_val(r.get("DESCRIPCION"))
                        if not rq_v or not desc or rq_v == "RQ":
                            continue
                        rows_pd.append({
                            "rq": rq_v, "item": int(clean_float(r.get("ITEM"), 1)),
                            "descripcion": desc, "cant": clean_float(r.get("CANT"), 1.0),
                            "und": clean_val(r.get("UND"), "UND")
                        })
                    if rows_pd:
                        pd.DataFrame(rows_pd).to_sql("ped", conn, if_exists="append", index=False, chunksize=600)
                    resumen_migracion.append(f"Pedidos RQ (PED): {len(rows_pd)}")

                # 7. ING
                if "ING" in xls.sheet_names:
                    df_ig = pd.read_excel(xls, "ING", header=1)
                    rows_ig = []
                    for _, r in df_ig.iterrows():
                        oc_v = clean_val(r.get("OC"))
                        if not oc_v:
                            continue
                        inf = reg_lookup.get(oc_v, {})
                        rows_ig.append({
                            "rq": clean_val(r.get("RQ")), "oc": oc_v,
                            "proveedor": inf.get("proveedor", clean_val(r.get("PROVEEDOR"), "PENDIENTE")),
                            "clasificacion": inf.get("clasificacion", clean_val(r.get("CLASIFICACION"), "PENDIENTE")),
                            "moneda": inf.get("moneda", clean_val(r.get("MONEDA"), "CORDOBAS")),
                            "fecha_ingreso": format_date_str(r.get("INGRESO")),
                            "factura": clean_val(r.get("FACTURA")),
                            "monto": clean_float(r.get("MONTO")),
                            "estatus": clean_val(r.get("ESTATUS"), "COMPLETO"),
                            "descripcion": inf.get("descripcion", clean_val(r.get("DESCRIPCION"))),
                            "stat": clean_val(r.get("STAT"), "CERRADA"),
                            "monto_oc": clean_float(r.get("MONTO OC"), inf.get("total", 0.0))
                        })
                    if rows_ig:
                        pd.DataFrame(rows_ig).to_sql("ing", conn, if_exists="append", index=False, chunksize=400)
                    resumen_migracion.append(f"Facturas (ING): {len(rows_ig)}")

            st.success("🎉 ¡Migración 100% completada sin errores! -> " + " | ".join(resumen_migracion))

    st.divider()
    st.markdown("#### 2️⃣ Exportar toda la Base de Datos a Excel (`.xlsx`)")
    if st.button("📥 Generar Archivo Excel (.xlsx) con Todas las Hojas"):
        buffer = io.BytesIO()
        with engine.connect() as conn:
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                pd.read_sql("SELECT * FROM reg", conn).to_excel(writer, sheet_name="REG", index=False)
                pd.read_sql("SELECT * FROM req", conn).to_excel(writer, sheet_name="REQ", index=False)
                pd.read_sql("SELECT * FROM sum_detalles", conn).to_excel(writer, sheet_name="SUM", index=False)
                pd.read_sql("SELECT * FROM ped", conn).to_excel(writer, sheet_name="PED", index=False)
                pd.read_sql("SELECT * FROM ing", conn).to_excel(writer, sheet_name="ING", index=False)
                pd.read_sql("SELECT * FROM proveedores", conn).to_excel(writer, sheet_name="PRO", index=False)
                pd.read_sql("SELECT * FROM inventario", conn).to_excel(writer, sheet_name="INV", index=False)
        st.download_button(
            label="⬇️ Descargar Libro Excel Completo (.xlsx)",
            data=buffer.getvalue(),
            file_name=f"CONTROL_COMPRAS_INSTACREDIT_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )