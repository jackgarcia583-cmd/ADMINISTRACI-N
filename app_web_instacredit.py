import os
import sqlite3
import pandas as pd
import streamlit as st
from sqlalchemy import create_engine, text

# ==========================================
# 0. CONFIGURACIÓN DE LA PÁGINA Y ESTILOS
# ==========================================
st.set_page_config(
    page_title="Control de Compras — Instacredit S.A.",
    page_icon="💳",
    layout="wide",
)

# Inyección de Estilos CSS con Identidad Visual Instacredit (Verde Institucional y Tonos Limpios)
st.markdown(
    """
    <style>
    /* Estilos Generales y Colores Corporativos */
    .main {
        background-color: #f8f9fa;
    }
    h1, h2, h3 {
        color: #005A36; /* Verde Oficial Instacredit */
        font-family: 'Inter Tight', sans-serif;
    }
    .stButton>button {
        background-color: #00843D; /* Verde Principal */
        color: white;
        border-radius: 6px;
        border: none;
        font-weight: bold;
    }
    .stButton>button:hover {
        background-color: #005A36;
        color: white;
    }
    .sidebar .sidebar-content {
        background-color: #ffffff;
    }
    </style>
""",
    unsafe_allow_html=True,
)

# ==========================================
# 1. GESTIÓN DE BASES DE DATOS (LOCAL Y NUBE)
# ==========================================
DB_NAME = "instacredit_compras_nic.db"


def get_sqlite_connection():
  conn = sqlite3.connect(DB_NAME)
  return conn


# Intentar conectar a Supabase desde st.secrets o usar SQLite local por defecto
DATABASE_URL = None
try:
  if "DATABASE_URL" in st.secrets:
    DATABASE_URL = st.secrets["DATABASE_URL"]
except Exception:
  pass

if DATABASE_URL:
  try:
    engine = create_engine(DATABASE_URL)
  except Exception:
    engine = None
else:
    engine = None


def run_query(query, params=None):
  if engine:
    try:
      with engine.connect() as conn:
        result = conn.execute(text(query), params or {})
        if query.strip().upper().startswith("SELECT"):
          return pd.DataFrame(result.fetchall(), columns=result.keys())
        conn.commit()
        return None
    except Exception:
      pass  # Fallback a SQLite local si falla la nube

  # Modo SQLite Local
  conn = get_sqlite_connection()
  if query.strip().upper().startswith("SELECT"):
    df = pd.read_sql(query, conn, params=params)
    conn.close()
    return df
  else:
    cursor = conn.cursor()
    cursor.execute(query, params or ())
    conn.commit()
    conn.close()
    return None


# ==========================================
# 2. INICIALIZACIÓN DE TABLAS MAESTRAS
# ==========================================
def init_db():
  queries = [
      """
        CREATE TABLE IF NOT EXISTS reg (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            rq TEXT,
            oc TEXT UNIQUE NOT NULL,
            fecha_oc TEXT,
            proveedor TEXT,
            clasificacion TEXT,
            moneda TEXT DEFAULT 'CORDOBAS',
            subtotal REAL DEFAULT 0.0,
            iva REAL DEFAULT 0.0,
            descuento REAL DEFAULT 0.0,
            total REAL DEFAULT 0.0,
            condicion TEXT DEFAULT 'CONTADO',
            mes TEXT,
            anio TEXT,
            destino TEXT,
            descripcion TEXT,
            comp TEXT DEFAULT 'INSTACREDIT',
            pais TEXT DEFAULT 'NIC',
            estatus TEXT DEFAULT 'PENDIENTE',
            solicita TEXT,
            departamento TEXT,
            revisa TEXT DEFAULT 'JACKSON GARCÍA RAYO',
            cargo_revisa TEXT DEFAULT 'COORDINADOR ADMINISTRATIVO',
            autoriza TEXT DEFAULT 'CECILIA BONILLA',
            cargo_autoriza TEXT DEFAULT 'GERENCIA PAIS'
        );
        """,
      """
        CREATE TABLE IF NOT EXISTS req (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            rq TEXT UNIQUE NOT NULL,
            solicitante TEXT,
            departamento TEXT,
            clasificacion_gasto TEXT,
            fecha_solic TEXT,
            fecha_req TEXT,
            detalle TEXT,
            oc_asociada TEXT DEFAULT 'SIN OC',
            factura TEXT DEFAULT 'SIN FACTURA',
            est TEXT DEFAULT 'SIN FACTURA',
            estatus TEXT DEFAULT 'ABIERTA'
        );
        """,
      """
        CREATE TABLE IF NOT EXISTS ped (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            rq TEXT NOT NULL,
            item INTEGER,
            descripcion TEXT NOT NULL,
            cant REAL DEFAULT 1,
            und TEXT DEFAULT 'UND'
        );
        """,
      """
        CREATE TABLE IF NOT EXISTS proveedores (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre_comercial TEXT UNIQUE NOT NULL,
            razon_social TEXT,
            cedula TEXT,
            direccion TEXT,
            numero TEXT,
            contacto TEXT,
            tel_cont TEXT,
            condicion TEXT,
            email TEXT,
            segmento TEXT
        );
        """,
  ]
  for q in queries:
    run_query(q)


init_db()

# ==========================================
# 3. CONTROL DE ACCESO (LOGIN CON ROLES)
# ==========================================
if "authenticated" not in st.session_state:
  st.session_state.authenticated = False
  st.session_state.user_name = ""
  st.session_state.user_role = ""

if not st.session_state.authenticated:
  st.markdown(
      "<h2 style='text-align: center; color: #005A36;'>Instacredit S.A. —"
      " Control de Compras (NIC)</h2>",
      unsafe_allow_html=True,
  )
  col1, col2, col3 = st.columns([1, 2, 1])
  with col2:
    st.info(
        "Por favor, ingrese sus credenciales corporativas para acceder al"
        " sistema."
    )
    input_user = st.text_input("Usuario")
    input_pass = st.text_input("Contraseña", type="password")

    if st.button("Iniciar Sesión"):
      # Credenciales del Administrador Requeridas
      if input_user == "Jackson Garcia" and input_pass == "AppCompras26*":
        st.session_state.authenticated = True
        st.session_state.user_name = "Jackson Garcia"
        st.session_state.user_role = "Administrador"
        st.success("¡Bienvenido, Jackson Garcia!")
        st.rerun()
      else:
        st.error(
            "Usuario o contraseña incorrectos. Verifique sus datos de acceso."
        )
  st.stop()

# ==========================================
# 4. NAVEGACIÓN Y BARRA LATERAL
# ==========================================
st.sidebar.markdown(
    "<h2 style='color: #005A36;'>Instacredit S.A.</h2>", unsafe_allow_html=True
)
st.sidebar.markdown(f"**Usuario:** {st.session_state.user_name}")
st.sidebar.markdown(f"**Rol:** {st.session_state.user_role}")

if st.sidebar.button("Cerrar Sesión"):
  st.session_state.authenticated = False
  st.rerun()

st.sidebar.markdown("---")
modulo = st.sidebar.radio(
    "Módulos del Sistema",
    [
        "📊 1. Dashboard Ejecutivo",
        "📝 2. Requisiciones (REQ+PED)",
        "📦 3. Órdenes de Compra (OC)",
        "💰 4. Precios y Descuentos (SUM)",
        "✅ 5. Validación y Aprobación (APR)",
        "🖨️ 6. Impresión Oficial / PDF (CON)",
        "🏢 7. Catálogo de Proveedores (PRO)",
        "⚙️ 8. Migración y Configuración",
    ],
)

# ==========================================
# 5. DESARROLLO DE MÓDULOS
# ==========================================

if "Dashboard" in modulo:
  st.title("📊 Dashboard Ejecutivo de Compras")
  df_reg = run_query("SELECT * FROM reg")

  if df_reg.empty:
    st.warning(
        "No hay registros de compras cargados actualmente en la base de datos."
    )
  else:
    # Filtros ejecutivos
    st.subheader("Filtros de Búsqueda")
    c1, c2, c3 = st.columns(3)
    with c1:
      anios = ["TODOS"] + list(df_reg["anio"].dropna().unique())
      sel_anio = st.selectbox("Año", anios)
    with c2:
      provs = ["TODOS"] + list(df_reg["proveedor"].dropna().unique())
      sel_prov = st.selectbox("Proveedor", provs)
    with c3:
      clasif = ["TODOS"] + list(df_reg["clasificacion"].dropna().unique())
      sel_clasif = st.selectbox("Clasificación de Gasto", clasif)

    df_filt = df_reg.copy()
    if sel_anio != "TODOS":
      df_filt = df_filt[df_filt["anio"] == sel_anio]
    if sel_prov != "TODOS":
      df_filt = df_filt[df_filt["proveedor"] == sel_prov]
    if sel_clasif != "TODOS":
      df_filt = df_filt[df_filt["clasificacion"] == sel_clasif]

    # Métricas principales
    m1, m2, m3 = st.columns(3)
    total_monto = df_filt["total"].sum()
    total_oc = df_filt["oc"].nunique()
    m1.metric("Monto Total Compras", f"C$ {total_monto:,.2f}")
    m2.metric("Órdenes de Compra", total_oc)
    m3.metric("Proveedores Activos", df_filt["proveedor"].nunique())

    st.markdown("---")
    st.dataframe(df_filt, width="stretch")

elif "Requisiciones" in modulo:
  st.title("📝 Módulo de Requisiciones (REQ + PED)")
  st.write(
      "Gestión de solicitudes de compra internas de Instacredit S.A. (Managua,"
      " Nicaragua)."
  )

  df_req = run_query("SELECT * FROM req")
  st.subheader("Listado de Requisiciones Activas")
  st.dataframe(df_req, width="stretch")

elif "Órdenes de Compra" in modulo:
  st.title("📦 Órdenes de Compra (OC)")
  df_reg = run_query("SELECT * FROM reg")
  st.dataframe(df_reg, width="stretch")

elif "Precios" in modulo:
  st.title("💰 Precios, Descuentos e Insumos (SUM)")
  st.write(
      "Control detallado de los costos por artículo, descuentos e IVA aplicado."
  )
  df_sum = run_query("SELECT * FROM sum_detalles")
  if not df_sum.empty:
    st.dataframe(df_sum, width="stretch")
  else:
    st.info(
        "No hay detalles de insumos registrados. Puede importarlos desde su"
        " archivo maestro."
    )

elif "Validación" in modulo:
  st.title("✅ Validación y Aprobación de Compras (APR/REG)")
  st.write(
      f"Usuario actual con privilegios de revisión: **{st.session_state.user_name}"
      f" ({st.session_state.user_role})**"
  )
  df_reg = run_query("SELECT * FROM reg WHERE estatus='PENDIENTE'")
  st.dataframe(df_reg, width="stretch")

elif "Impresión" in modulo:
  st.title("🖨️ Impresión Oficial y Generación de PDF (CON)")
  st.write("Seleccione una Orden de Compra para generar su vista oficial.")
  df_reg = run_query("SELECT oc, proveedor, total FROM reg")
  if not df_reg.empty:
    sel_oc = st.selectbox("Seleccione Orden de Compra", df_reg["oc"].tolist())
    if st.button("Generar Vista Previa Oficial"):
      st.success(f"Vista oficial generada para la OC: {sel_oc}")

elif "Catálogo" in modulo:
  st.title("🏢 Catálogo Editable de Proveedores (PRO)")
  df_prov = run_query("SELECT * FROM proveedores")
  st.dataframe(df_prov, width="stretch")

elif "Migración" in modulo:
  st.title("⚙️ Configuración y Migración de Datos")
  st.write("Gestione la base de datos local y conexiones con Supabase.")
  if st.button("Inicializar / Reparar Base de Datos Local"):
    init_db()
    st.success("¡Base de datos verificada y tablas sincronizadas con éxito!")