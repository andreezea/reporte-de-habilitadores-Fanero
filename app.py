"""
Dashboard Gerencial de Ventas - Streamlit App
================================================
Aplicación interactiva para analizar el desempeño de habilitadores de venta
(PDV Plus, Captura, Mercados, Ferias, Activaciones, Desarrolladores) por
producto y ubicación geográfica.

Estructura del archivo:
    1. Configuración y constantes
    2. Funciones de datos (carga / generación / cálculo / publicación)
    3. Funciones de presentación (KPIs, formato semáforo)
    4. Panel de administrador (acceso restringido)
    5. Interfaz principal (main)

Acceso de administrador:
    El panel de carga de datos está oculto para el público general. Solo es
    visible al abrir la app con el parámetro ?admin=1 en la URL, por ejemplo:
        https://tu-app.streamlit.app/?admin=1
    Las credenciales se leen de st.secrets (ver sección "Configuración de
    credenciales" más abajo). Sin ese parámetro, la app se ve exactamente
    igual para gerencia, sin ningún rastro del control de acceso.

Listo para desplegar en Streamlit Cloud: `streamlit run app.py`
"""

import calendar
import os
from datetime import datetime

import numpy as np
import pandas as pd
import streamlit as st

# =============================================================================
# 1. CONFIGURACIÓN Y CONSTANTES
# =============================================================================

st.set_page_config(
    page_title="Dashboard Gerencial de Ventas",
    page_icon="📊",
    layout="wide",
    # Sidebar oculta por defecto: ahí vive el panel de administrador, y así
    # nadie que reciba el enlace del dashboard nota que existe.
    initial_sidebar_state="collapsed",
)

# Ruta donde se guarda el último archivo publicado por el administrador.
# Persiste entre sesiones/usuarios mientras la app siga corriendo.
DATA_DIR = "data"
DATA_FILE = os.path.join(DATA_DIR, "ultima_carga.xlsx")

# Habilitadores disponibles para el filtro principal (selectbox)
HABILITADORES = [
    "PDV Plus",
    "Captura",
    "Mercados",
    "Ferias",
    "Activaciones",
    "Desarrolladores",
]

# Departamentos permitidos en el análisis
DEPARTAMENTOS = [
    "Amazonas",
    "Cajamarca",
    "Huancavelica",
    "Huánuco",
    "Junín",
    "Loreto",
    "Pasco",
    "San Martín",
    "Ucayali",
]

# Productos analizados
PRODUCTOS = ["Prepago", "Porta Flex", "Postpago", "OSS"]

# Distritos de referencia por departamento (usados para el dataset de ejemplo)
DISTRITOS_POR_DEPARTAMENTO = {
    "Amazonas": ["Chachapoyas", "Bagua", "Condorcanqui"],
    "Cajamarca": ["Cajamarca", "Jaén", "Cutervo"],
    "Huancavelica": ["Huancavelica", "Tayacaja", "Acobamba"],
    "Huánuco": ["Huánuco", "Leoncio Prado", "Ambo"],
    "Junín": ["Huancayo", "Tarma", "Chanchamayo"],
    "Loreto": ["Maynas", "Alto Amazonas", "Requena"],
    "Pasco": ["Pasco", "Oxapampa", "Daniel Alcides Carrión"],
    "San Martín": ["Moyobamba", "Tarapoto", "Rioja"],
    "Ucayali": ["Coronel Portillo", "Padre Abad", "Atalaya"],
}

# Nombres de referencia para el dataset sintético
NOMBRES_EJEMPLO = [
    "Carlos Ramírez", "María Torres", "Jorge Quispe", "Ana Flores",
    "Luis Mamani", "Rosa Huamán", "Pedro Vargas", "Karen Chávez",
    "Miguel Salazar", "Diana Rojas", "José Cárdenas", "Lucía Paredes",
    "Fernando Ríos", "Patricia Gómez", "Andrés Castillo", "Silvia Cruz",
    "Raúl Medina", "Carmen Delgado", "Víctor Herrera", "Elena Campos",
    "Manuel Reyes", "Gabriela Nuñez", "Sergio Fernández", "Cecilia Alvarado",
    "Diego Espinoza", "Verónica Guerrero", "Alberto Sánchez", "Milagros Ponce",
    "Ricardo Aguilar", "Susana Bravo",
]

COLUMNAS_REQUERIDAS = {
    "DNI", "Nombre", "Departamento", "Distrito", "Habilitador", "Producto",
    "Cuota", "Avance",
}


# =============================================================================
# 2. FUNCIONES DE DATOS
# =============================================================================

@st.cache_data
def generar_datos_ejemplo(n_gestores: int = 150, seed: int = 42) -> pd.DataFrame:
    """Genera un dataset sintético de ventas, con 4 filas (una por producto)
    por cada gestor, para poder probar el dashboard sin un archivo real."""
    rng = np.random.default_rng(seed)
    registros = []

    for i in range(n_gestores):
        dni = str(40000000 + int(rng.integers(0, 9_999_999)))
        nombre = rng.choice(NOMBRES_EJEMPLO)
        departamento = rng.choice(DEPARTAMENTOS)
        distrito = rng.choice(DISTRITOS_POR_DEPARTAMENTO[departamento])
        habilitador = rng.choice(HABILITADORES)

        for producto in PRODUCTOS:
            cuota = int(rng.integers(50, 500))
            # Factor de avance variable para simular sub y sobre cumplimiento
            factor_avance = rng.uniform(0.4, 1.3)
            avance = int(round(cuota * factor_avance))
            registros.append({
                "DNI": dni,
                "Nombre": nombre,
                "Departamento": departamento,
                "Distrito": distrito,
                "Habilitador": habilitador,
                "Producto": producto,
                "Cuota": cuota,
                "Avance": avance,
            })

    return pd.DataFrame(registros)


def cargar_datos_excel(archivo) -> pd.DataFrame | None:
    """Lee y valida un archivo Excel cargado por el administrador.

    Retorna None (y muestra un error en la UI) si faltan columnas requeridas.
    """
    try:
        df = pd.read_excel(archivo)
    except Exception as exc:  # noqa: BLE001 - se informa al usuario cualquier error de lectura
        st.error(f"No se pudo leer el archivo Excel: {exc}")
        return None

    faltantes = COLUMNAS_REQUERIDAS - set(df.columns)
    if faltantes:
        st.error(
            "El archivo no contiene las columnas requeridas: "
            + ", ".join(sorted(faltantes))
        )
        return None

    return df


@st.cache_data
def _leer_excel_publicado(path: str, mtime: float) -> pd.DataFrame:
    """Lee el Excel publicado. `mtime` forma parte de la clave de cache: si el
    administrador sube un archivo nuevo, el mtime cambia y el cache se invalida
    automáticamente para todos los usuarios."""
    return pd.read_excel(path)


def obtener_datos_publicados() -> pd.DataFrame:
    """Devuelve los datos que ve gerencia: el último archivo publicado por el
    administrador, o el dataset de ejemplo si todavía no se publicó ninguno."""
    if os.path.exists(DATA_FILE):
        return _leer_excel_publicado(DATA_FILE, os.path.getmtime(DATA_FILE))
    return generar_datos_ejemplo()


def publicar_datos(df: pd.DataFrame) -> None:
    """Guarda el archivo validado como la fuente de datos oficial del
    dashboard, visible para todos los usuarios en su próxima recarga."""
    os.makedirs(DATA_DIR, exist_ok=True)
    df.to_excel(DATA_FILE, index=False)
    _leer_excel_publicado.clear()  # invalida el cache de lectura


def calcular_metricas(df: pd.DataFrame, fecha_referencia: datetime | None = None) -> pd.DataFrame:
    """Calcula las columnas derivadas del análisis:

    - Cumplimiento % = Avance / Cuota
    - Proyección $   = Avance proyectado linealmente al cierre del mes,
                        según el día actual del periodo.
    """
    df = df.copy()

    # Cumplimiento %: evita división entre cero
    df["Cumplimiento %"] = np.where(df["Cuota"] > 0, df["Avance"] / df["Cuota"], 0.0)

    # Proyección $: se asume un avance lineal dentro del mes en curso
    fecha_referencia = fecha_referencia or datetime.now()
    dias_en_mes = calendar.monthrange(fecha_referencia.year, fecha_referencia.month)[1]
    dia_actual = max(fecha_referencia.day, 1)
    factor_proyeccion = dias_en_mes / dia_actual

    df["Proyección $"] = df["Avance"] * factor_proyeccion

    return df


# =============================================================================
# 3. FUNCIONES DE PRESENTACIÓN (FORMATO SEMÁFORO)
# =============================================================================

def color_semaforo(valor: float) -> str:
    """Devuelve el estilo CSS de fondo según el cumplimiento (semáforo)."""
    if pd.isna(valor):
        return ""
    if valor < 0.80:
        color = "#f8d7da"  # rojo
    elif valor < 1.00:
        color = "#fff3cd"  # amarillo
    else:
        color = "#d4edda"  # verde
    return f"background-color: {color}"


def aplicar_estilo_tabla(tabla: pd.DataFrame):
    """Aplica formato numérico y semáforo de cumplimiento a la tabla."""
    styler = tabla.style.format({
        "Cuota": "{:,.0f}",
        "Avance": "{:,.0f}",
        "Cumplimiento %": "{:.1%}",
        "Proyección $": "{:,.0f}",
    })

    # Compatibilidad entre versiones de pandas: Styler.map (>=2.1) vs applymap
    if hasattr(styler, "map"):
        styler = styler.map(color_semaforo, subset=["Cumplimiento %"])
    else:  # pragma: no cover - fallback para pandas antiguo
        styler = styler.applymap(color_semaforo, subset=["Cumplimiento %"])

    return styler


# =============================================================================
# 4. PANEL DE ADMINISTRADOR (ACCESO RESTRINGIDO)
# =============================================================================
#
# Configuración de credenciales (recomendado, no se sube al repositorio):
# En Streamlit Cloud → Settings → Secrets, agregar:
#
#   [admin]
#   usuario = "admin"
#   password = "coloca_aqui_una_clave_segura"
#
# Si no se configuran secrets (por ejemplo en pruebas locales), se usa un
# usuario/clave por defecto que se debe cambiar antes de publicar el enlace.

def _credenciales_admin() -> tuple[str, str]:
    try:
        return st.secrets["admin"]["usuario"], st.secrets["admin"]["password"]
    except Exception:  # noqa: BLE001 - no hay secrets configurados aún
        return "admin", "cambiar123"


def panel_admin() -> None:
    """Renderiza el control de acceso y la carga de datos. Solo se llama
    cuando la URL incluye ?admin=1, por lo que gerencia nunca lo ve."""
    st.header("🔒 Panel administrador")

    if not st.session_state.get("es_admin", False):
        with st.form("form_login_admin"):
            usuario = st.text_input("Usuario")
            clave = st.text_input("Contraseña", type="password")
            enviar = st.form_submit_button("Ingresar")

        if enviar:
            usuario_ok, clave_ok = _credenciales_admin()
            if usuario == usuario_ok and clave == clave_ok:
                st.session_state["es_admin"] = True
                st.rerun()
            else:
                st.error("Usuario o contraseña incorrectos.")
        return

    # --- Sesión de administrador activa ---
    st.success("Sesión de administrador activa.")

    with st.expander("Columnas requeridas del Excel"):
        st.write(sorted(COLUMNAS_REQUERIDAS))

    archivo = st.file_uploader("Cargar archivo Excel (.xlsx)", type=["xlsx"])
    if archivo is not None:
        df_validado = cargar_datos_excel(archivo)
        if df_validado is not None:
            publicar_datos(df_validado)
            st.success("Datos publicados. Todos los usuarios verán la actualización al recargar.")

    if os.path.exists(DATA_FILE):
        ultima_actualizacion = datetime.fromtimestamp(os.path.getmtime(DATA_FILE))
        st.caption(f"Última publicación: {ultima_actualizacion:%d/%m/%Y %H:%M}")

    if st.button("Cerrar sesión"):
        st.session_state["es_admin"] = False
        st.rerun()


# =============================================================================
# 5. INTERFAZ PRINCIPAL
# =============================================================================

def main():
    st.title("📊 Dashboard Gerencial de Ventas")
    st.caption(
        "Desempeño de habilitadores por producto y ubicación · "
        f"Actualizado al {datetime.now():%d/%m/%Y}"
    )

    # El panel administrador solo se renderiza con ?admin=1 en la URL.
    # Sin ese parámetro no queda ningún rastro visible (ni siquiera la
    # flechita de la sidebar aparece, porque nunca se le agrega contenido).
    if st.query_params.get("admin") == "1":
        with st.sidebar:
            panel_admin()

    # Fuente de datos: lo último publicado por el administrador,
    # o el dataset de ejemplo si aún no se publicó nada.
    df_raw = obtener_datos_publicados()

    # Se restringe el análisis únicamente a los departamentos y productos válidos
    df_raw = df_raw[df_raw["Departamento"].isin(DEPARTAMENTOS)]
    df_raw = df_raw[df_raw["Producto"].isin(PRODUCTOS)]

    if df_raw.empty:
        st.warning("No hay datos disponibles para los departamentos/productos configurados.")
        return

    df = calcular_metricas(df_raw)

    # --- Filtros ---
    st.subheader("Filtros")
    col_hab, col_dep, col_prod = st.columns(3)

    with col_hab:
        # Filtro principal: lista desplegable de habilitador
        habilitador_sel = st.selectbox("Habilitador", HABILITADORES)

    with col_dep:
        departamentos_sel = st.multiselect(
            "Departamento",
            options=sorted(df["Departamento"].unique()),
            default=sorted(df["Departamento"].unique()),
        )

    with col_prod:
        productos_sel = st.multiselect(
            "Producto",
            options=PRODUCTOS,
            default=PRODUCTOS,
        )

    df_filtrado = df[
        (df["Habilitador"] == habilitador_sel)
        & (df["Departamento"].isin(departamentos_sel))
        & (df["Producto"].isin(productos_sel))
    ]

    st.markdown("---")

    # --- KPIs gerenciales ---
    total_cuota = df_filtrado["Cuota"].sum()
    total_avance = df_filtrado["Avance"].sum()
    # % de cumplimiento promedio ponderado (avance total / cuota total)
    cumplimiento_prom = (total_avance / total_cuota) if total_cuota > 0 else 0.0

    kpi1, kpi2, kpi3 = st.columns(3)
    kpi1.metric("Total Cuota", f"{total_cuota:,.0f}")
    kpi2.metric("Total Avance", f"{total_avance:,.0f}")
    kpi3.metric("% Cumplimiento Promedio", f"{cumplimiento_prom:.1%}")

    st.markdown("---")

    # --- Pestañas ---
    (tab_detalle,) = st.tabs(["Detalle por Gestor"])

    with tab_detalle:
        st.subheader(f"Detalle por Gestor · {habilitador_sel}")

        columnas_mostrar = [
            "DNI", "Nombre", "Departamento", "Distrito", "Producto",
            "Cuota", "Avance", "Cumplimiento %", "Proyección $",
        ]

        if df_filtrado.empty:
            st.info("No hay registros para los filtros seleccionados.")
        else:
            tabla = (
                df_filtrado[columnas_mostrar]
                .sort_values(["Departamento", "Nombre", "Producto"])
                .reset_index(drop=True)
            )

            st.dataframe(
                aplicar_estilo_tabla(tabla),
                use_container_width=True,
                height=500,
            )

            st.download_button(
                "⬇️ Descargar tabla filtrada (CSV)",
                data=tabla.to_csv(index=False).encode("utf-8"),
                file_name=f"detalle_gestor_{habilitador_sel.replace(' ', '_').lower()}.csv",
                mime="text/csv",
            )

            st.caption("🟥 <80% · 🟨 80%–99% · 🟩 ≥100%")


if __name__ == "__main__":
    main()
