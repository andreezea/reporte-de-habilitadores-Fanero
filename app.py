"""
Reporte Habilitadores Fanero - Streamlit App
================================================
Dashboard gerencial para analizar el desempeño de habilitadores de venta
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

Lógica de proyección (Cuota y Avance son unidades, no montos en dinero):
    Al publicar los datos, el administrador indica el "día de corte": el
    último día del mes con ventas efectivamente cargadas. Con eso:
        Proy Unidades = Avance * (días del mes / día de corte)
        Proy %        = Proy Unidades / Cuota
        Días restantes = días del mes - día de corte
        Cuota diaria necesaria = (Cuota - Avance) / Días restantes

Listo para desplegar en Streamlit Cloud: `streamlit run app.py`
"""

import calendar
import json
import os
from datetime import datetime

import numpy as np
import pandas as pd
import streamlit as st

# =============================================================================
# 1. CONFIGURACIÓN Y CONSTANTES
# =============================================================================

st.set_page_config(
    page_title="Reporte Habilitadores Fanero",
    page_icon="📊",
    layout="wide",
    # Sidebar oculta por defecto: ahí vive el panel de administrador, y así
    # nadie que reciba el enlace del dashboard nota que existe.
    initial_sidebar_state="collapsed",
)

# Rutas donde se guarda el último archivo publicado por el administrador y
# sus metadatos (mes/año/día de corte). Persisten mientras la app siga
# corriendo, y las ve todo el que abra el enlace del dashboard.
DATA_DIR = "data"
DATA_FILE = os.path.join(DATA_DIR, "ultima_carga.xlsx")
DATA_META = os.path.join(DATA_DIR, "meta.json")

# Habilitadores disponibles para el filtro principal (selectbox)
HABILITADORES = [
    "PDV Plus",
    "Captura",
    "Mercados",
    "Ferias",
    "Activaciones",
    "Desarrolladores",
]

# Departamentos permitidos en el análisis (el orden aquí define el orden de
# las tablas que agrupan por departamento)
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

# Productos analizados (el orden aquí define el orden de las tablas)
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


def obtener_datos_publicados() -> tuple[pd.DataFrame, int, int, int]:
    """Devuelve lo que ve gerencia: (datos, día de corte, mes, año).

    Si el administrador ya publicó un archivo, se usan sus datos y los
    metadatos que declaró (mes / año / día de corte de las ventas). Si aún no
    se publicó nada, se usa el dataset de ejemplo con un día de corte
    razonable por defecto (ayer, respecto de la fecha del servidor).
    """
    ahora = datetime.now()

    if os.path.exists(DATA_FILE):
        df = _leer_excel_publicado(DATA_FILE, os.path.getmtime(DATA_FILE))
        try:
            with open(DATA_META, "r", encoding="utf-8") as f:
                meta = json.load(f)
            dia_corte = int(meta.get("dia_corte", max(ahora.day - 1, 1)))
            mes = int(meta.get("mes", ahora.month))
            anio = int(meta.get("anio", ahora.year))
        except Exception:  # noqa: BLE001 - metadatos no disponibles o corruptos
            dia_corte, mes, anio = max(ahora.day - 1, 1), ahora.month, ahora.year
        return df, dia_corte, mes, anio

    dia_corte = max(ahora.day - 1, 1)
    return generar_datos_ejemplo(), dia_corte, ahora.month, ahora.year


def publicar_datos(df: pd.DataFrame, dia_corte: int, mes: int, anio: int) -> None:
    """Guarda el archivo validado y sus metadatos como la fuente de datos
    oficial del dashboard, visible para todos los usuarios en su próxima
    recarga."""
    os.makedirs(DATA_DIR, exist_ok=True)
    df.to_excel(DATA_FILE, index=False)
    with open(DATA_META, "w", encoding="utf-8") as f:
        json.dump({"dia_corte": dia_corte, "mes": mes, "anio": anio}, f)
    _leer_excel_publicado.clear()  # invalida el cache de lectura


def calcular_metricas(df: pd.DataFrame, dias_en_mes: int, dia_corte: int) -> pd.DataFrame:
    """Calcula las columnas derivadas del análisis (Cuota/Avance son
    unidades, no montos en dinero):

    - Cumplimiento % = Avance / Cuota
    - Proy Unidades  = Avance * (días del mes / día de corte)
    - Proy %         = Proy Unidades / Cuota
    """
    df = df.copy()
    dia_corte = max(dia_corte, 1)  # evita división entre cero

    df["Cumplimiento %"] = np.where(df["Cuota"] > 0, df["Avance"] / df["Cuota"], 0.0)

    factor_proyeccion = dias_en_mes / dia_corte
    df["Proy Unidades"] = df["Avance"] * factor_proyeccion
    df["Proy %"] = np.where(df["Cuota"] > 0, df["Proy Unidades"] / df["Cuota"], 0.0)

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
    return f"background-color: {color}; color: #1a1a1a"


def _aplicar_semaforo(styler, columnas: list[str]):
    """Aplica color_semaforo a las columnas indicadas, compatible con
    distintas versiones de pandas (Styler.map vs applymap)."""
    if hasattr(styler, "map"):
        for col in columnas:
            styler = styler.map(color_semaforo, subset=[col])
    else:  # pragma: no cover - fallback para pandas antiguo
        for col in columnas:
            styler = styler.applymap(color_semaforo, subset=[col])
    return styler


def aplicar_estilo_tabla(tabla: pd.DataFrame):
    """Aplica formato numérico y semáforo de cumplimiento a la tabla de
    detalle por gestor."""
    styler = tabla.style.format({
        "Cuota": "{:,.0f}",
        "Avance": "{:,.0f}",
        "Cumplimiento %": "{:.1%}",
        "Proy Unidades": "{:,.0f}",
    })
    return _aplicar_semaforo(styler, ["Cumplimiento %"])


def aplicar_estilo_resumen_producto(tabla: pd.DataFrame):
    """Aplica formato numérico y semáforo (Cumplimiento % y Proy %) al
    resumen por departamento y producto."""
    styler = tabla.style.format({
        "Cuota": "{:,.0f}",
        "Avance": "{:,.0f}",
        "Cumplimiento %": "{:.1%}",
        "Proy Unidades": "{:,.0f}",
        "Proy %": "{:.1%}",
    })
    return _aplicar_semaforo(styler, ["Cumplimiento %", "Proy %"])


def aplicar_estilo_ritmo(tabla: pd.DataFrame):
    """Aplica formato numérico y semáforo (Cumplimiento %) a la tabla de
    ritmo diario."""
    styler = tabla.style.format({
        "Cuota": "{:,.0f}",
        "Avance": "{:,.0f}",
        "Cumplimiento %": "{:.1%}",
        "Diferencia": "{:,.0f}",
        "Cuota Diaria Necesaria": "{:,.1f}",
    })
    return _aplicar_semaforo(styler, ["Cumplimiento %"])


def resumen_por_producto(df_filtrado: pd.DataFrame, departamentos_sel: list[str],
                          productos_sel: list[str], dias_en_mes: int,
                          dia_corte: int) -> pd.DataFrame:
    """Agrega Cuota y Avance por Departamento + Producto (nunca se suman
    productos distintos entre sí) y recalcula Cumplimiento %, Proy Unidades y
    Proy % a partir de los totales agregados.

    Si hay varios departamentos seleccionados, cada uno aparece como su
    propia fila (una por cada combinación Departamento/Producto)."""
    resumen = (
        df_filtrado.groupby(["Departamento", "Producto"], as_index=False)
        .agg(Cuota=("Cuota", "sum"), Avance=("Avance", "sum"))
    )

    orden_dep = [d for d in DEPARTAMENTOS if d in departamentos_sel]
    orden_prod = [p for p in PRODUCTOS if p in productos_sel]
    combinaciones = pd.MultiIndex.from_product([orden_dep, orden_prod], names=["Departamento", "Producto"])

    resumen = resumen.set_index(["Departamento", "Producto"]).reindex(combinaciones).reset_index()
    resumen[["Cuota", "Avance"]] = resumen[["Cuota", "Avance"]].fillna(0)

    resumen["Cumplimiento %"] = np.where(resumen["Cuota"] > 0, resumen["Avance"] / resumen["Cuota"], 0.0)
    factor_proyeccion = dias_en_mes / max(dia_corte, 1)
    resumen["Proy Unidades"] = resumen["Avance"] * factor_proyeccion
    resumen["Proy %"] = np.where(resumen["Cuota"] > 0, resumen["Proy Unidades"] / resumen["Cuota"], 0.0)

    return resumen[["Departamento", "Producto", "Cuota", "Avance", "Cumplimiento %", "Proy Unidades", "Proy %"]]


def ritmo_diario_por_producto(df_filtrado: pd.DataFrame, productos_sel: list[str],
                               dias_restantes: int) -> pd.DataFrame:
    """Calcula, por producto, cuánto se necesita vender por día para cerrar
    la cuota en lo que resta del mes: (Cuota - Avance) / días restantes."""
    ritmo = (
        df_filtrado.groupby("Producto", as_index=False)
        .agg(Cuota=("Cuota", "sum"), Avance=("Avance", "sum"))
    )

    orden = [p for p in PRODUCTOS if p in productos_sel]
    ritmo = ritmo.set_index("Producto").reindex(orden).reset_index()
    ritmo[["Cuota", "Avance"]] = ritmo[["Cuota", "Avance"]].fillna(0)

    ritmo["Cumplimiento %"] = np.where(ritmo["Cuota"] > 0, ritmo["Avance"] / ritmo["Cuota"], 0.0)
    ritmo["Diferencia"] = ritmo["Cuota"] - ritmo["Avance"]

    if dias_restantes > 0:
        ritmo["Cuota Diaria Necesaria"] = ritmo["Diferencia"] / dias_restantes
    else:
        ritmo["Cuota Diaria Necesaria"] = np.nan

    return ritmo[["Producto", "Cuota", "Avance", "Cumplimiento %", "Diferencia", "Cuota Diaria Necesaria"]]


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

    ahora = datetime.now()
    col_mes, col_anio = st.columns(2)
    with col_mes:
        mes_sel = st.number_input("Mes de la cuota", min_value=1, max_value=12, value=ahora.month)
    with col_anio:
        anio_sel = st.number_input("Año", min_value=2020, max_value=2100, value=ahora.year, step=1)

    dias_en_mes_sel = calendar.monthrange(int(anio_sel), int(mes_sel))[1]
    dia_corte_defecto = min(max(ahora.day - 1, 1), dias_en_mes_sel)
    dia_corte_sel = st.number_input(
        "Día de corte (último día del mes con ventas cargadas)",
        min_value=1, max_value=dias_en_mes_sel, value=dia_corte_defecto,
    )

    archivo = st.file_uploader("Cargar archivo Excel (.xlsx)", type=["xlsx"])

    if archivo is not None and st.button("Publicar datos"):
        df_validado = cargar_datos_excel(archivo)
        if df_validado is not None:
            publicar_datos(df_validado, int(dia_corte_sel), int(mes_sel), int(anio_sel))
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
    st.title("📊 Reporte Habilitadores Fanero")

    # El panel administrador solo se renderiza con ?admin=1 en la URL.
    # Sin ese parámetro no queda ningún rastro visible (ni siquiera la
    # flechita de la sidebar aparece, porque nunca se le agrega contenido).
    if st.query_params.get("admin") == "1":
        with st.sidebar:
            panel_admin()

    # Fuente de datos: lo último publicado por el administrador (con su día
    # de corte), o el dataset de ejemplo si aún no se publicó nada.
    df_raw, dia_corte, mes, anio = obtener_datos_publicados()
    dias_en_mes = calendar.monthrange(anio, mes)[1]
    dias_restantes = max(dias_en_mes - dia_corte, 0)

    st.caption(
        "Desempeño de habilitadores por producto y ubicación · "
        f"Datos al día {dia_corte} de {dias_en_mes} ({mes:02d}/{anio})"
    )

    # Se restringe el análisis únicamente a los departamentos y productos válidos
    df_raw = df_raw[df_raw["Departamento"].isin(DEPARTAMENTOS)]
    df_raw = df_raw[df_raw["Producto"].isin(PRODUCTOS)]

    if df_raw.empty:
        st.warning("No hay datos disponibles para los departamentos/productos configurados.")
        return

    df = calcular_metricas(df_raw, dias_en_mes, dia_corte)

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

    # Salvaguarda: si el usuario deja un multiselect vacío (por ejemplo, borró
    # los chips por accidente), se interpreta como "sin filtro" en vez de
    # mostrar un dashboard vacío.
    if not departamentos_sel:
        departamentos_sel = sorted(df["Departamento"].unique())
    if not productos_sel:
        productos_sel = PRODUCTOS

    df_filtrado = df[
        (df["Habilitador"] == habilitador_sel)
        & (df["Departamento"].isin(departamentos_sel))
        & (df["Producto"].isin(productos_sel))
    ]

    st.markdown("---")

    # --- Resumen por producto (y departamento, si hay varios seleccionados) ---
    st.subheader("Resumen por Producto")
    if df_filtrado.empty:
        st.info("No hay datos para los filtros seleccionados.")
    else:
        tabla_resumen = resumen_por_producto(
            df_filtrado, departamentos_sel, productos_sel, dias_en_mes, dia_corte
        )
        st.dataframe(
            aplicar_estilo_resumen_producto(tabla_resumen),
            use_container_width=True,
            hide_index=True,
        )
        st.caption("🟥 <80% · 🟨 80%–99% · 🟩 ≥100% (aplica a Cumplimiento % y Proy %)")

    st.markdown("---")

    # --- Pestañas ---
    tab_detalle, tab_ritmo = st.tabs(["Detalle Habilitador", "Ritmo Diario"])

    with tab_detalle:
        st.subheader(f"Detalle habilitador. {habilitador_sel}")

        columnas_mostrar = [
            "DNI", "Nombre", "Departamento", "Distrito", "Producto",
            "Cuota", "Avance", "Cumplimiento %", "Proy Unidades",
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
                file_name=f"detalle_habilitador_{habilitador_sel.replace(' ', '_').lower()}.csv",
                mime="text/csv",
            )

            st.caption("🟥 <80% · 🟨 80%–99% · 🟩 ≥100%")

    with tab_ritmo:
        st.subheader(f"Ritmo diario necesario · {habilitador_sel}")
        st.caption(
            f"Quedan {dias_restantes} día(s) para el cierre del mes "
            f"(día {dia_corte} → día {dias_en_mes})."
        )

        if df_filtrado.empty:
            st.info("No hay datos para los filtros seleccionados.")
        else:
            tabla_ritmo = ritmo_diario_por_producto(df_filtrado, productos_sel, dias_restantes)

            if dias_restantes == 0:
                st.warning("El mes ya cerró; no quedan días para calcular el ritmo diario.")

            st.dataframe(
                aplicar_estilo_ritmo(tabla_ritmo),
                use_container_width=True,
                hide_index=True,
            )
            st.caption(
                "Cuota Diaria Necesaria = (Cuota - Avance) / días restantes · "
                "🟥 <80% · 🟨 80%–99% · 🟩 ≥100% (Cumplimiento %)"
            )


if __name__ == "__main__":
    main()
