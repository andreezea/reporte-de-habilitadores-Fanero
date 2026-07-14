"""
Reporte Habilitadores Fanero - Streamlit App
================================================
Dashboard gerencial para analizar el desempeño de habilitadores de venta
(PDV Plus, Captura, Mercados, Ferias, Activaciones, Desarrolladores) por
producto y ubicación geográfica.

Estructura del archivo:
    1. Configuración y constantes
    2. Funciones de datos (carga / generación / cálculo / publicación)
    3. Funciones de presentación (KPIs, formato semáforo, tablas dinámicas)
    4. Panel de administrador (acceso restringido)
    5. Plantillas Excel descargables
    6. Edición de avances por coordinador (acceso restringido)
    7. Interfaz principal (main)

Accesos ocultos (nadie los ve sin conocer la URL exacta):
    ?admin=1   → panel de administrador: publica el archivo Excel completo
                 (con día de corte, mes y año). Ej: tu-app.streamlit.app/?admin=1
    ?editar=1  → pestaña "Editar Avances": cada coordinador actualiza el
                 avance de sus propios departamentos, sin tocar el resto.
                 Ej: tu-app.streamlit.app/?editar=1
    Ambos se pueden combinar, por ejemplo ?admin=1&editar=1.

Lógica de proyección (Cuota y Avance son unidades, no montos en dinero):
    Al publicar los datos, el administrador indica el "día de corte": el
    último día del mes con ventas efectivamente cargadas. Con eso:
        Proy Unidades = Avance * (días del mes / día de corte)
        Proy %        = Proy Unidades / Cuota
        Días restantes = días del mes - día de corte
        Cuota diaria necesaria = (Cuota - Avance) / Días restantes

Puntos de Venta (PDV):
    Activaciones y Desarrolladores son habilitadores cuyos gestores manejan
    varios Puntos de Venta (PDV) por debajo, cada uno con su propia Cuota y
    Avance. El resto de habilitadores no usa este nivel (columna PDV vacía).
    La columna "PDV" en el Excel es opcional: si no existe, se asume vacía
    para todas las filas (compatibilidad con archivos anteriores). En la
    pestaña "Detalle Habilitador", cuando el habilitador seleccionado es uno
    de estos dos, cada gestor se muestra como una fila expandible: al hacer
    clic se despliega el resumen agregado y el detalle por PDV.

Listo para desplegar en Streamlit Cloud: `streamlit run app.py`
"""

import calendar
import json
import os
from datetime import datetime
from io import BytesIO

import numpy as np
import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.worksheet.datavalidation import DataValidation

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

# Rutas donde se guarda el último archivo publicado por el administrador, sus
# metadatos (mes/año/día de corte) y el registro de la última edición de
# avances. Persisten mientras la app siga corriendo, y las ve todo el que
# abra el enlace del dashboard.
DATA_DIR = "data"
DATA_FILE = os.path.join(DATA_DIR, "ultima_carga.xlsx")
DATA_META = os.path.join(DATA_DIR, "meta.json")
LOG_EDICION = os.path.join(DATA_DIR, "ultima_edicion.json")

# Habilitadores disponibles para el filtro principal (selectbox)
HABILITADORES = [
    "PDV Plus",
    "Captura",
    "Mercados",
    "Ferias",
    "Activaciones",
    "Desarrolladores",
]

# Habilitadores cuyos gestores tienen Puntos de Venta (PDV) por debajo, cada
# uno con su propia Cuota/Avance.
HABILITADORES_CON_PDV = ["Activaciones", "Desarrolladores"]

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

# Productos analizados (el orden aquí define el orden de las columnas)
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

# Columnas obligatorias en el Excel. "PDV" es opcional (ver HABILITADORES_CON_PDV).
COLUMNAS_REQUERIDAS = {
    "DNI", "Nombre", "Departamento", "Distrito", "Habilitador", "Producto",
    "Cuota", "Avance",
}

# Coordinadores que pueden editar avances (pestaña oculta ?editar=1) y los
# departamentos que cada uno tiene permitido tocar. Selección por nombre
# únicamente (sin PIN): pensado para uso interno de confianza, no como
# control de acceso fuerte.
COORDINADORES = {
    "Cinthya Maravi": ["Pasco"],
    "Diana Alvarado": ["Amazonas", "Cajamarca"],
    "Gian Ramirez": ["San Martín"],
    "Hernán Pizarro": ["Loreto"],
    "Jeferson Torres": ["Junín", "Huancavelica"],
    "Keiner Valdivia": ["Huánuco"],
    "Sheyla Piro": ["Ucayali"],
}


# =============================================================================
# 2. FUNCIONES DE DATOS
# =============================================================================

def _normalizar_pdv(df: pd.DataFrame) -> pd.DataFrame:
    """Garantiza que exista la columna PDV (opcional) y que quede como texto
    limpio, vacía ("") si no aplica. Se usa en todos los puntos donde entran
    datos nuevos: generación de ejemplo, carga de Excel y lectura publicada."""
    df = df.copy()
    if "PDV" not in df.columns:
        df["PDV"] = ""
    df["PDV"] = df["PDV"].fillna("").astype(str).str.strip()
    return df


@st.cache_data
def generar_datos_ejemplo(n_gestores: int = 150, seed: int = 42) -> pd.DataFrame:
    """Genera un dataset sintético de ventas, con 4 filas (una por producto)
    por cada gestor, para poder probar el dashboard sin un archivo real.

    Para Activaciones y Desarrolladores, cada producto se reparte entre 2 o 3
    PDV con su propia Cuota/Avance, simulando la estructura real de esos
    habilitadores."""
    rng = np.random.default_rng(seed)
    registros = []

    for i in range(n_gestores):
        dni = str(40000000 + int(rng.integers(0, 9_999_999)))
        nombre = rng.choice(NOMBRES_EJEMPLO)
        departamento = rng.choice(DEPARTAMENTOS)
        distrito = rng.choice(DISTRITOS_POR_DEPARTAMENTO[departamento])
        habilitador = rng.choice(HABILITADORES)
        tiene_pdv = habilitador in HABILITADORES_CON_PDV

        for producto in PRODUCTOS:
            cuota_total = int(rng.integers(50, 500))
            n_pdv = int(rng.integers(2, 4)) if tiene_pdv else 1
            pesos = rng.dirichlet(np.ones(n_pdv))

            for idx in range(n_pdv):
                cuota_pdv = max(int(round(cuota_total * pesos[idx])), 1)
                factor_avance = rng.uniform(0.4, 1.3)
                avance_pdv = int(round(cuota_pdv * factor_avance))
                registros.append({
                    "DNI": dni,
                    "Nombre": nombre,
                    "Departamento": departamento,
                    "Distrito": distrito,
                    "Habilitador": habilitador,
                    "Producto": producto,
                    "PDV": f"PDV {idx + 1}" if tiene_pdv else "",
                    "Cuota": cuota_pdv,
                    "Avance": avance_pdv,
                })

    return pd.DataFrame(registros)


def cargar_datos_excel(archivo) -> pd.DataFrame | None:
    """Lee y valida un archivo Excel cargado por el administrador.

    Retorna None (y muestra un error en la UI) si faltan columnas requeridas.
    La columna PDV es opcional: si no viene, se agrega vacía para todas las
    filas (compatibilidad con archivos que no usan Puntos de Venta).
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

    return _normalizar_pdv(df)


@st.cache_data
def _leer_excel_publicado(path: str, mtime: float) -> pd.DataFrame:
    """Lee el Excel publicado. `mtime` forma parte de la clave de cache: si
    cambia el archivo (nueva carga del admin o edición de avances), el cache
    se invalida automáticamente para todos los usuarios."""
    df = pd.read_excel(path)
    df["DNI"] = df["DNI"].astype(str).str.strip()
    df = _normalizar_pdv(df)
    return df


def obtener_datos_publicados() -> tuple[pd.DataFrame, int, int, int]:
    """Devuelve lo que ve gerencia: (datos, día de corte, mes, año).

    Si el administrador ya publicó un archivo, se usan sus datos y los
    metadatos que declaró (mes / año / día de corte de las ventas). Si aún no
    se publicó nada, se usa el dataset de ejemplo con un día de corte
    razonable por defecto (ayer, respecto de la fecha del servidor).

    El DNI siempre se normaliza a texto: al guardar/leer Excel, pandas suele
    inferir números en columnas que parecen enteras, y eso rompería el cruce
    DNI + Producto (+ PDV) al editar avances si no se normaliza en un solo
    lugar.
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
    df_ejemplo = generar_datos_ejemplo().copy()
    df_ejemplo["DNI"] = df_ejemplo["DNI"].astype(str).str.strip()
    df_ejemplo = _normalizar_pdv(df_ejemplo)
    return df_ejemplo, dia_corte, ahora.month, ahora.year


def publicar_datos(df: pd.DataFrame, dia_corte: int, mes: int, anio: int) -> None:
    """Guarda el archivo validado y sus metadatos como la fuente de datos
    oficial del dashboard, visible para todos los usuarios en su próxima
    recarga."""
    os.makedirs(DATA_DIR, exist_ok=True)
    df = df.copy()
    df["DNI"] = df["DNI"].astype(str).str.strip()
    df = _normalizar_pdv(df)
    df.to_excel(DATA_FILE, index=False)
    with open(DATA_META, "w", encoding="utf-8") as f:
        json.dump({"dia_corte": dia_corte, "mes": mes, "anio": anio}, f)
    _leer_excel_publicado.clear()  # invalida el cache de lectura


def calcular_metricas(df: pd.DataFrame, dias_en_mes: int, dia_corte: int) -> pd.DataFrame:
    """Calcula las columnas derivadas del análisis (Cuota/Avance son
    unidades, no montos en dinero). Se calculan siempre a nivel de fila (que
    puede ser un PDV o directamente un gestor, según el habilitador):

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
# 3. FUNCIONES DE PRESENTACIÓN (FORMATO SEMÁFORO Y TABLAS DINÁMICAS)
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


def _aplicar_semaforo(styler, columnas: list):
    """Aplica color_semaforo a las columnas indicadas (nombres simples o
    tuplas, para tablas con columnas MultiIndex), compatible con distintas
    versiones de pandas (Styler.map vs applymap)."""
    if hasattr(styler, "map"):
        for col in columnas:
            styler = styler.map(color_semaforo, subset=[col])
    else:  # pragma: no cover - fallback para pandas antiguo
        for col in columnas:
            styler = styler.applymap(color_semaforo, subset=[col])
    return styler


def aplicar_estilo_tabla(tabla: pd.DataFrame):
    """Aplica formato numérico y semáforo de cumplimiento a una tabla plana
    que tenga las columnas Cuota, Avance, Cumplimiento % y Proy Unidades
    (se usa tanto en el detalle por gestor como en los cuadros de PDV)."""
    fmt = {}
    for col, patron in (
        ("Cuota", "{:,.0f}"), ("Avance", "{:,.0f}"),
        ("Cumplimiento %", "{:.1%}"), ("Proy Unidades", "{:,.0f}"),
    ):
        if col in tabla.columns:
            fmt[col] = patron
    styler = tabla.style.format(fmt)
    columnas_semaforo = [c for c in ["Cumplimiento %"] if c in tabla.columns]
    return _aplicar_semaforo(styler, columnas_semaforo)


def resumen_por_producto(df_filtrado: pd.DataFrame, departamentos_sel: list,
                          productos_sel: list, dias_en_mes: int, dia_corte: int) -> pd.DataFrame:
    """Devuelve el resumen por Departamento con los productos como columnas
    agrupadas: debajo de cada producto van Cuota, Avance, Cumplimiento %,
    Proy Unidades y Proy %. Cuota y Avance se agregan primero (suma dentro
    del mismo producto, incluyendo todos los PDV si aplica) y los porcentajes
    se recalculan sobre esos totales; nunca se suman productos distintos
    entre sí."""
    largo = (
        df_filtrado.groupby(["Departamento", "Producto"], as_index=False)
        .agg(Cuota=("Cuota", "sum"), Avance=("Avance", "sum"))
    )

    orden_dep = [d for d in DEPARTAMENTOS if d in departamentos_sel]
    orden_prod = [p for p in PRODUCTOS if p in productos_sel]
    combinaciones = pd.MultiIndex.from_product([orden_dep, orden_prod], names=["Departamento", "Producto"])
    largo = largo.set_index(["Departamento", "Producto"]).reindex(combinaciones).reset_index()
    largo[["Cuota", "Avance"]] = largo[["Cuota", "Avance"]].fillna(0)

    largo["Cumplimiento %"] = np.where(largo["Cuota"] > 0, largo["Avance"] / largo["Cuota"], 0.0)
    factor_proyeccion = dias_en_mes / max(dia_corte, 1)
    largo["Proy Unidades"] = largo["Avance"] * factor_proyeccion
    largo["Proy %"] = np.where(largo["Cuota"] > 0, largo["Proy Unidades"] / largo["Cuota"], 0.0)

    metricas = ["Cuota", "Avance", "Cumplimiento %", "Proy Unidades", "Proy %"]
    ancho = largo.pivot_table(index="Departamento", columns="Producto", values=metricas, aggfunc="first")
    ancho = ancho.swaplevel(axis=1)
    columnas_orden = pd.MultiIndex.from_product([orden_prod, metricas])
    ancho = ancho.reindex(index=orden_dep, columns=columnas_orden)

    return ancho


def aplicar_estilo_resumen_producto(tabla: pd.DataFrame, orden_prod: list):
    """Aplica formato numérico y semáforo (Cumplimiento % y Proy %) al
    resumen por departamento con productos como columnas agrupadas."""
    fmt = {}
    for p in orden_prod:
        fmt[(p, "Cuota")] = "{:,.0f}"
        fmt[(p, "Avance")] = "{:,.0f}"
        fmt[(p, "Cumplimiento %")] = "{:.1%}"
        fmt[(p, "Proy Unidades")] = "{:,.0f}"
        fmt[(p, "Proy %")] = "{:.1%}"

    styler = tabla.style.format(fmt, na_rep="-")
    subset = [(p, "Cumplimiento %") for p in orden_prod] + [(p, "Proy %") for p in orden_prod]
    return _aplicar_semaforo(styler, subset)


def agregados_por_gestor(df_filtrado: pd.DataFrame, dias_en_mes: int, dia_corte: int) -> pd.DataFrame:
    """Para habilitadores con PDV: agrega Cuota y Avance por gestor + producto
    (sumando todos sus PDV) y recalcula Cumplimiento % y Proy Unidades sobre
    esos totales. Es el "resumen" que se ve dentro de cada fila expandible."""
    resumen = (
        df_filtrado.groupby(["DNI", "Nombre", "Departamento", "Distrito", "Producto"], as_index=False)
        .agg(Cuota=("Cuota", "sum"), Avance=("Avance", "sum"))
    )
    resumen["Cumplimiento %"] = np.where(resumen["Cuota"] > 0, resumen["Avance"] / resumen["Cuota"], 0.0)
    factor_proyeccion = dias_en_mes / max(dia_corte, 1)
    resumen["Proy Unidades"] = resumen["Avance"] * factor_proyeccion
    return resumen


def detalle_pdv_por_gestor(df_filtrado: pd.DataFrame, dias_en_mes: int, dia_corte: int) -> pd.DataFrame:
    """Detalle a nivel PDV (una fila por Producto + PDV) con sus propias
    métricas, sin agregar. Es lo que se ve al desplegar un gestor."""
    detalle = df_filtrado[["DNI", "Nombre", "Departamento", "Distrito", "Producto", "PDV", "Cuota", "Avance"]].copy()
    detalle["Cumplimiento %"] = np.where(detalle["Cuota"] > 0, detalle["Avance"] / detalle["Cuota"], 0.0)
    factor_proyeccion = dias_en_mes / max(dia_corte, 1)
    detalle["Proy Unidades"] = detalle["Avance"] * factor_proyeccion
    return detalle


def ritmo_diario_por_gestor(df_filtrado: pd.DataFrame, productos_sel: list,
                             dias_restantes: int) -> pd.DataFrame:
    """Arma el detalle por gestor (DNI / Nombre en filas) con los productos
    como columnas agrupadas: debajo de cada producto van Cuota Diaria
    Necesaria, Corte (avance registrado a la fecha de corte) y Cump %.

    Cuota y Avance se agregan primero por gestor + producto (sumando todos
    sus PDV si aplica) antes de calcular los indicadores, para no perder
    información cuando un gestor tiene varios PDV para el mismo producto."""
    agregado = (
        df_filtrado.groupby(["DNI", "Nombre", "Producto"], as_index=False)
        .agg(Cuota=("Cuota", "sum"), Avance=("Avance", "sum"))
    )
    agregado["Cump %"] = np.where(agregado["Cuota"] > 0, agregado["Avance"] / agregado["Cuota"], 0.0)
    agregado["Corte"] = agregado["Avance"]

    if dias_restantes > 0:
        agregado["Cuota Diaria"] = (agregado["Cuota"] - agregado["Avance"]) / dias_restantes
    else:
        agregado["Cuota Diaria"] = np.nan

    orden_prod = [p for p in PRODUCTOS if p in productos_sel]
    metricas = ["Cuota Diaria", "Corte", "Cump %"]

    ancho = agregado.pivot_table(index=["DNI", "Nombre"], columns="Producto", values=metricas, aggfunc="first")
    ancho = ancho.swaplevel(axis=1)
    columnas_orden = pd.MultiIndex.from_product([orden_prod, metricas])
    ancho = ancho.reindex(columns=columnas_orden)
    ancho = ancho.reset_index().sort_values(["Nombre", "DNI"]).set_index(["DNI", "Nombre"])

    return ancho


def aplicar_estilo_ritmo_gestor(tabla: pd.DataFrame, orden_prod: list):
    """Aplica formato numérico y semáforo (Cump %) a la tabla de ritmo
    diario por gestor."""
    fmt = {}
    for p in orden_prod:
        fmt[(p, "Cuota Diaria")] = "{:,.1f}"
        fmt[(p, "Corte")] = "{:,.0f}"
        fmt[(p, "Cump %")] = "{:.1%}"

    styler = tabla.style.format(fmt, na_rep="-")
    subset = [(p, "Cump %") for p in orden_prod]
    return _aplicar_semaforo(styler, subset)


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
        return "admin", "admin2025"


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

    st.download_button(
        "📥 Descargar plantilla Excel",
        data=generar_plantilla_excel(),
        file_name="plantilla_reporte_habilitadores_fanero.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    st.caption(
        "Elimina las filas de ejemplo antes de subir tu archivo real. Cada "
        "gestor debe tener una fila por producto (Prepago, Porta Flex, "
        "Postpago, OSS). Para Activaciones y Desarrolladores, agrega una "
        "fila por cada PDV (columna PDV); para los demás habilitadores deja "
        "esa columna vacía."
    )

    with st.expander("Columnas del Excel"):
        st.write("Obligatorias:", sorted(COLUMNAS_REQUERIDAS))
        st.write("Opcional: PDV (solo para Activaciones y Desarrolladores)")

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
# 5. PLANTILLAS EXCEL DESCARGABLES
# =============================================================================

def generar_plantilla_excel() -> bytes:
    """Genera en memoria la plantilla completa (para el admin) con los
    encabezados requeridos, filas de ejemplo (incluyendo un caso con PDV) y
    listas desplegables (validación de datos) para Departamento, Habilitador
    y Producto."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Datos"

    headers = ["DNI", "Nombre", "Departamento", "Distrito", "Habilitador", "Producto", "PDV", "Cuota", "Avance"]
    ws.append(headers)

    ejemplos = [
        # Habilitador sin PDV: la columna PDV queda vacía
        ["12345678", "Juan Pérez", "Amazonas", "Chachapoyas", "PDV Plus", "Prepago", "", 300, 150],
        ["12345678", "Juan Pérez", "Amazonas", "Chachapoyas", "PDV Plus", "Postpago", "", 100, 40],
        # Habilitador con PDV: mismo gestor y producto, repartido en 2 PDV
        ["87654321", "Rosa Huamán", "San Martín", "Tarapoto", "Activaciones", "Prepago", "PDV Norte", 180, 150],
        ["87654321", "Rosa Huamán", "San Martín", "Tarapoto", "Activaciones", "Prepago", "PDV Sur", 120, 95],
    ]
    for fila in ejemplos:
        ws.append(fila)

    ultima_fila = 1000
    dv_departamento = DataValidation(type="list", formula1='"' + ",".join(DEPARTAMENTOS) + '"', allow_blank=True)
    dv_habilitador = DataValidation(type="list", formula1='"' + ",".join(HABILITADORES) + '"', allow_blank=True)
    dv_producto = DataValidation(type="list", formula1='"' + ",".join(PRODUCTOS) + '"', allow_blank=True)

    ws.add_data_validation(dv_departamento)
    ws.add_data_validation(dv_habilitador)
    ws.add_data_validation(dv_producto)

    dv_departamento.add(f"C2:C{ultima_fila}")
    dv_habilitador.add(f"E2:E{ultima_fila}")
    dv_producto.add(f"F2:F{ultima_fila}")

    anchos = {"A": 12, "B": 22, "C": 16, "D": 18, "E": 16, "F": 12, "G": 16, "H": 10, "I": 10}
    for col, ancho in anchos.items():
        ws.column_dimensions[col].width = ancho

    buffer = BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


def dataframe_a_excel_bytes(df: pd.DataFrame, hoja: str = "Avances") -> bytes:
    """Convierte un DataFrame a bytes de un archivo .xlsx (para botones de
    descarga que arman una plantilla a partir de datos ya existentes)."""
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=hoja)
    return buffer.getvalue()


# =============================================================================
# 6. EDICIÓN DE AVANCES POR COORDINADOR (ACCESO RESTRINGIDO)
# =============================================================================
#
# Pestaña oculta: solo aparece con ?editar=1 en la URL. Cada coordinador
# elige su nombre en una lista desplegable y solo puede editar el Avance de
# los departamentos que tiene asignados (ver diccionario COORDINADORES). La
# selección de nombre no lleva contraseña: es un control de confianza para
# uso interno, no una autenticación real. Pueden editar en línea (tabla
# editable) o descargando su plantilla, editándola en Excel y subiéndola de
# vuelta; en ambos casos solo se aplican cambios a sus DNI/PDV/Producto
# permitidos, sin importar qué traiga el archivo subido.

def registrar_ultima_edicion(nombre: str) -> None:
    """Guarda quién fue la última persona en actualizar avances y cuándo."""
    os.makedirs(DATA_DIR, exist_ok=True)
    with open(LOG_EDICION, "w", encoding="utf-8") as f:
        json.dump({"nombre": nombre, "timestamp": datetime.now().isoformat()}, f)


def obtener_ultima_edicion() -> dict | None:
    """Devuelve {'nombre':..., 'timestamp':...} de la última edición, o None
    si todavía no se registró ninguna."""
    if os.path.exists(LOG_EDICION):
        try:
            with open(LOG_EDICION, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:  # noqa: BLE001 - archivo corrupto o incompleto
            return None
    return None


def actualizar_avances(cambios: pd.DataFrame, coordinador: str) -> None:
    """Aplica los nuevos valores de Avance (columnas DNI, PDV, Producto,
    Avance) sobre el dataset publicado completo -no solo el subconjunto
    visible del coordinador- y vuelve a publicar, preservando el día de
    corte, mes y año vigentes. También registra quién hizo el cambio y
    cuándo."""
    df_actual, dia_corte, mes, anio = obtener_datos_publicados()

    # DNI puede llegar como texto (tabla editable o Excel subido) o como
    # número (al releer el Excel publicado, pandas suele inferirlo como
    # entero). Se normaliza a texto en ambos lados, junto con PDV, para que
    # el cruce DNI + PDV + Producto siempre encuentre las filas correctas.
    df_actual = df_actual.copy()
    df_actual["DNI"] = df_actual["DNI"].astype(str).str.strip()
    df_actual = _normalizar_pdv(df_actual)

    cambios = cambios.copy()
    cambios["DNI"] = cambios["DNI"].astype(str).str.strip()
    cambios = _normalizar_pdv(cambios)

    df_actual = df_actual.set_index(["DNI", "PDV", "Producto"])
    cambios_idx = cambios.set_index(["DNI", "PDV", "Producto"])["Avance"]
    indices_validos = cambios_idx.index.intersection(df_actual.index)
    df_actual.loc[indices_validos, "Avance"] = cambios_idx.loc[indices_validos]
    df_actual = df_actual.reset_index()

    publicar_datos(df_actual, dia_corte, mes, anio)
    registrar_ultima_edicion(coordinador)


def panel_editar_avances(df_raw: pd.DataFrame) -> None:
    """Renderiza la pestaña de edición de avances. `df_raw` es el dataset
    publicado completo (sin filtrar por Habilitador), para que cada
    coordinador vea y edite todos sus registros sin importar el filtro de
    Habilitador que esté activo arriba en el dashboard."""
    ultima = obtener_ultima_edicion()
    if ultima:
        ts = datetime.fromisoformat(ultima["timestamp"])
        st.caption(f"Última conexión: {ultima['nombre']} · {ts:%d/%m/%Y %H:%M}")
    else:
        st.caption("Todavía no se registró ninguna edición de avances.")

    coordinador_sel = st.selectbox("Persona que cargará", sorted(COORDINADORES.keys()), key="coordinador_editor")
    departamentos_permitidos = COORDINADORES[coordinador_sel]
    st.caption(f"{coordinador_sel} solo puede editar: {', '.join(departamentos_permitidos)}")

    df_editable = (
        df_raw[df_raw["Departamento"].isin(departamentos_permitidos)]
        [["DNI", "Nombre", "Departamento", "Distrito", "Habilitador", "Producto", "PDV", "Cuota", "Avance"]]
        .sort_values(["Departamento", "Nombre", "Producto", "PDV"])
        .reset_index(drop=True)
    )
    df_editable["DNI"] = df_editable["DNI"].astype(str).str.strip()
    df_editable = _normalizar_pdv(df_editable)

    if df_editable.empty:
        st.info("No hay registros para los departamentos asignados a este coordinador.")
        return

    claves_permitidas = set(zip(df_editable["DNI"], df_editable["PDV"], df_editable["Producto"]))

    st.markdown("#### Opción 1 · Editar directamente en la tabla")
    columnas_bloqueadas = [c for c in df_editable.columns if c != "Avance"]
    editado = st.data_editor(
        df_editable,
        disabled=columnas_bloqueadas,
        hide_index=True,
        use_container_width=True,
        height=400,
        key=f"editor_avances_{coordinador_sel}",
    )

    if st.button("Guardar cambios de la tabla", key="guardar_avances_tabla"):
        actualizar_avances(editado[["DNI", "PDV", "Producto", "Avance"]], coordinador_sel)
        st.success(f"Avances actualizados por {coordinador_sel}.")
        st.rerun()

    st.markdown("---")
    st.markdown("#### Opción 2 · Descargar plantilla, editar en Excel y volver a subir")

    st.download_button(
        "📥 Descargar mis datos actuales (Excel)",
        data=dataframe_a_excel_bytes(df_editable),
        file_name=f"avances_{coordinador_sel.replace(' ', '_').lower()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="descargar_plantilla_avances",
    )
    st.caption("Solo modifica la columna Avance. No cambies DNI, Departamento, PDV ni Producto.")

    archivo_avances = st.file_uploader(
        "Subir Excel con avances actualizados", type=["xlsx"], key=f"subir_avances_{coordinador_sel}"
    )

    if archivo_avances is not None and st.button("Guardar cambios del archivo", key="guardar_avances_archivo"):
        try:
            df_subido = pd.read_excel(archivo_avances)
        except Exception as exc:  # noqa: BLE001 - se informa al usuario cualquier error de lectura
            st.error(f"No se pudo leer el archivo: {exc}")
            df_subido = None

        if df_subido is not None:
            faltantes = {"DNI", "Producto", "Avance"} - set(df_subido.columns)
            if faltantes:
                st.error("Al archivo le faltan columnas: " + ", ".join(sorted(faltantes)))
            else:
                df_subido = df_subido.copy()
                df_subido["DNI"] = df_subido["DNI"].astype(str).str.strip()
                df_subido = _normalizar_pdv(df_subido)
                mascara_permitida = df_subido.apply(
                    lambda fila: (fila["DNI"], fila["PDV"], fila["Producto"]) in claves_permitidas, axis=1
                )
                df_filtrado_subida = df_subido[mascara_permitida]
                ignoradas = len(df_subido) - len(df_filtrado_subida)

                if df_filtrado_subida.empty:
                    st.error(
                        f"Ninguna fila del archivo corresponde a los departamentos "
                        f"permitidos para {coordinador_sel} ({', '.join(departamentos_permitidos)})."
                    )
                else:
                    actualizar_avances(df_filtrado_subida[["DNI", "PDV", "Producto", "Avance"]], coordinador_sel)
                    if ignoradas > 0:
                        st.warning(
                            f"Se ignoraron {ignoradas} fila(s) que no pertenecen a los "
                            "departamentos permitidos para este coordinador."
                        )
                    st.success(f"Avances actualizados por {coordinador_sel} desde archivo.")
                    st.rerun()


# =============================================================================
# 7. INTERFAZ PRINCIPAL
# =============================================================================

def main():
    st.title("📊 Reporte Habilitadores Fanero")

    # Los paneles ocultos solo se renderizan con sus parámetros en la URL.
    # Sin ellos no queda ningún rastro visible para gerencia.
    if st.query_params.get("admin") == "1":
        with st.sidebar:
            panel_admin()
    mostrar_editor = st.query_params.get("editar") == "1"

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

    orden_prod_sel = [p for p in PRODUCTOS if p in productos_sel]

    df_filtrado = df[
        (df["Habilitador"] == habilitador_sel)
        & (df["Departamento"].isin(departamentos_sel))
        & (df["Producto"].isin(productos_sel))
    ]

    st.markdown("---")

    # --- Resumen por producto (productos como columnas agrupadas) ---
    st.subheader("Resumen por Producto")
    if df_filtrado.empty:
        st.info("No hay datos para los filtros seleccionados.")
    else:
        tabla_resumen = resumen_por_producto(
            df_filtrado, departamentos_sel, productos_sel, dias_en_mes, dia_corte
        )
        st.dataframe(
            aplicar_estilo_resumen_producto(tabla_resumen, orden_prod_sel),
            use_container_width=True,
        )
        st.caption("🟥 <80% · 🟨 80%–99% · 🟩 ≥100% (aplica a Cumplimiento % y Proy %)")

    st.markdown("---")

    # --- Pestañas ---
    tab_names = ["Detalle Habilitador", "Ritmo Diario"]
    if mostrar_editor:
        tab_names.append("Editar Avances")
    tabs = st.tabs(tab_names)
    tab_detalle, tab_ritmo = tabs[0], tabs[1]

    with tab_detalle:
        st.subheader(f"Detalle habilitador. {habilitador_sel}")

        if df_filtrado.empty:
            st.info("No hay registros para los filtros seleccionados.")
        elif habilitador_sel in HABILITADORES_CON_PDV:
            # --- Vista con Puntos de Venta: una fila expandible por gestor ---
            st.caption(
                "Este habilitador tiene Puntos de Venta (PDV) por debajo de cada "
                "gestor. Haz clic en un nombre para ver su resumen y el detalle "
                "por PDV."
            )

            resumen_gestor = agregados_por_gestor(df_filtrado, dias_en_mes, dia_corte)
            detalle_pdv = detalle_pdv_por_gestor(df_filtrado, dias_en_mes, dia_corte)

            gestores = (
                resumen_gestor[["DNI", "Nombre", "Departamento"]]
                .drop_duplicates()
                .sort_values(["Departamento", "Nombre"])
                .reset_index(drop=True)
            )
            st.caption(f"{len(gestores)} gestor(es) en {habilitador_sel}.")

            for _, gestor in gestores.iterrows():
                etiqueta = f"👤 {gestor['Nombre']} · DNI {gestor['DNI']} · {gestor['Departamento']}"
                with st.expander(etiqueta):
                    resumen_g = (
                        resumen_gestor[resumen_gestor["DNI"] == gestor["DNI"]]
                        [["Producto", "Cuota", "Avance", "Cumplimiento %", "Proy Unidades"]]
                        .sort_values("Producto")
                        .reset_index(drop=True)
                    )
                    st.markdown("**Resumen (suma de sus PDV)**")
                    st.dataframe(aplicar_estilo_tabla(resumen_g), use_container_width=True, hide_index=True)

                    detalle_g = (
                        detalle_pdv[detalle_pdv["DNI"] == gestor["DNI"]]
                        [["Producto", "PDV", "Cuota", "Avance", "Cumplimiento %", "Proy Unidades"]]
                        .sort_values(["Producto", "PDV"])
                        .reset_index(drop=True)
                    )
                    st.markdown("**Detalle por PDV**")
                    st.dataframe(aplicar_estilo_tabla(detalle_g), use_container_width=True, hide_index=True)

            columnas_pdv_csv = [
                "DNI", "Nombre", "Departamento", "Distrito", "Producto", "PDV",
                "Cuota", "Avance", "Cumplimiento %", "Proy Unidades",
            ]
            st.download_button(
                "⬇️ Descargar detalle por PDV (CSV)",
                data=detalle_pdv[columnas_pdv_csv].sort_values(["Departamento", "Nombre", "Producto", "PDV"])
                .to_csv(index=False).encode("utf-8"),
                file_name=f"detalle_pdv_{habilitador_sel.replace(' ', '_').lower()}.csv",
                mime="text/csv",
            )
            st.caption("🟥 <80% · 🟨 80%–99% · 🟩 ≥100%")
        else:
            # --- Vista plana habitual (sin PDV) ---
            columnas_mostrar = [
                "DNI", "Nombre", "Departamento", "Distrito", "Producto",
                "Cuota", "Avance", "Cumplimiento %", "Proy Unidades",
            ]
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
            f"(día {dia_corte} → día {dias_en_mes}). Detalle por gestor, "
            "con los productos como columnas agrupadas (suma de PDV si aplica)."
        )

        if df_filtrado.empty:
            st.info("No hay datos para los filtros seleccionados.")
        else:
            if dias_restantes == 0:
                st.warning("El mes ya cerró; no quedan días para calcular el ritmo diario.")

            tabla_ritmo = ritmo_diario_por_gestor(df_filtrado, productos_sel, dias_restantes)

            st.dataframe(
                aplicar_estilo_ritmo_gestor(tabla_ritmo, orden_prod_sel),
                use_container_width=True,
                height=500,
            )

            tabla_ritmo_csv = tabla_ritmo.copy()
            tabla_ritmo_csv.columns = [f"{p} - {m}" for p, m in tabla_ritmo_csv.columns]
            st.download_button(
                "⬇️ Descargar ritmo diario (CSV)",
                data=tabla_ritmo_csv.reset_index().to_csv(index=False).encode("utf-8"),
                file_name=f"ritmo_diario_{habilitador_sel.replace(' ', '_').lower()}.csv",
                mime="text/csv",
            )

            st.caption(
                "Cuota Diaria = (Cuota - Avance) / días restantes · Corte = avance a la fecha de corte · "
                "🟥 <80% · 🟨 80%–99% · 🟩 ≥100% (Cump %)"
            )

    if mostrar_editor:
        with tabs[2]:
            st.subheader("Editar Avances")
            panel_editar_avances(df_raw)


if __name__ == "__main__":
    main()
