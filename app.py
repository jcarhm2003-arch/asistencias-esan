"""
Control de Asistencia Académica - Streamlit App
=================================================
Compara listas de alumnos contra registros de asistencia de Zoom,
clasifica Asistencia / Tardanza / Falta, mantiene un historial acumulativo
por sesión, permite corrección manual con memoria persistente de
equivalencias de nombres, y exporta el resultado a Excel.
"""

import io
import json
import re
import unicodedata
from datetime import datetime, date

import pandas as pd
import pytz
import streamlit as st
from rapidfuzz import fuzz, process

# ---------------------------------------------------------------------------
# Configuración general de la página
# ---------------------------------------------------------------------------
st.set_page_config(page_title="Control de Asistencia Académica", page_icon="✅", layout="wide")

ZONA_MAP = {
    "Perú (UTC-5)": "America/Lima",
    "Estados Unidos (AM/PM)": "America/New_York",
    "UTC ISO": "UTC",
}
LIMA_TZ = pytz.timezone("America/Lima")


# ---------------------------------------------------------------------------
# FUNCIONES MODULARES
# ---------------------------------------------------------------------------

def normalizar_nombre(nombre: str) -> str:
    """Normaliza un nombre: minúsculas, sin tildes, sin espacios repetidos."""
    if not isinstance(nombre, str):
        return ""
    nombre = nombre.strip().lower()
    nombre = unicodedata.normalize("NFKD", nombre)
    nombre = "".join(c for c in nombre if not unicodedata.combining(c))
    nombre = re.sub(r"\s+", " ", nombre).strip()
    return nombre


def extraer_codigo(correo: str):
    """Extrae el código numérico inicial de un correo institucional (ej: 20100729@esan.edu.pe -> 20100729)."""
    if not isinstance(correo, str) or "@" not in correo:
        return None
    local = correo.split("@")[0].strip()
    match = re.match(r"^(\d+)", local)
    return match.group(1) if match else None


def convertir_hora(hora_str: str, zona_seleccionada: str, fecha_referencia: date = None):
    """
    Convierte un string de hora en formato HH:MM (24h), h:mm AM/PM o ISO UTC
    a un objeto datetime.time expresado en hora de Perú (America/Lima),
    que es la base usada para comparar contra los límites configurados.
    """
    if not isinstance(hora_str, str) or not hora_str.strip():
        raise ValueError("hora vacía o inválida")
    hora_str = hora_str.strip()
    fecha_referencia = fecha_referencia or datetime.now().date()

    # a) Formato ISO UTC (ej: 2026-06-30T09:27:00Z)
    if "T" in hora_str and (hora_str.upper().endswith("Z") or "+" in hora_str):
        iso_str = hora_str.replace("Z", "+00:00").replace("z", "+00:00")
        try:
            dt = datetime.fromisoformat(iso_str)
        except ValueError:
            raise ValueError(f"formato ISO UTC no reconocido: '{hora_str}'")
        if dt.tzinfo is None:
            dt = pytz.utc.localize(dt)
        return dt.astimezone(LIMA_TZ).time()

    # b) Formato AM/PM (ej: 9:15 PM)
    if "am" in hora_str.lower() or "pm" in hora_str.lower():
        limpio = hora_str.upper().replace(".", "").strip()
        parsed = None
        for fmt in ("%I:%M %p", "%I:%M%p", "%I %p"):
            try:
                parsed = datetime.strptime(limpio, fmt)
                break
            except ValueError:
                continue
        if parsed is None:
            raise ValueError(f"formato AM/PM no reconocido: '{hora_str}'")
        origen_tz = pytz.timezone(ZONA_MAP.get(zona_seleccionada, "America/New_York"))
        dt_naive = datetime.combine(fecha_referencia, parsed.time())
        dt_local = origen_tz.localize(dt_naive)
        return dt_local.astimezone(LIMA_TZ).time()

    # c) Formato HH:MM 24 horas (se asume ya en hora de Perú)
    for fmt in ("%H:%M:%S", "%H:%M"):
        try:
            return datetime.strptime(hora_str, fmt).time()
        except ValueError:
            continue

    raise ValueError(f"formato de hora no reconocido: '{hora_str}'")


def clasificar_asistencia(hora, config: dict) -> str:
    """Clasifica una hora ya convertida como 'A' (asistencia), 'T' (tardanza) o 'F' (falta)."""
    if not config.get("control_hora_activo"):
        return "A"

    hora_asistencia = config.get("hora_limite_asistencia")
    if hora_asistencia is not None and hora <= hora_asistencia:
        return "A"

    if config.get("marcar_tardanza_fuera"):
        hora_tardanza = config.get("hora_limite_tardanza")
        if hora_tardanza is not None and hora <= hora_tardanza:
            return "T"
        return "F"

    return "T"


def buscar_coincidencia(nombre_zoom: str, codigo_zoom, alumnos_df: pd.DataFrame, memoria: dict, umbral: int = 80):
    """
    Busca el alumno correspondiente a un registro de Zoom.
    Prioridad: 1) código exacto  2) memoria guardada  3) similaridad de nombre (RapidFuzz).
    Retorna: (codigo_alumno, nombre_alumno, tipo_coincidencia, score)
    """
    nombre_zoom_norm = normalizar_nombre(nombre_zoom)

    # 1. Código exacto
    if codigo_zoom:
        coincide = alumnos_df[alumnos_df["Código"].astype(str).str.strip() == str(codigo_zoom).strip()]
        coincide = coincide[coincide["Código"].astype(str).str.strip() != ""]
        if not coincide.empty:
            fila = coincide.iloc[0]
            return fila["Código"], fila["Alumno"], "CÓDIGO", 100

    # 2. Memoria guardada (equivalencias previas nombre_zoom -> nombre_real)
    for zoom_key, nombre_real in memoria.items():
        if normalizar_nombre(zoom_key) == nombre_zoom_norm:
            coincide = alumnos_df[alumnos_df["Alumno"] == nombre_real]
            if not coincide.empty:
                fila = coincide.iloc[0]
                return fila["Código"], fila["Alumno"], "MEMORIA", 100

    # 3. Similaridad por nombre
    if not alumnos_df.empty:
        alumnos_norm = {row["Alumno"]: normalizar_nombre(row["Alumno"]) for _, row in alumnos_df.iterrows()}
        opciones = list(alumnos_norm.values())
        if opciones:
            resultado = process.extractOne(nombre_zoom_norm, opciones, scorer=fuzz.token_sort_ratio)
            if resultado and resultado[1] >= umbral:
                nombre_match_norm = resultado[0]
                for nombre_real, norm in alumnos_norm.items():
                    if norm == nombre_match_norm:
                        fila = alumnos_df[alumnos_df["Alumno"] == nombre_real].iloc[0]
                        return fila["Código"], fila["Alumno"], "SIMILARIDAD", round(resultado[1], 1)

    return None, None, "SIN_COINCIDENCIA", 0


def procesar_zoom(texto_zoom: str, alumnos_df: pd.DataFrame, memoria: dict, config: dict, zona_seleccionada: str):
    """
    Parsea el texto pegado de Zoom (Nombre<TAB>Correo<TAB>Hora), ignora el encabezado,
    elimina duplicados (conservando la hora más temprana y contando ingresos),
    y busca coincidencia contra la tabla de alumnos.

    Retorna: (registros_zoom: dict, errores: list)
    registros_zoom key -> {nombre, correo, codigo, hora_min, ingresos, codigo_alumno, nombre_alumno, tipo, score}
    """
    lineas = [l for l in texto_zoom.strip().split("\n") if l.strip()]
    if len(lineas) <= 1:
        raise ValueError("No hay datos de Zoom para procesar (falta encabezado o filas de datos).")
    lineas = lineas[1:]  # ignorar encabezado

    crudos = {}
    errores = []

    for i, linea in enumerate(lineas, start=2):
        partes = linea.split("\t")
        if len(partes) < 3:
            errores.append(f"Fila {i}: se esperaban 3 columnas separadas por tabulación (nombre, correo, hora).")
            continue
        nombre, correo, hora_str = partes[0].strip(), partes[1].strip(), partes[2].strip()
        if not nombre:
            errores.append(f"Fila {i}: nombre vacío, fila ignorada.")
            continue
        codigo = extraer_codigo(correo)
        try:
            hora = convertir_hora(hora_str, zona_seleccionada)
        except ValueError as e:
            errores.append(f"Fila {i} ({nombre}): {e}")
            continue

        key = (normalizar_nombre(nombre), codigo or "")
        if key not in crudos:
            crudos[key] = {"nombre": nombre, "correo": correo, "codigo": codigo, "horas": [hora], "ingresos": 1}
        else:
            crudos[key]["horas"].append(hora)
            crudos[key]["ingresos"] += 1

    registros_zoom = {}
    for key, datos in crudos.items():
        hora_min = min(datos["horas"])
        codigo_al, nombre_al, tipo, score = buscar_coincidencia(datos["nombre"], datos["codigo"], alumnos_df, memoria)
        registros_zoom[datos["nombre"]] = {
            "nombre": datos["nombre"],
            "correo": datos["correo"],
            "codigo": datos["codigo"],
            "hora_min": hora_min,
            "ingresos": datos["ingresos"],
            "codigo_alumno": codigo_al,
            "nombre_alumno": nombre_al,
            "tipo": tipo,
            "score": score,
        }

    return registros_zoom, errores


def construir_resultados(alumnos_df: pd.DataFrame, registros_zoom: dict, config: dict):
    """
    Construye la tabla de resultados con una fila por alumno de la tabla oficial,
    cruzando contra los registros de Zoom ya matcheados.
    Retorna: (resultados_df, zoom_no_usados: dict de registros de zoom sin alumno asignado)
    """
    usados = set()
    filas = []
    for _, alumno in alumnos_df.iterrows():
        codigo_a, nombre_a = alumno["Código"], alumno["Alumno"]
        match = None
        for nombre_zoom, datos in registros_zoom.items():
            if datos["nombre_alumno"] == nombre_a and nombre_zoom not in usados:
                match = (nombre_zoom, datos)
                break
        if match:
            nombre_zoom, datos = match
            usados.add(nombre_zoom)
            estado = clasificar_asistencia(datos["hora_min"], config)
            filas.append({
                "Código": codigo_a,
                "Alumno": nombre_a,
                "Nombre_Zoom": nombre_zoom,
                "Hora": datos["hora_min"].strftime("%H:%M:%S"),
                "Ingresos": datos["ingresos"],
                "Tipo_Coincidencia": datos["tipo"],
                "Confianza": datos["score"],
                "Estado": estado,
            })
        else:
            filas.append({
                "Código": codigo_a,
                "Alumno": nombre_a,
                "Nombre_Zoom": None,
                "Hora": None,
                "Ingresos": 0,
                "Tipo_Coincidencia": "SIN_COINCIDENCIA",
                "Confianza": 0,
                "Estado": "SIN_COINCIDENCIA",
            })

    zoom_no_usados = {k: v for k, v in registros_zoom.items() if k not in usados}
    resultados_df = pd.DataFrame(filas)
    return resultados_df, zoom_no_usados


def exportar_excel(historial_df: pd.DataFrame) -> io.BytesIO:
    """Genera un archivo Excel en memoria a partir del historial acumulado."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        historial_df.to_excel(writer, index=False, sheet_name="Asistencia")
    output.seek(0)
    return output


# ---------------------------------------------------------------------------
# INICIALIZACIÓN DE ESTADO
# ---------------------------------------------------------------------------

def inicializar_estado():
    defaults = {
        "alumnos_df": pd.DataFrame({"Código": pd.Series(dtype="str"), "Alumno": pd.Series(dtype="str")}),
        "memoria": {},
        "historial_df": pd.DataFrame({"Código": pd.Series(dtype="str"), "Alumno": pd.Series(dtype="str")}),
        "sesion_actual": 0,
        "resultados_df": None,
        "zoom_no_usados": {},
        "errores_ultimo_proceso": [],
        "correccion_abierta": None,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


inicializar_estado()


# ---------------------------------------------------------------------------
# SIDEBAR: CONFIGURACIÓN
# ---------------------------------------------------------------------------

with st.sidebar:
    st.header("⚙️ Configuración")

    control_hora_activo = st.checkbox("Activar control por hora", value=True)

    hora_limite_asistencia = None
    hora_limite_tardanza = None
    marcar_tardanza_fuera = False

    if control_hora_activo:
        hora_limite_asistencia = st.time_input("Hora límite asistencia", value=datetime.strptime("08:15", "%H:%M").time())
        marcar_tardanza_fuera = st.checkbox("Marcar tardanza fuera de asistencia", value=True)
        if marcar_tardanza_fuera:
            hora_limite_tardanza = st.time_input("Hora límite tardanza (después de esto = Falta)", value=datetime.strptime("08:30", "%H:%M").time())
        st.caption("Sin 'marcar tardanza fuera de asistencia', todo lo posterior a la hora límite se marca como Tardanza (nunca Falta).")
    else:
        st.caption("Control por hora desactivado: toda coincidencia válida se marcará como Asistencia (A).")

    st.divider()
    st.subheader("🌍 Zona horaria")
    zona_seleccionada = st.selectbox("Formato de hora del reporte Zoom", list(ZONA_MAP.keys()), index=0)
    st.caption("Perú → HH:MM (24h) · Estados Unidos → h:mm AM/PM · UTC ISO → 2026-06-30T09:27:00Z")

    config = {
        "control_hora_activo": control_hora_activo,
        "hora_limite_asistencia": hora_limite_asistencia,
        "marcar_tardanza_fuera": marcar_tardanza_fuera,
        "hora_limite_tardanza": hora_limite_tardanza,
    }

    st.divider()
    st.subheader("🧠 Memoria de equivalencias")
    st.caption(f"{len(st.session_state['memoria'])} equivalencias guardadas.")

    memoria_json = json.dumps(st.session_state["memoria"], ensure_ascii=False, indent=2)
    st.download_button("⬇️ Exportar memoria (JSON)", data=memoria_json, file_name="memoria_asistencia.json", mime="application/json")

    archivo_memoria = st.file_uploader("⬆️ Importar memoria (JSON)", type=["json"], key="uploader_memoria")
    if archivo_memoria is not None:
        try:
            nueva_memoria = json.loads(archivo_memoria.read().decode("utf-8"))
            if isinstance(nueva_memoria, dict):
                st.session_state["memoria"].update(nueva_memoria)
                st.success(f"Memoria importada: {len(nueva_memoria)} equivalencias.")
            else:
                st.error("El JSON debe ser un objeto {nombre_zoom: nombre_real}.")
        except Exception as e:
            st.error(f"No se pudo leer el archivo: {e}")

    st.divider()
    if st.button("🗑️ Nuevo curso / Reiniciar", use_container_width=True):
        st.session_state["alumnos_df"] = pd.DataFrame({"Código": pd.Series(dtype="str"), "Alumno": pd.Series(dtype="str")})
        st.session_state["historial_df"] = pd.DataFrame({"Código": pd.Series(dtype="str"), "Alumno": pd.Series(dtype="str")})
        st.session_state["sesion_actual"] = 0
        st.session_state["resultados_df"] = None
        st.session_state["zoom_no_usados"] = {}
        st.session_state["errores_ultimo_proceso"] = []
        st.session_state["correccion_abierta"] = None
        st.success("Curso reiniciado. La memoria de equivalencias se conservó.")
        st.rerun()


# ---------------------------------------------------------------------------
# MAIN: TÍTULO
# ---------------------------------------------------------------------------

st.title("✅ Control de Asistencia Académica")
st.caption("Compara listas de alumnos contra reportes de Zoom, clasifica A/T/F y mantiene historial por sesión.")

tab_alumnos, tab_zoom, tab_resultados, tab_historial = st.tabs(
    ["👥 Tabla de alumnos", "📋 Pegado de Zoom", "📊 Resultados", "🗂️ Historial y exportación"]
)

# ---------------------------------------------------------------------------
# B. TABLA DE ALUMNOS
# ---------------------------------------------------------------------------
with tab_alumnos:
    st.subheader("Tabla de alumnos")
    st.caption("Edita manualmente, pega desde Excel (Ctrl+V dentro de la tabla) o agrega filas. El código puede quedar vacío.")

    edited = st.data_editor(
        st.session_state["alumnos_df"],
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "Código": st.column_config.TextColumn("Código"),
            "Alumno": st.column_config.TextColumn("Alumno"),
        },
        key="editor_alumnos",
    )
    st.session_state["alumnos_df"] = edited.fillna("")

# ---------------------------------------------------------------------------
# C. PEGADO DE ZOOM
# ---------------------------------------------------------------------------
with tab_zoom:
    st.subheader("Registro de asistencia de Zoom")
    st.caption("Pega las 3 columnas (Nombre, Correo, Hora) separadas por tabulación, con encabezado en la primera fila.")

    texto_zoom = st.text_area(
        "Pegar reporte de Zoom",
        height=280,
        placeholder="Nombre\tCorreo\tHora\nJuan Huaman\t20100729@esan.edu.pe\t08:03\nMaria Lopez\t20100730@esan.edu.pe\t2026-06-30T09:27:00Z",
    )

    col_a, col_b = st.columns([1, 3])
    with col_a:
        procesar = st.button("▶️ Procesar asistencia", type="primary", use_container_width=True)

    if procesar:
        alumnos_df = st.session_state["alumnos_df"]
        alumnos_df = alumnos_df[alumnos_df["Alumno"].astype(str).str.strip() != ""]
        if alumnos_df.empty:
            st.error("Primero agrega alumnos en la pestaña 'Tabla de alumnos'.")
        elif control_hora_activo and hora_limite_asistencia is None:
            st.error("Configura la hora límite de asistencia en la barra lateral.")
        else:
            try:
                registros_zoom, errores = procesar_zoom(
                    texto_zoom, alumnos_df, st.session_state["memoria"], config, zona_seleccionada
                )
                resultados_df, zoom_no_usados = construir_resultados(alumnos_df, registros_zoom, config)
                st.session_state["resultados_df"] = resultados_df
                st.session_state["zoom_no_usados"] = zoom_no_usados
                st.session_state["errores_ultimo_proceso"] = errores
                st.success(f"Procesado: {len(resultados_df)} alumnos evaluados, {len(zoom_no_usados)} registros de Zoom sin asignar.")
            except ValueError as e:
                st.error(str(e))

    if st.session_state["errores_ultimo_proceso"]:
        with st.expander(f"⚠️ {len(st.session_state['errores_ultimo_proceso'])} advertencias del último procesamiento"):
            for err in st.session_state["errores_ultimo_proceso"]:
                st.write(f"- {err}")

# ---------------------------------------------------------------------------
# D. RESULTADOS
# ---------------------------------------------------------------------------
with tab_resultados:
    st.subheader("Resultados de la sesión actual")

    resultados_df = st.session_state["resultados_df"]

    if resultados_df is None:
        st.info("Aún no se ha procesado ninguna sesión. Ve a 'Pegado de Zoom' y presiona 'Procesar asistencia'.")
    else:
        badge = {"A": "🟢", "T": "🟡", "F": "🔴", "SIN_COINCIDENCIA": "⚪"}
        resumen = resultados_df["Estado"].value_counts().to_dict()
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("🟢 Asistencia", resumen.get("A", 0))
        c2.metric("🟡 Tardanza", resumen.get("T", 0))
        c3.metric("🔴 Falta", resumen.get("F", 0))
        c4.metric("⚪ Sin coincidencia", resumen.get("SIN_COINCIDENCIA", 0))

        st.divider()

        for idx, row in resultados_df.iterrows():
            cols = st.columns([2.5, 1, 1, 1.2, 1.5, 1.2])
            cols[0].write(f"**{row['Alumno']}**  \n`{row['Código']}`")
            cols[1].write(f"{badge.get(row['Estado'], '⚪')} {row['Estado']}")
            cols[2].write(row["Hora"] or "—")
            cols[3].write(row["Tipo_Coincidencia"])
            cols[4].write(row["Nombre_Zoom"] or "—")

            if row["Estado"] == "SIN_COINCIDENCIA":
                if cols[5].button("Corregir", key=f"btn_corregir_{idx}"):
                    st.session_state["correccion_abierta"] = idx

            if st.session_state["correccion_abierta"] == idx:
                opciones = list(st.session_state["zoom_no_usados"].keys())
                if not opciones:
                    st.warning("No hay registros de Zoom disponibles para asignar.")
                else:
                    seleccion = st.selectbox(
                        f"Nombre en Zoom que corresponde a **{row['Alumno']}**",
                        opciones,
                        key=f"select_corr_{idx}",
                    )
                    cg1, cg2 = st.columns([1, 1])
                    if cg1.button("💾 Guardar equivalencia", key=f"guardar_corr_{idx}"):
                        st.session_state["memoria"][seleccion] = row["Alumno"]
                        datos = st.session_state["zoom_no_usados"].pop(seleccion)
                        estado = clasificar_asistencia(datos["hora_min"], config)
                        resultados_df.at[idx, "Nombre_Zoom"] = seleccion
                        resultados_df.at[idx, "Hora"] = datos["hora_min"].strftime("%H:%M:%S")
                        resultados_df.at[idx, "Ingresos"] = datos["ingresos"]
                        resultados_df.at[idx, "Tipo_Coincidencia"] = "MANUAL"
                        resultados_df.at[idx, "Confianza"] = 100
                        resultados_df.at[idx, "Estado"] = estado
                        st.session_state["resultados_df"] = resultados_df
                        st.session_state["correccion_abierta"] = None
                        st.success(f"'{seleccion}' asignado a {row['Alumno']} y guardado en memoria.")
                        st.rerun()
                    if cg2.button("Cancelar", key=f"cancelar_corr_{idx}"):
                        st.session_state["correccion_abierta"] = None
                        st.rerun()
            st.divider()

        st.subheader("💾 Guardar sesión en historial")
        st.caption("Los alumnos que sigan 'Sin coincidencia' se guardarán como Falta (F).")
        if st.button("Guardar sesión actual en el historial", type="primary"):
            st.session_state["sesion_actual"] += 1
            nueva_col = f"S{st.session_state['sesion_actual']}"

            historial_df = st.session_state["historial_df"].copy()
            for col_existente in historial_df.columns:
                if col_existente not in ("Código", "Alumno"):
                    pass

            valores = {}
            for _, row in resultados_df.iterrows():
                estado_final = row["Estado"] if row["Estado"] != "SIN_COINCIDENCIA" else "F"
                valores[row["Código"] if row["Código"] else row["Alumno"]] = (row["Alumno"], row["Código"], estado_final)

            for llave, (alumno, codigo, estado_final) in valores.items():
                match_mask = (historial_df["Alumno"] == alumno)
                if match_mask.any():
                    historial_df.loc[match_mask, nueva_col] = estado_final
                else:
                    nueva_fila = {c: "-" for c in historial_df.columns}
                    nueva_fila["Código"] = codigo
                    nueva_fila["Alumno"] = alumno
                    nueva_fila[nueva_col] = estado_final
                    historial_df = pd.concat([historial_df, pd.DataFrame([nueva_fila])], ignore_index=True)

            if nueva_col not in historial_df.columns:
                historial_df[nueva_col] = "-"
            historial_df[nueva_col] = historial_df[nueva_col].fillna("-")

            st.session_state["historial_df"] = historial_df
            st.success(f"Sesión {nueva_col} guardada en el historial.")
            st.rerun()

# ---------------------------------------------------------------------------
# E. HISTORIAL Y EXPORTACIÓN
# ---------------------------------------------------------------------------
with tab_historial:
    st.subheader("Historial acumulado de asistencia")

    historial_df = st.session_state["historial_df"]
    if historial_df.empty or len(historial_df.columns) <= 2:
        st.info("Todavía no hay sesiones guardadas en el historial.")
    else:
        st.dataframe(historial_df, use_container_width=True, hide_index=True)

        excel_bytes = exportar_excel(historial_df)
        st.download_button(
            "⬇️ Exportar historial a Excel",
            data=excel_bytes,
            file_name=f"asistencia_historial_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
