import streamlit as st
import pandas as pd
import json
import re
import io
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from unidecode import unidecode

# ─────────────────────────────────────────────
# CONFIGURACIÓN
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Control de Asistencias ESAN",
    page_icon="📋",
    layout="wide"
)

HISTORIAL_FILE = "historial_relaciones.json"

# ─────────────────────────────────────────────
# UTILIDADES
# ─────────────────────────────────────────────

def normalizar(texto):
    """Normaliza texto: sin tildes, sin paréntesis, minúsculas, sin puntos."""
    if not texto:
        return ""
    texto = str(texto)
    texto = re.sub(r'\(.*?\)', '', texto)       # quitar alias entre paréntesis
    texto = unidecode(texto)                    # quitar tildes
    texto = texto.lower()
    texto = re.sub(r'[^a-z\s]', '', texto)     # quitar puntos y símbolos
    return texto.strip()

def palabras(texto):
    """Devuelve set de palabras significativas (>= 2 letras)."""
    return set(w for w in normalizar(texto).split() if len(w) >= 2)

def coincidencias(nombre_zoom, nombre_excel):
    """Cuenta palabras en común entre nombre Zoom y nombre Excel."""
    p_zoom = palabras(nombre_zoom)
    p_excel = palabras(nombre_excel)
    return len(p_zoom & p_excel)

def extraer_codigo_correo(correo):
    """Extrae código de alumno de correo tipo 2607391@esan.edu.pe"""
    if pd.isna(correo) or not str(correo).strip():
        return None
    match = re.match(r'^(\d{6,7})@esan\.edu\.pe$', str(correo).strip())
    return match.group(1) if match else None

# ─────────────────────────────────────────────
# LECTURA DE ARCHIVOS
# ─────────────────────────────────────────────

def leer_excel_asistencia(archivo):
    """Lee el Excel maestro y retorna metadata + DataFrame de alumnos."""
    wb = load_workbook(archivo, read_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))

    metadata = {}
    header_row = None

    for i, row in enumerate(rows):
        if row[0] == 'Promoción':
            metadata['promocion'] = row[3]
        elif row[0] == 'Curso':
            metadata['curso'] = row[3]
        elif row[0] == 'Profesor(a)':
            metadata['profesor'] = row[3]
        elif row[0] == 'No':
            header_row = i
            break

    if header_row is None:
        st.error("No se encontró la fila de encabezados (No, Código, Apellidos y Nombres).")
        return None, None

    alumnos = []
    for row in rows[header_row + 1:]:
        if row[0] is None or not str(row[0]).strip():
            continue
        try:
            numero = int(row[0])
        except (ValueError, TypeError):
            continue
        codigo = str(row[1]).strip() if row[1] else ""
        nombre = str(row[2]).strip() if row[2] else ""
        if nombre:
            alumnos.append({
                "No": numero,
                "Codigo": codigo,
                "Nombre": nombre
            })

    df = pd.DataFrame(alumnos)
    wb.close()
    return metadata, df

def leer_csv_zoom(archivo):
    """Lee CSV de Zoom y retorna DataFrame con participantes únicos."""
    try:
        df = pd.read_csv(archivo, sep=';', encoding='utf-8')
    except Exception:
        df = pd.read_csv(archivo, sep=';', encoding='latin-1')

    col_nombre = 'Nombre (nombre original)'
    col_correo = 'Correo electrónico'

    if col_nombre not in df.columns:
        st.error(f"No se encontró la columna '{col_nombre}' en el CSV.")
        return None

    # Participantes únicos por nombre+correo (pueden repetirse por entrar/salir)
    df = df[[col_nombre, col_correo]].copy()
    df.columns = ['nombre_zoom', 'correo']
    df['nombre_zoom'] = df['nombre_zoom'].astype(str).str.strip()
    df['correo'] = df['correo'].astype(str).str.strip().replace('nan', '')

    # Quitar filas vacías o con nombres de sala
    df = df[df['nombre_zoom'].str.len() > 1]
    df = df[~df['nombre_zoom'].str.lower().str.contains('sala online|sala esan', na=False)]

    # Deduplicar — si un participante entró/salió varias veces, quedarse con 1 fila
    # Preferir la fila que tiene correo
    df['tiene_correo'] = df['correo'].str.contains('@esan', na=False)
    df = df.sort_values('tiene_correo', ascending=False)
    df = df.drop_duplicates(subset='nombre_zoom', keep='first')
    df = df.drop(columns='tiene_correo').reset_index(drop=True)

    return df

# ─────────────────────────────────────────────
# HISTORIAL
# ─────────────────────────────────────────────

def cargar_historial():
    try:
        with open(HISTORIAL_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return {}

def guardar_historial(historial):
    with open(HISTORIAL_FILE, 'w', encoding='utf-8') as f:
        json.dump(historial, f, ensure_ascii=False, indent=2)

def clave_historial(nombre_zoom):
    return normalizar(nombre_zoom)

# ─────────────────────────────────────────────
# LÓGICA DE CRUCE
# ─────────────────────────────────────────────

def cruzar_asistencia(df_alumnos, df_zoom, historial, curso_id):
    """
    Retorna:
      - resultado: dict {codigo_alumno: 'A' | 'F' | 'P'}
      - pendientes: lista de dicts con nombre_zoom sin match automático
      - auto_matches: lista de dicts con matches automáticos para mostrar
    """
    resultado = {row['Codigo']: 'F' for _, row in df_alumnos.iterrows()}
    pendientes = []
    auto_matches = []

    historial_curso = historial.get(curso_id, {})

    for _, pax in df_zoom.iterrows():
        nombre_zoom = pax['nombre_zoom']
        correo = pax['correo']
        codigo_correo = extraer_codigo_correo(correo)
        clave = clave_historial(nombre_zoom)

        # Nivel 1: match por código en correo
        if codigo_correo and codigo_correo in resultado:
            resultado[codigo_correo] = 'A'
            alumno_nombre = df_alumnos[df_alumnos['Codigo'] == codigo_correo]['Nombre'].values[0]
            auto_matches.append({
                'nombre_zoom': nombre_zoom,
                'alumno': alumno_nombre,
                'metodo': '📧 Correo'
            })
            continue

        # Nivel 2: historial guardado
        if clave in historial_curso:
            codigo_guardado = historial_curso[clave]
            if codigo_guardado == '__IGNORAR__':
                continue  # nombre ignorado (ej: profesor, sala)
            if codigo_guardado in resultado:
                resultado[codigo_guardado] = 'A'
                alumno_nombre = df_alumnos[df_alumnos['Codigo'] == codigo_guardado]['Nombre'].values[0]
                auto_matches.append({
                    'nombre_zoom': nombre_zoom,
                    'alumno': alumno_nombre,
                    'metodo': '💾 Historial'
                })
            continue  # ya está en historial, no mostrar en pendientes aunque el código no esté

        # Nivel 3: match por palabras (>= 2 coincidencias)
        mejor_score = 0
        mejor_codigo = None
        mejor_nombre = None
        for _, alumno in df_alumnos.iterrows():
            score = coincidencias(nombre_zoom, alumno['Nombre'])
            if score > mejor_score:
                mejor_score = score
                mejor_codigo = alumno['Codigo']
                mejor_nombre = alumno['Nombre']

        if mejor_score >= 2:
            resultado[mejor_codigo] = 'A'
            auto_matches.append({
                'nombre_zoom': nombre_zoom,
                'alumno': mejor_nombre,
                'metodo': f'🔤 Palabras ({mejor_score})'
            })
        else:
            pendientes.append({
                'nombre_zoom': nombre_zoom,
                'correo': correo,
                'mejor_sugerencia': mejor_nombre,
                'mejor_score': mejor_score
            })

    return resultado, pendientes, auto_matches

# ─────────────────────────────────────────────
# EXPORTAR EXCEL
# ─────────────────────────────────────────────

def exportar_excel(archivo_original, sesiones, metadata):
    """
    Genera Excel con columnas S1, S2, S3... añadidas al lado del nombre.
    sesiones: lista de dicts {label: 'S1', resultado: {codigo: 'A'|'F'}}
    """
    wb = load_workbook(archivo_original)
    ws = wb.active

    # Encontrar fila header y columna nombre
    header_row_idx = None
    for i, row in enumerate(ws.iter_rows(values_only=True), 1):
        if row[0] == 'No':
            header_row_idx = i
            break

    if not header_row_idx:
        return None

    # Columna donde escribir S1, S2... (después de Apellidos y Nombres = col C=3, combinada hasta F=6)
    # Usamos columna 7 en adelante para las sesiones
    col_inicio_sesiones = 7

    # Colores
    verde = PatternFill("solid", fgColor="C6EFCE")
    rojo = PatternFill("solid", fgColor="FFC7CE")
    amarillo = PatternFill("solid", fgColor="FFEB9C")
    azul_header = PatternFill("solid", fgColor="BDD7EE")

    # Escribir encabezados de sesiones
    for s_idx, sesion in enumerate(sesiones):
        col = col_inicio_sesiones + s_idx
        cell = ws.cell(row=header_row_idx, column=col, value=sesion['label'])
        cell.fill = azul_header
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center')

    # Escribir resultados por alumno
    for row_idx in range(header_row_idx + 1, ws.max_row + 1):
        codigo_cell = ws.cell(row=row_idx, column=2)
        if not codigo_cell.value:
            continue
        codigo = str(codigo_cell.value).strip()

        for s_idx, sesion in enumerate(sesiones):
            col = col_inicio_sesiones + s_idx
            valor = sesion['resultado'].get(codigo, '—')
            cell = ws.cell(row=row_idx, column=col, value=valor)
            cell.alignment = Alignment(horizontal='center')
            if valor == 'A':
                cell.fill = verde
                cell.font = Font(color="276221", bold=True)
            elif valor == 'F':
                cell.fill = rojo
                cell.font = Font(color="9C0006", bold=True)
            elif valor == 'P':
                cell.fill = amarillo

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# ─────────────────────────────────────────────
# INTERFAZ PRINCIPAL
# ─────────────────────────────────────────────

def main():
    st.title("📋 Control de Asistencias ESAN")

    historial = cargar_historial()

    # ── SIDEBAR ──
    with st.sidebar:
        st.header("⚙️ Configuración")
        st.markdown("---")

        archivo_excel = st.file_uploader(
            "1️⃣ Sube el Excel maestro de asistencias",
            type=["xlsx"],
            key="excel_upload"
        )

        if archivo_excel:
            excel_bytes = archivo_excel.read()
            metadata, df_alumnos = leer_excel_asistencia(io.BytesIO(excel_bytes))
            if metadata and df_alumnos is not None:
                st.success(f"✅ {len(df_alumnos)} alumnos cargados")
                st.markdown(f"**Curso:** {metadata.get('curso', '—')}")
                st.markdown(f"**Profesor:** {metadata.get('profesor', '—')}")

                # ID único del curso para historial
                curso_id = metadata.get('curso', 'sin_curso')
                st.session_state['metadata'] = metadata
                st.session_state['df_alumnos'] = df_alumnos
                st.session_state['curso_id'] = curso_id
                st.session_state['archivo_excel_bytes'] = excel_bytes

        st.markdown("---")

        if 'sesiones' not in st.session_state:
            st.session_state['sesiones'] = []

        if 'df_alumnos' in st.session_state:
            n_sesiones = len(st.session_state['sesiones'])
            st.markdown(f"**Sesiones procesadas:** {n_sesiones}")
            for i, s in enumerate(st.session_state['sesiones']):
                st.markdown(f"- S{i+1}: {s['label']}")

            st.markdown("---")
            if st.button("🗑️ Reiniciar todo", type="secondary"):
                for key in ['sesiones', 'df_alumnos', 'metadata', 'curso_id',
                            'archivo_excel_bytes', 'pendientes_actuales',
                            'resultado_actual', 'auto_matches_actuales', 'label_sesion']:
                    st.session_state.pop(key, None)
                st.rerun()

    # ── ÁREA PRINCIPAL ──
    if 'df_alumnos' not in st.session_state:
        st.info("👈 Comienza subiendo el Excel maestro de asistencias en el panel izquierdo.")
        return

    df_alumnos = st.session_state['df_alumnos']
    curso_id = st.session_state['curso_id']
    metadata = st.session_state['metadata']

    # ── TABS ──
    tab1, tab2, tab3 = st.tabs(["📥 Procesar sesión", "✅ Resultados", "📤 Exportar"])

    # ─── TAB 1: PROCESAR SESIÓN ───
    with tab1:
        st.subheader("Cargar CSV de Zoom")

        col1, col2 = st.columns([2, 1])
        with col1:
            archivo_csv = st.file_uploader(
                "Sube el CSV de Zoom para esta sesión",
                type=["csv"],
                key="csv_upload"
            )
        with col2:
            label_sesion = st.text_input(
                "Etiqueta de sesión (ej: 16/06)",
                value=st.session_state.get('label_sesion', ''),
                key="label_input"
            )

        if archivo_csv and label_sesion:
            df_zoom = leer_csv_zoom(archivo_csv)
            if df_zoom is not None:
                if st.button("🔄 Procesar sesión", type="primary"):
                    resultado, pendientes, auto_matches = cruzar_asistencia(
                        df_alumnos, df_zoom, historial, curso_id
                    )
                    st.session_state['pendientes_actuales'] = pendientes
                    st.session_state['resultado_actual'] = resultado
                    st.session_state['auto_matches_actuales'] = auto_matches
                    st.session_state['label_sesion'] = label_sesion
                    st.rerun()

        # ── MATCHES AUTOMÁTICOS ──
        if 'auto_matches_actuales' in st.session_state:
            auto_matches = st.session_state['auto_matches_actuales']
            pendientes = st.session_state['pendientes_actuales']
            resultado = st.session_state['resultado_actual']

            col_a, col_b, col_c = st.columns(3)
            total_asistio = sum(1 for v in resultado.values() if v == 'A')
            col_a.metric("✅ Asistieron (auto)", total_asistio)
            col_b.metric("⚠️ Pendientes", len(pendientes))
            col_c.metric("❌ Faltas", len(df_alumnos) - total_asistio - len(pendientes))

            if auto_matches:
                with st.expander(f"✅ Matches automáticos ({len(auto_matches)})", expanded=False):
                    df_auto = pd.DataFrame(auto_matches)
                    df_auto.columns = ['Nombre en Zoom', 'Alumno en lista', 'Método']
                    st.dataframe(df_auto, use_container_width=True, hide_index=True)

            # ── PENDIENTES: REVISIÓN MANUAL ──
            if pendientes:
                st.markdown("---")
                st.subheader(f"⚠️ Pendientes de revisión manual ({len(pendientes)})")
                st.caption("Estos nombres no coincidieron automáticamente. Enlázalos o ignóralos.")

                # Solo mostrar alumnos que aún no tienen asistencia marcada
                codigos_con_a = {cod for cod, val in resultado.items() if val == 'A'}
                opciones_alumnos = ["— No enlazar / Ignorar"] + [
                    f"{row['Nombre']} ({row['Codigo']})"
                    for _, row in df_alumnos.iterrows()
                    if row['Codigo'] not in codigos_con_a
                ]

                relaciones_manuales = {}

                for i, pax in enumerate(pendientes):
                    nombre_zoom = pax['nombre_zoom']
                    sugerencia = pax['mejor_sugerencia']

                    with st.container():
                        col_zoom, col_flecha, col_select = st.columns([3, 0.5, 4])

                        with col_zoom:
                            st.markdown(
                                f"<div style='background:#fff3cd;padding:10px;border-radius:8px;"
                                f"border-left:4px solid #ffc107;margin:4px 0;color:#1a1a1a'>"
                                f"<b style='color:#1a1a1a'>Zoom:</b> {nombre_zoom}<br>"
                                f"<small style='color:#555'>Correo: {pax['correo'] or '—'}</small>"
                                f"</div>",
                                unsafe_allow_html=True
                            )

                        with col_flecha:
                            st.markdown("<div style='padding-top:18px;text-align:center;font-size:20px'>→</div>",
                                        unsafe_allow_html=True)

                        with col_select:
                            # Pre-seleccionar sugerencia si existe
                            idx_default = 0
                            if sugerencia:
                                for j, op in enumerate(opciones_alumnos):
                                    if sugerencia in op:
                                        idx_default = j
                                        break

                            seleccion = st.selectbox(
                                f"Alumno para: {nombre_zoom[:30]}...",
                                opciones_alumnos,
                                index=idx_default,
                                key=f"select_{i}",
                                label_visibility="collapsed"
                            )
                            relaciones_manuales[nombre_zoom] = seleccion

                        st.markdown("<hr style='margin:4px 0;opacity:0.2'>", unsafe_allow_html=True)

                col_rep, col_ok = st.columns([1, 1])

                with col_rep:
                    if st.button("🔄 Reprocesar con estas relaciones", type="secondary"):
                        # Aplicar relaciones al resultado
                        for nombre_zoom, seleccion in relaciones_manuales.items():
                            if seleccion == "— No enlazar / Ignorar":
                                clave = clave_historial(nombre_zoom)
                                if curso_id not in historial:
                                    historial[curso_id] = {}
                                historial[curso_id][clave] = '__IGNORAR__'
                            else:
                                match = re.search(r'\((\d+)\)$', seleccion)
                                if match:
                                    codigo = match.group(1)
                                    resultado[codigo] = 'A'
                                    clave = clave_historial(nombre_zoom)
                                    if curso_id not in historial:
                                        historial[curso_id] = {}
                                    historial[curso_id][clave] = codigo

                        guardar_historial(historial)
                        st.session_state['resultado_actual'] = resultado
                        st.session_state['pendientes_actuales'] = []
                        st.success("✅ Relaciones aplicadas y guardadas en historial.")
                        st.rerun()

                with col_ok:
                    if st.button("✅ Confirmar sesión y guardar", type="primary"):
                        # Guardar relaciones en historial
                        for nombre_zoom, seleccion in relaciones_manuales.items():
                            clave = clave_historial(nombre_zoom)
                            if curso_id not in historial:
                                historial[curso_id] = {}
                            if seleccion == "— No enlazar / Ignorar":
                                historial[curso_id][clave] = '__IGNORAR__'
                            else:
                                match = re.search(r'\((\d+)\)$', seleccion)
                                if match:
                                    codigo = match.group(1)
                                    resultado[codigo] = 'A'
                                    historial[curso_id][clave] = codigo

                        guardar_historial(historial)

                        # Guardar sesión
                        st.session_state['sesiones'].append({
                            'label': st.session_state['label_sesion'],
                            'resultado': resultado
                        })
                        st.session_state.pop('pendientes_actuales', None)
                        st.session_state.pop('resultado_actual', None)
                        st.session_state.pop('auto_matches_actuales', None)
                        st.success(f"✅ Sesión '{st.session_state['label_sesion']}' guardada.")
                        st.rerun()

            else:
                # Sin pendientes
                if st.button("✅ Confirmar sesión y guardar", type="primary"):
                    st.session_state['sesiones'].append({
                        'label': st.session_state['label_sesion'],
                        'resultado': resultado
                    })
                    st.session_state.pop('pendientes_actuales', None)
                    st.session_state.pop('resultado_actual', None)
                    st.session_state.pop('auto_matches_actuales', None)
                    st.success(f"✅ Sesión '{st.session_state['label_sesion']}' guardada sin pendientes.")
                    st.rerun()

    # ─── TAB 2: RESULTADOS ───
    with tab2:
        sesiones = st.session_state.get('sesiones', [])
        if not sesiones:
            st.info("Aún no hay sesiones procesadas.")
        else:
            st.subheader("Tabla de asistencia acumulada")

            rows = []
            for _, alumno in df_alumnos.iterrows():
                fila = {
                    'No': alumno['No'],
                    'Código': alumno['Codigo'],
                    'Apellidos y Nombres': alumno['Nombre']
                }
                total_a = 0
                for i, s in enumerate(sesiones):
                    val = s['resultado'].get(alumno['Codigo'], 'F')
                    fila[f"S{i+1}\n{s['label']}"] = val
                    if val == 'A':
                        total_a += 1
                fila['Total asistencias'] = total_a
                rows.append(fila)

            df_resultado = pd.DataFrame(rows)

            def colorear(val):
                if val == 'A':
                    return 'background-color: #C6EFCE; color: #276221; font-weight:bold'
                elif val == 'F':
                    return 'background-color: #FFC7CE; color: #9C0006; font-weight:bold'
                return ''

            cols_sesion = [c for c in df_resultado.columns if c.startswith('S')]
            st.dataframe(
                df_resultado.style.map(colorear, subset=cols_sesion),
                use_container_width=True,
                hide_index=True,
                height=500
            )

            # Resumen
            st.markdown("---")
            col1, col2, col3 = st.columns(3)
            col1.metric("Total alumnos", len(df_alumnos))
            col2.metric("Sesiones procesadas", len(sesiones))
            promedio = sum(
                sum(1 for s in sesiones if s['resultado'].get(a['Codigo']) == 'A')
                for _, a in df_alumnos.iterrows()
            ) / max(len(df_alumnos), 1)
            col3.metric("Promedio de asistencia", f"{promedio:.1f} / {len(sesiones)}")

    # ─── TAB 3: EXPORTAR ───
    with tab3:
        sesiones = st.session_state.get('sesiones', [])
        if not sesiones:
            st.info("Procesa al menos una sesión antes de exportar.")
        else:
            st.subheader("Descargar Excel con asistencias marcadas")
            st.markdown(
                "El archivo descargado será el Excel original con columnas **S1, S2, S3...** "
                "añadidas al lado de los nombres, con colores verde (A) y rojo (F)."
            )

            archivo_excel_bytes = st.session_state.get('archivo_excel_bytes')
            if archivo_excel_bytes:
                sesiones_export = [
                    {'label': f"S{i+1}\n{s['label']}", 'resultado': s['resultado']}
                    for i, s in enumerate(sesiones)
                ]
                output = exportar_excel(
                    io.BytesIO(archivo_excel_bytes),
                    sesiones_export,
                    metadata
                )
                if output:
                    nombre_curso = re.sub(r'[^\w]', '_', metadata.get('curso', 'CURSO'))[:30]
                    st.download_button(
                        label="📥 Descargar Excel con asistencias",
                        data=output,
                        file_name=f"ASISTENCIAS_{nombre_curso}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary"
                    )
            else:
                st.warning("Vuelve a subir el Excel original para poder exportar.")

        st.markdown("---")
        st.subheader("📚 Historial de relaciones guardadas")
        historial_actual = cargar_historial()
        if historial_actual:
            col_hist1, col_hist2 = st.columns([2, 1])
            with col_hist2:
                if st.button("🗑️ Limpiar todo el historial", type="secondary"):
                    guardar_historial({})
                    st.success("Historial limpiado.")
                    st.rerun()

            for curso, relaciones in historial_actual.items():
                with st.expander(f"Curso: {curso[:60]}"):
                    col_c1, col_c2 = st.columns([3, 1])
                    with col_c2:
                        if st.button(f"🗑️ Limpiar este curso", key=f"del_{curso[:20]}"):
                            del historial_actual[curso]
                            guardar_historial(historial_actual)
                            st.success("Curso eliminado del historial.")
                            st.rerun()
                    filas = []
                    for clave, codigo in relaciones.items():
                        if codigo == '__IGNORAR__':
                            filas.append({'Nombre Zoom (normalizado)': clave, 'Alumno asignado': '🚫 Ignorado'})
                        else:
                            alumno_match = df_alumnos[df_alumnos['Codigo'] == codigo]
                            nombre_alumno = alumno_match['Nombre'].values[0] if len(alumno_match) else codigo
                            filas.append({'Nombre Zoom (normalizado)': clave, 'Alumno asignado': nombre_alumno})
                    st.dataframe(pd.DataFrame(filas), use_container_width=True, hide_index=True)
        else:
            st.info("No hay historial guardado aún.")

if __name__ == "__main__":
    main()
