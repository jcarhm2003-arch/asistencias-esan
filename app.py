import streamlit as st
import pandas as pd
import json
import re
import io
import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
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
    if not texto:
        return ""
    texto = str(texto)
    texto = re.sub(r'\(.*?\)', '', texto)
    texto = unidecode(texto)
    texto = texto.lower()
    texto = re.sub(r'[^a-z\s]', '', texto)
    return texto.strip()

def palabras(texto):
    return set(w for w in normalizar(texto).split() if len(w) >= 2)

def coincidencias(nombre_zoom, nombre_excel):
    p_zoom = palabras(nombre_zoom)
    p_excel = palabras(nombre_excel)
    return len(p_zoom & p_excel)

def extraer_codigo_correo(correo):
    if not correo or str(correo).strip() in ('', 'nan'):
        return None
    match = re.match(r'^(\d{6,7})@esan\.edu\.pe$', str(correo).strip())
    return match.group(1) if match else None

def parsear_hora(val):
    """Detecta formato ISO o peruano y convierte a hora Lima."""
    if not val or str(val).strip() in ('', 'nan'):
        return pd.NaT
    val = str(val).strip()
    try:
        dt = pd.to_datetime(val, utc=True)
        return dt.tz_convert('America/Lima').tz_localize(None)
    except Exception:
        pass
    try:
        return pd.to_datetime(val, format='%d/%m/%Y %H:%M', errors='coerce')
    except Exception:
        return pd.NaT

# ─────────────────────────────────────────────
# PARSEO DE TEXTO PEGADO
# ─────────────────────────────────────────────

def parsear_alumnos(texto):
    """
    Parsea texto pegado desde Excel con columnas:
    Código (opcional) + Nombre
    Detecta automáticamente si hay código o solo nombre.
    """
    lineas = [l.strip() for l in texto.strip().splitlines() if l.strip()]
    alumnos = []
    for i, linea in enumerate(lineas):
        partes = re.split(r'\t|  +', linea.strip())
        partes = [p.strip() for p in partes if p.strip()]
        if not partes:
            continue
        # Si la primera parte es numérica (7 dígitos o menos) → es código
        if len(partes) >= 2 and re.match(r'^\d{5,8}$', partes[0]):
            codigo = partes[0]
            nombre = ' '.join(partes[1:])
        else:
            codigo = str(i + 1)  # código interno si no hay
            nombre = ' '.join(partes)
        if nombre:
            alumnos.append({'No': i + 1, 'Codigo': codigo, 'Nombre': nombre.upper()})
    return pd.DataFrame(alumnos) if alumnos else None

def parsear_zoom(texto):
    """
    Parsea texto pegado con columnas:
    Nombre (obligatorio) + Correo (opcional) + Hora entrada (opcional)
    Detecta automáticamente cuántas columnas hay.
    Retorna df con participantes únicos + dict de primeras horas.
    """
    lineas = [l.strip() for l in texto.strip().splitlines() if l.strip()]
    registros = []
    for linea in lineas:
        partes = re.split(r'\t|  +', linea.strip())
        partes = [p.strip() for p in partes if p.strip()]
        if not partes:
            continue
        nombre = partes[0]
        correo = partes[1] if len(partes) >= 2 else ''
        hora = partes[2] if len(partes) >= 3 else ''

        # Si la segunda columna parece una hora y no un correo, reajustar
        if correo and re.match(r'\d{4}-\d{2}|^\d{2}/\d{2}/', correo):
            hora = correo
            correo = ''

        registros.append({
            'nombre_zoom': nombre.strip(),
            'correo': correo.strip(),
            'hora_entrada': hora.strip()
        })

    if not registros:
        return None, {}

    df_raw = pd.DataFrame(registros)
    df_raw['nombre_zoom'] = df_raw['nombre_zoom'].astype(str).str.strip()

    # Filtrar filas vacías o salas
    df_raw = df_raw[df_raw['nombre_zoom'].str.len() > 1]
    df_raw = df_raw[~df_raw['nombre_zoom'].str.lower().str.contains('sala online|sala esan', na=False)]

    # Calcular primera hora de entrada por nombre
    primeras_horas = {}
    if df_raw['hora_entrada'].str.strip().ne('').any():
        df_horas = df_raw[['nombre_zoom', 'hora_entrada']].copy()
        df_horas['hora_entrada'] = df_horas['hora_entrada'].apply(parsear_hora)
        primeras_horas = df_horas.groupby('nombre_zoom')['hora_entrada'].min().to_dict()

    # Deduplicar — preferir fila con correo
    df_raw['tiene_correo'] = df_raw['correo'].str.contains('@esan', na=False)
    df_raw = df_raw.sort_values('tiene_correo', ascending=False)
    df_raw = df_raw.drop_duplicates(subset='nombre_zoom', keep='first')
    df_raw = df_raw.drop(columns=['tiene_correo', 'hora_entrada']).reset_index(drop=True)

    return df_raw, primeras_horas

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

def cruzar_asistencia(df_alumnos, df_zoom, historial, curso_id, primeras_horas=None, hora_limite=None):
    resultado = {row['Codigo']: 'F' for _, row in df_alumnos.iterrows()}
    pendientes = []
    auto_matches = []
    historial_curso = historial.get(curso_id, {})

    def calcular_estado(nombre_zoom):
        if not hora_limite or not primeras_horas:
            return 'A'
        hora_entrada = primeras_horas.get(nombre_zoom)
        if hora_entrada is None or pd.isna(hora_entrada):
            return 'A'
        if hora_entrada.time() > hora_limite:
            return 'T'
        return 'A'

    for _, pax in df_zoom.iterrows():
        nombre_zoom = pax['nombre_zoom']
        correo = pax.get('correo', '')
        codigo_correo = extraer_codigo_correo(correo)
        clave = clave_historial(nombre_zoom)
        estado = calcular_estado(nombre_zoom)

        # Nivel 1: match por código en correo
        if codigo_correo and codigo_correo in resultado:
            resultado[codigo_correo] = estado
            alumno_nombre = df_alumnos[df_alumnos['Codigo'] == codigo_correo]['Nombre'].values[0]
            auto_matches.append({
                'nombre_zoom': nombre_zoom,
                'alumno': alumno_nombre,
                'metodo': '📧 Correo' + (' ⏰' if estado == 'T' else '')
            })
            continue

        # Nivel 2: historial guardado
        if clave in historial_curso:
            codigo_guardado = historial_curso[clave]
            if codigo_guardado == '__IGNORAR__':
                continue
            if codigo_guardado in resultado:
                resultado[codigo_guardado] = estado
                alumno_nombre = df_alumnos[df_alumnos['Codigo'] == codigo_guardado]['Nombre'].values[0]
                auto_matches.append({
                    'nombre_zoom': nombre_zoom,
                    'alumno': alumno_nombre,
                    'metodo': '💾 Historial' + (' ⏰' if estado == 'T' else '')
                })
            continue

        # Nivel 3: match por palabras (sin empates)
        mejor_score = 0
        mejor_codigo = None
        mejor_nombre = None
        empatados = []
        for _, alumno in df_alumnos.iterrows():
            score = coincidencias(nombre_zoom, alumno['Nombre'])
            if score > mejor_score:
                mejor_score = score
                mejor_codigo = alumno['Codigo']
                mejor_nombre = alumno['Nombre']
                empatados = [alumno['Nombre']]
            elif score == mejor_score and score > 0:
                empatados.append(alumno['Nombre'])

        if mejor_score >= 2 and len(empatados) == 1:
            resultado[mejor_codigo] = estado
            auto_matches.append({
                'nombre_zoom': nombre_zoom,
                'alumno': mejor_nombre,
                'metodo': f'🔤 Palabras ({mejor_score})' + (' ⏰' if estado == 'T' else '')
            })
        elif mejor_score >= 2 and len(empatados) > 1:
            pendientes.append({
                'nombre_zoom': nombre_zoom,
                'correo': correo,
                'mejor_sugerencia': None,
                'mejor_score': mejor_score,
                'estado': estado
            })
        else:
            pendientes.append({
                'nombre_zoom': nombre_zoom,
                'correo': correo,
                'mejor_sugerencia': mejor_nombre,
                'mejor_score': mejor_score,
                'estado': estado
            })

    return resultado, pendientes, auto_matches

# ─────────────────────────────────────────────
# EXPORTAR EXCEL (tabla nueva limpia)
# ─────────────────────────────────────────────

def exportar_excel(df_alumnos, sesiones, nombre_curso):
    wb = __import__('openpyxl').Workbook()
    ws = wb.active
    ws.title = "Asistencias"

    verde = PatternFill("solid", fgColor="C6EFCE")
    rojo = PatternFill("solid", fgColor="FFC7CE")
    amarillo = PatternFill("solid", fgColor="FFEB9C")
    azul_header = PatternFill("solid", fgColor="BDD7EE")
    gris_header = PatternFill("solid", fgColor="D9D9D9")

    # Encabezados
    headers = ['No', 'Código', 'Apellidos y Nombres'] + \
              [f"S{i+1}\n{s['label']}" for i, s in enumerate(sesiones)] + \
              ['Total asistencias']

    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.fill = azul_header if col_idx > 3 else gris_header
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    ws.row_dimensions[1].height = 30

    # Datos
    for row_idx, (_, alumno) in enumerate(df_alumnos.iterrows(), 2):
        ws.cell(row=row_idx, column=1, value=alumno['No'])
        ws.cell(row=row_idx, column=2, value=alumno['Codigo'])
        ws.cell(row=row_idx, column=3, value=alumno['Nombre'])

        total = 0
        for s_idx, sesion in enumerate(sesiones):
            col = 4 + s_idx
            valor = sesion['resultado'].get(alumno['Codigo'], 'F')
            cell = ws.cell(row=row_idx, column=col, value=valor)
            cell.alignment = Alignment(horizontal='center')
            if valor == 'A':
                cell.fill = verde
                cell.font = Font(color="276221", bold=True)
                total += 1
            elif valor == 'T':
                cell.fill = amarillo
                cell.font = Font(color="7a5c00", bold=True)
                total += 1
            elif valor == 'F':
                cell.fill = rojo
                cell.font = Font(color="9C0006", bold=True)

        total_cell = ws.cell(row=row_idx, column=4 + len(sesiones), value=total)
        total_cell.alignment = Alignment(horizontal='center')
        total_cell.font = Font(bold=True)

    # Ancho de columnas
    ws.column_dimensions['A'].width = 5
    ws.column_dimensions['B'].width = 12
    ws.column_dimensions['C'].width = 35
    for i in range(len(sesiones) + 1):
        ws.column_dimensions[__import__('openpyxl').utils.get_column_letter(4 + i)].width = 10

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
        st.header("⚙️ Configuración del curso")
        st.markdown("---")

        curso_nombre = st.text_input(
            "Nombre del curso",
            value=st.session_state.get('curso_nombre', ''),
            placeholder="Ej: SEGURIDAD LABORAL 2026-1",
            key="curso_input"
        )

        st.markdown("**Lista de alumnos** — pega desde Excel:")
        st.caption("Formato: `Código    Nombre` o solo `Nombre` (una fila por alumno)")

        texto_alumnos = st.text_area(
            "Pega aquí los alumnos",
            height=200,
            key="texto_alumnos",
            placeholder="2607391\tADVINCULA ROMERO JOSE MANUEL\n2607372\tAGUIRRE GARCIA FRANCISCO\n..."
        )

        if st.button("✅ Cargar lista de alumnos", type="primary"):
            if not curso_nombre.strip():
                st.error("Escribe el nombre del curso primero.")
            elif not texto_alumnos.strip():
                st.error("Pega la lista de alumnos.")
            else:
                df = parsear_alumnos(texto_alumnos)
                if df is not None and len(df) > 0:
                    st.session_state['df_alumnos'] = df
                    st.session_state['curso_id'] = normalizar(curso_nombre)
                    st.session_state['curso_nombre'] = curso_nombre
                    st.session_state['sesiones'] = []
                    for key in ['pendientes_actuales', 'resultado_actual', 'auto_matches_actuales', 'label_sesion']:
                        st.session_state.pop(key, None)
                    st.success(f"✅ {len(df)} alumnos cargados")
                else:
                    st.error("No se pudo parsear la lista. Verifica el formato.")

        st.markdown("---")

        if 'df_alumnos' in st.session_state:
            df_alumnos_sidebar = st.session_state['df_alumnos']
            st.markdown(f"**Curso:** {st.session_state.get('curso_nombre', '—')}")
            st.markdown(f"**Alumnos:** {len(df_alumnos_sidebar)}")
            n_sesiones = len(st.session_state.get('sesiones', []))
            st.markdown(f"**Sesiones procesadas:** {n_sesiones}")
            for i, s in enumerate(st.session_state.get('sesiones', [])):
                st.markdown(f"- S{i+1}: {s['label']}")

            st.markdown("---")
            if st.button("🗑️ Reiniciar todo", type="secondary"):
                for key in ['sesiones', 'df_alumnos', 'curso_id', 'curso_nombre',
                            'pendientes_actuales', 'resultado_actual',
                            'auto_matches_actuales', 'label_sesion']:
                    st.session_state.pop(key, None)
                st.rerun()

    # ── ÁREA PRINCIPAL ──
    if 'df_alumnos' not in st.session_state:
        st.info("👈 Comienza pegando la lista de alumnos en el panel izquierdo.")
        return

    df_alumnos = st.session_state['df_alumnos']
    curso_id = st.session_state['curso_id']

    # ── TABS ──
    tab1, tab2, tab3 = st.tabs(["📥 Procesar sesión", "✅ Resultados", "📤 Exportar"])

    # ─── TAB 1: PROCESAR SESIÓN ───
    with tab1:
        st.subheader("Datos de participantes Zoom")
        st.caption("Pega desde el CSV de Zoom. Columnas: **Nombre** (obligatorio) · **Correo** (opcional) · **Hora de entrada** (opcional)")

        usar_hora_limite = st.checkbox("⏰ ¿Aplicar hora límite de tardanza?", key="check_hora_limite")
        hora_limite_valor = None
        if usar_hora_limite:
            hora_limite_input = st.time_input(
                "Hora límite (después de esta hora se marca T)",
                value=datetime.time(18, 10),
                key="hora_limite_input"
            )
            hora_limite_valor = hora_limite_input

        col1, col2 = st.columns([3, 1])
        with col1:
            texto_zoom = st.text_area(
                "Pega aquí los participantes de Zoom",
                height=200,
                key="texto_zoom",
                placeholder="Wilson Monzon Araujo\t2607391@esan.edu.pe\t18/06/2026 17:25\nLUCINDA BUSTINZA MORALES\t2607365@esan.edu.pe\t18/06/2026 18:30\nTANIA\t\t..."
            )
        with col2:
            label_sesion = st.text_input(
                "Etiqueta de sesión",
                value=st.session_state.get('label_sesion', ''),
                placeholder="Ej: 16/06",
                key="label_input"
            )

        if texto_zoom.strip() and label_sesion.strip():
            if st.button("🔄 Procesar sesión", type="primary"):
                df_zoom, primeras_horas = parsear_zoom(texto_zoom)
                if df_zoom is not None and len(df_zoom) > 0:
                    resultado, pendientes, auto_matches = cruzar_asistencia(
                        df_alumnos, df_zoom, historial, curso_id,
                        primeras_horas=primeras_horas,
                        hora_limite=hora_limite_valor
                    )
                    st.session_state['pendientes_actuales'] = pendientes
                    st.session_state['resultado_actual'] = resultado
                    st.session_state['auto_matches_actuales'] = auto_matches
                    st.session_state['label_sesion'] = label_sesion
                    st.rerun()
                else:
                    st.error("No se pudo parsear los datos de Zoom. Verifica el formato.")

        # ── REVISIÓN ──
        if 'auto_matches_actuales' in st.session_state:
            auto_matches = st.session_state['auto_matches_actuales']
            pendientes = st.session_state['pendientes_actuales']
            resultado = st.session_state['resultado_actual']

            col_a, col_b, col_c, col_d = st.columns(4)
            total_asistio = sum(1 for v in resultado.values() if v == 'A')
            total_tarde = sum(1 for v in resultado.values() if v == 'T')
            col_a.metric("✅ A tiempo", total_asistio)
            col_b.metric("⏰ Tardanzas", total_tarde)
            col_c.metric("⚠️ Pendientes", len(pendientes))
            col_d.metric("❌ Faltas", len(df_alumnos) - total_asistio - total_tarde - len(pendientes))

            st.markdown("---")
            st.subheader("👥 Revisión de participantes")
            st.caption("Verde = enlazado automáticamente. Amarillo = requiere revisión manual.")

            codigos_con_asistencia = {cod for cod, val in resultado.items() if val in ('A', 'T')}

            opciones_alumnos = ["— No enlazar / Ignorar"] + [
                f"{'✅ ' if row['Codigo'] in codigos_con_asistencia else '❌ F  '}{row['Nombre']} ({row['Codigo']})"
                for _, row in df_alumnos.iterrows()
            ]

            relaciones_manuales = {}

            todos_zoom = []
            for m in auto_matches:
                todos_zoom.append({
                    'nombre_zoom': m['nombre_zoom'],
                    'correo': '',
                    'mejor_sugerencia': m['alumno'],
                    'metodo': m['metodo'],
                    'es_auto': True
                })
            for pax in pendientes:
                todos_zoom.append({
                    'nombre_zoom': pax['nombre_zoom'],
                    'correo': pax.get('correo', ''),
                    'mejor_sugerencia': pax['mejor_sugerencia'],
                    'metodo': None,
                    'es_auto': False,
                    'estado': pax.get('estado', 'A')
                })

            for i, pax in enumerate(todos_zoom):
                nombre_zoom = pax['nombre_zoom']
                es_auto = pax['es_auto']

                with st.container():
                    col_zoom, col_flecha, col_select = st.columns([3, 0.5, 4])

                    with col_zoom:
                        if es_auto:
                            color_borde = "#28a745"
                            color_bg = "#d4edda"
                            badge = f"<small style='color:#155724'>{pax['metodo']}</small>"
                        else:
                            color_borde = "#ffc107"
                            color_bg = "#fff3cd"
                            correo_display = pax['correo'] if pax['correo'] and pax['correo'] != 'nan' else '—'
                            badge = f"<small style='color:#856404'>⚠️ Sin enlazar &nbsp;|&nbsp; Correo: {correo_display}</small>"

                        st.markdown(
                            f"<div style='background:{color_bg};padding:10px;border-radius:8px;"
                            f"border-left:4px solid {color_borde};margin:4px 0;color:#1a1a1a'>"
                            f"<b style='color:#1a1a1a'>Zoom:</b> {nombre_zoom}<br>{badge}</div>",
                            unsafe_allow_html=True
                        )

                    with col_flecha:
                        st.markdown("<div style='padding-top:18px;text-align:center;font-size:20px'>→</div>",
                                    unsafe_allow_html=True)

                    with col_select:
                        if es_auto:
                            st.markdown(
                                f"<div style='background:#d4edda;padding:10px;border-radius:8px;"
                                f"border:1px solid #c3e6cb;margin:4px 0;color:#155724;font-weight:bold'>"
                                f"✅ {pax['mejor_sugerencia']}</div>",
                                unsafe_allow_html=True
                            )
                        else:
                            idx_default = 0
                            sugerencia = pax['mejor_sugerencia']
                            if sugerencia:
                                for j, op in enumerate(opciones_alumnos):
                                    if sugerencia in op:
                                        idx_default = j
                                        break
                            seleccion = st.selectbox(
                                f"Alumno para: {nombre_zoom[:30]}",
                                opciones_alumnos,
                                index=idx_default,
                                key=f"select_{i}",
                                label_visibility="collapsed"
                            )
                            relaciones_manuales[nombre_zoom] = seleccion

                    st.markdown("<hr style='margin:4px 0;opacity:0.2'>", unsafe_allow_html=True)

            # ── BOTONES ──
            st.markdown("---")
            col_rep, col_ok = st.columns(2)

            with col_rep:
                if pendientes and st.button("🔄 Reprocesar con estas relaciones", type="secondary"):
                    estados_pendientes = {p['nombre_zoom']: p.get('estado', 'A') for p in pendientes}
                    for nombre_zoom, seleccion in relaciones_manuales.items():
                        if seleccion == "— No enlazar / Ignorar":
                            clave = clave_historial(nombre_zoom)
                            historial.setdefault(curso_id, {})[clave] = '__IGNORAR__'
                        else:
                            match = re.search(r'\(([^)]+)\)$', seleccion)
                            if match:
                                codigo = match.group(1)
                                resultado[codigo] = estados_pendientes.get(nombre_zoom, 'A')
                                historial.setdefault(curso_id, {})[clave_historial(nombre_zoom)] = codigo
                    guardar_historial(historial)
                    st.session_state['resultado_actual'] = resultado
                    st.session_state['pendientes_actuales'] = []
                    st.success("✅ Relaciones aplicadas.")
                    st.rerun()

            with col_ok:
                if st.button("✅ Confirmar sesión y guardar", type="primary"):
                    estados_pendientes = {p['nombre_zoom']: p.get('estado', 'A') for p in pendientes}
                    for nombre_zoom, seleccion in relaciones_manuales.items():
                        clave = clave_historial(nombre_zoom)
                        if seleccion == "— No enlazar / Ignorar":
                            historial.setdefault(curso_id, {})[clave] = '__IGNORAR__'
                        else:
                            match = re.search(r'\(([^)]+)\)$', seleccion)
                            if match:
                                codigo = match.group(1)
                                resultado[codigo] = estados_pendientes.get(nombre_zoom, 'A')
                                historial.setdefault(curso_id, {})[clave] = codigo
                    guardar_historial(historial)
                    st.session_state['sesiones'].append({
                        'label': st.session_state['label_sesion'],
                        'resultado': resultado
                    })
                    for key in ['pendientes_actuales', 'resultado_actual', 'auto_matches_actuales']:
                        st.session_state.pop(key, None)
                    st.success(f"✅ Sesión '{st.session_state['label_sesion']}' guardada.")
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
                    fila[f"S{i+1} {s['label']}"] = val
                    if val in ('A', 'T'):
                        total_a += 1
                fila['Total'] = total_a
                rows.append(fila)

            df_resultado = pd.DataFrame(rows)

            def colorear(val):
                if val == 'A':
                    return 'background-color: #C6EFCE; color: #276221; font-weight:bold'
                elif val == 'T':
                    return 'background-color: #FFEB9C; color: #7a5c00; font-weight:bold'
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

            st.markdown("---")
            col1, col2, col3 = st.columns(3)
            col1.metric("Total alumnos", len(df_alumnos))
            col2.metric("Sesiones procesadas", len(sesiones))
            promedio = sum(
                sum(1 for s in sesiones if s['resultado'].get(a['Codigo']) in ('A', 'T'))
                for _, a in df_alumnos.iterrows()
            ) / max(len(df_alumnos), 1)
            col3.metric("Promedio asistencia", f"{promedio:.1f} / {len(sesiones)}")

    # ─── TAB 3: EXPORTAR ───
    with tab3:
        sesiones = st.session_state.get('sesiones', [])
        if not sesiones:
            st.info("Procesa al menos una sesión antes de exportar.")
        else:
            st.subheader("Descargar Excel con asistencias")
            st.markdown("Tabla limpia con **No, Código, Nombre, S1, S2...** y total de asistencias.")

            sesiones_export = [
                {'label': s['label'], 'resultado': s['resultado']}
                for s in sesiones
            ]
            nombre_curso = st.session_state.get('curso_nombre', 'CURSO')
            output = exportar_excel(df_alumnos, sesiones_export, nombre_curso)

            nombre_archivo = re.sub(r'[^\w\s-]', '', nombre_curso).strip().replace(' ', '_')
            st.download_button(
                label="📥 Descargar Excel",
                data=output,
                file_name=f"ASISTENCIAS_{nombre_archivo}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )

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
                        if st.button("🗑️ Limpiar este curso", key=f"del_{curso[:20]}"):
                            del historial_actual[curso]
                            guardar_historial(historial_actual)
                            st.success("Curso eliminado.")
                            st.rerun()
                    filas = []
                    for clave, codigo in relaciones.items():
                        if codigo == '__IGNORAR__':
                            filas.append({'Nombre Zoom': clave, 'Alumno': '🚫 Ignorado'})
                        else:
                            alumno_match = df_alumnos[df_alumnos['Codigo'] == codigo]
                            nombre_alumno = alumno_match['Nombre'].values[0] if len(alumno_match) else codigo
                            filas.append({'Nombre Zoom': clave, 'Alumno': nombre_alumno})
                    st.dataframe(pd.DataFrame(filas), use_container_width=True, hide_index=True)
        else:
            st.info("No hay historial guardado aún.")

if __name__ == "__main__":
    main()
