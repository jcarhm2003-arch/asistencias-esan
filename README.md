# 📋 Control de Asistencias ESAN

App web para cruzar asistencias de Zoom con el Excel maestro de alumnos.

---

## 🚀 Cómo desplegar (una sola vez)

### Paso 1 — Crear cuenta en GitHub
1. Ve a https://github.com y crea una cuenta gratuita
2. Crea un repositorio nuevo llamado `asistencias-esan` (público o privado)

### Paso 2 — Subir los archivos
Sube estos 3 archivos a tu repositorio:
- `app.py`
- `requirements.txt`
- `README.md`

Puedes hacerlo directamente desde la web de GitHub (botón "Add file" → "Upload files")

### Paso 3 — Crear cuenta en Streamlit Cloud
1. Ve a https://streamlit.io/cloud
2. Regístrate con tu cuenta de GitHub
3. Haz clic en "New app"
4. Selecciona tu repositorio `asistencias-esan`
5. En "Main file path" escribe: `app.py`
6. Haz clic en "Deploy"

En 2-3 minutos tendrás tu app en una URL como:
`https://asistencias-esan-XXXXX.streamlit.app`

---

## 📖 Cómo usar la app

### Para cada curso nuevo:
1. Sube el Excel maestro de asistencias (panel izquierdo)
2. Ve a la pestaña **"Procesar sesión"**
3. Sube el CSV de Zoom de la sesión 1 y ponle una etiqueta (ej: `16/06`)
4. Haz clic en **"Procesar sesión"**
5. Revisa los matches automáticos
6. Para los pendientes: selecciona el alumno correcto en el dropdown
7. Haz clic en **"Reprocesar"** para confirmar, luego **"Confirmar sesión"**
8. Repite desde el paso 3 para la sesión 2, 3, etc.
9. Al terminar todas las sesiones, ve a **"Exportar"** y descarga el Excel

### Lógica de matching (en orden):
1. **Por correo** — si el correo del participante tiene `codigo@esan.edu.pe`, match directo
2. **Por historial** — si ese nombre Zoom ya fue enlazado antes en este curso, usa la relación guardada
3. **Por palabras** — si 2 o más palabras coinciden con el nombre en el Excel, match automático
4. **Manual** — si no coincide nada, aparece en la lista de pendientes para enlazar a mano

### Notas importantes:
- Los nombres de sala (`Sala Online 153 ESAN`) se ignoran automáticamente
- Si un participante entró y salió varias veces, se cuenta como 1 asistencia
- Las relaciones manuales se guardan y se sugieren en futuras sesiones del mismo curso
- Cada curso maneja su propio historial (no se mezclan)
- Si alguien no aparece en el CSV de esa sesión, figura como **F** aunque haya asistido a otras

---

## 📁 Estructura de archivos

```
asistencias-esan/
├── app.py                      # App principal
├── requirements.txt            # Dependencias
├── README.md                   # Este archivo
└── historial_relaciones.json   # Se crea automáticamente
```

---

## ⚠️ Limitación del historial en Streamlit Cloud

Streamlit Community Cloud **no persiste archivos entre reinicios del servidor**. El historial de relaciones se guarda durante la sesión activa, pero si el servidor se reinicia (cada ~7 días de inactividad), se pierde.

**Solución recomendada:** Al terminar cada sesión de trabajo, descarga el `historial_relaciones.json` desde la pestaña Exportar (próxima versión) o cópialo manualmente al repo de GitHub.

Si esto es un problema frecuente, avísame y lo conectamos a Google Drive para persistencia permanente.
