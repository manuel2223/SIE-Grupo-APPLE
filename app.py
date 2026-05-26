import streamlit as st
import pandas as pd
import random
import tempfile
import os
from fpdf import FPDF
import gspread
import re
from oauth2client.service_account import ServiceAccountCredentials

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Generador SAEL Web", page_icon="📝", layout="wide")

st.markdown("""
    <style>
    /* Quitamos el fondo forzado de .main para respetar tu modo oscuro */
    .stButton>button { width: 100%; border-radius: 10px; height: 3em; font-weight: bold; }
    .stDownloadButton>button { background-color: #2e7d32 !important; color: white !important; }
    
    /* Estilo adaptable para el menú de la izquierda */
    div[role="radiogroup"] label {
        padding: 10px;
        background-color: rgba(128, 128, 128, 0.1); /* Gris semi-transparente */
        border: 1px solid rgba(128, 128, 128, 0.3);
        border-radius: 8px;
        margin-bottom: 5px;
        transition: 0.3s;
    }
    div[role="radiogroup"] label:hover {
        background-color: rgba(128, 128, 128, 0.2);
        border-color: #4ade80; /* Borde verdoso al pasar el ratón */
    }
    </style>
    """, unsafe_allow_html=True)

# ==========================================
# 1. LÓGICA DE PDF (CON CORRECCIÓN DE MÁRGENES)
# ==========================================
class PDF_Examen(FPDF):
    def __init__(self, tipo_examen="", es_plantilla=False):
        super().__init__()
        self.tipo_examen = tipo_examen
        self.es_plantilla = es_plantilla
        self.add_font('Verdana', '', 'verdana.ttf', uni=True)
        self.add_font('Verdana', 'B', 'verdanab.ttf', uni=True)

    def header(self):
        try: self.image('Imagen1.png', 30, 6, 65) 
        except: pass 
            
        if self.page_no() > 1 and not self.es_plantilla:
            self.set_y(28)
            if getattr(self, 'tipo_examen', ''):
                self.set_font('Verdana', 'B', 12)
                self.set_left_margin(0)
                self.set_right_margin(0)
                self.set_x(0)
                self.cell(210, 10, f'MODELO {str(self.tipo_examen).upper()}', 0, 1, 'C')
                self.set_left_margin(30)
                self.set_right_margin(20)
            self.ln(2)
            self.set_x(10)
    def escribir_instrucciones_markdown(self, texto):
        """Traductor en tiempo real de Markdown a FPDF"""
        parrafos = texto.split('\n')
        for parrafo in parrafos:
            # Cortamos la frase en fragmentos aislando los que tienen símbolos
            fragmentos = re.split(r'(\*_.*?_\*|_\*.*?\*_|\*.*?\*|_.*?_)', parrafo)
            
            for frag in fragmentos:
                if not frag: continue
                
                # Detectamos qué símbolos tiene el fragmento para aplicarle la fuente
                if (frag.startswith('*_') and frag.endswith('_*')) or (frag.startswith('_*') and frag.endswith('*_')):
                    self.set_font('Verdana', 'BU', 11) # B = Bold, U = Underline
                    texto_limpio = frag[2:-2]
                elif frag.startswith('*') and frag.endswith('*'):
                    self.set_font('Verdana', 'B', 11)
                    texto_limpio = frag[1:-1]
                elif frag.startswith('_') and frag.endswith('_'):
                    self.set_font('Verdana', 'U', 11)
                    texto_limpio = frag[1:-1]
                else:
                    self.set_font('Verdana', '', 11)
                    texto_limpio = frag
                
                # Usamos write() en vez de multi_cell() para que las palabras se escriban seguidas
                self.write(7, texto_limpio)
            
            self.ln(7) # Salto de línea al terminar el párrafo

def generar_archivos_examen(lista_preguntas, tipo_examen, instrucciones):
    preguntas_normales = []
    preguntas_reserva = []

    for fila in lista_preguntas:
        for i in range(1, 21): 
            col_enunciado = f'Enunciado de la Pregunta {i}'
            col_enunciado_reserva = f'Enunciado de la Pregunta de Reserva {i}'

            if col_enunciado in fila and fila[col_enunciado]:
                correcta = fila.get(f'Opción CORRECTA (Pregunta {i})', '')
                inc1 = fila.get(f'Opción Falsa 1 (Pregunta {i})', '')
                inc2 = fila.get(f'Opción Falsa 2 (Pregunta {i}) - Opcional', '')
                inc3 = fila.get(f'Opción Falsa 3 (Pregunta {i}) - Opcional', '')
                
                justificacion = ""
                for k, v in fila.items():
                    k_lower = str(k).lower()
                    if "justificaci" in k_lower and str(i) in k_lower and "reserv" not in k_lower:
                        justificacion = v
                        break

                opciones = [{'texto': correcta, 'es_correcta': True}, {'texto': inc1, 'es_correcta': False}]
                if inc2: opciones.append({'texto': inc2, 'es_correcta': False})
                if inc3: opciones.append({'texto': inc3, 'es_correcta': False})
                    
                preguntas_normales.append({'enunciado': str(fila[col_enunciado]).replace('\n', ' '), 'opciones': opciones, 'justificacion': justificacion})

            if col_enunciado_reserva in fila and fila[col_enunciado_reserva]:
                correcta_reserva = fila.get(f'Opción CORRECTA (Pregunta de Reserva {i})', '')
                inc4 = fila.get(f'Opción Falsa 1 (Pregunta de Reserva {i})', '')
                inc5 = fila.get(f'Opción Falsa 2 (Pregunta de Reserva {i}) - Opcional', '')
                inc6 = fila.get(f'Opción Falsa 3 (Pregunta de Reserva {i}) - Opcional', '')
                
                justificacion_reserva = ""
                for k, v in fila.items():
                    k_lower = str(k).lower()
                    if "justificaci" in k_lower and str(i) in k_lower and "reserv" in k_lower:
                        justificacion_reserva = v
                        break

                opciones_res = [{'texto': correcta_reserva, 'es_correcta': True}, {'texto': inc4, 'es_correcta': False}]
                if inc5: opciones_res.append({'texto': inc5, 'es_correcta': False})
                if inc6: opciones_res.append({'texto': inc6, 'es_correcta': False})
                    
                preguntas_reserva.append({'enunciado': str(fila[col_enunciado_reserva]).replace('\n', ' '), 'opciones': opciones_res, 'justificacion': justificacion_reserva})

    if not preguntas_normales: return None, None

    random.shuffle(preguntas_normales)
    random.shuffle(preguntas_reserva)
    letras = ['A', 'B', 'C', 'D']
    plantilla_respuestas = {}
    plantilla_respuestas_reserva = {}

    pdf = PDF_Examen(tipo_examen, es_plantilla=False)
    pdf.add_page()
    pdf.set_y(40)
    pdf.set_left_margin(30)
    pdf.set_right_margin(20)
    if instrucciones.strip():
        pdf.escribir_instrucciones_markdown(instrucciones)
    
    pdf.add_page()
    pdf.set_font("Verdana", 'BU', 12)
    pdf.set_left_margin(0); pdf.set_right_margin(0); pdf.set_x(0)
    pdf.cell(210, 10, 'PREGUNTAS', 0, 1, 'C')
    pdf.set_left_margin(30); pdf.set_right_margin(20)
    pdf.ln(5)
    pdf.set_font("Verdana", size=11)

    for numero, pregunta in enumerate(preguntas_normales, start=1):
        texto_enun = f"{numero}.- {pregunta['enunciado']}"
        altura_estimada = (((len(texto_enun) // 70) + 1) * 8) + sum((((len(op['texto']) // 60) + 1) * 6) for op in pregunta['opciones']) + 18 
        if pdf.get_y() + altura_estimada > 255: pdf.add_page()

        pdf.set_x(30)
        pdf.multi_cell(0, 6, f"{numero}.- {pregunta['enunciado']}")
        pdf.ln(4)
        random.shuffle(pregunta['opciones'])
        for indice, opcion in enumerate(pregunta['opciones']):
            letra = letras[indice]
            y_actual = pdf.get_y()
            pdf.set_left_margin(44)
            pdf.set_xy(38, y_actual)
            pdf.cell(6, 6, f"{letra}.")
            pdf.multi_cell(0, 6, opcion['texto'])
            pdf.set_left_margin(30)
            if opcion['es_correcta']: 
                plantilla_respuestas[numero] = {'letra': letra, 'justificacion': pregunta.get('justificacion', '')}
        pdf.ln(4)

    if preguntas_reserva:
        pdf.add_page()
        pdf.set_font("Verdana", 'BU', 12)
        pdf.set_left_margin(0); pdf.set_right_margin(0); pdf.set_x(0)
        pdf.cell(210, 10, 'PREGUNTAS DE RESERVA', 0, 1, 'C')
        pdf.set_left_margin(30); pdf.set_right_margin(20)
        pdf.ln(5)
        pdf.set_font("Verdana", size=10.5)

        for numero, pregunta in enumerate(preguntas_reserva, start=1):
            texto_enun = f"{numero}.- {pregunta['enunciado']}"
            altura_estimada = (((len(texto_enun) // 75) + 1) * 8) + sum((((len(op['texto']) // 65) + 1) * 6) for op in pregunta['opciones']) + 12
            if pdf.get_y() + altura_estimada > 270: pdf.add_page()

            pdf.set_x(30)
            pdf.multi_cell(0, 6, f"{numero}.- {pregunta['enunciado']}")
            pdf.ln(4)
            random.shuffle(pregunta['opciones'])
            for indice, opcion in enumerate(pregunta['opciones']):
                letra = letras[indice]
                
                # 1. Guardamos la altura actual antes de escribir nada
                y_actual = pdf.get_y()
                
                # 2. Cambiamos el margen PRIMERO para que el texto largo sepa dónde rebotar (38 + 6 = 44)
                pdf.set_left_margin(44)
                
                # 3. Retrocedemos el lápiz manualmente a 38 (fuera del margen) para poner la letra
                pdf.set_xy(38, y_actual)
                pdf.cell(6, 6, f"{letra}.") 
                
                # 4. Imprimimos el texto. La 1ª línea sale junto a la letra, la 2ª rebota en 44.
                pdf.multi_cell(0, 6, opcion['texto'])
                
                # 5. Restauramos el margen a su sitio original (30) para la siguiente línea
                pdf.set_left_margin(30)
                
                if opcion['es_correcta']: 
                    plantilla_respuestas_reserva[numero] = {'letra': letra, 'justificacion': pregunta.get('justificacion', '')}
            pdf.ln(4)

    # --- CREACIÓN DE LA PLANTILLA ---
    pdf_plantilla = PDF_Examen(tipo_examen, es_plantilla=True)
    pdf_plantilla.add_page()
    pdf_plantilla.set_y(28) # 🟢 EMPUJÓN PARA NO PISAR EL LOGO
    pdf_plantilla.set_font("Verdana", 'B', 14)
    pdf_plantilla.cell(0, 10, f'PLANTILLA DE CORRECCION - MODELO {str(tipo_examen).upper()}', 0, 1, 'C')
    pdf_plantilla.ln(10)
    pdf_plantilla.set_font("Verdana", size=12)
    
    for numero, respuestas in plantilla_respuestas.items():
        if respuestas['justificacion']: pdf_plantilla.multi_cell(0, 8, f"Pregunta {numero}  -------  Respuesta: {respuestas['letra']}\n   Justificación: {respuestas['justificacion']}")
        else: pdf_plantilla.cell(0, 8, f"Pregunta {numero}  -------  Respuesta: {respuestas['letra']}", 0, 1)

    if plantilla_respuestas_reserva:
        pdf_plantilla.add_page()
        pdf_plantilla.set_y(28) # 🟢 EMPUJÓN PARA NO PISAR EL LOGO
        pdf_plantilla.set_font("Verdana", 'B', 14)
        pdf_plantilla.cell(0, 10, f'PREGUNTAS DE RESERVA - MODELO {str(tipo_examen).upper()}', 0, 1, 'C')
        pdf_plantilla.ln(10)
        pdf_plantilla.set_font("Verdana", size=12)
        for numero, respuestas in plantilla_respuestas_reserva.items():
            if respuestas['justificacion']: pdf_plantilla.multi_cell(0, 8, f"Pregunta de Reserva {numero}  -------  Respuesta: {respuestas['letra']}\n   Justificación: {respuestas['justificacion']}")
            else: pdf_plantilla.cell(0, 8, f"Pregunta de Reserva {numero}  -------  Respuesta: {respuestas['letra']}", 0, 1)

    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_ex, \
         tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pl:
        
        pdf.output(tmp_ex.name)
        pdf_plantilla.output(tmp_pl.name)
        
        with open(tmp_ex.name, "rb") as f: bytes_examen = f.read()
        with open(tmp_pl.name, "rb") as f: bytes_plantilla = f.read()
        
    os.remove(tmp_ex.name)
    os.remove(tmp_pl.name)
    
    return bytes_examen, bytes_plantilla


# ==========================================
# 2. CONEXIÓN A LA NUBE
# ==========================================
def obtener_datos():
    try:
        scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/drive']
        creds_dict = dict(st.secrets["gcp_service_account"])
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        cliente = gspread.authorize(creds)
        hoja = cliente.open("Banco de Preguntas Examenes SAEL (respuestas)").sheet1
        
        datos_brutos = hoja.get_all_values() 
        encabezados_originales = datos_brutos[0] 
        filas = datos_brutos[1:]             
        
        encabezados_unicos = []
        conteos = {}
        for col in encabezados_originales:
            if col in conteos:
                conteos[col] += 1
                encabezados_unicos.append(f"{col}_{conteos[col]}")
            else:
                conteos[col] = 0
                encabezados_unicos.append(col)
        
        df = pd.DataFrame(filas, columns=encabezados_unicos)
        return df, hoja
    except Exception as e:
        st.error(f"Error de conexión: {e}")
        return None, None


# ==========================================
# 3. INTERFAZ WEB (STREAMLIT)
# ==========================================
st.title("Generador de Exámenes SAEL")
st.sidebar.header("⚙️ Panel de Control")

if 'df' not in st.session_state:
    if st.sidebar.button("🚀 CONECTAR A LA NUBE"):
        with st.spinner("Accediendo al Banco de Preguntas..."):
            df, hoja = obtener_datos()
            if df is not None:
                st.session_state.df = df
                st.session_state.hoja = hoja
                st.success("¡Conexión Exitosa!")

if 'df' in st.session_state:
    convocatorias = [c for c in st.session_state.df['Código de la Convocatoria'].unique() if str(c).strip() != '']
    st.session_state.conv_activa = st.sidebar.selectbox("Selecciona Convocatoria", convocatorias)
    
    if "ultima_conv" not in st.session_state:
        st.session_state.ultima_conv = st.session_state.conv_activa

    if st.session_state.ultima_conv != st.session_state.conv_activa:
        st.session_state.bytes_ex = None
        st.session_state.bytes_pl = None
        st.session_state.ultima_conv = st.session_state.conv_activa
        
    tab1, tab2 = st.tabs(["📄 Generador", "🛠️ Gestor de Preguntas"])

    # --- PESTAÑA 1: GENERAR ---
    with tab1:
        col_izq, col_der = st.columns([1, 1])
        with col_izq:
            modelo = st.text_input("Modelo de Examen (A, B, C...)", "A")
            st.caption("💡 **Ayuda de formato:** Usa `*texto*` para **negrita**, `_texto_` para subrayado, y `*_texto_*` para ambas.")
            instrucciones = st.text_area("Instrucciones de Portada (Opcional)", height=200)
        
        with col_der:
            st.info("Al pulsar el botón, se mezclará el banco de preguntas y se generarán los documentos finales listos para imprimir.")
            
            if "bytes_ex" not in st.session_state:
                st.session_state.bytes_ex = None
                st.session_state.bytes_pl = None
                st.session_state.modelo_guardado = ""

            if st.button("GENERAR EXAMEN COMPLETO"):
                df_filtrado = st.session_state.df[st.session_state.df['Código de la Convocatoria'] == st.session_state.conv_activa]
                diccionario_preguntas = df_filtrado.to_dict('records')
                
                with st.spinner("Construyendo PDFs aleatorios..."):
                    b_ex, b_pl = generar_archivos_examen(diccionario_preguntas, modelo, instrucciones)
                    st.session_state.bytes_ex = b_ex
                    st.session_state.bytes_pl = b_pl
                    st.session_state.modelo_guardado = modelo

            if st.session_state.bytes_ex and st.session_state.bytes_pl:
                st.success("¡Documentos listos! Puedes descargar ambos archivos.")
                
                d_col1, d_col2 = st.columns(2)
                with d_col1:
                    st.download_button(
                        label="📄 DESCARGAR EXAMEN", data=st.session_state.get("bytes_ex"),
                        file_name=f"Examen_{st.session_state.conv_activa}_modelo-{st.session_state.modelo_guardado}.pdf",
                        mime="application/pdf", key="btn_descarga_examen"
                    )
                with d_col2:
                    st.download_button(
                        label="✅ DESCARGAR PLANTILLA", data=st.session_state.get("bytes_pl"),
                        file_name=f"Plantilla_{st.session_state.conv_activa}_modelo-{st.session_state.modelo_guardado}.pdf",
                        mime="application/pdf", key="btn_descarga_plantilla"
                    )

    # --- PESTAÑA 2: GESTOR MAESTRO-DETALLE ---
    with tab2:
        df_conv = st.session_state.df[st.session_state.df['Código de la Convocatoria'] == st.session_state.conv_activa]
        
        with st.expander("🚨 Eliminar Convocatoria Entera"):
            st.warning("Esta acción purgará físicamente todas las preguntas de esta convocatoria en Google Sheets.")
            confirmar_borrado = st.checkbox("Confirmo que deseo ELIMINAR el examen completo.")
            if st.button("🗑️ BORRAR EXAMEN", type="primary", disabled=not confirmar_borrado):
                with st.spinner("Borrando en la nube..."):
                    try:
                        indices_a_borrar = sorted(df_conv.index.tolist(), reverse=True)
                        for idx in indices_a_borrar: st.session_state.hoja.delete_rows(idx + 2)
                        st.success("Convocatoria eliminada.")
                    except Exception as e: st.error(f"Error: {e}")

        # 🟢 EXTRACCIÓN DE PREGUNTAS PARA LA LISTA
        preguntas_ui = []
        for index, row in df_conv.iterrows():
            fila_excel = index + 2
            correo = str(row.get('Dirección de correo electrónico', 'Correo no registrado'))

            # Bucle Preguntas Normales
            for i in range(1, 21):
                c_enun = f'Enunciado de la Pregunta {i}'
                if c_enun in row and str(row[c_enun]).strip():
                    preguntas_ui.append({
                        'label': f"[P{i}] {str(row[c_enun])[:35]}...", 'fila_excel': fila_excel, 'bloque': i, 'tipo': 'Normal', 'correo': correo,
                        'enunciado': str(row[c_enun]), 'correcta': str(row.get(f'Opción CORRECTA (Pregunta {i})', '')),
                        'f1': str(row.get(f'Opción Falsa 1 (Pregunta {i})', '')), 'f2': str(row.get(f'Opción Falsa 2 (Pregunta {i}) - Opcional', '')),
                        'f3': str(row.get(f'Opción Falsa 3 (Pregunta {i}) - Opcional', '')),
                        'c_enun': c_enun, 'c_corr': f'Opción CORRECTA (Pregunta {i})', 'c_f1': f'Opción Falsa 1 (Pregunta {i})',
                        'c_f2': f'Opción Falsa 2 (Pregunta {i}) - Opcional', 'c_f3': f'Opción Falsa 3 (Pregunta {i}) - Opcional'
                    })
            # Bucle Preguntas Reserva
            for i in range(1, 6):
                c_enun = f'Enunciado de la Pregunta de Reserva {i}'
                if c_enun in row and str(row[c_enun]).strip():
                    preguntas_ui.append({
                        'label': f"[R{i}] {str(row[c_enun])[:35]}...", 'fila_excel': fila_excel, 'bloque': i, 'tipo': 'Reserva', 'correo': correo,
                        'enunciado': str(row[c_enun]), 'correcta': str(row.get(f'Opción CORRECTA (Pregunta de Reserva {i})', '')),
                        'f1': str(row.get(f'Opción Falsa 1 (Pregunta de Reserva {i})', '')), 'f2': str(row.get(f'Opción Falsa 2 (Pregunta de Reserva {i}) - Opcional', '')),
                        'f3': str(row.get(f'Opción Falsa 3 (Pregunta de Reserva {i}) - Opcional', '')),
                        'c_enun': c_enun, 'c_corr': f'Opción CORRECTA (Pregunta de Reserva {i})', 'c_f1': f'Opción Falsa 1 (Pregunta de Reserva {i})',
                        'c_f2': f'Opción Falsa 2 (Pregunta de Reserva {i}) - Opcional', 'c_f3': f'Opción Falsa 3 (Pregunta de Reserva {i}) - Opcional'
                    })

        st.divider()

        if not preguntas_ui:
            st.info("Aún no hay preguntas para esta convocatoria.")
        else:
            # 🟢 LA MAGIA DEL DISEÑO MAESTRO-DETALLE
            col_lista, col_detalle = st.columns([1, 2])

            with col_lista:
                st.markdown("#### Banco de Preguntas")
                opciones = [p['label'] for p in preguntas_ui]
                # Radio button con estilo inyectado arriba para parecer una lista clickable
                seleccion = st.radio("Lista", opciones, label_visibility="collapsed")

            with col_detalle:
                if seleccion:
                    preg = next(p for p in preguntas_ui if p['label'] == seleccion)

                    st.markdown(f"#### Editando ({preg['tipo'].upper()}) - Fila {preg['fila_excel']} - Bloque {preg['bloque']}")
                    
                    # 🟢 EL CORREO BLOQUEADO
                    st.text_input("Autor (Correo original del formulario):", value=preg['correo'], disabled=True)

                    with st.form("form_edicion"):
                        n_enun = st.text_area("Enunciado:", value=preg['enunciado'], height=100)
                        n_corr = st.text_input("Opción CORRECTA:", value=preg['correcta'])
                        n_f1 = st.text_input("Falsa 1:", value=preg['f1'])
                        n_f2 = st.text_input("Falsa 2:", value=preg['f2'])
                        n_f3 = st.text_input("Falsa 3:", value=preg['f3'])

                        c1, c2 = st.columns(2)
                        btn_guardar = c1.form_submit_button("💾 GUARDAR EN LA NUBE", type="primary")
                        btn_eliminar = c2.form_submit_button("🗑️ ELIMINAR PREGUNTA")

                    if btn_guardar:
                        with st.spinner("Enviando cambios a Google Sheets..."):
                            try:
                                # Buscamos la columna exacta
                                col_enun_idx = st.session_state.df.columns.get_loc(preg['c_enun']) + 1
                                col_corr_idx = st.session_state.df.columns.get_loc(preg['c_corr']) + 1
                                col_f1_idx = st.session_state.df.columns.get_loc(preg['c_f1']) + 1
                                col_f2_idx = st.session_state.df.columns.get_loc(preg['c_f2']) + 1
                                col_f3_idx = st.session_state.df.columns.get_loc(preg['c_f3']) + 1
                                fila_ex = preg['fila_excel']

                                st.session_state.hoja.update_cell(fila_ex, col_enun_idx, n_enun)
                                st.session_state.hoja.update_cell(fila_ex, col_corr_idx, n_corr)
                                st.session_state.hoja.update_cell(fila_ex, col_f1_idx, n_f1)
                                st.session_state.hoja.update_cell(fila_ex, col_f2_idx, n_f2)
                                st.session_state.hoja.update_cell(fila_ex, col_f3_idx, n_f3)

                                st.success("¡Guardado! Refresca la conexión a la nube para ver los cambios reflejados.")
                            except Exception as e:
                                st.error(f"Error: {e}")

                    if btn_eliminar:
                        with st.spinner("Vaciando celdas en la nube..."):
                            try:
                                col_enun_idx = st.session_state.df.columns.get_loc(preg['c_enun']) + 1
                                col_corr_idx = st.session_state.df.columns.get_loc(preg['c_corr']) + 1
                                col_f1_idx = st.session_state.df.columns.get_loc(preg['c_f1']) + 1
                                col_f2_idx = st.session_state.df.columns.get_loc(preg['c_f2']) + 1
                                col_f3_idx = st.session_state.df.columns.get_loc(preg['c_f3']) + 1
                                fila_ex = preg['fila_excel']

                                st.session_state.hoja.update_cell(fila_ex, col_enun_idx, "")
                                st.session_state.hoja.update_cell(fila_ex, col_corr_idx, "")
                                st.session_state.hoja.update_cell(fila_ex, col_f1_idx, "")
                                st.session_state.hoja.update_cell(fila_ex, col_f2_idx, "")
                                st.session_state.hoja.update_cell(fila_ex, col_f3_idx, "")

                                st.success("¡Pregunta eliminada! Refresca la conexión a la nube para limpiar la lista.")
                            except Exception as e:
                                st.error(f"Error: {e}")