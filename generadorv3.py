import pandas as pd
import random
import os
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from fpdf import FPDF
from docxtpl import DocxTemplate, RichText
from docx2pdf import convert
from pypdf import PdfWriter
import customtkinter as ctk
from tkinter import messagebox, filedialog
import re
import sys
import os

# Escudo anti-errores para el modo --noconsole
if sys.stdout is None:
    sys.stdout = open(os.devnull, 'w')
if sys.stderr is None:
    sys.stderr = open(os.devnull, 'w')

def resource_path(relative_path):
    """ Obtiene la ruta absoluta para recursos dentro del .exe de PyInstaller """
    try:
        # PyInstaller crea una carpeta temporal en AppData\Local\Temp\_MEIxxxxxx
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

    
# --- CONFIGURACIÓN DE LA NUBE ---
NOMBRE_EXCEL_NUBE = "Banco de Preguntas Examenes SAEL (respuestas)"

# ==========================================
# 1. FUNCIONES DE GENERACIÓN DE PDF (Sin cambios)
# ==========================================
class PDF(FPDF):
    def __init__(self, tipo_examen=""):
        super().__init__()
        self.tipo_examen = tipo_examen
        
        self.set_left_margin(30)
        self.set_right_margin(20)
        
        self.add_font('Verdana', '', resource_path('verdana.ttf'), uni = True)
        self.add_font('Verdana', 'B', resource_path('verdanab.ttf'), uni = True)

    def header(self):
        try:
            # 1. POSICIONAMOS EL LOGO CENTRADO
            # X = 59.7 (para que quede en el centro exacto del folio)
            # Y = 10 (margen superior)
            # Ancho = 90.6 (tu medida de Word)
            self.image(resource_path('Imagen1.png'), 30, 6, 65) 
        except:
            pass 
            
        # 2. POSICIONAMOS EL MODELO JUSTO DEBAJO
        # Si el logo empieza en Y=10 y mide aprox 30mm de alto, 
        # bajamos el "lápiz" a Y=45 para que no se pisen.
        self.set_y(28)
        
        if getattr(self, 'tipo_examen', ''):
            self.set_font('Verdana', 'B', 12)
            
            # 1. Apagamos los márgenes
            self.set_left_margin(0)
            self.set_right_margin(0)
            
            # 2. Nos pegamos al borde izquierdo absoluto
            self.set_x(0)
            
            # 3. Usamos el ancho total del folio (210)
            self.cell(210, 10, f'MODELO {str(self.tipo_examen).upper()}', 0, 1, 'C')
            
            # 4. Volvemos a encender los márgenes oficiales
            self.set_left_margin(30)
            self.set_right_margin(20)
            
        # 3. ESPACIO DE SEGURIDAD
        # Dejamos un hueco antes de que empiecen las preguntas
        self.ln(2)
        self.set_x(10) # ¡SEGURO ANTI-ERRORES! Forzamos el lápiz a la izquierda

    def footer(self):
        self.set_y(-15)
        self.set_font('Verdana', '', 8) 
        # texto_pagina = 'Página ' + str(self.page_no()) + '/{nb}'
        # self.cell(0, 10, texto_pagina, 0, 0, 'C')
        
        
        
def generar_examen_y_pdf(lista_preguntas, conv, ruta_destino, tipo_examen=""):
    preguntas_normales = []
    preguntas_reserva = []

# 1. EXTRACCIÓN DE DATOS
    for fila in lista_preguntas:
        for i in range(1, 21): 
            col_enunciado = f'Enunciado de la Pregunta {i}'
            col_enunciado_reserva = f'Enunciado de la Pregunta de Reserva {i}'

            # --- PREGUNTAS NORMALES ---
            if col_enunciado in fila and fila[col_enunciado]:
                enunciado = fila[col_enunciado]
                correcta = fila.get(f'Opción CORRECTA (Pregunta {i})', '')
                inc1 = fila.get(f'Opción Falsa 1 (Pregunta {i})', '')
                inc2 = fila.get(f'Opción Falsa 2 (Pregunta {i}) - Opcional', '')
                inc3 = fila.get(f'Opción Falsa 3 (Pregunta {i}) - Opcional', '')
                
                # 🟢 NUEVO: Buscador inteligente para la justificación normal 'i'
                justificacion = ""
                for k, v in fila.items():
                    k_lower = str(k).lower()
                    if "justificaci" in k_lower and str(i) in k_lower and "reserv" not in k_lower:
                        justificacion = v
                        break

                opciones_mezclar = [{'texto': correcta, 'es_correcta': True}, {'texto': inc1, 'es_correcta': False}]
                if inc2: opciones_mezclar.append({'texto': inc2, 'es_correcta': False})
                if inc3: opciones_mezclar.append({'texto': inc3, 'es_correcta': False})
                    
                preguntas_normales.append({
                    'enunciado': str(enunciado).replace('\n', ' '), 
                    'opciones': opciones_mezclar,
                    'justificacion': justificacion # 🟢 Guardamos la justificación limpia
                })

            # --- PREGUNTAS DE RESERVA ---
            if col_enunciado_reserva in fila and fila[col_enunciado_reserva]:
                enunciado_reserva = fila[col_enunciado_reserva]
                correcta_reserva = fila.get(f'Opción CORRECTA (Pregunta de Reserva {i})', '')
                inc4 = fila.get(f'Opción Falsa 1 (Pregunta de Reserva {i})', '')
                inc5 = fila.get(f'Opción Falsa 2 (Pregunta de Reserva {i}) - Opcional', '')
                inc6 = fila.get(f'Opción Falsa 3 (Pregunta de Reserva {i}) - Opcional', '')
                
                # 🟢 NUEVO: Buscador inteligente para la justificación de reserva 'i'
                justificacion_reserva = ""
                for k, v in fila.items():
                    k_lower = str(k).lower()
                    if "justificaci" in k_lower and str(i) in k_lower and "reserv" in k_lower:
                        justificacion_reserva = v
                        break

                opciones_mezclar_reserva = [{'texto': correcta_reserva, 'es_correcta': True}, {'texto': inc4, 'es_correcta': False}]
                if inc5: opciones_mezclar_reserva.append({'texto': inc5, 'es_correcta': False})
                if inc6: opciones_mezclar_reserva.append({'texto': inc6, 'es_correcta': False})
                    
                preguntas_reserva.append({
                    'enunciado': str(enunciado_reserva).replace('\n', ' '), 
                    'opciones': opciones_mezclar_reserva,
                    'justificacion': justificacion_reserva # 🟢 Guardamos la justificación limpia
                })

    if not preguntas_normales: return False

    random.shuffle(preguntas_normales)
    random.shuffle(preguntas_reserva)
    letras = ['A', 'B', 'C', 'D']
    plantilla_respuestas = {}
    plantilla_respuestas_reserva = {}

    # 2. CREACIÓN DEL PDF DEL EXAMEN
    pdf_examen = PDF(tipo_examen)
    pdf_examen.add_page()
    
# --- TÍTULO DE PREGUNTAS (Opción Nuclear) ---
    pdf_examen.set_font("Verdana", 'BU', 12)
    
    # Apagamos los márgenes
    pdf_examen.set_left_margin(0)
    pdf_examen.set_right_margin(0)
    pdf_examen.set_x(0)
    
    # Imprimimos usando el ancho total del folio (210)
    pdf_examen.cell(210, 10, 'PREGUNTAS', 0, 1, 'C')
    
    # Restauramos márgenes para que los enunciados no se rompan
    pdf_examen.set_left_margin(30)
    pdf_examen.set_right_margin(20)
    
    pdf_examen.ln(5)
    pdf_examen.set_font("Verdana", size=11)

    # --- BUCLE 1: NORMALES ---
    for numero, pregunta in enumerate(preguntas_normales, start=1):
        
        # 🟢 RADAR AJUSTADO A MARGEN 30 🟢
        # Al tener más margen, caben menos letras: bajamos a 75 y 65 caracteres
        texto_enun = f"{numero}.- {pregunta['enunciado']}"
        lineas_enun = (len(texto_enun) // 70) + 1 
        altura_estimada = (lineas_enun * 8)
        
        for opcion in pregunta['opciones']:
            lineas_op = (len(opcion['texto']) // 60) + 1 
            altura_estimada += (lineas_op * 6)
            
        altura_estimada += 18 
        
        if pdf_examen.get_y() + altura_estimada > 255:
            pdf_examen.add_page()

        pdf_examen.set_x(30) # Margen izquierdo 30
        # CAMBIO: Formato 1.- 
        pdf_examen.multi_cell(0, 8, f"{numero}.- {pregunta['enunciado']}")
        
        opciones = pregunta['opciones']
        random.shuffle(opciones)
        for indice, opcion in enumerate(opciones):
            letra = letras[indice]
            # Ponemos la respuesta con un pequeño sangrado respecto al 30 (ej: 38)
            pdf_examen.set_x(38) 
            # CAMBIO: Formato A. 
            pdf_examen.multi_cell(0, 6, f"{letra}. {opcion['texto']}")
            if opcion['es_correcta']: 
                plantilla_respuestas[numero] = {
                    'letra': letra,
                    'justificacion': pregunta.get('justificacion', '') 
                }
            
        pdf_examen.ln(4)

    # --- BUCLE 2: RESERVAS ---
    if preguntas_reserva:
        # --- TÍTULO DE RESERVA (Opción Nuclear) ---
        pdf_examen.add_page()
        pdf_examen.set_font("Verdana", 'BU', 12)
        
        pdf_examen.set_left_margin(0)
        pdf_examen.set_right_margin(0)
        pdf_examen.set_x(0)
        
        pdf_examen.cell(210, 10, 'PREGUNTAS DE RESERVA', 0, 1, 'C')
        
        pdf_examen.set_left_margin(30)
        pdf_examen.set_right_margin(20)
        
        pdf_examen.ln(5)
        pdf_examen.set_font("Verdana", size=10.5)

        for numero, pregunta in enumerate(preguntas_reserva, start=1):
            
            # Radar ajustado para reservas
            texto_enun = f"{numero}.- {pregunta['enunciado']}"
            lineas_enun = (len(texto_enun) // 75) + 1
            altura_estimada = (lineas_enun * 8)
            for opcion in pregunta['opciones']:
                lineas_op = (len(opcion['texto']) // 65) + 1
                altura_estimada += (lineas_op * 6)
            altura_estimada += 12
            
            if pdf_examen.get_y() + altura_estimada > 270:
                pdf_examen.add_page()

            pdf_examen.set_x(30)
            # CAMBIO: Formato 1.-
            pdf_examen.multi_cell(0, 8, f"{numero}.- {pregunta['enunciado']}")
            
            opciones = pregunta['opciones']
            random.shuffle(opciones)
            for indice, opcion in enumerate(opciones):
                letra = letras[indice]
                pdf_examen.set_x(38) 
                # CAMBIO: Formato A.
                pdf_examen.multi_cell(0, 6, f"{letra}. {opcion['texto']}")
                if opcion['es_correcta']: 
                    plantilla_respuestas_reserva[numero] = {
                        'letra': letra,
                        'justificacion': pregunta.get('justificacion', '') 
                    }
                
            pdf_examen.ln(4)

    pdf_examen.output("Examen_Oficial.pdf")

    # 3. CREACIÓN DE LA PLANTILLA DE CORRECCIÓN
    pdf_plantilla = PDF(tipo_examen)
    pdf_plantilla.add_page()
    pdf_plantilla.set_font("Verdana", 'B', 14)
    pdf_plantilla.cell(0, 10, 'PLANTILLA DE CORRECCION', 0, 1, 'C')
    pdf_plantilla.ln(10)
    pdf_plantilla.set_font("Verdana", size=12)
    
    for numero, respuestas in plantilla_respuestas.items():
        letra = respuestas['letra']
        justificacion = respuestas['justificacion']
        if justificacion:
            pdf_plantilla.multi_cell(0, 8, f"Pregunta {numero}  -------  Respuesta: {letra}\n   Justificación: {justificacion}")
        else:
            pdf_plantilla.cell(0, 8, f"Pregunta {numero}  -------  Respuesta: {letra}", 0, 1)

    if plantilla_respuestas_reserva:
        pdf_plantilla.add_page()
        pdf_plantilla.set_font("Verdana", 'B', 14)
        pdf_plantilla.cell(0, 10, 'RESPUESTAS - PREGUNTAS DE RESERVA', 0, 1, 'C')
        pdf_plantilla.ln(10)
        pdf_plantilla.set_font("Verdana", size=12)
        
        for numero, respuestas in plantilla_respuestas_reserva.items():
            letra = respuestas['letra']
            justificacion = respuestas['justificacion']
            if justificacion:
                pdf_plantilla.multi_cell(0, 8, f"Pregunta de Reserva {numero}  -------  Respuesta: {letra}\n   Justificación: {justificacion}")
            else:
                pdf_plantilla.cell(0, 8, f"Pregunta de Reserva {numero}  -------  Respuesta: {letra}", 0, 1)

    # Guardamos la plantilla directamente en la carpeta que eligió el jefe
    ruta_plantilla = os.path.join(ruta_destino, f"PlantillaCorreccion_{conv}_modelo-{tipo_examen}.pdf")
    pdf_plantilla.output(ruta_plantilla)
    return True


def procesar_texto_enriquecido(texto):
    """
    Escanea el texto y convierte:
    *texto* -> Negrita
    _texto_ -> Subrayado
    *_texto_* -> Negrita Y Subrayado
    Todo forzado a Verdana 10.5 (size=21 en XML de Word)
    """
    rt = RichText()
    
    # Nuevo escáner más inteligente que pilla combinaciones
    partes = re.split(r'(\*_[^_]+_\*|\*[^*]+\*|_[^_]+_)', texto)
    
    for parte in partes:
        if not parte:
            continue
            
        # 1. NEGRITA Y SUBRAYADO a la vez (*_texto_*)
        if parte.startswith('*_') and parte.endswith('_*'):
            rt.add(parte[2:-2], bold=True, underline=True, font='Verdana', size=21)
            
        # 2. SOLO NEGRITA (*texto*)
        elif parte.startswith('*') and parte.endswith('*'):
            rt.add(parte[1:-1], bold=True, font='Verdana', size=21)
            
        # 3. SOLO SUBRAYADO (_texto_)
        elif parte.startswith('_') and parte.endswith('_'):
            rt.add(parte[1:-1], underline=True, font='Verdana', size=21)
            
        # 4. TEXTO NORMAL (Respeta los saltos de línea \n)
        else:
            rt.add(parte, font='Verdana', size=21)
            
    return rt

def generar_portada_desde_word(datos_variables):
    doc = DocxTemplate(resource_path("Plantilla_Portada.docx"))    
    # --- INTERCEPTAMOS LAS INSTRUCCIONES ---
    # Si hay instrucciones, las pasamos por nuestro traductor antes de inyectarlas
    if "instrucciones" in datos_variables and datos_variables["instrucciones"].strip():
        texto_crudo = datos_variables["instrucciones"]
        datos_variables["instrucciones"] = procesar_texto_enriquecido(texto_crudo)
    # ---------------------------------------
    
    doc.render(datos_variables)
    doc.save("portada_temp.docx")
    convert("portada_temp.docx", "portada_final.pdf")
    os.remove("portada_temp.docx")
    return "portada_final.pdf"

def fusionar_pdfs(ruta_portada, ruta_preguntas, nombre_final):
    writer = PdfWriter()
    for ruta in [ruta_portada, ruta_preguntas]:
        with open(ruta, "rb") as f: writer.append(f)
    with open(nombre_final, "wb") as f: writer.write(f)

# ==========================================
# 2. INTERFAZ GRÁFICA (App Principal)
# ==========================================
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class AppExamen(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Generador de Exámenes SAEL - Diputación de Cádiz")
        self.geometry("1000x750") # Ventana más ancha para el modo edición
        
        self.cliente = None
        self.hoja = None
        self.df_maestro = None
        self.encabezados = []
        self.pregunta_seleccionada = None

        # --- PANEL SUPERIOR (CONEXIÓN Y FILTRO) ---
        self.frame_top = ctk.CTkFrame(self)
        self.frame_top.pack(fill="x", padx=10, pady=10)
        
        self.btn_conectar = ctk.CTkButton(self.frame_top, text="1. CONECTAR A LA NUBE", command=self.conectar_nube, fg_color="#d35400")
        self.btn_conectar.pack(side="left", padx=10, pady=10)

        ctk.CTkLabel(self.frame_top, text="Convocatoria activa:").pack(side="left", padx=(20, 5))
        self.combo_convocatoria = ctk.CTkComboBox(self.frame_top, values=["Conecta primero..."], state="disabled", width=250, command=self.al_cambiar_convocatoria)
        self.combo_convocatoria.pack(side="left", padx=5)

        # --- SISTEMA DE PESTAÑAS ---
        self.tabview = ctk.CTkTabview(self)
        self.tabview.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        
        self.tab_generar = self.tabview.add("Generar Examen")
        self.tab_editar = self.tabview.add("Gestor de Preguntas")

        self.construir_tab_generar()
        self.construir_tab_editar()

    # --- PESTAÑA 1: GENERAR ---
    def construir_tab_generar(self):
        self.scroll_generar = ctk.CTkScrollableFrame(self.tab_generar)
        self.scroll_generar.pack(fill="both", expand=True, padx=10, pady=10)

        # --- CUADRO INFORMATIVO DE FORMATO ---
        self.frame_info = ctk.CTkFrame(self.scroll_generar, fg_color="#34495e")
        self.frame_info.pack(fill="x", padx=10, pady=10)
        
        info_text = (
            "Ayuda de Formato:\n"
            "• Para NEGRITA: pon el texto entre asteriscos -> *texto*\n"
            "• Para SUBRAYADO: pon el texto entre guiones bajos -> _texto_\n"
            "• Para AMBOS: combínalos así -> *_texto_*"
        )
        self.lbl_ayuda = ctk.CTkLabel(self.frame_info, text=info_text, justify="left", font=ctk.CTkFont(size=12))
        self.lbl_ayuda.pack(pady=10, padx=10)

        # --- CAMPOS DE DATOS ---
        self.inputs = {}
        # Aquí mantenemos el Modelo (Tipo A, B...)
        campos = [
            
            ("tipo_examen", "Modelo (A, B, C...)") # Este es el que genera Modelo A, B...
        ]
        
        for key, placeholder in campos:
            ctk.CTkLabel(self.scroll_generar, text=f"{placeholder}:").pack(anchor="w", padx=10)
            entry = ctk.CTkEntry(self.scroll_generar, width=400)
            entry.pack(pady=5, padx=10)
            self.inputs[key] = entry

        # --- EL MEGA CAMPO DE INSTRUCCIONES ---
        ctk.CTkLabel(self.scroll_generar, text="Instrucciones del Examen:").pack(anchor="w", padx=10)
        self.txt_instrucciones = ctk.CTkTextbox(self.scroll_generar, width=400, height=200) # Más alto
        self.txt_instrucciones.pack(pady=5, padx=10)

        self.btn_generar = ctk.CTkButton(self.scroll_generar, text="2. GENERAR PDF FINAL", 
                                         command=self.ejecutar_generacion, height=50, state="disabled")
        self.btn_generar.pack(pady=20)

    # --- PESTAÑA 2: GESTOR DE PREGUNTAS (MAESTRO-DETALLE) ---
    def construir_tab_editar(self):
        # Frame izquierdo (La Lista Maestra)
        self.frame_lista = ctk.CTkScrollableFrame(self.tab_editar, width=300)
        self.frame_lista.pack(side="left", fill="y", padx=5, pady=5)
        
        # Frame derecho (El Editor de Detalle)
        self.frame_detalle = ctk.CTkFrame(self.tab_editar)
        self.frame_detalle.pack(side="right", fill="both", expand=True, padx=5, pady=5)

        self.lbl_edit_titulo = ctk.CTkLabel(self.frame_detalle, text="Selecciona una pregunta de la izquierda", font=ctk.CTkFont(weight="bold", size=16))
        self.lbl_edit_titulo.pack(pady=10)

        # Campos de edición
        ctk.CTkLabel(self.frame_detalle, text="Enunciado:").pack(anchor="w", padx=10)
        self.txt_edit_enun = ctk.CTkTextbox(self.frame_detalle, height=80)
        self.txt_edit_enun.pack(fill="x", padx=10, pady=2)

        ctk.CTkLabel(self.frame_detalle, text="Correcta:").pack(anchor="w", padx=10)
        self.entry_edit_corr = ctk.CTkEntry(self.frame_detalle)
        self.entry_edit_corr.pack(fill="x", padx=10, pady=2)

        ctk.CTkLabel(self.frame_detalle, text="Falsa 1:").pack(anchor="w", padx=10)
        self.entry_edit_f1 = ctk.CTkEntry(self.frame_detalle)
        self.entry_edit_f1.pack(fill="x", padx=10, pady=2)

        ctk.CTkLabel(self.frame_detalle, text="Falsa 2:").pack(anchor="w", padx=10)
        self.entry_edit_f2 = ctk.CTkEntry(self.frame_detalle)
        self.entry_edit_f2.pack(fill="x", padx=10, pady=2)

        ctk.CTkLabel(self.frame_detalle, text="Falsa 3:").pack(anchor="w", padx=10)
        self.entry_edit_f3 = ctk.CTkEntry(self.frame_detalle)
        self.entry_edit_f3.pack(fill="x", padx=10, pady=2)

        self.btn_guardar_cambios = ctk.CTkButton(self.frame_detalle, text="💾 GUARDAR CAMBIOS EN LA NUBE", command=self.guardar_edicion, fg_color="#27ae60", hover_color="#2ecc71", state="disabled")
        self.btn_guardar_cambios.pack(pady=20)
        
        self.btn_eliminar = ctk.CTkButton(
            self.frame_detalle,
            text="🗑️ ELIMINAR PREGUNTA",
            command=self.eliminar_pregunta,
            fg_color="#c0392b",
            hover_color="#e74c3c",
            state="disabled"
        )
        self.btn_eliminar.pack(pady=5)

    # ==========================================
    # LÓGICA DE DATOS Y CONEXIÓN
    # ==========================================
    def conectar_nube(self):
        self.btn_conectar.configure(text="Conectando...", state="disabled")
        self.update()
        try:
            scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/drive']
            creds = ServiceAccountCredentials.from_json_keyfile_name(resource_path('examenes.json'), scope)
            self.cliente = gspread.authorize(creds)
            self.hoja = self.cliente.open(NOMBRE_EXCEL_NUBE).sheet1
            
            # --- NUEVA LECTURA A PRUEBA DE DUPLICADOS ---
            datos_brutos = self.hoja.get_all_values() # Leemos todo en bruto
            
            self.encabezados = datos_brutos[0] # La fila 1 (los nombres)
            filas = datos_brutos[1:] # El resto (los datos)
            
            # Pandas se traga las columnas repetidas sin quejarse
            self.df_maestro = pd.DataFrame(filas, columns=self.encabezados)
            self.datos_sheet = self.df_maestro.to_dict('records')
            
            # Sacamos las convocatorias (quitando posibles celdas vacías)
            convocatorias = [c for c in self.df_maestro['Código de la Convocatoria'].unique() if str(c).strip() != '']
            # ---------------------------------------------
            self.combo_convocatoria.configure(values=[str(c) for c in convocatorias], state="normal")
            
            self.btn_generar.configure(state="normal", fg_color="#2980b9")
            self.btn_conectar.configure(text="✅ Conectado", fg_color="#27ae60")
            messagebox.showinfo("Nube", "¡Conexión establecida con éxito!")
        except Exception as e:
            self.btn_conectar.configure(text="1. CONECTAR A LA NUBE", state="normal")
            messagebox.showerror("Error Nube", f"No se pudo conectar: {e}")

    def al_cambiar_convocatoria(self, nueva_convocatoria):
        # Limpiar la lista maestra anterior
        for widget in self.frame_lista.winfo_children():
            widget.destroy()

        # Rellenar la lista izquierda con las preguntas de esta convocatoria
        for index, fila_dict in enumerate(self.datos_sheet):
            if str(fila_dict.get('Código de la Convocatoria', '')) == nueva_convocatoria:
                fila_excel = index + 2 # La fila 1 son encabezados, el index 0 de get_all_records() es la Fila 2
                
                for i in range(1, 21):
                    col_enun = f'Enunciado de la Pregunta {i}'
                    if col_enun in fila_dict and fila_dict[col_enun]:
                        # Empaquetamos los datos en un diccionario
                        preg = {
                            'fila_excel': fila_excel,
                            'bloque': i,
                            'enunciado': fila_dict[col_enun],
                            'correcta': fila_dict.get(f'Opción CORRECTA (Pregunta {i})', ''),
                            'f1': fila_dict.get(f'Opción Falsa 1 (Pregunta {i})', ''),
                            'f2': fila_dict.get(f'Opción Falsa 2 (Pregunta {i}) - Opcional', ''),
                            'f3': fila_dict.get(f'Opción Falsa 3 (Pregunta {i}) - Opcional', '')
                        }
                        
                        # Crear botón en la lista (truncamos el texto para que quepa)
                        texto_btn = f"[P{i}] {str(preg['enunciado'])[:40]}..."
                        btn = ctk.CTkButton(self.frame_lista, text=texto_btn, anchor="w",
                                            fg_color="transparent", text_color=("black", "white"), hover_color="gray",
                                            command=lambda p=preg: self.cargar_detalle(p))
                        btn.pack(fill="x", pady=2)
                
                                # ===== PREGUNTAS DE RESERVA =====
                for j in range(1, 6):

                    col_enun_res = f'Enunciado de la Pregunta de Reserva {j}'

                    if col_enun_res in fila_dict and fila_dict[col_enun_res]:

                        preg = {
                            'fila_excel': fila_excel,
                            'bloque': j,
                            'tipo': 'reserva',
                            'enunciado': fila_dict[col_enun_res],
                            'correcta': fila_dict.get(f'Opción CORRECTA (Pregunta de Reserva {j})', ''),
                            'f1': fila_dict.get(f'Opción Falsa 1 (Pregunta de Reserva {j})', ''),
                            'f2': fila_dict.get(f'Opción Falsa 2 (Pregunta de Reserva {j}) - Opcional', ''),
                            'f3': fila_dict.get(f'Opción Falsa 3 (Pregunta de Reserva {j}) - Opcional', '')
                        }

                        texto_btn = f"[R{j}] {str(preg['enunciado'])[:40]}..."

                        btn = ctk.CTkButton(
                            self.frame_lista,
                            text=texto_btn,
                            anchor="w",
                            fg_color="transparent",
                            text_color=("black", "white"),
                            hover_color="#8e44ad",  # 💜 color distinto
                            command=lambda p=preg: self.cargar_detalle(p)
                        )
                        btn.pack(fill="x", pady=2)

    def cargar_detalle(self, preg):
        self.pregunta_seleccionada = preg
        tipo = preg.get('tipo', 'normal')
        # self.lbl_edit_titulo.configure(text=f"Editando Excel (Fila {preg['fila_excel']} - Bloque {preg['bloque']})")
        self.lbl_edit_titulo.configure(
        text=f"Editando ({tipo.upper()}) - Fila {preg['fila_excel']} - Bloque {preg['bloque']}")
        
        # Limpiar y rellenar campos
        self.txt_edit_enun.delete("1.0", "end")
        self.txt_edit_enun.insert("1.0", str(preg['enunciado']))
        
        entradas = [self.entry_edit_corr, self.entry_edit_f1, self.entry_edit_f2, self.entry_edit_f3]
        claves = ['correcta', 'f1', 'f2', 'f3']
        
        for entry, clave in zip(entradas, claves):
            entry.delete(0, "end")
            entry.insert(0, str(preg[clave]))
        
        self.btn_eliminar.configure(state="normal")
        self.btn_guardar_cambios.configure(state="normal")

    def guardar_edicion(self):
        if not self.pregunta_seleccionada: return
        self.btn_guardar_cambios.configure(text="Guardando...", state="disabled")
        self.update()

        try:
            preg = self.pregunta_seleccionada
            tipo = preg.get('tipo', 'normal')
            f_excel = preg['fila_excel']
            b = preg['bloque']

            # Buscar la columna exacta (índice 1-based para gspread)
            if tipo == 'normal':
                col_enun = self.encabezados.index(f'Enunciado de la Pregunta {b}') + 1
                col_corr = self.encabezados.index(f'Opción CORRECTA (Pregunta {b})') + 1
                col_f1 = self.encabezados.index(f'Opción Falsa 1 (Pregunta {b})') + 1
                col_f2 = self.encabezados.index(f'Opción Falsa 2 (Pregunta {b}) - Opcional') + 1
                col_f3 = self.encabezados.index(f'Opción Falsa 3 (Pregunta {b}) - Opcional') + 1

            else:  #RESERVA
                col_enun = self.encabezados.index(f'Enunciado de la Pregunta de Reserva {b}') + 1
                col_corr = self.encabezados.index(f'Opción CORRECTA (Pregunta de Reserva {b})') + 1
                col_f1 = self.encabezados.index(f'Opción Falsa 1 (Pregunta de Reserva {b})') + 1
                col_f2 = self.encabezados.index(f'Opción Falsa 2 (Pregunta de Reserva {b}) - Opcional') + 1
                col_f3 = self.encabezados.index(f'Opción Falsa 3 (Pregunta de Reserva {b}) - Opcional') + 1

            # Actualizar en la nube
            self.hoja.update_cell(f_excel, col_enun, self.txt_edit_enun.get("1.0", "end-1c"))
            self.hoja.update_cell(f_excel, col_corr, self.entry_edit_corr.get())
            self.hoja.update_cell(f_excel, col_f1, self.entry_edit_f1.get())
            self.hoja.update_cell(f_excel, col_f2, self.entry_edit_f2.get())
            self.hoja.update_cell(f_excel, col_f3, self.entry_edit_f3.get())

            messagebox.showinfo("Éxito", "Pregunta actualizada en la nube correctamente.\n(Refresca la conexión para ver los cambios)")
            self.btn_guardar_cambios.configure(text="💾 GUARDAR CAMBIOS EN LA NUBE", state="normal")
            
        except Exception as e:
            messagebox.showerror("Error", f"Fallo al guardar: {e}")
            self.btn_guardar_cambios.configure(text="💾 GUARDAR CAMBIOS EN LA NUBE", state="normal")

    def ejecutar_generacion(self):
        conv = self.combo_convocatoria.get()
        df_filtrado = pd.DataFrame(self.datos_sheet)
        df_filtrado = df_filtrado[df_filtrado['Código de la Convocatoria'].astype(str) == conv]
        
        datos = {k: v.get() for k, v in self.inputs.items()}
        datos["instrucciones"] = self.txt_instrucciones.get("1.0", "end-1c")

        # Capturamos el tipo de examen
        tipo_examen = self.inputs.get("tipo_examen").get() if "tipo_examen" in self.inputs else ""

        # --- 1. ABRIR VENTANA PARA ELEGIR CARPETA ---
        ruta_destino = filedialog.askdirectory(title="Selecciona la carpeta para guardar los PDFs")
        
        # Si el usuario cierra la ventana o le da a Cancelar, paramos la función en seco
        if not ruta_destino: 
            return 
        # ---------------------------------------------

        try:
            # 2. Generamos la portada temporal
            generar_portada_desde_word(datos)
            
            # 3. ¡LA LLAMADA CLAVE! Le pasamos el orden correcto: lista, conv, ruta_destino, tipo_examen
            if generar_examen_y_pdf(df_filtrado.to_dict('records'), conv, ruta_destino, tipo_examen):
                
                # 4. Preparamos el Examen Final
                sufijo = f"_Tipo_{tipo_examen}" if tipo_examen else ""
                nombre_final = f"Examen_{conv}_modelo-{datos['tipo_examen']}.pdf"
                
                # Unimos la ruta elegida con el nombre del archivo final
                ruta_examen_final = os.path.join(ruta_destino, nombre_final)
                
                # Fusionamos la portada y el examen directamente en la ruta final
                fusionar_pdfs("portada_final.pdf", "Examen_Oficial.pdf", ruta_examen_final)
                
                messagebox.showinfo("Éxito", f"¡Archivos generados correctamente en:\n{ruta_destino}")
                
                # 5. Limpiamos la basura temporal
                try: os.remove("portada_final.pdf")
                except: pass
                try: os.remove("Examen_Oficial.pdf")
                except: pass
                
        except Exception as e:
            messagebox.showerror("Error", str(e))
            
    def eliminar_pregunta(self):
        if not self.pregunta_seleccionada:
            return

        confirmar = messagebox.askyesno(
            "Confirmar eliminación",
            "¿Seguro que quieres eliminar esta pregunta?\n\n(Esta acción no se puede deshacer)"
        )

        if not confirmar:
            return

        try:
            preg = self.pregunta_seleccionada
            tipo = preg.get('tipo', 'normal')
            f_excel = preg['fila_excel']
            b = preg['bloque']

            # Detectar columnas
            if tipo == 'normal':
                col_enun = self.encabezados.index(f'Enunciado de la Pregunta {b}') + 1
                col_corr = self.encabezados.index(f'Opción CORRECTA (Pregunta {b})') + 1
                col_f1 = self.encabezados.index(f'Opción Falsa 1 (Pregunta {b})') + 1
                col_f2 = self.encabezados.index(f'Opción Falsa 2 (Pregunta {b}) - Opcional') + 1
                col_f3 = self.encabezados.index(f'Opción Falsa 3 (Pregunta {b}) - Opcional') + 1
            else:
                col_enun = self.encabezados.index(f'Enunciado de la Pregunta de Reserva {b}') + 1
                col_corr = self.encabezados.index(f'Opción CORRECTA (Pregunta de Reserva {b})') + 1
                col_f1 = self.encabezados.index(f'Opción Falsa 1 (Pregunta de Reserva {b})') + 1
                col_f2 = self.encabezados.index(f'Opción Falsa 2 (Pregunta de Reserva {b}) - Opcional') + 1
                col_f3 = self.encabezados.index(f'Opción Falsa 3 (Pregunta de Reserva {b}) - Opcional') + 1

            # 🔥 BORRAR = dejar vacío
            self.hoja.update_cell(f_excel, col_enun, "")
            self.hoja.update_cell(f_excel, col_corr, "")
            self.hoja.update_cell(f_excel, col_f1, "")
            self.hoja.update_cell(f_excel, col_f2, "")
            self.hoja.update_cell(f_excel, col_f3, "")

            messagebox.showinfo("Eliminado", "Pregunta eliminada correctamente.\n(Recarga para ver cambios)")

            # Limpiar interfaz
            self.txt_edit_enun.delete("1.0", "end")
            self.entry_edit_corr.delete(0, "end")
            self.entry_edit_f1.delete(0, "end")
            self.entry_edit_f2.delete(0, "end")
            self.entry_edit_f3.delete(0, "end")

            self.btn_guardar_cambios.configure(state="disabled")
            self.btn_eliminar.configure(state="disabled")

        except Exception as e:
            messagebox.showerror("Error", f"No se pudo eliminar: {e}")

if __name__ == "__main__":
    app = AppExamen()
    app.mainloop()
    
    
    
    
    