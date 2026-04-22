import pandas as pd
import random
import os
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from fpdf import FPDF
from docxtpl import DocxTemplate
from docx2pdf import convert
from pypdf import PdfWriter
import customtkinter as ctk
from tkinter import messagebox

# --- CONFIGURACIÓN DE LA NUBE ---
NOMBRE_EXCEL_NUBE = "Banco de Preguntas SAEL (Definitivo) (respuestas)"

# ==========================================
# 1. FUNCIONES DE GENERACIÓN DE PDF (Sin cambios)
# ==========================================
class PDF(FPDF):
    def header(self):
        self.set_font('Arial', 'B', 15)
        self.cell(0, 10, 'EXAMEN DE CAPTACION - DIPUTACION DE CADIZ', 0, 1, 'C')
        self.ln(5)

def generar_examen_y_pdf(lista_preguntas):
    preguntas_normales = []
    preguntas_reserva = []

    for fila in lista_preguntas:
        for i in range(1, 21): 
            col_enunciado = f'Enunciado de la Pregunta {i}'
            col_enunciado_reserva = f'Enunciado de la Pregunta de Reserva {i}'

                        # ===== NORMAL =====
            if col_enunciado in fila and fila[col_enunciado]:

                enunciado = fila[col_enunciado]
                correcta = fila.get(f'Opción CORRECTA (Pregunta {i})', '')
                inc1 = fila.get(f'Opción Falsa 1 (Pregunta {i})', '')
                inc2 = fila.get(f'Opción Falsa 2 (Pregunta {i}) - Opcional', '')
                inc3 = fila.get(f'Opción Falsa 3 (Pregunta {i}) - Opcional', '')

                opciones_mezclar = [
                    {'texto': correcta, 'es_correcta': True},
                    {'texto': inc1, 'es_correcta': False}
                ]

                if inc2: opciones_mezclar.append({'texto': inc2, 'es_correcta': False})
                if inc3: opciones_mezclar.append({'texto': inc3, 'es_correcta': False})

                preguntas_normales.append({
                    'enunciado': str(enunciado).replace('\n', ' '),
                    'opciones': opciones_mezclar
                })


            # ===== RESERVA =====
            if col_enunciado_reserva in fila and fila[col_enunciado_reserva]:

                enunciado_reserva = fila[col_enunciado_reserva]
                correcta_reserva = fila.get(f'Opción CORRECTA (Pregunta de Reserva {i})', '')
                inc4 = fila.get(f'Opción Falsa 1 (Pregunta de Reserva {i})', '')
                inc5 = fila.get(f'Opción Falsa 2 (Pregunta de Reserva {i}) - Opcional', '')
                inc6 = fila.get(f'Opción Falsa 3 (Pregunta de Reserva {i}) - Opcional', '')

                opciones_mezclar_reserva = [
                    {'texto': correcta_reserva, 'es_correcta': True},
                    {'texto': inc4, 'es_correcta': False}
                ]

                if inc5: opciones_mezclar_reserva.append({'texto': inc5, 'es_correcta': False})
                if inc6: opciones_mezclar_reserva.append({'texto': inc6, 'es_correcta': False})

                preguntas_reserva.append({
                    'enunciado': str(enunciado_reserva).replace('\n', ' '),
                    'opciones': opciones_mezclar_reserva
                })

    if not preguntas_normales: return False

    random.shuffle(preguntas_normales)
    random.shuffle(preguntas_reserva)
    letras = ['A', 'B', 'C', 'D']
    plantilla_respuestas = {}
    plantilla_respuestas_reserva = {}

    pdf_examen = PDF()
    pdf_examen.add_page()
    pdf_examen.set_font("Arial", size=11)

    for numero, pregunta in enumerate(preguntas_normales, start=1):
        pdf_examen.multi_cell(0, 8, txt=f"{numero}. {pregunta['enunciado']}".encode('latin-1', 'replace').decode('latin-1'))
        opciones = pregunta['opciones']
        random.shuffle(opciones)
        for indice, opcion in enumerate(opciones):
            letra = letras[indice]
            pdf_examen.multi_cell(0, 6, txt=f"    {letra}) {opcion['texto']}".encode('latin-1', 'replace').decode('latin-1'))
            if opcion['es_correcta']: plantilla_respuestas[numero] = letra
        pdf_examen.ln(4)

    pdf_examen.add_page()
    pdf_examen.cell(0, 10, 'PREGUNTAS DE RESERVA', 0, 1, 'C')
    pdf_examen.ln(10)

    for numero, pregunta in enumerate(preguntas_reserva, start=1):
        pdf_examen.multi_cell(0, 8, txt=f"{numero}. {pregunta['enunciado']}".encode('latin-1', 'replace').decode('latin-1'))
        opciones = pregunta['opciones']
        random.shuffle(opciones)
        for indice, opcion in enumerate(opciones):
            letra = letras[indice]
            pdf_examen.multi_cell(0, 6, txt=f"    {letra}) {opcion['texto']}".encode('latin-1', 'replace').decode('latin-1'))
            if opcion['es_correcta']: plantilla_respuestas_reserva[numero] = letra
        pdf_examen.ln(4)

    pdf_examen.output("Examen_Oficial.pdf")

    pdf_plantilla = PDF()
    pdf_plantilla.add_page()
    pdf_plantilla.set_font("Arial", 'B', 14)
    pdf_plantilla.cell(0, 10, 'PLANTILLA DE CORRECCION', 0, 1, 'C')
    pdf_plantilla.ln(10)
    pdf_plantilla.set_font("Arial", size=12)
    for numero, letra in plantilla_respuestas.items():
        pdf_plantilla.cell(0, 8, txt=f"Pregunta {numero}  -------  Respuesta: {letra}", ln=1)

    pdf_plantilla.cell(0, 10, 'PREGUNTAS DE RESERVA', 0, 1, 'C')
    pdf_plantilla.ln(10)

    for numero, letra in plantilla_respuestas_reserva.items():
        pdf_plantilla.cell(0, 8, txt=f"Pregunta {numero}  -------  Respuesta: {letra}", ln=1)
    pdf_plantilla.output("Plantilla_Correccion.pdf")
    return True

def generar_portada_desde_word(datos_variables):
    doc = DocxTemplate("Plantilla_Portada.docx")
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
        self.title("Suite SAEL - Diputación de Cádiz")
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

        self.inputs = {}
        campos = [("oficio", "Puesto"), ("localidad", "Municipio"), ("n_plazas", "Plazas"), ("n_BOP", "BOP")]
        for key, placeholder in campos:
            ctk.CTkLabel(self.scroll_generar, text=f"{placeholder}:").pack(anchor="w", padx=10)
            entry = ctk.CTkEntry(self.scroll_generar, width=400)
            entry.pack(pady=5, padx=10)
            self.inputs[key] = entry

        ctk.CTkLabel(self.scroll_generar, text="Instrucciones del Examen:").pack(anchor="w", padx=10)
        self.txt_instrucciones = ctk.CTkTextbox(self.scroll_generar, width=400, height=100)
        self.txt_instrucciones.pack(pady=5, padx=10)

        self.btn_generar = ctk.CTkButton(self.scroll_generar, text="2. GENERAR PDF", command=self.ejecutar_generacion, height=50, state="disabled")
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

    # ==========================================
    # LÓGICA DE DATOS Y CONEXIÓN
    # ==========================================
    def conectar_nube(self):
        self.btn_conectar.configure(text="Conectando...", state="disabled")
        self.update()
        try:
            scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/drive']
            creds = ServiceAccountCredentials.from_json_keyfile_name('credenciales.json', scope)
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

        try:
            generar_portada_desde_word(datos)
            if generar_examen_y_pdf(df_filtrado.to_dict('records')):
                nombre_final = f"Examen_{datos['oficio']}_{conv}.pdf"
                fusionar_pdfs("portada_final.pdf", "Examen_Oficial.pdf", nombre_final)
                messagebox.showinfo("Éxito", f"Examen generado: {nombre_final}")
                os.remove("portada_final.pdf")
                os.remove("Examen_Oficial.pdf")
        except Exception as e:
            messagebox.showerror("Error", str(e))

if __name__ == "__main__":
    app = AppExamen()
    app.mainloop()