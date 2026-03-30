import pandas as pd
import random
from fpdf import FPDF
from docxtpl import DocxTemplate
from docx2pdf import convert
import os
from pypdf import PdfWriter
import customtkinter as ctk
from tkinter import messagebox
from tkcalendar import DateEntry

# Creamos una clase personalizada para el PDF para poder poner acentos sin que explote
class PDF(FPDF):
    def header(self):
        self.set_font('Arial', 'B', 15)
        self.cell(0, 10, 'EXAMEN DE CAPTACION - DIPUTACION DE CADIZ', 0, 1, 'C')
        self.ln(5)

def generar_examen_y_pdf(ruta_excel):
    print("Leyendo el Excel de Google Forms...")
    try:
        df = pd.read_excel(ruta_excel)
    except FileNotFoundError:
        print(f"Error: No se encuentra el archivo {ruta_excel}.")
        return

    preguntas_validas = []

    # Recorremos cada fila del Excel
    for index, fila in df.iterrows():
        
        # Revisamos los 20 bloques posibles del formulario
        for i in range(1, 21): 
            # ¡AQUÍ ESTÁ LA CLAVE! Usamos los nombres EXACTOS de las nuevas columnas
            col_enunciado = f'Enunciado de la Pregunta {i}'
            
            # Si la columna no existe en el Excel (porque alguien borró algo), saltamos
            if col_enunciado not in fila:
                continue
                
            enunciado = fila[col_enunciado]

            # Si está vacío, pasamos de largo
            if pd.isna(enunciado) or str(enunciado).strip() == "":
                continue

            # Extraemos las opciones
            correcta = fila[f'Opción CORRECTA (Pregunta {i})']
            inc1 = fila[f'Opción Falsa 1 (Pregunta {i})']
            inc2 = fila[f'Opción Falsa 2 (Pregunta {i}) - Opcional']
            inc3 = fila[f'Opción Falsa 3 (Pregunta {i}) - Opcional']

            # Metemos la correcta y la falsa 1 (que son obligatorias)
            opciones_mezclar = [
                {'texto': correcta, 'es_correcta': True},
                {'texto': inc1, 'es_correcta': False}
            ]

            # Si rellenaron la Falsa 2, la añadimos (aguanta Verdadero/Falso)
            if pd.notna(inc2) and str(inc2).strip() != "":
                opciones_mezclar.append({'texto': inc2, 'es_correcta': False})
                
            # Si rellenaron la Falsa 3, la añadimos
            if pd.notna(inc3) and str(inc3).strip() != "":
                opciones_mezclar.append({'texto': inc3, 'es_correcta': False})

            preguntas_validas.append({
                'enunciado': str(enunciado).replace('\n', ' '), # Limpiamos saltos de línea
                'opciones': opciones_mezclar
            })

    total_preguntas = len(preguntas_validas)
    if total_preguntas == 0:
        print("El Excel está vacío o no coincide el nombre de las columnas.")
        return

    print(f"¡Éxito! Se han detectado {total_preguntas} preguntas válidas.")
    print("Generando PDFs...")

    # Aleatorizamos el orden de las preguntas
    random.shuffle(preguntas_validas)
    letras = ['A', 'B', 'C', 'D']
    plantilla_respuestas = {}

    # --- INICIAMOS EL PDF DEL EXAMEN ---
    pdf_examen = PDF()
    pdf_examen.add_page()
    pdf_examen.set_font("Arial", size=11)

    for numero, pregunta in enumerate(preguntas_validas, start=1):
        # Escribimos el enunciado (multi_cell permite textos largos que saltan de línea)
        texto_pregunta = f"{numero}. {pregunta['enunciado']}"
        pdf_examen.multi_cell(0, 8, txt=texto_pregunta.encode('latin-1', 'replace').decode('latin-1'))

        # Mezclamos las opciones de esta pregunta
        opciones = pregunta['opciones']
        random.shuffle(opciones)

        for indice, opcion in enumerate(opciones):
            letra = letras[indice]
            texto_opcion = f"    {letra}) {opcion['texto']}"
            pdf_examen.multi_cell(0, 6, txt=texto_opcion.encode('latin-1', 'replace').decode('latin-1'))

            # Cazamos la correcta para la plantilla
            if opcion['es_correcta']:
                plantilla_respuestas[numero] = letra
        
        pdf_examen.ln(4) # Espacio extra entre preguntas

    # Guardamos el PDF del examen
    pdf_examen.output("Examen_Oficial.pdf")

    # --- INICIAMOS EL PDF DE LA PLANTILLA ---
    pdf_plantilla = PDF()
    pdf_plantilla.add_page()
    pdf_plantilla.set_font("Arial", 'B', 14)
    pdf_plantilla.cell(0, 10, 'PLANTILLA DE CORRECCION (PARA EL TRIBUNAL)', 0, 1, 'C')
    pdf_plantilla.ln(10)
    
    pdf_plantilla.set_font("Arial", size=12)
    for numero in range(1, total_preguntas + 1):
        texto_plantilla = f"Pregunta {numero}  ----------------------  Respuesta:  {plantilla_respuestas[numero]}"
        pdf_plantilla.cell(0, 8, txt=texto_plantilla, ln=1)

    # Guardamos la plantilla
    pdf_plantilla.output("Plantilla_Correccion.pdf")
    print("¡Terminado! Revisa tu carpeta, tienes dos archivos .pdf nuevos.")

def generar_portada_desde_word(datos_variables):
    # 1. Cargamos la plantilla de Word
    doc = DocxTemplate("plantilla_portada.docx")
    
    # 2. Sustituimos las etiquetas {{...}} por los datos reales
    # datos_variables es un diccionario: {'aspirante': 'Juan Pérez', ...}
    doc.render(datos_variables)
    
    # 3. Guardamos el Word temporal ya relleno
    ruta_word_relleno = "portada_rellena.docx"
    doc.save(ruta_word_relleno)
    
    # 4. Lo convertimos a PDF
    # NOTA: Esto requiere tener Word instalado en el PC
    convert(ruta_word_relleno, "portada_final.pdf")
    
    try:                
        # 3. Borramos el Word temporal para que solo quede el PDF limpio
        os.remove(ruta_word_relleno) 
    except Exception as e_pdf:
        pass
    return "portada_final.pdf"

def fusionar_pdfs(ruta_portada, ruta_preguntas, nombre_final):
    writer = PdfWriter()

    # Añadimos la portada
    with open(ruta_portada, "rb") as f_portada:
        writer.append(f_portada)

    # Añadimos las preguntas
    with open(ruta_preguntas, "rb") as f_preguntas:
        writer.append(f_preguntas)

    # Guardamos el resultado final
    with open(nombre_final, "wb") as f_salida:
        writer.write(f_salida)
    
    print(f"¡Éxito! Examen completo generado en: {nombre_final}")

# Configuración de apariencia
ctk.set_appearance_mode("System")  # "Dark" o "Light"
ctk.set_default_color_theme("blue")

class AppExamen(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Generador de Exámenes - Diputación de Cádiz")
        self.geometry("500x700")

        # Cargar municipios desde Excel
        try:
            df_muni = pd.read_excel('municipios.xlsx')
            self.lista_municipios = df_muni['Municipio'].astype(str).tolist()
        except:
            self.lista_municipios = ["Error cargando municipios.xlsx"]

        # --- TÍTULO ---
        self.label_titulo = ctk.CTkLabel(self, text="CONFIGURACIÓN DEL EXAMEN", font=ctk.CTkFont(size=20, weight="bold"))
        self.label_titulo.pack(pady=20)

        # --- FORMULARIO ---
        self.scroll_frame = ctk.CTkScrollableFrame(self, width=450, height=500)
        self.scroll_frame.pack(pady=10, padx=10, fill="both", expand=True)

        self.inputs = {}
        
        # Campo Especial: Localidad (Menú desplegable)
        ctk.CTkLabel(self.scroll_frame, text="Localidad:").pack(anchor="w", padx=10)
        self.combo_localidad = ctk.CTkComboBox(self.scroll_frame, values=self.lista_municipios, width=400)
        self.combo_localidad.pack(pady=(0, 10), padx=10)
        self.combo_localidad.set("")

        # Resto de campos dinámicos
        campos = [
            ("oficio", "Puesto / Oficio (ej: Policia)"),
            ("n_plazas", "Número de Plazas"),
            ("n_BOP", "Número de BOP"),
            ("fecha_BOP", "Fecha del BOP (dd/mm/yyyy)"),
            ("n_preguntas", "Número de Preguntas"),
            ("n_reserva", "Preguntas de Reserva"),
            ("n_horas", "Tiempo (Horas)"),
            ("puntuacion", "Puntos por acierto"),
            ("error", "Puntos que resta el fallo"),
            ("condicion_aprobado", "Condición para aprobar")
        ]

        # --- SECCIÓN FECHA (CALENDARIO) ---
        # ctk.CTkLabel(self.scroll_frame, text="Fecha del BOP").pack(anchor="w", padx=10)
        # self.cal = DateEntry(self.scroll_frame, width=12, background='darkblue',
        #                      foreground='white', borderwidth=2, locale='es_ES', date_pattern='dd/mm/yyyy')
        # # El DateEntry es de Tkinter clásico, lo empaquetamos con cuidado
        # self.cal.pack(pady=(0, 15), padx=10, anchor="w")

        for key, placeholder in campos:
            ctk.CTkLabel(self.scroll_frame, text=f"{placeholder}:").pack(anchor="w", padx=10)
            entry = ctk.CTkEntry(self.scroll_frame, width=400, placeholder_text=placeholder)
            entry.pack(pady=(0, 10), padx=10)
            self.inputs[key] = entry

        # --- BOTÓN GENERAR ---
        self.btn_generar = ctk.CTkButton(self, text="GENERAR EXAMEN COMPLETO", command=self.ejecutar_proceso, 
                                         fg_color="#2c3e50", hover_color="#34495e", height=50)
        self.btn_generar.pack(pady=20)

    def filtrar_municipios(self, event):
        # Capturamos lo escrito
        escrito = self.combo_localidad.get().lower()
        
        if escrito == "":
            self.combo_localidad.configure(values=self.lista_municipios)
        else:
            filtrados = [m for m in self.lista_municipios if escrito in m.lower()]
            if filtrados:
                self.combo_localidad.configure(values=filtrados)
            else:
                self.combo_localidad.configure(values=["Sin coincidencias"])
        
        # Forzar que el menú se vea desplegado mientras escribes
        self.combo_localidad._canvas.focus_set()

    def ejecutar_proceso(self):
        # 1. Recopilar datos de la interfaz
        datos = {k: v.get() for k, v in self.inputs.items()}
        datos["localidad"] = self.combo_localidad.get()
        datos["n_respuestas"] = 4 # Valor fijo o añadir a la interfaz

        # Validación básica
        if not datos["oficio"] or not datos["localidad"]:
            messagebox.showwarning("Atención", "Por favor, rellena al menos el Oficio y la Localidad.")
            return

        try:
            # 2. Llamar a tus funciones (deben estar definidas en el mismo archivo)
            print(f"Iniciando proceso para {datos['localidad']}...")
            
            ruta_portada = generar_portada_desde_word(datos)
            generar_examen_y_pdf('respuestas1.xlsx')
            
            nombre_final = f"Examen_{datos['oficio']}_{datos['localidad']}.pdf"
            fusionar_pdfs("portada_final.pdf", "Examen_Oficial.pdf", nombre_final)

            nombre_plantilla_final = f"Plantilla_Examen_{datos['oficio']}_{datos['localidad']}.pdf"

            # Limpieza
            for f in ["portada_final.pdf", "Examen_Oficial.pdf"]:
                try: os.remove(f)
                except: pass

            messagebox.showinfo("¡Éxito!", f"El examen se ha generado correctamente:\n{nombre_final}")
        
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo generar el examen:\n{str(e)}")

if __name__ == "__main__":
    app = AppExamen()
    app.mainloop()