import tkinter as tk
from tkinter import messagebox
from tkinter import filedialog
import customtkinter as ctk
import pandas as pd
from docxtpl import DocxTemplate
import os
import sys
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

# --- CONFIGURACIÓN DE APARIENCIA (ESTILO MODERNO) ---
ctk.set_appearance_mode("System")  # Se adapta al modo oscuro/claro de Windows
ctk.set_default_color_theme("blue")

# --- RUTAS ---
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

EXCEL_DATOS = os.path.join(BASE_DIR, "Datos_oficioremision.xlsx")
WORD_PLANTILLA = os.path.join(BASE_DIR, "Plantilla_oficioremision.docx")

# --- NUESTRO BUSCADOR FLOTANTE PRO (ADAPTADO A CUSTOMTKINTER) ---
class BuscadorPro(ctk.CTkEntry):
    def __init__(self, parent, lista_opciones, **kwargs):
        super().__init__(parent, **kwargs)
        self.lista_opciones = [str(x) for x in lista_opciones]
        self.popup = None
        self.listbox = None
        
        self.bind('<KeyRelease>', self.al_teclear)
        self.bind('<FocusOut>', lambda e: self.after(200, self.cerrar_popup))
        self.bind('<Button-1>', self.al_clicar)
        self.bind('<Down>', self.bajar_lista)

    def al_clicar(self, event):
        self.mostrar_popup(self.lista_opciones)

    def al_teclear(self, event):
        if event.keysym in ('Up', 'Down', 'Return', 'Escape', 'Tab'):
            return
            
        texto = self.get().lower()
        if not texto:
            filtrados = self.lista_opciones
        else:
            filtrados = [item for item in self.lista_opciones if texto in item.lower()]
            
        self.mostrar_popup(filtrados)

    def mostrar_popup(self, datos):
        if not datos:
            self.cerrar_popup()
            return
            
        if self.popup is None:
            self.popup = tk.Toplevel(self.winfo_toplevel())
            self.popup.wm_overrideredirect(True)
            self.popup.attributes('-topmost', True)
            
            # --- MAGIA CAMALEÓNICA (Detectar modo oscuro/claro) ---
            modo = ctk.get_appearance_mode()
            if modo == "Dark":
                bg_color = "#343638"  # Gris oscuro (igual que la caja de texto)
                fg_color = "#ffffff"  # Texto blanco
                sel_color = "#1f538d" # Azul de selección CustomTkinter
                borde = "#565b5e"     # Borde gris
            else:
                bg_color = "#f9f9fa"  # Blanco roto
                fg_color = "#242424"  # Texto casi negro
                sel_color = "#1f538d" # Azul de selección
                borde = "#979da2"     # Borde gris claro

            self.popup.configure(bg=borde) # Usamos el fondo del popup como borde
            
            # Marco interior con los colores adaptados
            marco = tk.Frame(self.popup, bg=bg_color)
            marco.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)
            
            self.listbox = tk.Listbox(marco, font=("Arial", 11), selectmode=tk.SINGLE, 
                                      activestyle='none', exportselection=False, 
                                      bg=bg_color, fg=fg_color, selectbackground=sel_color, 
                                      selectforeground="#ffffff", highlightthickness=0, relief="flat")
            self.listbox.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)
            
            self.listbox.bind('<ButtonRelease-1>', self.seleccionar)
            self.listbox.bind('<Return>', self.seleccionar)
            
        self.listbox.delete(0, tk.END)
        for item in datos:
            self.listbox.insert(tk.END, item)
            
        altura = min(6, len(datos))
        self.listbox.config(height=altura)
        
        x = self.winfo_rootx()
        y = self.winfo_rooty() + self.winfo_height()
        w = self.winfo_width()
        # Le sumamos 4 píxeles al alto para que encaje perfecto con el nuevo borde
        self.popup.geometry(f"{w}x{self.listbox.winfo_reqheight()+4}+{x}+{y+2}")

    def cerrar_popup(self):
        if self.popup:
            self.popup.destroy()
            self.popup = None

    def seleccionar(self, event=None):
        if not self.listbox.curselection(): return
        seleccion = self.listbox.get(self.listbox.curselection())
        self.delete(0, tk.END)
        self.insert(0, seleccion)
        self.cerrar_popup()
        self.focus_set()
        self.icursor(tk.END)

    def bajar_lista(self, event):
        if self.popup and self.listbox and self.listbox.size() > 0:
            self.listbox.focus_set()
            self.listbox.selection_set(0)

# --- ESTILOS DE WORD ---
def aplicar_verdana(paragraph, texto, negrita=False, tamaño=10):
    paragraph.clear() 
    run = paragraph.add_run(texto)
    run.font.name = 'Verdana'
    run.font.size = Pt(tamaño)
    run.bold = negrita
    run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Verdana')

def aplicar_borde_grueso(cell):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = tcPr.find(qn('w:tcBorders'))
    if tcBorders is None:
        tcBorders = OxmlElement('w:tcBorders')
        tcPr.append(tcBorders)
    right = tcBorders.find(qn('w:right'))
    if right is None:
        right = OxmlElement('w:right')
        tcBorders.append(right)
    right.set(qn('w:val'), 'single')
    right.set(qn('w:sz'), '16') 
    right.set(qn('w:space'), '0')
    right.set(qn('w:color'), '000000')

# --- EL PROGRAMA PRINCIPAL ---
class GeneradorOficiosApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Generador de Oficios Pro")
        self.geometry("750x750")

        self.municipios, self.personas, self.cargos = self.cargar_datos_excel()

        # --- PANEL SUPERIOR ---
        self.frame_superior = ctk.CTkFrame(self)
        self.frame_superior.pack(fill="x", padx=20, pady=20)
        
        ctk.CTkLabel(self.frame_superior, text="📄 Datos del Oficio", font=ctk.CTkFont(weight="bold", size=16)).pack(pady=(10, 5))

        # Contenedor para alinear los inputs
        form_frame = ctk.CTkFrame(self.frame_superior, fg_color="transparent")
        form_frame.pack(fill="x", padx=20, pady=10)

        ctk.CTkLabel(form_frame, text="Nº Expediente:").grid(row=0, column=0, sticky="w", pady=10, padx=10)
        self.expediente_entry = ctk.CTkEntry(form_frame, width=300)
        self.expediente_entry.grid(row=0, column=1, sticky="ew", padx=10, pady=10)

        ctk.CTkLabel(form_frame, text="Municipio:").grid(row=1, column=0, sticky="w", pady=10, padx=10)
        self.municipio_combo = BuscadorPro(form_frame, lista_opciones=self.municipios, width=300)
        self.municipio_combo.grid(row=1, column=1, sticky="ew", padx=10, pady=10)

        ctk.CTkLabel(form_frame, text="Asunto:").grid(row=2, column=0, sticky="w", pady=10, padx=10)
        self.asunto_entry = ctk.CTkEntry(form_frame, width=300)
        self.asunto_entry.grid(row=2, column=1, sticky="ew", padx=10, pady=10)

        form_frame.columnconfigure(1, weight=1)

        # --- PANEL INFERIOR (SCROLLABLE FRAME MÁGICO) ---
        self.frame_tabla = ctk.CTkScrollableFrame(self, label_text="👥 Asistentes", label_font=ctk.CTkFont(weight="bold", size=14))
        self.frame_tabla.pack(fill="both", expand=True, padx=20, pady=10)

        self.asistentes = [] 
        self.filas_widgets = [] 

        btn_add = ctk.CTkButton(self, text="+ Añadir Persona", fg_color="transparent", border_width=2, text_color=("black", "white"), command=self.añadir_fila_asistente)
        btn_add.pack(pady=10)

        btn_generar = ctk.CTkButton(self, text="🚀 Generar Documento PDF", font=ctk.CTkFont(weight="bold", size=14), fg_color="#27ae60", hover_color="#2ecc71", height=50, command=self.generar_documento)
        btn_generar.pack(fill="x", padx=40, pady=20)

        self.añadir_fila_asistente() 

    def cargar_datos_excel(self):
        try:
            if not os.path.exists(EXCEL_DATOS):
                messagebox.showerror("Error", f"No se encuentra el archivo {EXCEL_DATOS}.")
                self.destroy()
                return [], [], []
            
            df_municipios = pd.read_excel(EXCEL_DATOS, sheet_name="Municipios").dropna()
            df_municipios.columns = df_municipios.columns.str.strip()
            
            df_personas = pd.read_excel(EXCEL_DATOS, sheet_name="Personas").dropna()
            df_cargos = pd.read_excel(EXCEL_DATOS, sheet_name="Cargos").dropna()
            
            municipios = [str(x) for x in df_municipios['Municipio'].tolist()]
            self.diccionario_atributos = dict(zip(df_municipios['Municipio'], df_municipios['Atributo']))
            personas = [str(x) for x in df_personas['Nombre'].tolist()]
            cargos = [str(x) for x in df_cargos['Cargo'].tolist()]
            
            return municipios, personas, cargos
        except Exception as e:
            messagebox.showerror("Error Critico", f"No se pudo leer el Excel: {e}")
            self.destroy()
            return [], [], []

    def añadir_fila_asistente(self):
        row_frame = ctk.CTkFrame(self.frame_tabla, fg_color="transparent")
        row_frame.pack(fill="x", pady=5)
        
        combo_persona = BuscadorPro(row_frame, lista_opciones=self.personas, width=350, placeholder_text="Nombre de la persona...")
        combo_persona.pack(side="left", padx=5, fill="x", expand=True)

        combo_cargo = BuscadorPro(row_frame, lista_opciones=self.cargos, width=200, placeholder_text="Cargo...")
        combo_cargo.pack(side="left", padx=5)

        btn_del = ctk.CTkButton(row_frame, text="✖", fg_color="#e74c3c", hover_color="#c0392b", width=40, command=lambda f=row_frame, c_p=combo_persona, c_c=combo_cargo: self.eliminar_fila(f, c_p, c_c))
        btn_del.pack(side="left", padx=5)
        
        self.filas_widgets.append(row_frame)
        self.asistentes.append({'combo_p': combo_persona, 'combo_c': combo_cargo})

    def eliminar_fila(self, frame, combo_p, combo_c):
        frame.destroy()
        for asistente in self.asistentes:
            if asistente['combo_p'] == combo_p and asistente['combo_c'] == combo_c:
                self.asistentes.remove(asistente)
                break
        self.filas_widgets.remove(frame)

    def generar_documento(self):
        try:
            expediente = self.expediente_entry.get().strip()
            municipio_original = self.municipio_combo.get().strip()
            asunto = self.asunto_entry.get().strip().upper()
            
            municipio = municipio_original.upper()
            tipo_entidad = self.diccionario_atributos.get(municipio_original, "Municipio")

            if str(tipo_entidad).strip().lower() == "municipio":
                municipio_texto_final = f"AYUNTAMIENTO DE {municipio}"
            else:
                municipio_texto_final = municipio

            if not (expediente and municipio and asunto):
                messagebox.showwarning("Atención", "Rellena los campos Expediente, Municipio y Asunto.")
                return

            if not os.path.exists(WORD_PLANTILLA):
                messagebox.showerror("Error", f"No se encuentra la plantilla {WORD_PLANTILLA}.")
                return

            titulares = []
            suplentes = []
            
            for asistente in self.asistentes:
                nombre = asistente['combo_p'].get().strip()
                cargo = asistente['combo_c'].get().strip()
                if nombre and cargo:
                    if "suplente" in cargo.lower():
                        suplentes.append({'nombre': nombre, 'cargo': cargo})
                    else:
                        titulares.append({'nombre': nombre, 'cargo': cargo})

            if not titulares and not suplentes:
                messagebox.showwarning("Atención", "Añade al menos una persona.")
                return

            context = {
                'nexpediente': expediente,
                'municipio': municipio_texto_final,
                'asunto': asunto
            }
            doc = DocxTemplate(WORD_PLANTILLA)
            doc.render(context)

            for table in doc.docx.tables:
                if len(table.columns) >= 2: 
                    tblPr = table._element.xpath('w:tblPr')
                    if tblPr:
                        tblW = tblPr[0].xpath('w:tblW')
                        if tblW:
                            tblW[0].set(qn('w:type'), 'auto')
                            tblW[0].set(qn('w:w'), '0')

                    ancho_c1 = Inches(1.8)
                    ancho_c2 = Inches(2.0)
                    
                    table.columns[0].width = ancho_c1
                    table.columns[1].width = ancho_c2
                    table.alignment = WD_TABLE_ALIGNMENT.CENTER

                    table.rows[0].cells[0].width = ancho_c1
                    table.rows[0].cells[1].width = ancho_c2

                    table.rows[0].cells[1].text = "" 
                    p_tit = table.rows[0].cells[0].paragraphs[0]
                    p_tit.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    aplicar_verdana(p_tit, "TITULARES", negrita=True, tamaño=10)

                    for t in titulares:
                        nueva_fila = table.add_row()
                        nueva_fila.cells[0].width = ancho_c1
                        nueva_fila.cells[1].width = ancho_c2

                        p_cargo = nueva_fila.cells[0].paragraphs[0]
                        aplicar_verdana(p_cargo, t['cargo'].upper(), negrita=True, tamaño=9)
                        
                        p_nombre = nueva_fila.cells[1].paragraphs[0]
                        aplicar_verdana(p_nombre, t['nombre'], negrita=False, tamaño=9)
                        
                        aplicar_borde_grueso(nueva_fila.cells[0])

                    if suplentes:
                        fila_titulo_sup = table.add_row()
                        fila_titulo_sup.cells[0].width = ancho_c1
                        fila_titulo_sup.cells[1].width = ancho_c2

                        fila_titulo_sup.cells[1].text = "" 
                        p_sup = fila_titulo_sup.cells[0].paragraphs[0]
                        p_sup.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        aplicar_verdana(p_sup, "SUPLENTES", negrita=True, tamaño=10)

                        for s in suplentes:
                            fila_sup = table.add_row()
                            fila_sup.cells[0].width = ancho_c1
                            fila_sup.cells[1].width = ancho_c2

                            fila_sup.cells[0].text = "" 
                            p_nom_sup = fila_sup.cells[1].paragraphs[0]
                            aplicar_verdana(p_nom_sup, s['nombre'], negrita=False, tamaño=9)
                            
                            aplicar_borde_grueso(fila_sup.cells[0])
                            
                    break

            nombre_sugerido = f"Oficio - {municipio} - Expediente {expediente}.pdf"
            
            ruta_pdf_final = filedialog.asksaveasfilename(
                title="Guardar Oficio como...",
                initialdir=r"G:\Mi unidad", 
                initialfile=nombre_sugerido,
                defaultextension=".pdf",
                filetypes=[("Archivo PDF", "*.pdf")]
            )

            if not ruta_pdf_final: return
            ruta_word_temp = ruta_pdf_final.replace(".pdf", ".docx")
            doc.save(ruta_word_temp)
            
            try:
                if sys.stdout is None: sys.stdout = open(os.devnull, "w")
                if sys.stderr is None: sys.stderr = open(os.devnull, "w")

                from docx2pdf import convert
                convert(ruta_word_temp, ruta_pdf_final)
                os.remove(ruta_word_temp) 
                
                messagebox.showinfo("¡Éxito Total!", "¡Documento guardado maravillosamente!")
            except Exception as e_pdf:
                messagebox.showwarning("Casi perfecto", f"Se generó el archivo Word, pero falló la conversión a PDF.\nSe ha guardado el Word en:\n{ruta_word_temp}\n\nError: {e_pdf}")

        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error inesperado al generar: {e}")

if __name__ == "__main__":
    app = GeneradorOficiosApp()
    app.mainloop()