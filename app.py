import tkinter as tk
from tkinter import filedialog, messagebox
from docxtpl import DocxTemplate
import os

class DocxApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Generador de Documentos")
        self.root.geometry("500x600")
        
        self.template_path = None
        self.variables = set()
        self.entries = {}
        self.json_config = {}
        self.json_path = None
        
        # Título
        self.lbl_title_main = tk.Label(root, text="Generador desde Plantilla DOCX", font=("Arial", 16, "bold"))
        self.lbl_title_main.pack(pady=15)

        # UI Elements
        self.lbl_info = tk.Label(root, text="1. Selecciona una plantilla DOCX", font=("Arial", 12))
        self.lbl_info.pack(pady=10)
        
        self.btn_select = tk.Button(root, text="Seleccionar Plantilla", command=self.select_template, bg="#4CAF50", fg="white", font=("Arial", 10, "bold"))
        self.btn_select.pack(pady=5)
        
        self.lbl_template = tk.Label(root, text="Ninguna plantilla seleccionada", fg="gray")
        self.lbl_template.pack(pady=5)
        
        self.frame_form = tk.Frame(root)
        self.frame_form.pack(pady=10, fill=tk.BOTH, expand=True)
        
        self.btn_generate = tk.Button(root, text="Generar Documento", command=self.generate_doc, state=tk.DISABLED, bg="#2196F3", fg="white", font=("Arial", 10, "bold"))
        self.btn_generate.pack(pady=20)
        
    def select_template(self):
        self.template_path = filedialog.askopenfilename(
            title="Seleccionar Plantilla",
            filetypes=[("Word Documents", "*.docx")]
        )
        if self.template_path:
            self.lbl_template.config(text=os.path.basename(self.template_path), fg="blue")
            self.load_variables()
            
    def load_variables(self):
        try:
            doc = DocxTemplate(self.template_path)
            # Extraer las etiquetas automáticamente
            all_vars = doc.get_undeclared_template_variables()
            
            import json
            self.json_path = os.path.splitext(self.template_path)[0] + ".json"
            self.json_config = {}
            if os.path.exists(self.json_path):
                try:
                    with open(self.json_path, 'r', encoding='utf-8') as f:
                        self.json_config = json.load(f)
                except Exception as e_json:
                    messagebox.showwarning("Aviso JSON", f"No se pudo leer el archivo JSON asociado:\n{e_json}")
            
            # Quitar de la UI las variables que ya vienen en el JSON
            self.variables = {v for v in all_vars if v not in self.json_config}
            
            self.build_form()
            self.btn_generate.config(state=tk.NORMAL)
        except Exception as e:
            if "expected token" in str(e) or "TemplateSyntaxError" in str(type(e)):
                messagebox.showerror("Error en las Etiquetas", 
                    "Hay un problema con el nombre de una etiqueta en tu plantilla.\n\n"
                    "Recuerda que las etiquetas dentro de {{ }} NO deben contener espacios ni caracteres especiales (como acentos o la letra ñ).\n\n"
                    "Por ejemplo:\n"
                    "❌ Incorrecto: {{ tipo de memoría }}\n"
                    "✅ Correcto: {{ tipo_de_memoria }}\n\n"
                    f"Detalle técnico: {str(e)}")
            else:
                messagebox.showerror("Error", f"No se pudo leer la plantilla:\n{str(e)}")
            
    def build_form(self):
        # Limpiar formulario anterior
        for widget in self.frame_form.winfo_children():
            widget.destroy()
        
        self.entries = {}
        
        if not self.variables:
            lbl = tk.Label(self.frame_form, text="No se encontraron etiquetas {{etiqueta}} en la plantilla.")
            lbl.pack()
            return
            
        lbl_title = tk.Label(self.frame_form, text="2. Completa los campos:", font=("Arial", 12))
        lbl_title.pack(pady=10)
        
        # Crear canvas para scroll en caso de muchas variables
        canvas = tk.Canvas(self.frame_form)
        scrollbar = tk.Scrollbar(self.frame_form, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(
                scrollregion=canvas.bbox("all")
            )
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        for var in sorted(self.variables):
            frame_row = tk.Frame(scrollable_frame)
            frame_row.pack(fill=tk.X, pady=5, padx=20)
            
            lbl = tk.Label(frame_row, text=var + ":", width=20, anchor="e", font=("Arial", 10))
            lbl.pack(side=tk.LEFT, padx=5)
            
            entry = tk.Entry(frame_row, width=30, font=("Arial", 10))
            entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
            self.entries[var] = entry
            
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
    def generate_doc(self):
        if not self.template_path:
            return
            
        context = {var: entry.get() for var, entry in self.entries.items()}
        
        # Procesar variables automáticas desde el JSON
        import json
        json_updated = False
        if self.json_config:
            for var_name, var_info in self.json_config.items():
                if isinstance(var_info, dict):
                    if var_info.get("tipo") == "incremento":
                        val = var_info.get("valor", 1)
                        fmt = str(var_info.get("formato", ""))
                        # Si hay formato (ej. "000"), rellenar con ceros según la longitud
                        if fmt:
                            context[var_name] = str(val).zfill(len(fmt))
                        else:
                            context[var_name] = val
                            
                        var_info["valor"] = val + 1
                        json_updated = True
                    else:
                        val = var_info.get("valor", "")
                        fmt = str(var_info.get("formato", ""))
                        if fmt and str(val).isnumeric():
                            context[var_name] = str(val).zfill(len(fmt))
                        else:
                            context[var_name] = val
                else:
                    context[var_name] = var_info
        
        save_path = filedialog.asksaveasfilename(
            title="Guardar Documento",
            defaultextension=".docx",
            filetypes=[("Word Documents", "*.docx"), ("PDF Documents", "*.pdf")]
        )
        
        if not save_path:
            return
            
        try:
            doc = DocxTemplate(self.template_path)
            doc.render(context)
            
            if save_path.lower().endswith(".pdf"):
                import tempfile
                from docx2pdf import convert
                # Crear un archivo docx temporal
                fd, temp_docx_path = tempfile.mkstemp(suffix=".docx")
                os.close(fd)
                doc.save(temp_docx_path)
                
                # Convertir a pdf
                try:
                    convert(temp_docx_path, save_path)
                    messagebox.showinfo("Éxito", f"Documento PDF generado correctamente en:\n{save_path}")
                except Exception as e_conv:
                    messagebox.showwarning("Aviso", "Falló la conversión a PDF (¿tienes Microsoft Word instalado?). Se guardó como DOCX en su lugar.")
                    doc.save(save_path[:-4] + ".docx")
                finally:
                    if os.path.exists(temp_docx_path):
                        try:
                            os.remove(temp_docx_path)
                        except:
                            pass
            else:
                doc.save(save_path)
                messagebox.showinfo("Éxito", f"Documento guardado correctamente en:\n{save_path}")
                
            # Actualizar el JSON si hubo incrementos
            if json_updated and self.json_path:
                try:
                    with open(self.json_path, 'w', encoding='utf-8') as f:
                        json.dump(self.json_config, f, indent=4, ensure_ascii=False)
                except Exception as e_json:
                    messagebox.showerror("Error JSON", f"El documento se generó, pero falló la actualización del JSON:\n{e_json}")
                
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo generar el documento:\n{str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = DocxApp(root)
    root.mainloop()
