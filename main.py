'''
Python GUI para convertir archivos .docx a PDF en Windows usando docx2pdf.

Características:
- Permite seleccionar múltiples archivos .docx o una carpeta completa.
- Permite elegir carpeta de salida para los PDFs.
- Muestra un registro de progreso en la interfaz.
- Requiere Python y docx2pdf instalado (pip install docx2pdf).

Ejecuta este script en Windows.
'''
import os
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from docx2pdf import convert

class WordToPdfGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Convertidor de Word a PDF")
        self.root.resizable(False, False)

        # Variables de estado
        self.selected_files = []        # Lista de rutas de archivos .docx
        self.selected_folder = ""      # Carpeta seleccionada para convertir todos los .docx
        self.output_folder = ""        # Carpeta donde se guardarán los PDFs

        # Creamos el marco principal
        main_frame = ttk.Frame(root, padding=10)
        main_frame.grid(row=0, column=0, sticky="NSEW")

        # Opción: seleccionar archivos o carpeta
        mode_frame = ttk.LabelFrame(main_frame, text="Modo de conversión", padding=10)
        mode_frame.grid(row=0, column=0, columnspan=2, sticky="EW", pady=(0, 10))

        self.mode = tk.StringVar(value="files")
        ttk.Radiobutton(mode_frame, text="Archivos individuales", variable=self.mode, value="files", command=self._toggle_mode).grid(row=0, column=0, sticky="W", padx=5)
        ttk.Radiobutton(mode_frame, text="Carpeta completa", variable=self.mode, value="folder", command=self._toggle_mode).grid(row=0, column=1, sticky="W", padx=5)

        # Botones de selección
        self.files_button = ttk.Button(main_frame, text="Agregar archivos .docx", command=self._select_files)
        self.files_button.grid(row=1, column=0, sticky="EW", padx=(0,5))

        self.folder_button = ttk.Button(main_frame, text="Seleccionar carpeta", command=self._select_folder)
        self.folder_button.grid(row=1, column=1, sticky="EW", padx=(5,0))

        # Listbox para mostrar archivos seleccionados o carpeta
        list_frame = ttk.Frame(main_frame)
        list_frame.grid(row=2, column=0, columnspan=2, sticky="NSEW", pady=(10, 10))

        self.list_label = ttk.Label(list_frame, text="Archivos seleccionados:")
        self.list_label.grid(row=0, column=0, sticky="W")

        self.listbox = tk.Listbox(list_frame, width=60, height=8)
        self.listbox.grid(row=1, column=0, sticky="NSEW", pady=(5, 0))

        # Botón para limpiar selección
        self.clear_button = ttk.Button(main_frame, text="Limpiar selección", command=self._clear_selection)
        self.clear_button.grid(row=3, column=0, columnspan=2, sticky="EW", pady=(5, 10))

        # Carpeta de salida
        out_frame = ttk.Frame(main_frame)
        out_frame.grid(row=4, column=0, columnspan=2, sticky="EW", pady=(0, 10))

        ttk.Label(out_frame, text="Carpeta de salida:").grid(row=0, column=0, sticky="W")
        self.output_entry = ttk.Entry(out_frame, width=45)
        self.output_entry.grid(row=0, column=1, sticky="W", padx=(5,0))
        self.out_button = ttk.Button(out_frame, text="Seleccionar", command=self._select_output_folder)
        self.out_button.grid(row=0, column=2, sticky="W", padx=(5,0))

        # Botón de conversión
        self.convert_button = ttk.Button(main_frame, text="Convertir a PDF", command=self._start_conversion)
        self.convert_button.grid(row=5, column=0, columnspan=2, sticky="EW")

        # Área de registro de progreso
        log_frame = ttk.LabelFrame(main_frame, text="Registro", padding=10)
        log_frame.grid(row=6, column=0, columnspan=2, sticky="NSEW", pady=(10, 0))

        self.log_text = tk.Text(log_frame, width=60, height=8, state="disabled")
        self.log_text.grid(row=0, column=0, sticky="NSEW")

        # Inicializar UI según modo
        self._toggle_mode()

    def _toggle_mode(self):
        modo = self.mode.get()
        if modo == "files":
            self.files_button.state(["!disabled"])
            self.folder_button.state(["disabled"])
            self.list_label.config(text="Archivos seleccionados:")
        else:
            self.files_button.state(["disabled"])
            self.folder_button.state(["!disabled"])
            self.list_label.config(text="Carpeta seleccionada:")
        self._clear_selection()

    def _select_files(self):
        archivos = filedialog.askopenfilenames(title="Seleccionar archivos .docx",
                                               filetypes=[("Word Documents", "*.docx")])
        if archivos:
            self.selected_files = list(archivos)
            self.selected_folder = ""
            self._update_listbox(files=True)

    def _select_folder(self):
        carpeta = filedialog.askdirectory(title="Seleccionar carpeta que contiene .docx")
        if carpeta:
            self.selected_folder = carpeta
            self.selected_files = []
            self._update_listbox(files=False)

    def _select_output_folder(self):
        carpeta = filedialog.askdirectory(title="Seleccionar carpeta de salida para PDFs")
        if carpeta:
            self.output_folder = carpeta
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, carpeta)

    def _update_listbox(self, files=True):
        self.listbox.delete(0, tk.END)
        if files:
            for f in self.selected_files:
                self.listbox.insert(tk.END, f)
        else:
            self.listbox.insert(tk.END, self.selected_folder)

    def _clear_selection(self):
        self.selected_files = []
        self.selected_folder = ""
        self.listbox.delete(0, tk.END)

    def _log(self, mensaje):
        self.log_text.config(state="normal")
        self.log_text.insert(tk.END, mensaje + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state="disabled")

    def _start_conversion(self):
        # Verificar selección y carpeta de salida
        modo = self.mode.get()
        if modo == "files" and not self.selected_files:
            messagebox.showwarning("Atención", "Debes seleccionar al menos un archivo .docx.")
            return
        if modo == "folder" and not self.selected_folder:
            messagebox.showwarning("Atención", "Debes seleccionar una carpeta.")
            return
        if not self.output_entry.get():
            messagebox.showwarning("Atención", "Debes seleccionar una carpeta de salida.")
            return

        # Desactivar botón mientras se convierte
        self.convert_button.state(["disabled"])
        self._log("Iniciando conversión...")

        # Ejecutar en hilo separado para no bloquear la UI
        hilo = threading.Thread(target=self._convertir, daemon=True)
        hilo.start()

    def _convertir(self):
        modo = self.mode.get()
        salida = self.output_entry.get()
        try:
            if modo == "files":
                for ruta in self.selected_files:
                    nombre_pdf = os.path.splitext(os.path.basename(ruta))[0] + ".pdf"
                    destino = os.path.join(salida, nombre_pdf)
                    self._log(f"Convirtiendo: {ruta}")
                    convert(ruta, destino)
                    self._log(f"Guardado: {destino}")
            else:
                self._log(f"Convirtiendo carpeta: {self.selected_folder}")
                # docx2pdf convertirá todos los .docx de la carpeta al destino
                convert(self.selected_folder, salida)
                self._log(f"Archivos convertidos en: {salida}")

            self._log("Conversión finalizada.")
            messagebox.showinfo("Éxito", "La conversión ha finalizado correctamente.")
        except Exception as e:
            self._log(f"Error: {e}")
            messagebox.showerror("Error", f"Ocurrió un error: {e}")
        finally:
            self.convert_button.state(["!disabled"])

if __name__ == "__main__":
    root = tk.Tk()
    app = WordToPdfGUI(root)
    root.mainloop()
