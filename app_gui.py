import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path

from app import convert_file, convert_files


class MarkdownConverterGUI(tk.Tk):
    """Ventana gráfica para preparar archivos para IA en Markdown."""

    def __init__(self):
        super().__init__()
        self.title("Preparador de archivos para IA")
        self.geometry("720x360")
        try:
            self.iconbitmap(default="icon.ico")
        except Exception:
            pass
        self.minsize(680, 320)
        self.configure(bg="#f4f7fb")

        self.input_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.status_text = tk.StringVar(value="Selecciona uno o varios archivos para comenzar")
        self.input_paths: list[str] = []

        self._build_ui()

    def _build_ui(self):
        """Construye la interfaz con los campos de entrada, salida y botones."""
        main = ttk.Frame(self, padding=20)
        main.pack(fill="both", expand=True)
        main.configure(style="Card.TFrame")

        title = ttk.Label(main, text="Preparador de archivos para IA", font=("Segoe UI", 16, "bold"))
        title.pack(anchor="w", pady=(0, 12))

        subtitle = ttk.Label(main, text="Convierte archivos TXT, HTML, PDF, Word, Excel y PowerPoint a Markdown para IA.", wraplength=640)
        subtitle.pack(anchor="w", pady=(0, 12))

        ttk.Label(main, text="Archivos de entrada:").pack(anchor="w")
        input_frame = ttk.Frame(main)
        input_frame.pack(fill="x", pady=(4, 8))
        input_entry = ttk.Entry(input_frame, textvariable=self.input_path, width=90)
        input_entry.pack(side="left", fill="x", expand=True)
        ttk.Button(input_frame, text="Explorar", command=self.select_input_file).pack(side="left", padx=(8, 0))

        ttk.Label(main, text="Carpeta de salida:").pack(anchor="w")
        output_frame = ttk.Frame(main)
        output_frame.pack(fill="x", pady=(4, 8))
        output_entry = ttk.Entry(output_frame, textvariable=self.output_path, width=90)
        output_entry.pack(side="left", fill="x", expand=True)
        ttk.Button(output_frame, text="Elegir carpeta", command=self.select_output_folder).pack(side="left", padx=(8, 0))

        button_frame = ttk.Frame(main)
        button_frame.pack(fill="x", pady=(14, 10))
        ttk.Button(button_frame, text="Convertir", command=self.convert_file).pack(side="left")
        ttk.Button(button_frame, text="Salir", command=self.destroy).pack(side="left", padx=(8, 0))

        ttk.Label(main, textvariable=self.status_text, wraplength=640, foreground="#1f4e79").pack(anchor="w", pady=(8, 0))

        self.bind("<Control-v>", self.handle_drop)
        self.bind("<Button-1>", self.clear_selection)
        self.drop_target = ttk.Label(main, text="Arrastra archivos aquí o usa Ctrl+V para pegarlos", foreground="#666666")
        self.drop_target.pack(anchor="w", pady=(10, 0))
        self.drop_target.bind("<Enter>", lambda event: self.drop_target.configure(foreground="#1f4e79"))
        self.drop_target.bind("<Leave>", lambda event: self.drop_target.configure(foreground="#666666"))
        self.drop_target.bind("<Button-1>", lambda event: self.select_input_file())

    def select_input_file(self):
        """Permite seleccionar uno o varios archivos desde el explorador."""
        file_paths = filedialog.askopenfilenames(
            title="Selecciona uno o varios archivos",
            filetypes=[
                ("Archivos soportados", "*.txt;*.html;*.htm;*.docx;*.pdf;*.xlsx;*.pptx"),
                ("Todos los archivos", "*.*"),
            ],
        )
        if file_paths:
            self.input_paths = list(file_paths)
            self.input_path.set("; ".join(self.input_paths))
            self.output_path.set(str(Path(file_paths[0]).parent))
            self.status_text.set(f"Seleccionados {len(self.input_paths)} archivos. Puedes ajustar la carpeta de salida.")

    def select_output_folder(self):
        """Permite elegir la carpeta donde se guardarán los archivos Markdown."""
        folder_path = filedialog.askdirectory(title="Selecciona una carpeta de salida")
        if folder_path:
            self.output_path.set(folder_path)

    def clear_selection(self, event=None):
        """Limpia la selección actual cuando el usuario hace clic fuera."""
        if event is None:
            return
        if self.input_path.get():
            self.input_path.set("")
            self.output_path.set("")
            self.input_paths = []
            self.status_text.set("Selección limpiada")

    def handle_drop(self, event=None):
        """Recibe datos pegados desde el portapapeles para cargar archivos."""
        if event is None:
            return
        data = self.clipboard_get()
        if data:
            paths = [path for path in data.splitlines() if path.strip()]
            if paths:
                self.input_paths = paths
                self.input_path.set("; ".join(paths))
                self.output_path.set(str(Path(paths[0]).parent))
                self.status_text.set(f"Archivos pegados desde el portapapeles: {len(paths)}")

    def convert_file(self):
        """Inicia la conversión para uno o varios archivos."""
        input_values = [value.strip() for value in self.input_path.get().split(";") if value.strip()]
        output_value = self.output_path.get().strip()

        if not input_values:
            messagebox.showwarning("Falta archivo", "Selecciona uno o varios archivos de entrada antes de convertir.")
            return

        missing = [path for path in input_values if not Path(path).exists()]
        if missing:
            messagebox.showerror("Archivo no encontrado", f"No existe el archivo:\n{missing[0]}")
            return

        try:
            if len(input_values) == 1:
                result_path = convert_file(input_values[0], output_value or None)
                self.status_text.set(f"Conversión completada: {result_path}")
                messagebox.showinfo("Listo", f"Archivo convertido correctamente en:\n{result_path}")
            else:
                result_paths = convert_files(input_values, output_dir=output_value or None)
                self.status_text.set(f"Conversión completada: {len(result_paths)} archivos")
                messagebox.showinfo("Listo", "Archivos convertidos correctamente")
        except Exception as exc:  # pragma: no cover - UI feedback
            self.status_text.set(f"Error: {exc}")
            messagebox.showerror("Error al convertir", str(exc))


if __name__ == "__main__":
    app = MarkdownConverterGUI()
    app.mainloop()
