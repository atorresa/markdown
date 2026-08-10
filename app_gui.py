import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path

from PIL import Image, ImageTk
from tkinterdnd2 import DND_FILES, TkinterDnD

from app import __version__, convert_file, convert_files, is_url


def resource_path(relative_path: str) -> Path:
    """Resuelve una ruta de recurso tanto en desarrollo como empaquetada con PyInstaller."""
    base_path = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base_path / relative_path


TERMS_AND_CONDITIONS_TEXT = """TÉRMINOS Y CONDICIONES DE USO Y POLÍTICA DE PRIVACIDAD.

Última actualización: 9 de agosto de 2026

Al descargar, instalar, acceder o utilizar esta aplicación (en adelante, la "Aplicación"), usted (en adelante, el "Usuario") acepta de manera plena y sin reservas los presentes Términos y Condiciones de Uso. Si no está de acuerdo con las condiciones aquí establecidas, deberá abstenerse de instalar o utilizar la Aplicación.

1. TITULARIDAD Y PROPIEDAD INTELECTUAL
La Aplicación, su código fuente, arquitectura, diseño, interfaz de usuario, marcas, logotipos y cualquier documentación técnica son propiedad exclusiva de VIRTUALTELCO S.A.S. BIC (sociedad comercial identificada bajo las leyes de la República de Colombia).

Todos los derechos de propiedad intelectual e industrial están reservados a favor de VIRTUALTELCO S.A.S. BIC. Nada en estos términos se interpretará como transferencia o cesión de la propiedad de la Aplicación al Usuario.

2. CONCESIÓN DE LICENCIA DE USO
VIRTUALTELCO S.A.S. BIC concede al Usuario una licencia de uso:

Carácter: Limitada, no exclusiva, personal, intransferible, gratuita, revocable y no sublicenciable.

Finalidad: Exclusivamente para uso personal y no comercial.

Restricciones: El Usuario no podrá vender, alquilar, sublicenciar, descompilar, realizar ingeniería inversa, modificar o crear obras derivadas basadas en la Aplicación sin autorización previa y por escrito de VIRTUALTELCO S.A.S. BIC.

3. AUSENCIA DE COMPRAS EN LA APLICACIÓN Y GRATUIDAD TOTAL
La Aplicación es completamente gratuita. NO contiene ni ofrece compras dentro de la aplicación (In-App Purchases), suscripciones pagadas, microtransacciones ni cargos ocultos de ningún tipo. El acceso a todas las funcionalidades disponibles en la Aplicación se otorga sin costo alguno para el Usuario.

4. AUSENCIA DE PUBLICIDAD DE TERCEROS
La Aplicación NO contiene publicidad de terceros, banners publicitarios, anuncios de video, ventanas emergentes ni ningún tipo de red publicitaria (Ad Networks). La experiencia de uso está 100% libre de anuncios comerciales.

5. PRINCIPIO DE PRIVACIDAD Y AUSENCIA DE RECOLECCIÓN DE DATOS
En cumplimiento de la Ley Estatutaria 1581 de 2012 de Colombia, el Reglamento General de Protección de Datos (RGPD/GDPR) de la Unión Europea, la CCPA/CPRA de California (EE. UU.) y la COPPA (protección de menores):

Declaramos de forma categórica que la Aplicación opera bajo el principio de Privacidad desde el Diseño y NO recolecta, almacena, procesa, transmite, vende ni comparte ningún tipo de dato personal (como nombres, correos electrónicos, identificadores de dispositivo, ubicaciones GPS o listas de contactos).

Al no incluir publicidad ni compras dentro de la app, tampoco se emplean cookies de rastreo, SDKs de analítica comercial ni herramientas de perfilamiento de usuarios.

Cualquier procesamiento técnico, cálculo o almacenamiento de archivos se realiza exclusivamente de manera local en el dispositivo del Usuario.

6. RESPONSABILIDAD ÚNICA DEL USUARIO Y EXENCIÓN DE RESPONSABILIDAD
Uso bajo propio riesgo: El uso de la Aplicación y de los datos, cálculos o resultados generados a través de ella se realiza bajo el exclusivo criterio, cuenta y riesgo del Usuario.

Responsabilidad por los resultados: VIRTUALTELCO S.A.S. BIC no garantiza la idoneidad, precisión o exhaustividad de los resultados obtenidos con la Aplicación. El Usuario es el único y exclusivo responsable de las decisiones o acciones que adopte con base en el uso de la Aplicación.

Garantía "Tal Cual" (As Is): La Aplicación se proporciona "TAL CUAL" y "SEGÚN DISPONIBILIDAD", sin garantías de funcionamiento ininterrumpido ni libre de errores.

Limitación de responsabilidad: En la máxima medida permitida por la ley aplicable, VIRTUALTELCO S.A.S. BIC no será responsable por ningún daño directo, indirecto, incidental, especial o consecuencial derivado del uso o la imposibilidad de uso de la Aplicación.

7. MODIFICACIONES Y TERMINACIÓN
VIRTUALTELCO S.A.S. BIC se reserva el derecho de modificar, actualizar o discontinuar la Aplicación o los presentes Términos en cualquier momento. La versión actualizada estará disponible dentro de la Aplicación o en la ficha de la tienda correspondiente.

8. LEY APLICABLE Y JURISDICCIÓN
Estos Términos y Condiciones se rigen e interpretan de conformidad con las leyes de la República de Colombia. Cualquier disputa o controversia será sometida a la jurisdicción de los jueces y tribunales competentes de Colombia.

9. CONTACTO
Para cualquier inquietud relacionada con estos Términos o el funcionamiento de la Aplicación, puede comunicarse con:

Razón Social: VIRTUALTELCO S.A.S. BIC

www.virtualtelco.com.co

País: Colombia"""


class MarkdownConverterGUI(TkinterDnD.Tk):
    """Ventana gráfica para preparar archivos para IA en Markdown."""

    BRAND_BLUE = "#04225e"  # Azul de "TELCO" muestreado del logo.
    BRAND_WHITE = "#ffffff"  # Blanco de fondo del logo.

    def __init__(self):
        """Crea la ventana principal, sus variables de estado y construye la interfaz."""
        super().__init__()
        self.title(f"Preparador de archivos para IA en general o 3CX IA — v{__version__}")
        self.geometry("720x480")
        try:
            self.iconbitmap(default="icon.ico")
        except Exception:
            pass
        self.minsize(680, 440)
        self.configure(bg=self.BRAND_WHITE)
        self._configure_styles()

        self.input_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.status_text = tk.StringVar(value="Selecciona uno o varios archivos para comenzar")
        self.input_paths: list[str] = []

        self._build_menu()
        self._build_ui()

    def _configure_styles(self):
        """Configura los estilos ttk: fondo blanco de marca y la barra de progreso en azul."""
        style = ttk.Style(self)
        # El tema nativo de Windows ("vista"/"xpnative") ignora los colores personalizados
        # en la Progressbar; "clam" sí los respeta y mantiene el resto de widgets estándar.
        style.theme_use("clam")

        # Fondo blanco (el mismo del logo) para los frames y etiquetas de la ventana.
        style.configure("TFrame", background=self.BRAND_WHITE)
        style.configure("Card.TFrame", background=self.BRAND_WHITE)
        style.configure("TLabel", background=self.BRAND_WHITE)

        style.configure(
            "Brand.Horizontal.TProgressbar",
            troughcolor="#dfe6ee",
            background=self.BRAND_BLUE,
            bordercolor="#dfe6ee",
            lightcolor=self.BRAND_BLUE,
            darkcolor=self.BRAND_BLUE,
            font=("Arial", 16,),
        )

    def _build_logo(self, parent):
        """Muestra el logo de la empresa en la cabecera; si no está disponible, usa un título de texto."""
        logo_path = resource_path("assets/logo.jpg")
        try:
            image = Image.open(logo_path)
            max_width = 320
            ratio = max_width / image.width
            image = image.resize((max_width, round(image.height * ratio)), Image.LANCZOS)
            # Se guarda como atributo para que Tkinter no la recolecte como basura.
            self._logo_image = ImageTk.PhotoImage(image)
            ttk.Label(parent, image=self._logo_image).pack(anchor="w", pady=(0, 12))
        except Exception:
            ttk.Label(parent, text="Preparador de archivos para IA", font=("Arial", 16, "bold")).pack(
                anchor="w", pady=(0, 12)
            )

    def _build_menu(self):
        """Crea la barra de menú superior para navegar entre Inicio, Términos y Licencia."""
        menubar = tk.Menu(self)
        menubar.add_command(label="Inicio", command=lambda: self._show_page("home"))
        menubar.add_command(label="Términos y condiciones", command=lambda: self._show_page("terms"))
        menubar.add_command(label="Licencia", command=lambda: self._show_page("license"))
        self.config(menu=menubar)

    def _build_ui(self):
        """Construye las páginas de la aplicación (Inicio, Términos y condiciones, Licencia)."""
        self.container = ttk.Frame(self)
        self.container.pack(fill="both", expand=True)

        self.pages = {
            "home": self._build_home_page(),
            "terms": self._build_text_page("Términos y condiciones", TERMS_AND_CONDITIONS_TEXT),
            "license": self._build_text_page("Licencia", self._load_license_text()),
        }
        self._show_page("home")

    def _show_page(self, name: str):
        """Muestra la página indicada ('home', 'terms' o 'license') y oculta las demás."""
        for page_name, frame in self.pages.items():
            if page_name == name:
                frame.pack(fill="both", expand=True)
            else:
                frame.pack_forget()

    def _load_license_text(self) -> str:
        """Lee el archivo LICENSE del repositorio para mostrarlo en la página de Licencia."""
        try:
            return resource_path("LICENSE").read_text(encoding="utf-8")
        except Exception:
            return "No se encontró el archivo LICENSE."

    def _build_text_page(self, title: str, text: str) -> ttk.Frame:
        """Construye una página genérica con un título y un texto largo, legible y de solo lectura."""
        frame = ttk.Frame(self.container, padding=20, style="Card.TFrame")

        ttk.Label(frame, text=title, font=("Arial", 16, "bold")).pack(anchor="w", pady=(0, 10))

        text_frame = ttk.Frame(frame)
        text_frame.pack(fill="both", expand=True)

        scrollbar = ttk.Scrollbar(text_frame, orient="vertical")
        scrollbar.pack(side="right", fill="y")

        text_widget = tk.Text(
            text_frame,
            wrap="word",
            font=("Arial", 10),
            background=self.BRAND_WHITE,
            relief="flat",
            padx=4,
            pady=4,
            yscrollcommand=scrollbar.set,
        )
        text_widget.insert("1.0", text)
        # insert() deja el cursor al final del texto insertado; sin esto, el widget arranca
        # con scroll casi al final en vez de mostrar el principio.
        text_widget.mark_set("insert", "1.0")
        text_widget.see("1.0")
        text_widget.configure(state="disabled")
        text_widget.pack(side="left", fill="both", expand=True)
        scrollbar.configure(command=text_widget.yview)

        return frame

    def _build_home_page(self) -> ttk.Frame:
        """Construye la página de Inicio: los campos de entrada, salida y botones de conversión."""
        main = ttk.Frame(self.container, padding=20)
        main.configure(style="Card.TFrame")

        self._build_logo(main)

        subtitle = ttk.Label(
            main,
            text=(
                "Convierte archivos TXT, HTML, PDF, Word, Excel, PowerPoint, OpenDocument "
                "(LibreOffice/OpenOffice), RTF o una página web (URL) a Markdown para 3CX o IA."
            ),
            font=("Arial", 16),
            wraplength=640,
        )
        subtitle.pack(anchor="w", pady=(0, 12))

        ttk.Label(main, text="Archivos de entrada (o pega una URL, ej. https://ejemplo.com):").pack(anchor="w")
        input_frame = ttk.Frame(main)
        input_frame.pack(fill="x", pady=(4, 8))
        input_entry = ttk.Entry(input_frame, textvariable=self.input_path, width=90)
        input_entry.pack(side="left", fill="x", expand=True)
        ttk.Button(input_frame, text="Explorar", command=self.select_input_file).pack(side="left", padx=(8, 0))

        ttk.Label(main, text="Carpeta de salida del archivo markdown:").pack(anchor="w")
        output_frame = ttk.Frame(main)
        output_frame.pack(fill="x", pady=(4, 8))
        output_entry = ttk.Entry(output_frame, textvariable=self.output_path, width=90)
        output_entry.pack(side="left", fill="x", expand=True)
        ttk.Button(output_frame, text="Elegir carpeta", command=self.select_output_folder).pack(side="left", padx=(8, 0))

        button_frame = ttk.Frame(main)
        button_frame.pack(fill="x", pady=(14, 10))
        self.convert_button = ttk.Button(button_frame, text="Convertir", command=self.convert_file)
        self.convert_button.pack(side="left")
        ttk.Button(button_frame, text="Salir", command=self.destroy).pack(side="left", padx=(8, 0))

        ttk.Label(main, textvariable=self.status_text, wraplength=640, foreground="#1f4e79").pack(anchor="w", pady=(8, 0))

        self.progress = ttk.Progressbar(main, mode="indeterminate", style="Brand.Horizontal.TProgressbar")

        self.bind("<Control-v>", self.handle_paste)
        self.drop_target = ttk.Label(
            main,
            text="Arrastra archivos aquí o usa Ctrl+V para pegarlos",
            foreground="#666666",
            relief="groove",
            padding=10,
            anchor="center",
        )
        self.drop_target.pack(fill="x", pady=(10, 0))
        self.drop_target.bind("<Enter>", lambda event: self.drop_target.configure(foreground="#1f4e79"))
        self.drop_target.bind("<Leave>", lambda event: self.drop_target.configure(foreground="#666666"))
        self.drop_target.bind("<Button-1>", lambda event: self.select_input_file())

        # Registra la ventana y el área destacada como zonas donde soltar archivos.
        for widget in (self, self.drop_target):
            widget.drop_target_register(DND_FILES)
            widget.dnd_bind("<<Drop>>", self.handle_file_drop)

        return main

    def select_input_file(self):
        """Permite seleccionar uno o varios archivos desde el explorador."""
        file_paths = filedialog.askopenfilenames(
            title="Selecciona uno o varios archivos",
            filetypes=[
                (
                    "Archivos soportados",
                    "*.txt;*.html;*.htm;*.docx;*.pdf;*.xlsx;*.pptx;*.odt;*.ods;*.odp;*.rtf",
                ),
                ("Todos los archivos", "*.*"),
            ],
        )
        if file_paths:
            self._apply_selected_paths(list(file_paths), origen="Seleccionados")

    def select_output_folder(self):
        """Permite elegir la carpeta donde se guardarán los archivos Markdown."""
        folder_path = filedialog.askdirectory(title="Selecciona una carpeta de salida")
        if folder_path:
            self.output_path.set(folder_path)

    def _apply_selected_paths(self, paths: list[str], origen: str):
        """Actualiza los campos de entrada/salida con la lista de rutas o URLs recibida."""
        paths = [path for path in paths if path.strip()]
        if not paths:
            return
        self.input_paths = paths
        self.input_path.set("; ".join(paths))
        # Solo autocompletamos la carpeta de salida si el primer elemento es un archivo local;
        # una URL no tiene una carpeta "padre" en disco.
        if not is_url(paths[0]):
            self.output_path.set(str(Path(paths[0]).parent))
        self.status_text.set(f"{origen} {len(paths)} archivo(s)/URL. Puedes ajustar la carpeta de salida.")

    def handle_paste(self, event=None):
        """Recibe rutas de archivo pegadas desde el portapapeles (Ctrl+V)."""
        if event is None:
            return
        data = self.clipboard_get()
        if data:
            paths = [path for path in data.splitlines() if path.strip()]
            self._apply_selected_paths(paths, origen="Pegados desde el portapapeles:")

    def handle_file_drop(self, event):
        """Recibe archivos soltados directamente sobre la ventana (arrastrar y soltar)."""
        paths = [path for path in self.tk.splitlist(event.data) if path.strip()]
        self._apply_selected_paths(paths, origen="Soltados:")

    def convert_file(self):
        """Inicia la conversión para uno o varios archivos."""
        input_values = [value.strip() for value in self.input_path.get().split(";") if value.strip()]
        output_value = self.output_path.get().strip()

        if not input_values:
            messagebox.showwarning("Falta archivo", "Selecciona uno o varios archivos de entrada antes de convertir.")
            return

        missing = [path for path in input_values if not is_url(path) and not Path(path).exists()]
        if missing:
            messagebox.showerror("Archivo no encontrado", f"No existe el archivo:\n{missing[0]}")
            return

        self.convert_button.configure(state="disabled")
        self.status_text.set("Convirtiendo, por favor espera...")
        # "before" asegura que la barra siempre ocupe el mismo lugar fijo (justo antes del
        # recuadro de arrastrar), en vez de reinsertarse al final tras cada pack_forget().
        self.progress.pack(fill="x", pady=(6, 0), before=self.drop_target)
        self.progress.start(10)

        thread = threading.Thread(target=self._convert_in_background, args=(input_values, output_value), daemon=True)
        thread.start()

    def _convert_in_background(self, input_values, output_value):
        """Ejecuta la conversión en un hilo aparte para no congelar la ventana."""
        try:
            if len(input_values) == 1:
                result_path = convert_file(input_values[0], output_value or None)
                status = f"Conversión completada: {result_path}"
                info = ("Listo", f"Archivo convertido correctamente en:\n{result_path}")
            else:
                result_paths = convert_files(input_values, output_dir=output_value or None)
                status = f"Conversión completada: {len(result_paths)} archivos"
                info = ("Listo", "Archivos convertidos correctamente")
        except Exception as exc:  # pragma: no cover - UI feedback
            self.after(0, self._on_conversion_error, str(exc))
        else:
            self.after(0, self._on_conversion_success, status, info)

    def _on_conversion_success(self, status, info):
        """Se ejecuta en el hilo principal cuando la conversión termina correctamente."""
        self._stop_progress()
        self.status_text.set(status)
        messagebox.showinfo(*info)

    def _on_conversion_error(self, error_text):
        """Se ejecuta en el hilo principal cuando la conversión falla con un error."""
        self._stop_progress()
        self.status_text.set(f"Error: {error_text}")
        messagebox.showerror("Error al convertir", error_text)

    def _stop_progress(self):
        """Detiene y oculta la barra de progreso, y reactiva el botón."""
        self.progress.stop()
        self.progress.pack_forget()
        self.convert_button.configure(state="normal")


if __name__ == "__main__":
    app = MarkdownConverterGUI()
    app.mainloop()
