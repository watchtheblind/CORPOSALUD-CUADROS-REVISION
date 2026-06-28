
import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import date
from openpyxl import load_workbook
from core.configs_loader import _source_path
from core.file_reader import WorkbookReader
from core.mapper import Mapper
from core.template_writer import TemplateWriter
from utils.loader import LoadingUI
from utils.updater import GitHubUpdater
import os
import sys 

GITHUB_USER = "watchtheblind"
GITHUB_REPO = "CORPOSALUD-CUADROS-REVISION"


class PayrollApp:
    """Arma la interfaz de usuario sin tocar la lógica de negocio"""

    def __init__(self):
        self.root = tk.Tk()
        self.root.withdraw()
        self.root.attributes("-topmost", True)

    def obtener_ruta_externa(self, archivo):
        # Si es un ejecutable de PyInstaller
        if hasattr(sys, '_MEIPASS'):
            # sys.executable es la ruta del .exe. Subimos un nivel para estar en su carpeta.
            return os.path.join(os.path.dirname(sys.executable), archivo)
        # Si es el script .py normal
        return os.path.join(os.path.dirname(os.path.abspath(__file__)), archivo)

    def run(self):
        if self._check_for_updates():
            return  # Actualización en proceso, no continúe

        template_path = self.obtener_ruta_externa("plantilla2.xlsx")
        if not os.path.exists(template_path):
            messagebox.showerror(
                "Error",
                f"No se encontró la plantilla en:\n{template_path}"
            )
            return

        data_path = filedialog.askopenfilename(
            title="Seleccionar LIBRO CARGA",
            parent=self.root
        )
        if not data_path:
            self.root.destroy()
            return

        loading_ui = CargaUI(self.root, "Procesando Nómina y Fórmulas...")

        threading.Thread(
            target=self._process,
            args=(data_path, template_path, loading_ui),
            daemon=True
        ).start()

        self.root.mainloop()

    # --- Lógica de actualización ---

    def _check_for_updates(self) -> bool:
        updater = GitHubUpdater(GITHUB_USER, GITHUB_REPO)
        has_update, url, tag = updater.verify()

        if not has_update:
            return False

        if messagebox.askyesno(
            "Actualización Nueva",
            f"Hay una versión más reciente disponible ({tag}).\n"
            "¿Deseas actualizar el programa ahora?"
        ):
            loading_ui = CargaUI(
                self.root,
                "Descargando e instalando...\n"
                "El programa se reiniciará solo."
            )
            threading.Thread(
                target=lambda: updater.ejecutar_reemplazo(url),
                daemon=True
            ).start()
            self.root.mainloop()
            return True

        return False
    
    # MÉTODOS PRIVADOS

    def _process(self, data_path, template_path, loading_ui):
        try:
            # 1. Leer los datos fuente
            reader = WorkbookReader(data_path).leer()

            # 2. Abrir plantilla
            workbook = load_workbook(template_path)
            sheet = workbook.active

            # 3. Mapear columnas del excel
            mapper = Mapeador(reader.idx, sheet)

            # 4. Escribir datos
            writer = TemplateWriter(sheet, mapper, reader.idx)
            writer.escribir(reader.filas)

            # 5. Save
            self.root.after(
                0,
                lambda: self._save(workbook, loading_ui)
            )

        except Exception as e:
            self.root.after(
                0, lambda: messagebox.showerror("Error", str(e))
            )
            self.root.after(0, loading_ui.cerrar)

    def _save(self, workbook, loading_ui):
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile=f"Nomina_Procesada_{date.today().strftime('%d_%m_%Y')}",
            title="Guardar Resultado Final",
            parent=self.root
        )

        if not save_path:
            loading_ui.cerrar()
            self.root.destroy()
            return

        # Update message
        for widget in loading_ui.top.winfo_children():
            if isinstance(widget, tk.Label):
                widget.config(
                    text="Escribiendo archivo en disco...\n"
                         "Por favor, no cierre el programa."
                )

        def _save_thread():
            try:
                workbook.save(save_path)
                self.root.after(0, lambda: self._on_success(loading_ui, save_path))
            except Exception as e:
                self.root.after(
                    0, lambda: messagebox.showerror("Error al guardar", str(e))
                )
                self.root.after(0, loading_ui.cerrar)

        threading.Thread(target=_save_thread, daemon=True).start()

    def _on_success(self, loading_ui, save_path):
        loading_ui.cerrar()
        messagebox.showinfo("Éxito", "El archivo se ha guardado correctamente.")
        os.startfile(save_path)
        self.root.destroy()