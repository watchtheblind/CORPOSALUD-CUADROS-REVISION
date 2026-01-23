import tkinter as tk
from tkinter import ttk, messagebox
import threading

class CargaUI:
    """Ventana de carga profesional y reutilizable."""
    def __init__(self, parent, mensaje="Procesando..."):
        self.top = tk.Toplevel(parent)
        self.top.title("Procesador de Datos")
        self.top.geometry("350x120")
        self.top.resizable(False, False)
        self.top.attributes("-topmost", True)
        self.top.protocol("WM_DELETE_WINDOW", lambda: None)
        
        # Centrado
        self.top.update_idletasks()
        x = (self.top.winfo_screenwidth() // 2) - 175
        y = (self.top.winfo_screenheight() // 2) - 60
        self.top.geometry(f"+{x}+{y}")

        tk.Label(self.top, text=mensaje, font=("Segoe UI", 10, "bold")).pack(pady=15)
        self.pb = ttk.Progressbar(self.top, mode='indeterminate', length=280)
        self.pb.pack(pady=5)
        self.pb.start(15)

    def cerrar(self):
        self.top.destroy()

def ejecutar_tarea_con_carga(root, mensaje, funcion_objetivo, *args):
    """Lanza cualquier función en un hilo con pantalla de carga."""
    loading = CargaUI(root, mensaje)
    
    def wrapper():
        try:
            funcion_objetivo(*args)
        except Exception as e:
            messagebox.showerror("Error de Proceso", f"Fallo en la tarea:\n{str(e)}")
        finally:
            root.after(0, loading.cerrar)

    thread = threading.Thread(target=wrapper, daemon=True)
    thread.start()