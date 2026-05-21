import tkinter as tk
from tkinter import ttk, messagebox
import threading

class LoadingUI:
    """Ventana de carga reutilizable"""
    def __init__(self, parent, message="Procesando..."):
        self.top = tk.Toplevel(parent)
        self.top.title("Procesador de Datos")
        self.top.geometry("350x120")
        self.top.resizable(False, False)
        self.top.attributes("-topmost", True)
        
        # Disable close button
        self.top.protocol("WM_DELETE_WINDOW", lambda: None)
        
        # Centering logic
        self.top.update_idletasks()
        width = 350
        height = 120
        x = (self.top.winfo_screenwidth() // 2) - (width // 2)
        y = (self.top.winfo_screenheight() // 2) - (height // 2)
        self.top.geometry(f"+{x}+{y}")

        tk.Label(self.top, text=message, font=("Segoe UI", 10, "bold")).pack(pady=15)
        
        self.progress_bar = ttk.Progressbar(self.top, mode='indeterminate', length=280)
        self.progress_bar.pack(pady=5)
        self.progress_bar.start(15)

    def close(self):
        self.top.destroy()

def run_task_with_loading(root, message, target_function, *args):
    """Launches any function in a thread with a loading screen."""
    loading_screen = LoadingUI(root, message)
    
    def wrapper():
        try:
            target_function(*args)
        except Exception as e:
            messagebox.showerror("Error de Proceso", f"Fallo en la tarea:\n{str(e)}")
        finally:
            # Ensure UI closure happens on the main thread
            root.after(0, loading_screen.close)

    worker_thread = threading.Thread(target=wrapper, daemon=True)
    worker_thread.start()