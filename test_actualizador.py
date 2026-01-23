from actualizador import Actualizador
import tkinter as tk

def probar_sistema():
    # Simulamos que nuestra app está en la 1.0.0
    VERSION_LOCAL = "1.0.0"
    # URL de prueba (puedes usar un Gist de GitHub o un JSON en tu PC)
    URL_JSON = "https://raw.githubusercontent.com/tu_usuario/tu_repo/main/update.json"
    
    app_updater = Actualizador(URL_JSON, VERSION_LOCAL, "ProcesadorNomina")
    
    print("Verificando si hay actualizaciones...")
    hay_nueva, url, version = app_updater.verificar_actualizacion()
    
    if hay_nueva:
        print(f"¡Nueva versión encontrada! {version}")
        # Aquí llamarías a la UI de carga para mostrar que descarga
    else:
        print("El software está al día.")

if __name__ == "__main__":
    probar_sistema()