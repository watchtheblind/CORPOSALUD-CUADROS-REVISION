import requests
import os
import sys
import datetime

class ActualizadorGitHub:
    """
    SISTEMA DE ACTUALIZACIÓN AUTÓNOMA POR LÍNEA DE TIEMPO (CORPOSALUD)
    -----------------------------------------------------------------
    Lógica de funcionamiento:

    1. COMPARACIÓN TEMPORAL (Cero Hardcoding):
       En lugar de usar números de versión (1.0.1, etc.), este script compara la fecha 
       de última modificación del archivo .exe local contra la fecha de publicación 
       del 'Latest Release' en GitHub. Si la nube es más reciente, se activa el update.

    2. EL TRUCO DEL RELEVO (.BAT):
       Windows prohíbe que un programa se borre o modifique a sí mismo mientras se ejecuta. 
       Para saltar esta restricción:
       a) Descargamos el nuevo .exe con un nombre temporal (temp_update.exe).
       b) Creamos un script de procesamiento por lotes (update.bat) externo.
       c) La aplicación principal se cierra inmediatamente para liberar el archivo .exe.

    3. EL SCRIPT 'ASESINO' (update.bat):
       - Espera 2 segundos (timeout) para asegurar que el proceso de Python terminó.
       - Elimina el ejecutable viejo (ya desbloqueado).
       - Renombra el 'temp_update.exe' al nombre original del programa.
       - Lanza la nueva versión.
       - Se auto-elimina (del %~f0) para no dejar rastro.

    REQUISITO: El nombre del .exe en el Release de GitHub DEBE SER IDÉNTICO al local.
    """
    def __init__(self, usuario, repo):
        self.usuario = usuario
        self.repo = repo
        self.url_api = f"https://api.github.com/repos/{usuario}/{repo}/releases/latest"

    def verificar(self):
        try:
            # 1. Obtener fecha de creación del .exe actual
            ruta_exe = sys.executable
            fecha_local = os.path.getmtime(ruta_exe)
            dt_local = datetime.datetime.fromtimestamp(fecha_local, datetime.timezone.utc)

            # 2. Consultar GitHub
            response = requests.get(self.url_api, timeout=5)
            if response.status_code == 200:
                data = response.json()
                
                # Fecha de publicación del release (ISO format)
                fecha_github_str = data.get("published_at")
                dt_github = datetime.datetime.strptime(fecha_github_str, "%Y-%m-%dT%H:%M:%SZ").replace(tzinfo=datetime.timezone.utc)

                # COMPARACIÓN: ¿El de la nube es más nuevo que el mío?
                if dt_github > dt_local:
                    assets = data.get("assets", [])
                    url_descarga = next((a["browser_download_url"] for a in assets if a["name"].endswith(".exe")), None)
                    return True, url_descarga, data.get("tag_name")
            
            return False, None, None
        except Exception as e:
            print(f"Error verificando: {e}")
            return False, None, None

    def ejecutar_reemplazo(self, url_descarga):
        exe_actual = sys.executable
        ruta_dir = os.path.dirname(exe_actual)
        temp_exe = os.path.join(ruta_dir, "temp_update.exe")

        # Descarga
        r = requests.get(url_descarga, stream=True)
        with open(temp_exe, 'wb') as f:
            for chunk in r.iter_content(chunk_size=8192):
                f.write(chunk)

        # BAT de relevo
        bat_path = os.path.join(ruta_dir, "update.bat")
        with open(bat_path, "w") as f:
            f.write(f'@echo off\n')
            f.write(f'timeout /t 2 /nobreak > nul\n')
            f.write(f'del /f /q "{exe_actual}"\n')
            f.write(f'move /y "{temp_exe}" "{exe_actual}"\n')
            f.write(f'start "" "{exe_actual}"\n')
            f.write(f'del "%~f0"\n')

        os.startfile(bat_path)
        sys.exit()