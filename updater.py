import requests
import os
import subprocess
import sys

class GitHubUpdater:
    def __init__(self, repo_user, repo_name, local_version):
        self.repo_user = repo_user
        self.repo_name = repo_name
        self.local_version = local_version
        self.url_api = f"https://api.github.com/repos/{repo_user}/{repo_name}/releases/latest"

    def chequear_actualizacion(self):
        """Consulta la API de GitHub por la última versión."""
        try:
            response = requests.get(self.url_api, timeout=10)
            if response.status_code == 200:
                data = response.json()
                tag_name = data['tag_name'] # Ejemplo: "v1.1.0"
                
                if tag_name > self.local_version:
                    # Retornamos la URL del primer asset (el archivo que subas al release)
                    return True, tag_name, data['assets'][0]['browser_download_url']
            return False, None, None
        except Exception as e:
            print(f"Error de conexión: {e}")
            return False, None, None

    def descargar_y_reemplazar(self, url_descarga, nuevo_archivo):
        """Descarga el nuevo binario/script."""
        try:
            r = requests.get(url_descarga, stream=True)
            with open(nuevo_archivo, 'wb') as f:
                for chunk in r.iter_content(chunk_size=8192):
                    f.write(chunk)
            return True
        except Exception as e:
            print(f"Error en descarga: {e}")
            return False