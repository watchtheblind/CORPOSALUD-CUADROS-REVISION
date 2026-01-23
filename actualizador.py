import requests

class ActualizadorGitHub:
    def __init__(self, usuario, repo, version_local):
        # URL del manifiesto JSON
        self.url_control = f"https://raw.githubusercontent.com/{usuario}/{repo}/main/versiones.json"
        self.version_local = version_local

    def verificar(self):
        try:
            r = requests.get(self.url_json, timeout=5)
            if r.status_code == 200:
                data = r.json()
                v_remota = data['nomina']['version']
                url_app = data['nomina']['url']
                return v_remota > self.version_local, v_remota, url_app
            return False, None, None
        except:
            return False, None, None

    def descargar(self, url, destino):
        """Descarga el archivo bloqueando lo mínimo posible."""
        try:
            r = requests.get(url, timeout=20)
            if r.status_code == 200:
                with open(ruta_destino, 'wb') as f:
                    f.write(r.content)
                return True
            return False
        except:
            return False