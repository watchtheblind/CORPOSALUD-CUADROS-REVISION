import requests

class ActualizadorGitHub:
    def __init__(self, usuario, repo, version_local):
        # URL del manifiesto JSON
        self.url_control = f"https://raw.githubusercontent.com/{usuario}/{repo}/main/versiones.json"
        self.version_local = version_local

    def verificar(self):
        """Consulta GitHub para ver si hay una versión superior."""
        try:
            respuesta = requests.get(self.url_control, timeout=10)
            if respuesta.status_code == 200:
                datos = respuesta.json()
                # Accedemos a la clave 'nomina' que definimos en el JSON
                info = datos.get('nomina', {})
                version_remota = info.get('version', "0.0.0")
                url_descarga = info.get('url')
                
                # Comparamos versiones (ejemplo: "1.0.1" > "1.0.0")
                if version_remota > self.version_local:
                    return True, url_descarga, version_remota
            return False, None, None
        except Exception as e:
            print(f"Error al conectar con GitHub: {e}")
            return False, None, None

    def descargar(self, url):
        """Descarga el archivo desde la URL y lo guarda localmente."""
        try:
            r = requests.get(url, timeout=15)
            if r.status_code == 200:
                # Lo guardamos con un nombre distinto para no sobreescribir el que está abierto
                with open("ejecutable_NUEVO.py", 'wb') as f:
                    f.write(r.content)
                return True
            return False
        except Exception as e:
            print(f"Error en la descarga: {e}")
            return False