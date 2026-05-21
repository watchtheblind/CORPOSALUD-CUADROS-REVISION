import json
import os
import sys
import re


def _source_path(file: str) -> str:
    """Resuelve rutas tanto en desarrollo como empaquetado con PyInstaller."""
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, file)
    return os.path.join(os.path.abspath("."), file)

def _load_json_resource(path: str) -> dict | list:
    """Método genérico para cargar recursos JSON con manejo de rutas."""
    try:
        full_path = _source_path(path)
        with open(full_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except FileNotFoundError:
        logging.error(f"No se encontró el archivo en: {path}")
        return {} # O lanzar una excepción personalizada
    except json.JSONDecodeError:
        logging.error(f"El archivo {path} no tiene un formato JSON válido.")
        return {}

def clean_text(text) -> str:
    """Normaliza texto eliminando todo excepto letras y números."""
    return re.sub(r'[^A-Z0-9]', '', str(text).upper()) if text else ""


def load_mapping(path: str = "config/column_mapping.json") -> dict:
    """
    Carga el JSON de mapeo y devuelve:
    {
      "CEDULA": ["CEDULA", "CEDULA"],
      ...
    }
    """
    data = _load_json_resource(path)
    # Eliminar claves que empiecen con "_" (comentarios)
    return {
        column_key: synonyms 
        for column_key, synonyms in data.items() 
        if not column_key.startswith("_")
    }


def load_concepts_with_factors(path: str = "config/concepts_with_factors.json") -> list:
    """Carga la lista de conceptos que tienen factor adyacente."""
    return _load_json_resource(path)