"""
Módulo de utilidades para manejo seguro de rutas de archivos y nombres de exportación.
"""

import os
import sys
import re
from typing import Optional

# Configurar automáticamente la salida UTF-8 en Windows al importar
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass


def resolve_file_path(file_path: str) -> str:
    """
    Resuelve una ruta de archivo ya sea que se pase:
    - Ruta absoluta (ej: C:/Users/.../archivo.xlsx)
    - Ruta relativa al directorio de trabajo actual (ej: samples/archivo.xlsx o ./archivo.xlsx)
    - Nombre simple de archivo en la misma carpeta del script o en la carpeta samples.
    """
    if not file_path:
        raise ValueError("La ruta del archivo no puede estar vacía.")

    # 1. Si es absoluta y existe
    if os.path.isabs(file_path) and os.path.exists(file_path):
        return os.path.abspath(file_path)
    
    # 2. Relativa directa al directorio de trabajo actual (CWD)
    if os.path.exists(file_path):
        return os.path.abspath(file_path)

    # 3. Relativo al directorio raíz del proyecto
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    alt_path = os.path.join(base_dir, file_path)
    if os.path.exists(alt_path):
        return os.path.abspath(alt_path)

    # 4. Dentro de la carpeta samples
    samples_path = os.path.join(base_dir, "samples", file_path)
    if os.path.exists(samples_path):
        return os.path.abspath(samples_path)

    # Si no se encuentra, retornar la ruta absoluta esperada para que el error sea claro
    return os.path.abspath(file_path)


def sanitize_filename(name: str, replacement: str = "_") -> str:
    """
    Limpia una cadena para que sea un nombre de archivo válido en Windows/Linux/Mac.
    """
    # Eliminar caracteres no permitidos en nombres de archivo
    sanitized = re.sub(r'[\\/*?:"<>|]', replacement, name)
    # Reducir espacios y guiones repetidos
    sanitized = re.sub(r'\s+', '_', sanitized).strip('._ ')
    return sanitized if sanitized else "tabla"


def ensure_dir(dir_path: str) -> str:
    """Asegura que el directorio exista creándolo recursivamente si es necesario."""
    os.makedirs(dir_path, exist_ok=True)
    return os.path.abspath(dir_path)
