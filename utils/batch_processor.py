"""
Módulo para procesamiento y consolidación por lotes de carpetas de archivos (PDF, Excel, CSV, SQLite).
Permite unir decenas de archivos periódicos (facturas, reportes mensuales, inventarios) en un único DataFrame maestro.
"""

from typing import List, Optional, Union, Dict, Any
import os
import glob
import pandas as pd
from .file_utils import resolve_file_path


def unir_archivos_carpeta(
    carpeta: str,
    extension: str = "xlsx",
    tabla: Union[int, str] = 1,
    fila_encabezado: Optional[Union[int, List[int], str]] = 1,
    celda_inicio: Optional[str] = None,
    hoja: Optional[Union[str, int]] = None,
    patron: Optional[str] = None,
    recursivo: bool = False,
    agregar_columna_origen: bool = True,
    nombre_col_origen: str = "Archivo_Origen",
    archivo_abierto: Optional[bool] = False,
    skip_footer: int = 0
) -> pd.DataFrame:
    """
    Lee todos los archivos que coincidan con la extensión en una carpeta, extrae la tabla
    especificada de cada uno y los une en un solo DataFrame consolidado.

    Parámetros:
    -----------
    carpeta : str
        Ruta del directorio a escanear (relativa o absoluta).
    extension : str (por defecto 'xlsx')
        Extensión de los archivos a buscar ('xlsx', 'pdf', 'csv', 'db').
    tabla : int o str (por defecto 1)
        Número de tabla o nombre de tabla a extraer en cada archivo.
    fila_encabezado : int, 'auto' o None (por defecto 1)
        Fila donde están los encabezados en cada archivo (ej: 4 si filas 1-3 son basura).
    celda_inicio : str, opcional (ej: 'C4')
        Para Excel, si la tabla comienza en una coordenada específica.
    hoja : str o int, opcional
        Pestaña de Excel a extraer.
    patron : str, opcional
        Patrón comodín adicional (ej: 'factura_*.pdf' o 'reporte_2026_*.xlsx').
    recursivo : bool (por defecto False)
        Si busca también dentro de subcarpetas.
    agregar_columna_origen : bool (por defecto True)
        Si agrega una columna indicando de qué archivo proviene cada fila.

    Retorna:
    --------
    pd.DataFrame
        DataFrame consolidado con la unión de todos los archivos encontrados.
    """
    # Importación diferida para evitar ciclos
    from helpers.table_manager import obtener_tabla

    dir_path = resolve_file_path(carpeta)
    if not os.path.isdir(dir_path):
        raise NotADirectoryError(f"No se encontró el directorio: {dir_path}")

    # Construir patrón de búsqueda
    clean_ext = extension.replace(".", "").lower()
    if patron:
        search_pattern = patron if patron.endswith(f".{clean_ext}") else f"{patron}.{clean_ext}"
    else:
        search_pattern = f"*.{clean_ext}"

    if recursivo:
        search_path = os.path.join(dir_path, "**", search_pattern)
        archivos = glob.glob(search_path, recursive=True)
    else:
        search_path = os.path.join(dir_path, search_pattern)
        archivos = glob.glob(search_path)

    # Filtrar archivos temporales de Office (~$archivo.xlsx)
    archivos = [f for f in archivos if not os.path.basename(f).startswith("~$")]

    if not archivos:
        print(f"⚠️ No se encontraron archivos con extensión '.{clean_ext}' en {dir_path}")
        return pd.DataFrame()

    dfs_list: List[pd.DataFrame] = []

    for file_item in sorted(archivos):
        file_name = os.path.basename(file_item)
        try:
            df_temp = obtener_tabla(
                archivo=file_item,
                tabla=tabla,
                fila_encabezado=fila_encabezado,
                celda_inicio=celda_inicio,
                hoja=hoja,
                skip_footer=skip_footer,
                archivo_abierto=archivo_abierto
            )
            if df_temp is not None and not df_temp.empty:
                if agregar_columna_origen:
                    df_temp[nombre_col_origen] = file_name
                dfs_list.append(df_temp)
        except Exception as e:
            print(f"⚠️ Error al procesar '{file_name}': {e}")

    if not dfs_list:
        return pd.DataFrame()

    # Concatenar todos los DataFrames
    df_consolidado = pd.concat(dfs_list, ignore_index=True)
    return df_consolidado
