"""
tablas-python: Suite integral para extracción, transformación, conciliación y exportación de tablas en pandas.
"""

from helpers.table_manager import (
    TableManager,
    obtener_tabla,
    inspeccionar_archivo,
    exportar_archivo_a_csv,
)
from helpers.display_helper import DisplayHelper
from utils.exporter import TableExporter, guardar_csv, guardar_excel
from utils.excel_writer import escribir_en_excel
from utils.batch_processor import unir_archivos_carpeta
from utils.validator import DataValidator, validar_dataframe, reporte_calidad, detectar_duplicados
from utils.data_helpers import (
    limpiar_numero,
    limpiar_columnas_numericas,
    normalizar_fechas,
    formato_moneda,
    formato_porcentaje,
    formato_miles,
    formatear_dataframe,
    agregar_fila_totales,
    calcular_participacion,
    calcular_variacion,
    aplicar_impuesto,
    agrupar_y_resumir,
    obtener_celda,
    modificar_celda,
    buscar_v,
    conciliar_tablas,
)

__version__ = "0.1.0"

__all__ = [
    "TableManager",
    "obtener_tabla",
    "inspeccionar_archivo",
    "exportar_archivo_a_csv",
    "DisplayHelper",
    "TableExporter",
    "guardar_csv",
    "guardar_excel",
    "escribir_en_excel",
    "unir_archivos_carpeta",
    "DataValidator",
    "validar_dataframe",
    "reporte_calidad",
    "detectar_duplicados",
    "limpiar_numero",
    "limpiar_columnas_numericas",
    "normalizar_fechas",
    "formato_moneda",
    "formato_porcentaje",
    "formato_miles",
    "formatear_dataframe",
    "agregar_fila_totales",
    "calcular_participacion",
    "calcular_variacion",
    "aplicar_impuesto",
    "agrupar_y_resumir",
    "obtener_celda",
    "modificar_celda",
    "buscar_v",
    "conciliar_tablas",
]
