"""
Módulo de utilidades para extracción, limpieza, formateo, cálculo, validación,
escritura en Excel y consolidación por lotes de tablas.
"""

from .file_utils import resolve_file_path, sanitize_filename, ensure_dir
from .table_cleaner import TableCleaner, clean_raw_table
from .pdf_extractor import PDFTableExtractor, RawTableInfo
from .excel_extractor import ExcelTableExtractor, RawExcelTableInfo
from .sqlite_extractor import SQLiteTableExtractor, RawSQLiteTableInfo
from .exporter import TableExporter, guardar_csv, guardar_excel
from .excel_writer import escribir_en_excel
from .batch_processor import unir_archivos_carpeta
from .validator import DataValidator, validar_dataframe, reporte_calidad, detectar_duplicados
from .data_helpers import (
    limpiar_numero,
    limpiar_columnas_numericas,
    normalizar_fechas,
    formato_moneda,
    formato_clp,
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

__all__ = [
    "resolve_file_path",
    "sanitize_filename",
    "ensure_dir",
    "TableCleaner",
    "clean_raw_table",
    "PDFTableExtractor",
    "RawTableInfo",
    "ExcelTableExtractor",
    "RawExcelTableInfo",
    "SQLiteTableExtractor",
    "RawSQLiteTableInfo",
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
    "formato_clp",
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
