"""
Módulo principal de gestión de tablas para interactuar fácilmente desde main.py.
Permite abrir archivos PDF (pdfplumber), Excel (xlwings - abierto o cerrado) y SQLite (.db/.sqlite),
inspeccionar qué tablas existen y extraer DataFrames limpios en pandas indicando
la tabla y la fila del encabezado o ejecutando consultas.
"""

from typing import List, Optional, Union, Any, Dict
import os
import pandas as pd

from utils.pdf_extractor import PDFTableExtractor, RawTableInfo
from utils.excel_extractor import ExcelTableExtractor, RawExcelTableInfo
from utils.sqlite_extractor import SQLiteTableExtractor, RawSQLiteTableInfo
from utils.table_cleaner import TableCleaner
from utils.file_utils import resolve_file_path
from utils.exporter import TableExporter, guardar_csv, guardar_excel
from helpers.display_helper import DisplayHelper


class TableManager:
    """
    Gestor unificado de tablas para PDFs, Excel/CSV (xlwings) y bases de datos SQLite.
    
    Ejemplos de uso en main.py:
    --------------------------
    # 1. PDF
    doc = TableManager("samples/ejemplo_facturas.pdf")
    df = doc.get_df(tabla=3, fila_encabezado=4)

    # 2. Excel (detecta si está abierto o cerrado automáticamente)
    doc_xl = TableManager("inventario.xlsx", archivo_abierto=None)
    df_inv = doc_xl.get_df(tabla=1, fila_encabezado=4)

    # 3. SQLite
    db = TableManager("empresa.db")
    df_ventas = db.get_df("ventas") # o db.get_df(tabla=1)
    """

    def __init__(
        self,
        file_path: str,
        archivo_abierto: Optional[bool] = None,
        preferir_xlwings: bool = True,
        auto_load: bool = True
    ):
        """
        Parámetros:
        -----------
        file_path : str
            Ruta del archivo (relativa o absoluta).
        archivo_abierto : bool o None (para Excel)
            - True: se conecta al libro ya abierto en Excel.
            - False: abre el libro cerrado en segundo plano.
            - None: autodetecta si el archivo ya está abierto en Excel.
        preferir_xlwings : bool (por defecto True)
            Utiliza xlwings para Excel con fallback a openpyxl.
        auto_load : bool (por defecto True)
            Carga las tablas automáticamente al instanciar.
        """
        self.file_path = resolve_file_path(file_path)
        if not os.path.exists(self.file_path):
            raise FileNotFoundError(f"No se encontró el archivo: {self.file_path}")
        
        self.ext = os.path.splitext(self.file_path)[1].lower()
        self.archivo_abierto = archivo_abierto
        self.preferir_xlwings = preferir_xlwings
        self.tables: List[Union[RawTableInfo, RawExcelTableInfo, RawSQLiteTableInfo]] = []
        
        valid_extensions = ['.pdf', '.xlsx', '.xls', '.xlsm', '.csv', '.db', '.sqlite', '.sqlite3', '.db3']
        if self.ext not in valid_extensions:
            raise ValueError(f"Extensión no soportada: {self.ext}. Formatos soportados: {valid_extensions}")

        if auto_load:
            self.cargar_tablas()

    def cargar_tablas(
        self,
        paginas: Optional[List[int]] = None,
        hojas: Optional[List[Union[str, int]]] = None,
        pdf_settings: Optional[Dict[str, Any]] = None,
        dividir_bloques_excel: bool = False
    ) -> List[Any]:
        """
        Carga o recarga las tablas del archivo aplicando filtros de páginas u hojas.
        """
        if self.ext == '.pdf':
            extractor = PDFTableExtractor(self.file_path)
            self.tables = extractor.extract_all_tables(pages=paginas, table_settings=pdf_settings)
        elif self.ext in ['.db', '.sqlite', '.sqlite3', '.db3']:
            extractor = SQLiteTableExtractor(self.file_path)
            self.tables = extractor.extract_all_tables()
        else:
            extractor = ExcelTableExtractor(
                self.file_path,
                archivo_abierto=self.archivo_abierto,
                preferir_xlwings=self.preferir_xlwings
            )
            self.tables = extractor.extract_all_tables(sheets=hojas, split_blank_blocks=dividir_bloques_excel)
        
        return self.tables

    @property
    def total_tablas(self) -> int:
        """Cantidad total de tablas detectadas en el documento o base de datos."""
        return len(self.tables)

    def resumen(self):
        """Muestra en consola un resumen claro de todas las tablas encontradas."""
        if not self.tables:
            print(f"⚠️ No se detectaron tablas en {self.file_path}.")
            return
        DisplayHelper.print_tables_summary(self.tables, self.file_path)

    def ver_crudo(self, tabla: Union[int, str] = 1, max_filas: int = 8):
        """
        Muestra la vista previa cruda de una tabla con sus números de fila (1-indexed).
        Muy útil para identificar exactamente en qué fila están los encabezados reales.
        """
        table_info = self._get_raw_table_info(tabla)
        DisplayHelper.print_raw_preview(table_info, max_rows=max_filas)

    def _get_raw_table_info(self, tabla: Union[int, str]) -> Any:
        """Obtiene el objeto de información cruda de una tabla por su ID (1-indexed) o nombre."""
        if isinstance(tabla, int):
            if tabla < 1 or tabla > len(self.tables):
                raise IndexError(
                    f"Tabla {tabla} no válida. El archivo tiene {len(self.tables)} tabla(s) detectada(s)."
                )
            return self.tables[tabla - 1]
        
        # Buscar por nombre (para SQLite o Excel)
        target_name = str(tabla).strip().lower()
        for t in self.tables:
            if hasattr(t, 'table_name') and t.table_name.lower() == target_name:
                return t
            if hasattr(t, 'sheet_name') and t.sheet_name.lower() == target_name:
                return t
        
        raise ValueError(f"No se encontró la tabla o pestaña '{tabla}' en {self.file_path}")

    def get_df(
        self,
        tabla: Union[int, str] = 1,
        fila_encabezado: Optional[Union[int, List[int], str]] = 1,
        celda_inicio: Optional[str] = None,
        celda: Optional[str] = None,
        rango: Optional[str] = None,
        hoja: Optional[Union[str, int]] = None,
        skip_footer: int = 0,
        eliminar_filas_vacias: bool = True,
        eliminar_columnas_vacias: bool = True,
        auto_inferir_tipos: bool = True,
    ) -> pd.DataFrame:
        """
        Extrae y limpia la tabla deseada devolviendo un pandas DataFrame limpio.

        Parámetros:
        -----------
        tabla : int o str (por defecto 1)
            - int: Número de tabla (1-indexed). Ej: tabla=3.
            - str: Nombre oficial de tabla en Excel (ej: 'TablaProductos'),
                   nombre de tabla en SQLite (ej: 'ventas'), o nombre de hoja.
        fila_encabezado : int, list[int], 'auto' o None (por defecto 1)
            Número de fila que contiene los nombres de columnas (1-indexed).
            Ejemplo: fila_encabezado=4 descarta filas 1, 2, 3 como basura.
            (En SQLite o cuando se usa celda_inicio/rango exacto, se suele usar 1).
        celda_inicio / celda : str, opcional (ej: 'C4', 'B3')
            En Excel, coordenada de la celda donde inicia la tabla.
        rango : str, opcional (ej: 'C4:F20')
            En Excel, rango exacto a extraer.
        hoja : str o int, opcional
            Nombre o índice de la hoja en Excel donde se ubica la celda o rango.
        skip_footer : int (por defecto 0)
            Cantidad de filas finales a descartar (por ejemplo notas al pie o totales).
        eliminar_filas_vacias : bool (por defecto True)
            Descarta filas completamente vacías.
        eliminar_columnas_vacias : bool (por defecto True)
            Descarta columnas completamente vacías.
        auto_inferir_tipos : bool (por defecto True)
            Limpia signos monetarios y convierte valores numéricos cuando sea posible.

        Retorna:
        --------
        pd.DataFrame
            DataFrame de pandas listo para ser usado.
        """
        # Caso especial para SQLite
        if self.ext in ['.db', '.sqlite', '.sqlite3', '.db3']:
            extractor = SQLiteTableExtractor(self.file_path)
            return extractor.get_table_df(tabla)

        # Caso extracción directa por celda_inicio o rango en Excel
        target_cell = celda_inicio or celda
        if target_cell or rango:
            extractor = ExcelTableExtractor(
                self.file_path,
                archivo_abierto=self.archivo_abierto,
                preferir_xlwings=self.preferir_xlwings
            )
            raw_matrix = extractor.extract_by_cell_or_range(
                celda_inicio=target_cell,
                rango=rango,
                hoja=hoja
            )
            return TableCleaner.clean(
                raw_data=raw_matrix,
                header_row=fila_encabezado,
                skip_footer=skip_footer,
                drop_empty_rows=eliminar_filas_vacias,
                drop_empty_cols=eliminar_columnas_vacias,
                auto_clean_types=auto_inferir_tipos,
            )

        # Caso estándar por ID o por nombre de tabla/hoja
        table_info = self._get_raw_table_info(tabla)
        
        # Si la tabla es una Named Table de Excel que ya parte exactamente en su rango,
        # los encabezados están en la fila 1 de dicha tabla
        header_to_use = fila_encabezado
        if hasattr(table_info, 'start_cell') and table_info.start_cell and getattr(table_info, 'table_name', None) != table_info.sheet_name:
            # Es una tabla con nombre oficial que inicia en start_cell
            if fila_encabezado == 1 or fila_encabezado == 4: # si el usuario no especificó otra cosa
                header_to_use = 1

        df = TableCleaner.clean(
            raw_data=table_info.raw_data,
            header_row=header_to_use,
            skip_footer=skip_footer,
            drop_empty_rows=eliminar_filas_vacias,
            drop_empty_cols=eliminar_columnas_vacias,
            auto_clean_types=auto_inferir_tipos,
        )
        return df

    # Alias para máxima flexibilidad
    get_tabla = get_df
    get_table = get_df

    def query(self, sql_query: str, params: Optional[Union[tuple, dict]] = None) -> pd.DataFrame:
        """
        Ejecuta una consulta SQL si el archivo cargado es una base de datos SQLite.
        """
        if self.ext not in ['.db', '.sqlite', '.sqlite3', '.db3']:
            raise ValueError("El método query() solo está disponible para bases de datos SQLite.")
        extractor = SQLiteTableExtractor(self.file_path)
        return extractor.query_df(sql_query, params=params)

    def exportar(
        self,
        df: pd.DataFrame,
        ruta_salida: str,
        formato: Optional[str] = None,
        sep: str = ";",
        encoding: str = "utf-8-sig",
        index: bool = False
    ) -> str:
        """
        Exporta un DataFrame a Excel o CSV.
        """
        if formato is None:
            _, ext = os.path.splitext(ruta_salida)
            formato = ext.replace(".", "").lower()
            if not formato:
                formato = "csv" if ruta_salida.lower().endswith(".csv") else "xlsx"

        if formato in ['csv']:
            return TableExporter.to_csv(df, output_path=ruta_salida, sep=sep, encoding=encoding, index=index)
        elif formato in ['xlsx', 'excel']:
            return TableExporter.to_excel(df, output_path=ruta_salida, index=index)
        else:
            raise ValueError(f"Formato no soportado: {formato}. Use 'csv' o 'xlsx'.")

    def exportar_csv(
        self,
        df: pd.DataFrame,
        ruta_salida: str,
        sep: str = ";",
        encoding: str = "utf-8-sig",
        index: bool = False
    ) -> str:
        """Exporta un DataFrame a CSV con separador y codificación optimizados para Excel."""
        return TableExporter.to_csv(df, output_path=ruta_salida, sep=sep, encoding=encoding, index=index)

    def exportar_todas_a_csv(
        self,
        directorio_salida: str,
        fila_encabezado: Optional[Union[int, List[int], str]] = 1,
        sep: str = ";",
        encoding: str = "utf-8-sig",
        index: bool = False,
        **clean_kwargs
    ) -> List[str]:
        """
        Extrae y exporta TODAS las tablas encontradas en el archivo a archivos CSV individuales.

        Parámetros:
        -----------
        directorio_salida : str
            Directorio donde se guardarán los archivos CSV.
        fila_encabezado : int, 'auto', etc.
            Fila de encabezado a aplicar para tablas de PDF/Excel.

        Retorna:
        --------
        List[str]
            Lista con las rutas de todos los archivos CSV generados.
        """
        tables_dict: Dict[str, pd.DataFrame] = {}
        base_name = os.path.splitext(os.path.basename(self.file_path))[0]

        for idx, t_info in enumerate(self.tables, start=1):
            if self.ext in ['.db', '.sqlite', '.sqlite3', '.db3']:
                t_name = getattr(t_info, 'table_name', f"tabla_{idx}")
                df = self.get_df(tabla=t_name)
                key = f"{base_name}_{t_name}"
            else:
                t_name = getattr(t_info, 'table_name', None) or getattr(t_info, 'sheet_name', None)
                if t_name and t_name != "CSV":
                    key = f"{base_name}_{t_name}"
                elif hasattr(t_info, 'page_number'):
                    key = f"{base_name}_pag_{t_info.page_number}_tab_{t_info.table_in_page}"
                else:
                    key = f"{base_name}_tabla_{idx}"

                df = self.get_df(tabla=idx, fila_encabezado=fila_encabezado, **clean_kwargs)

            tables_dict[key] = df

        saved_files = TableExporter.export_batch_to_csv(
            tables_dict=tables_dict,
            output_dir=directorio_salida,
            sep=sep,
            encoding=encoding,
            index=index
        )
        return saved_files


def obtener_tabla(
    archivo: str,
    tabla: Union[int, str] = 1,
    fila_encabezado: Optional[Union[int, List[int], str]] = 1,
    celda_inicio: Optional[str] = None,
    celda: Optional[str] = None,
    rango: Optional[str] = None,
    hoja: Optional[Union[str, int]] = None,
    skip_footer: int = 0,
    archivo_abierto: Optional[bool] = None,
    **kwargs
) -> pd.DataFrame:
    """
    Función de una sola línea para extraer directamente un DataFrame desde cualquier PDF, Excel o SQLite.

    Ejemplos:
    ---------
    # PDF (Tabla 3, encabezado en fila 4):
    df = obtener_tabla("facturas.pdf", tabla=3, fila_encabezado=4)

    # Excel por Celda de Inicio (ej: tabla parte en la celda C4):
    df = obtener_tabla("inventario.xlsx", celda_inicio="C4", hoja="Inventario")

    # Excel por Nombre Oficial de Tabla:
    df = obtener_tabla("inventario.xlsx", tabla="TablaProductos")

    # Excel por Rango exacto:
    df = obtener_tabla("inventario.xlsx", rango="B3:F15")

    # SQLite por Nombre de Tabla:
    df = obtener_tabla("empresa.db", tabla="ventas")
    """
    manager = TableManager(archivo, archivo_abierto=archivo_abierto)
    return manager.get_df(
        tabla=tabla,
        fila_encabezado=fila_encabezado,
        celda_inicio=celda_inicio,
        celda=celda,
        rango=rango,
        hoja=hoja,
        skip_footer=skip_footer,
        **kwargs
    )


def inspeccionar_archivo(archivo: str, max_filas_preview: int = 8, archivo_abierto: Optional[bool] = None):
    """
    Función de ayuda rápida para ver todas las tablas y sus primeras filas con números de fila.
    """
    manager = TableManager(archivo, archivo_abierto=archivo_abierto)
    manager.resumen()
    for i in range(1, manager.total_tablas + 1):
        manager.ver_crudo(tabla=i, max_filas=max_filas_preview)


def exportar_archivo_a_csv(
    archivo: str,
    carpeta_salida: str,
    fila_encabezado: Optional[Union[int, List[int], str]] = 1,
    sep: str = ";",
    encoding: str = "utf-8-sig",
    archivo_abierto: Optional[bool] = None,
    **kwargs
) -> List[str]:
    """
    Función de una sola línea para extraer y guardar TODAS las tablas de un archivo en archivos CSV individuales.

    Ejemplo:
    --------
    archivos_guardados = exportar_archivo_a_csv("facturas.pdf", carpeta_salida="exports/facturas", fila_encabezado=4)
    """
    manager = TableManager(archivo, archivo_abierto=archivo_abierto)
    return manager.exportar_todas_a_csv(
        directorio_salida=carpeta_salida,
        fila_encabezado=fila_encabezado,
        sep=sep,
        encoding=encoding,
        **kwargs
    )
