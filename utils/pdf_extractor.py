"""
Módulo para extracción de tablas desde archivos PDF usando pdfplumber.
Soporta detección de múltiples tablas por página, configuraciones de extracción
y escaneo página por página o completo.
"""

from typing import List, Dict, Any, Optional
from dataclasses import dataclass
import os
import pdfplumber
from .file_utils import resolve_file_path


@dataclass
class RawTableInfo:
    """Información y contenido crudo de una tabla detectada."""
    table_id: int
    source_file: str
    page_number: int
    table_in_page: int
    raw_data: List[List[Any]]
    num_rows: int
    num_cols: int
    bbox: Optional[tuple] = None

    def get_preview(self, max_rows: int = 5) -> List[List[Any]]:
        """Retorna las primeras filas crudas para inspección visual."""
        return self.raw_data[:max_rows]


class PDFTableExtractor:
    """
    Extractor de tablas especializado en documentos PDF.
    """

    def __init__(self, file_path: str):
        self.file_path = resolve_file_path(file_path)
        if not os.path.exists(self.file_path):
            raise FileNotFoundError(f"El archivo PDF no existe: {self.file_path}")

    def extract_all_tables(
        self,
        pages: Optional[List[int]] = None,
        table_settings: Optional[Dict[str, Any]] = None,
        fallback_text_strategy: bool = True
    ) -> List[RawTableInfo]:
        """
        Extrae todas las tablas del PDF.

        Parámetros:
        -----------
        pages : list of int, opcional
            Lista de números de página a procesar (1-indexed). Si es None, procesa todo el PDF.
        table_settings : dict, opcional
            Configuración personalizada para pdfplumber (ej. vertical_strategy, horizontal_strategy).
        fallback_text_strategy : bool (por defecto True)
            Si con la estrategia estándar no se detectan tablas, intenta con estrategia de texto.

        Retorna:
        --------
        List[RawTableInfo]
            Lista con la información y datos crudos de cada tabla encontrada.
        """
        tables_found: List[RawTableInfo] = []
        global_table_id = 1

        with pdfplumber.open(self.file_path) as pdf:
            total_pages = len(pdf.pages)
            page_indices = range(total_pages)
            
            if pages is not None:
                # Convertir a 0-indexed y filtrar válidas
                page_indices = [p - 1 for p in pages if 1 <= p <= total_pages]

            for p_idx in page_indices:
                page = pdf.pages[p_idx]
                page_num = p_idx + 1
                
                # 1. Intento con settings provistos o estándar
                extracted = page.extract_tables(table_settings) if table_settings else page.extract_tables()
                
                # 2. Si no encontró y fallback está activo, probar estrategia text
                if not extracted and fallback_text_strategy:
                    alt_settings = {
                        "vertical_strategy": "text",
                        "horizontal_strategy": "text",
                        "snap_tolerance": 3,
                    }
                    extracted = page.extract_tables(alt_settings)

                # Procesar cada tabla detectada en la página
                for table_idx, raw_table in enumerate(extracted, start=1):
                    # Validar que la tabla tenga contenido útil
                    if not raw_table or len(raw_table) == 0:
                        continue
                    
                    # Calcular filas y columnas
                    n_rows = len(raw_table)
                    n_cols = max((len(r) for r in raw_table if isinstance(r, list)), default=0)
                    
                    info = RawTableInfo(
                        table_id=global_table_id,
                        source_file=self.file_path,
                        page_number=page_num,
                        table_in_page=table_idx,
                        raw_data=raw_table,
                        num_rows=n_rows,
                        num_cols=n_cols
                    )
                    tables_found.append(info)
                    global_table_id += 1

        return tables_found

    def extract_table_by_id(
        self,
        table_id: int,
        table_settings: Optional[Dict[str, Any]] = None
    ) -> RawTableInfo:
        """
        Extrae y devuelve directamente la tabla con el ID global indicado (1-indexed).
        """
        all_tables = self.extract_all_tables(table_settings=table_settings)
        for t in all_tables:
            if t.table_id == table_id:
                return t
        raise IndexError(
            f"No se encontró la tabla {table_id}. Tablas disponibles: {len(all_tables)}"
        )
