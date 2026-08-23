"""
Módulo para extracción de tablas y consultas desde bases de datos SQLite (.db, .sqlite, .sqlite3).
Permite listar tablas, inspeccionar esquemas y extraer DataFrames de tablas completas o consultas SQL.
"""

from typing import List, Dict, Any, Optional, Union
from dataclasses import dataclass
import os
import sqlite3
import pandas as pd
from .file_utils import resolve_file_path


@dataclass
class RawSQLiteTableInfo:
    """Información de una tabla o vista detectada en una base de datos SQLite."""
    table_id: int
    source_file: str
    table_name: str
    table_type: str  # 'table' o 'view'
    columns: List[str]
    num_rows: int
    raw_data: List[List[Any]]
    num_cols: int

    def get_preview(self, max_rows: int = 5) -> List[List[Any]]:
        """Retorna los encabezados y las primeras filas de datos."""
        if not self.raw_data:
            return [self.columns]
        return [self.columns] + self.raw_data[:max_rows]


class SQLiteTableExtractor:
    """
    Extractor de datos para bases de datos SQLite.
    """

    def __init__(self, file_path: str):
        self.file_path = resolve_file_path(file_path)
        if not os.path.exists(self.file_path):
            raise FileNotFoundError(f"No se encontró la base de datos SQLite en: {self.file_path}")

    def list_tables(self) -> List[Dict[str, str]]:
        """Lista todas las tablas y vistas disponibles en la base de datos."""
        with sqlite3.connect(self.file_path) as conn:
            cursor = conn.cursor()
            cursor.execute(
                "SELECT name, type FROM sqlite_master "
                "WHERE type IN ('table', 'view') AND name NOT LIKE 'sqlite_%' "
                "ORDER BY name"
            )
            rows = cursor.fetchall()
            return [{"name": r[0], "type": r[1]} for r in rows]

    def extract_all_tables(self) -> List[RawSQLiteTableInfo]:
        """
        Extrae la información y una muestra de datos de todas las tablas en la base de datos.
        """
        tables = self.list_tables()
        tables_found: List[RawSQLiteTableInfo] = []

        with sqlite3.connect(self.file_path) as conn:
            for idx, item in enumerate(tables, start=1):
                t_name = item["name"]
                t_type = item["type"]
                
                # Obtener columnas
                cursor = conn.cursor()
                cursor.execute(f"PRAGMA table_info('{t_name}')")
                col_info = cursor.fetchall()
                cols = [c[1] for c in col_info] if col_info else []

                # Obtener conteo de filas
                try:
                    cursor.execute(f"SELECT COUNT(*) FROM '{t_name}'")
                    row_count = cursor.fetchone()[0]
                except Exception:
                    row_count = 0

                # Obtener primeras filas
                cursor.execute(f"SELECT * FROM '{t_name}' LIMIT 20")
                sample_data = [list(r) for r in cursor.fetchall()]

                info = RawSQLiteTableInfo(
                    table_id=idx,
                    source_file=self.file_path,
                    table_name=t_name,
                    table_type=t_type,
                    columns=cols,
                    num_rows=row_count,
                    raw_data=sample_data,
                    num_cols=len(cols)
                )
                tables_found.append(info)

        return tables_found

    def get_table_df(self, table_name_or_id: Union[str, int]) -> pd.DataFrame:
        """
        Obtiene un DataFrame con la tabla completa por su nombre o su ID (1-indexed).
        """
        tables = self.list_tables()
        selected_name = None

        if isinstance(table_name_or_id, int):
            if 1 <= table_name_or_id <= len(tables):
                selected_name = tables[table_name_or_id - 1]["name"]
            else:
                raise IndexError(f"ID de tabla {table_name_or_id} fuera de rango. Hay {len(tables)} tablas.")
        else:
            table_str = str(table_name_or_id).strip()
            # Buscar coincidencia exacta o insensible a mayúsculas
            for t in tables:
                if t["name"].lower() == table_str.lower():
                    selected_name = t["name"]
                    break
            if not selected_name:
                raise ValueError(f"No se encontró la tabla '{table_name_or_id}' en {self.file_path}")

        with sqlite3.connect(self.file_path) as conn:
            df = pd.read_sql_query(f"SELECT * FROM '{selected_name}'", conn)
        return df

    def query_df(self, sql_query: str, params: Optional[Union[tuple, dict]] = None) -> pd.DataFrame:
        """
        Ejecuta una consulta SQL personalizada y devuelve el resultado en un DataFrame.
        """
        with sqlite3.connect(self.file_path) as conn:
            if params is not None:
                df = pd.read_sql_query(sql_query, conn, params=params)
            else:
                df = pd.read_sql_query(sql_query, conn)
        return df
