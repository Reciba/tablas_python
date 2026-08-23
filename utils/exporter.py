"""
Módulo dedicado a la exportación de DataFrames y lotes de tablas a archivos CSV y Excel.
Optimizado con codificación utf-8-sig y separadores configurables para máxima compatibilidad con Microsoft Excel.
"""

from typing import Dict, List, Optional, Union, Any
import os
import pandas as pd
from .file_utils import ensure_dir, sanitize_filename


class TableExporter:
    """
    Clase de exportación para guardar DataFrames individuales o colecciones completas de tablas.
    """

    @staticmethod
    def to_csv(
        df: pd.DataFrame,
        output_path: str,
        sep: str = ";",
        encoding: str = "utf-8-sig",
        index: bool = False,
        decimal: str = ",",
        **kwargs
    ) -> str:
        """
        Exporta un DataFrame a archivo CSV.

        Parámetros:
        -----------
        df : pd.DataFrame
            El DataFrame a exportar.
        output_path : str
            Ruta del archivo CSV destino.
        sep : str (por defecto ';')
            Separador de campos (';' abre perfectamente en Excel en español sin desfasar columnas).
        encoding : str (por defecto 'utf-8-sig')
            Codificación con BOM para que tildes, caracteres especiales y la 'ñ' se vean bien en Excel.
        index : bool (por defecto False)
            Si incluye la columna de índice numérico de pandas.
        decimal : str (por defecto ',')
            Separador decimal para valores numéricos en el CSV.

        Retorna:
        --------
        str
            Ruta absoluta del archivo CSV generado.
        """
        if not output_path.lower().endswith(".csv"):
            output_path += ".csv"

        ensure_dir(os.path.dirname(os.path.abspath(output_path)))
        
        df.to_csv(
            output_path,
            sep=sep,
            encoding=encoding,
            index=index,
            decimal=decimal,
            **kwargs
        )
        return os.path.abspath(output_path)

    @staticmethod
    def to_excel(
        df: pd.DataFrame,
        output_path: str,
        sheet_name: str = "Datos",
        index: bool = False,
        **kwargs
    ) -> str:
        """
        Exporta un DataFrame a un archivo Excel (.xlsx).
        """
        if not output_path.lower().endswith((".xlsx", ".xls")):
            output_path += ".xlsx"

        ensure_dir(os.path.dirname(os.path.abspath(output_path)))
        
        df.to_excel(
            output_path,
            sheet_name=sheet_name,
            index=index,
            engine="openpyxl",
            **kwargs
        )
        return os.path.abspath(output_path)

    @classmethod
    def export_batch_to_csv(
        cls,
        tables_dict: Dict[str, pd.DataFrame],
        output_dir: str,
        sep: str = ";",
        encoding: str = "utf-8-sig",
        index: bool = False
    ) -> List[str]:
        """
        Exporta múltiples DataFrames en archivos CSV individuales dentro de un directorio.

        Parámetros:
        -----------
        tables_dict : dict
            Diccionario {nombre_tabla: df} con los DataFrames a exportar.
        output_dir : str
            Directorio donde se guardarán los archivos CSV.

        Retorna:
        --------
        List[str]
            Lista con las rutas absolutas de todos los archivos CSV creados.
        """
        ensure_dir(output_dir)
        generated_files: List[str] = []

        for name, df in tables_dict.items():
            if df is None or df.empty:
                continue
            
            clean_name = sanitize_filename(str(name))
            file_name = f"{clean_name}.csv"
            out_file = os.path.join(output_dir, file_name)
            
            saved_path = cls.to_csv(
                df=df,
                output_path=out_file,
                sep=sep,
                encoding=encoding,
                index=index
            )
            generated_files.append(saved_path)

        return generated_files


# Funciones directas de conveniencia
def guardar_csv(
    df: pd.DataFrame,
    ruta: str,
    sep: str = ";",
    encoding: str = "utf-8-sig",
    index: bool = False,
    **kwargs
) -> str:
    """Función de una línea para guardar un DataFrame a CSV optimizado para Excel."""
    return TableExporter.to_csv(df, output_path=ruta, sep=sep, encoding=encoding, index=index, **kwargs)


def guardar_excel(
    df: pd.DataFrame,
    ruta: str,
    sheet_name: str = "Datos",
    index: bool = False,
    **kwargs
) -> str:
    """Función de una línea para guardar un DataFrame a Excel."""
    return TableExporter.to_excel(df, output_path=ruta, sheet_name=sheet_name, index=index, **kwargs)
