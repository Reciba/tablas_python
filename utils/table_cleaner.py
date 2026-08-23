"""
Módulo para limpiar y estructurar datos tabulares crudos en DataFrames de pandas.
Permite descartar filas de 'basura' o metadatos superiores indicando la fila exacta
del encabezado (por ejemplo, fila 4), eliminar filas/columnas vacías y normalizar nombres.
"""

from typing import List, Any, Optional, Union
import pandas as pd
import numpy as np


class TableCleaner:
    """
    Clase utilitaria para transformar matrices de datos crudos (listas de listas o DataFrames sucios)
    en DataFrames limpios y listos para análisis.
    """

    @staticmethod
    def _normalize_cell(val: Any) -> Any:
        """Limpia un valor individual: elimina saltos de línea molestos y espacios extra."""
        if val is None:
            return np.nan
        if isinstance(val, str):
            val_str = val.strip().replace("\r\n", " ").replace("\n", " ")
            # Reducir múltiples espacios consecutivos
            val_str = " ".join(val_str.split())
            return val_str if val_str != "" else np.nan
        return val

    @classmethod
    def _matrix_to_dataframe(cls, raw_data: Union[List[List[Any]], pd.DataFrame]) -> pd.DataFrame:
        """Convierte una matriz o DataFrame existente en un DataFrame estandarizado."""
        if isinstance(raw_data, pd.DataFrame):
            df = raw_data.copy()
        elif isinstance(raw_data, list):
            if not raw_data:
                return pd.DataFrame()
            # Asegurar longitud uniforme en todas las filas
            max_cols = max((len(r) for r in raw_data if isinstance(r, (list, tuple))), default=0)
            if max_cols == 0:
                return pd.DataFrame()
            
            padded_rows = []
            for row in raw_data:
                if isinstance(row, (list, tuple)):
                    r_list = list(row)
                    if len(r_list) < max_cols:
                        r_list.extend([None] * (max_cols - len(r_list)))
                    padded_rows.append(r_list[:max_cols])
                else:
                    padded_rows.append([row] + [None] * (max_cols - 1))
            df = pd.DataFrame(padded_rows)
        else:
            raise TypeError(f"Tipo de datos no soportado: {type(raw_data)}. Debe ser list o pd.DataFrame.")
        
        # Limpieza básica celda por celda
        return df.map(cls._normalize_cell)

    @classmethod
    def auto_detect_header_row(cls, df: pd.DataFrame, max_search_rows: int = 10) -> int:
        """
        Heurística para detectar automáticamente qué fila contiene los encabezados.
        Busca la primera fila con alta densidad de texto y pocos valores nulos.
        Retorna el número de fila basado en 1 (1-indexed).
        """
        best_row = 1
        best_score = -1.0
        
        limit = min(len(df), max_search_rows)
        for idx in range(limit):
            row_vals = df.iloc[idx].dropna().tolist()
            if not row_vals:
                continue
            
            # Cantidad de celdas no nulas
            non_null_ratio = len(row_vals) / max(len(df.columns), 1)
            # Celdas que son strings y tienen longitud representativa
            str_count = sum(1 for v in row_vals if isinstance(v, str) and len(v.strip()) > 0 and not v.strip().replace('.', '', 1).isdigit())
            str_ratio = str_count / max(len(row_vals), 1)
            
            # Puntuación combinada
            score = (non_null_ratio * 0.6) + (str_ratio * 0.4)
            if score > best_score:
                best_score = score
                best_row = idx + 1  # 1-indexed

        return best_row

    @classmethod
    def clean(
        cls,
        raw_data: Union[List[List[Any]], pd.DataFrame],
        header_row: Optional[Union[int, List[int], str]] = 1,
        skip_footer: int = 0,
        drop_empty_rows: bool = True,
        drop_empty_cols: bool = True,
        auto_clean_types: bool = True,
    ) -> pd.DataFrame:
        """
        Limpia y estructura una tabla cruda.

        Parámetros:
        -----------
        raw_data : list of lists o pd.DataFrame
            Los datos crudos extraídos del PDF o Excel.
        header_row : int, list of int, 'auto' o None (por defecto 1)
            - int: Fila (1-indexed) que contiene los nombres de columnas. Todo lo que esté
                   arriba de esta fila se descarta como metadatos/basura.
            - list[int]: Por ejemplo [3, 4] si los encabezados ocupan múltiples filas continuas.
            - 'auto': Detecta automáticamente la fila más probable de encabezado.
            - None / 0: No usa encabezado, asigna nombres genéricos (Col_1, Col_2, ...).
        skip_footer : int (por defecto 0)
            Cantidad de filas finales a descartar (por ejemplo notas al pie, totales agregados, etc.).
        drop_empty_rows : bool (por defecto True)
            Si elimina filas que estén completamente vacías.
        drop_empty_cols : bool (por defecto True)
            Si elimina columnas que estén completamente vacías.
        auto_clean_types : bool (por defecto True)
            Intenta convertir números y formatos evidentes.

        Retorna:
        --------
        pd.DataFrame
            DataFrame limpio listo para ser utilizado en Python.
        """
        df_raw = cls._matrix_to_dataframe(raw_data)
        if df_raw.empty:
            return pd.DataFrame()

        # Si se solicita detección automática
        if header_row == 'auto':
            header_row = cls.auto_detect_header_row(df_raw)

        # Tratar footer si aplica
        if skip_footer > 0 and len(df_raw) > skip_footer:
            df_raw = df_raw.iloc[:-skip_footer]

        if header_row is None or header_row == 0:
            # Sin fila de encabezado: las columnas serán numéricas o genéricas
            clean_df = df_raw.copy()
            clean_df.columns = [f"Col_{i+1}" for i in range(len(clean_df.columns))]
        elif isinstance(header_row, (list, tuple)):
            # Encabezado multi-fila (e.g. [3, 4])
            # Convertir a 0-indexed
            h_indices = [h - 1 for h in header_row if 1 <= h <= len(df_raw)]
            if not h_indices:
                clean_df = df_raw.copy()
                clean_df.columns = [f"Col_{i+1}" for i in range(len(clean_df.columns))]
            else:
                header_parts = []
                for h_idx in h_indices:
                    header_parts.append(df_raw.iloc[h_idx].fillna("").astype(str).tolist())
                
                # Combinar partes del encabezado
                combined_cols = []
                for col_idx in range(len(df_raw.columns)):
                    parts = [header_parts[row_i][col_idx].strip() for row_i in range(len(header_parts))]
                    parts = [p for p in parts if p]
                    col_name = " _ ".join(parts) if parts else f"Col_{col_idx+1}"
                    combined_cols.append(col_name)
                
                last_header_idx = max(h_indices)
                clean_df = df_raw.iloc[last_header_idx + 1:].copy()
                clean_df.columns = combined_cols
        else:
            # header_row es un entero (1-indexed)
            h_idx = int(header_row) - 1
            if h_idx < 0 or h_idx >= len(df_raw):
                raise ValueError(
                    f"La fila de encabezado {header_row} está fuera de rango. La tabla tiene {len(df_raw)} filas."
                )
            
            raw_headers = df_raw.iloc[h_idx].tolist()
            # Las filas de datos empiezan después de header_row
            clean_df = df_raw.iloc[h_idx + 1:].copy()
            
            # Formatear nombres de columnas
            cols = []
            for i, h in enumerate(raw_headers):
                if pd.isna(h) or str(h).strip() == "":
                    cols.append(f"Col_{i+1}")
                else:
                    cols.append(str(h).strip())
            clean_df.columns = cols

        # Resolver nombres de columnas duplicados añadiendo sufijo
        seen_cols = {}
        unique_cols = []
        for col in clean_df.columns:
            if col in seen_cols:
                seen_cols[col] += 1
                unique_cols.append(f"{col}_{seen_cols[col]}")
            else:
                seen_cols[col] = 0
                unique_cols.append(col)
        clean_df.columns = unique_cols

        # Eliminar filas completamente vacías
        if drop_empty_rows:
            clean_df = clean_df.dropna(how='all')

        # Eliminar columnas completamente vacías
        if drop_empty_cols:
            clean_df = clean_df.dropna(axis=1, how='all')

        # Reiniciar índice
        clean_df = clean_df.reset_index(drop=True)

        # Conversión automática de tipos numéricos si aplica
        if auto_clean_types:
            clean_df = cls._try_infer_types(clean_df)

        return clean_df

    @classmethod
    def _try_infer_types(cls, df: pd.DataFrame) -> pd.DataFrame:
        """Intenta convertir columnas numéricas limpiando signos de moneda o separadores."""
        df_out = df.copy()
        for col in df_out.columns:
            series = df_out[col]
            if series.dtype == object:
                # Probar si es convertible a numérico tras remover formato común
                try:
                    # Limpiar comas/puntos comunes en español/inglés si es puro número
                    cleaned_series = series.astype(str).str.replace('$', '', regex=False).str.replace('€', '', regex=False).str.strip()
                    # Si tiene formato 1.000,00 cambiar por 1000.00
                    # o formato 1,000.00
                    sample_non_null = cleaned_series.dropna()
                    if not sample_non_null.empty:
                        # Intentar conversión directa
                        converted = pd.to_numeric(cleaned_series.str.replace(',', ''), errors='coerce')
                        if converted.notna().sum() > len(sample_non_null) * 0.7:
                            df_out[col] = converted
                except Exception:
                    pass
        return df_out


def clean_raw_table(
    raw_data: Union[List[List[Any]], pd.DataFrame],
    header_row: Optional[Union[int, List[int], str]] = 1,
    skip_footer: int = 0,
    **kwargs
) -> pd.DataFrame:
    """Función de acceso directo para TableCleaner.clean()"""
    return TableCleaner.clean(raw_data, header_row=header_row, skip_footer=skip_footer, **kwargs)
