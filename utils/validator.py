"""
Módulo para validación y auditoría de calidad de datos en DataFrames.
Permite verificar esquemas requeridos, detectar valores nulos, registros duplicados
y generar reportes de integridad antes de procesar o exportar.
"""

from typing import List, Dict, Any, Optional, Union
import pandas as pd
import numpy as np


class DataValidator:
    """
    Validador de calidad e integridad de datos para DataFrames.
    """

    @staticmethod
    def validar(
        df: pd.DataFrame,
        columnas_requeridas: Optional[List[str]] = None,
        no_nulos: Optional[List[str]] = None,
        tipos_esperados: Optional[Dict[str, str]] = None,
        min_filas: int = 1
    ) -> Dict[str, Any]:
        """
        Valida que un DataFrame cumpla con una serie de reglas de negocio.

        Parámetros:
        -----------
        df : pd.DataFrame
            El DataFrame a validar.
        columnas_requeridas : list of str, opcional
            Lista de nombres de columnas que DEBEN existir.
        no_nulos : list of str, opcional
            Lista de columnas que no pueden tener valores vacíos/nulos.
        tipos_esperados : dict, opcional
            Diccionario {'columna': 'numeric'|'string'|'datetime'} para verificar tipos.
        min_filas : int (por defecto 1)
            Cantidad mínima de filas esperadas.

        Retorna:
        --------
        dict con:
            - 'es_valido': bool (True si pasó todas las pruebas)
            - 'errores': list[str] (detalles de cada fallo)
            - 'alertas': list[str] (advertencias menores)
        """
        errores = []
        alertas = []

        if df is None:
            return {"es_valido": False, "errores": ["El DataFrame es None."], "alertas": []}

        # 1. Cantidad de filas
        if len(df) < min_filas:
            errores.append(f"El DataFrame tiene {len(df)} filas (se esperaban al menos {min_filas}).")

        # 2. Columnas requeridas
        if columnas_requeridas:
            columnas_actuales = set(df.columns)
            faltantes = [col for col in columnas_requeridas if col not in columnas_actuales]
            if faltantes:
                errores.append(f"Faltan columnas requeridas: {faltantes}")

        # 3. Columnas sin nulos
        if no_nulos:
            for col in no_nulos:
                if col in df.columns:
                    n_nulos = df[col].isna().sum()
                    if n_nulos > 0:
                        errores.append(f"La columna '{col}' tiene {n_nulos} valores nulos/vacíos.")

        # 4. Tipos de datos esperados
        if tipos_esperados:
            for col, tipo in tipos_esperados.items():
                if col in df.columns:
                    if tipo == 'numeric' and not pd.api.types.is_numeric_dtype(df[col]):
                        errores.append(f"La columna '{col}' no es de tipo numérico (tipo actual: {df[col].dtype}).")
                    elif tipo == 'datetime' and not pd.api.types.is_datetime64_any_dtype(df[col]):
                        errores.append(f"La columna '{col}' no es de tipo fecha/datetime.")

        es_valido = len(errores) == 0
        return {
            "es_valido": es_valido,
            "errores": errores,
            "alertas": alertas,
            "total_filas": len(df),
            "total_columnas": len(df.columns)
        }

    @staticmethod
    def reporte_calidad(df: pd.DataFrame) -> pd.DataFrame:
        """
        Genera una tabla de diagnóstico de calidad de datos con conteo de nulos,
        porcentaje de completitud, valores únicos y tipos.
        """
        if df.empty:
            return pd.DataFrame(columns=["Columna", "Tipo", "No_Nulos", "Nulos", "% Nulos", "Valores_Unicos", "Muestra"])

        total_filas = len(df)
        reporte = []

        for col in df.columns:
            serie = df[col]
            n_nulos = serie.isna().sum()
            pct_nulos = round((n_nulos / total_filas) * 100, 1)
            n_unicos = serie.nunique(dropna=True)
            tipo = str(serie.dtype)
            
            # Muestra del primer valor no nulo
            muestra_val = serie.dropna().iloc[0] if not serie.dropna().empty else None

            reporte.append({
                "Columna": col,
                "Tipo": tipo,
                "No_Nulos": total_filas - n_nulos,
                "Nulos": n_nulos,
                "% Nulos": f"{pct_nulos}%",
                "Valores_Unicos": n_unicos,
                "Muestra": str(muestra_val)[:30] if muestra_val is not None else ""
            })

        return pd.DataFrame(reporte)

    @staticmethod
    def detectar_duplicados(df: pd.DataFrame, columnas_clave: Optional[List[str]] = None) -> pd.DataFrame:
        """
        Extrae las filas duplicadas del DataFrame para auditoría.
        """
        if columnas_clave:
            duplicados = df[df.duplicated(subset=columnas_clave, keep=False)]
        else:
            duplicados = df[df.duplicated(keep=False)]
        return duplicados.sort_values(by=columnas_clave) if columnas_clave else duplicados


def validar_dataframe(
    df: pd.DataFrame,
    columnas_requeridas: Optional[List[str]] = None,
    no_nulos: Optional[List[str]] = None,
    tipos_esperados: Optional[Dict[str, str]] = None,
    min_filas: int = 1
) -> Dict[str, Any]:
    """Acceso directo a DataValidator.validar()"""
    return DataValidator.validar(df, columnas_requeridas, no_nulos, tipos_esperados, min_filas)


def reporte_calidad(df: pd.DataFrame) -> pd.DataFrame:
    """Acceso directo a DataValidator.reporte_calidad()"""
    return DataValidator.reporte_calidad(df)


def detectar_duplicados(df: pd.DataFrame, columnas_clave: Optional[List[str]] = None) -> pd.DataFrame:
    """Acceso directo a DataValidator.detectar_duplicados()"""
    return DataValidator.detectar_duplicados(df, columnas_clave)
