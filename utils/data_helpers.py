"""
Módulo de utilidades y helpers para cálculos rápidos, limpieza y formateo de datos en pandas.
Permite ahorrar código en main.py para formateo de monedas, porcentajes, agregación de totales,
cálculo de variaciones, márgenes e impuestos.
"""

from typing import Any, Dict, List, Optional, Union
import re
import pandas as pd
import numpy as np


# ==========================================
# 1. PARSEO Y LIMPIEZA DE NÚMEROS Y FECHAS
# ==========================================

def limpiar_numero(val: Any, default: Any = np.nan) -> Union[float, int, Any]:
    """
    Convierte cualquier valor numérico, moneda o porcentaje a float o int de forma precisa.
    Maneja formatos latinoamericanos (1.234,56 / 0,056), anglosajones (1,234.56 / 0.056),
    símbolos ($ / € / USD / CLP / %), y negativos entre paréntesis (100) -> -100.
    
    Ejemplos:
        limpiar_numero("0,056")       -> 0.056
        limpiar_numero("9,999")       -> 9.999
        limpiar_numero("89,949")      -> 89.949
        limpiar_numero("0,056%")      -> 0.056
        limpiar_numero("$ 1.250.000") -> 1250000.0
        limpiar_numero("(450,50)")    -> -450.50
    """
    if pd.isna(val) or val is None:
        return default
    if isinstance(val, (int, float, np.number)):
        return float(val) if isinstance(val, float) else val

    s = str(val).strip()
    if not s:
        return default

    # 1. Detectar negativos con paréntesis: (123.45) o (123,45)
    es_negativo = False
    if s.startswith("(") and s.endswith(")"):
        es_negativo = True
        s = s[1:-1].strip()
    elif s.startswith("-"):
        es_negativo = True
        s = s[1:].strip()

    # 2. Remover símbolos de monedas y texto innecesario
    s = re.sub(r'[\$\€\£\¥\s]|USD|CLP|EUR|UF', '', s, flags=re.IGNORECASE)
    # Remover %
    s = s.replace('%', '').strip()

    if not s:
        return default

    # 3. Determinar separadores decimales y de miles
    # Caso A: Tiene comas y puntos (ej: 1.234.567,89 o 1,234,567.89)
    if '.' in s and ',' in s:
        if s.rfind(',') > s.rfind('.'):
            # Formato latino: 1.234.567,89 -> quitar puntos y reemplazar coma por punto
            s = s.replace('.', '').replace(',', '.')
        else:
            # Formato anglo: 1,234,567.89 -> quitar comas
            s = s.replace(',', '')

    # Caso B: Solo tiene coma(s)
    elif ',' in s:
        num_comas = s.count(',')
        if num_comas == 1:
            # 1 sola coma: SIEMPRE es separador decimal en español (ej: 0,056 | 9,999 | 89,949 | 123,45)
            s = s.replace(',', '.')
        else:
            # Múltiples comas: separador de miles anglosajón (ej: 1,000,000)
            s = s.replace(',', '')

    # Caso C: Solo tiene punto(s)
    elif '.' in s:
        num_puntos = s.count('.')
        if num_puntos > 1:
            # Múltiples puntos: separador de miles latino (ej: 1.000.000)
            s = s.replace('.', '')

    try:
        num = float(s)
        if es_negativo:
            num = -num
        return num
    except ValueError:
        return default


def limpiar_columnas_numericas(df: pd.DataFrame, columnas: Union[str, List[str]]) -> pd.DataFrame:
    """
    Convierte una o más columnas a tipo numérico (float) limpiando caracteres extra.
    """
    df_out = df.copy()
    if isinstance(columnas, str):
        columnas = [columnas]
    for col in columnas:
        if col in df_out.columns:
            df_out[col] = df_out[col].apply(limpiar_numero)
            df_out[col] = pd.to_numeric(df_out[col], errors='coerce')
    return df_out


def normalizar_fechas(df: pd.DataFrame, columnas: Union[str, List[str]], formato_salida: str = "%Y-%m-%d") -> pd.DataFrame:
    """
    Estandariza columnas de fechas al formato deseado (por defecto YYYY-MM-DD).
    """
    df_out = df.copy()
    if isinstance(columnas, str):
        columnas = [columnas]
    for col in columnas:
        if col in df_out.columns:
            fechas = pd.to_datetime(df_out[col], errors='coerce', dayfirst=True)
            df_out[col] = fechas.dt.strftime(formato_salida).fillna(df_out[col])
    return df_out


# ==========================================
# 2. FORMATEO DE MONEDAS, PORCENTAJES Y MILES
# ==========================================

def formato_moneda(val: Any, simbolo: str = "$", decimales: int = 0, separador_miles: str = ".") -> str:
    """
    Formatea un número a representación de moneda.
    Ejemplo:
        formato_moneda(1500000) -> "$ 1.500.000"
        formato_moneda(1250.75, decimales=2) -> "$ 1.250,75"
    """
    num = limpiar_numero(val)
    if pd.isna(num) or not isinstance(num, (int, float)):
        return "" if pd.isna(val) else str(val)

    if decimales == 0:
        texto = f"{round(num):,}".replace(",", separador_miles)
    else:
        fmt = f"{{:,.{decimales}f}}"
        partes = fmt.format(num).split(".")
        entero = partes[0].replace(",", separador_miles)
        dec = partes[1]
        sep_dec = "," if separador_miles == "." else "."
        texto = f"{entero}{sep_dec}{dec}"

    return f"{simbolo} {texto}" if simbolo else texto


def formato_clp(val: Any, simbolo: str = "$") -> str:
    """
    Formatea un número o texto a Pesos Chilenos (CLP) con separador de miles y sin decimales.
    
    Ejemplos:
        formato_clp(1500000)              -> "$ 1.500.000"
        formato_clp("1500000")            -> "$ 1.500.000"
        formato_clp(250000, simbolo="CLP") -> "CLP 250.000"
    """
    return formato_moneda(val, simbolo=simbolo, decimales=0, separador_miles=".")


def formato_porcentaje(
    val: Any,
    decimales: Optional[int] = None,
    multiplicar_por_100: bool = False,
    separador_decimal: str = ","
) -> str:
    """
    Formatea un valor a porcentaje de manera exacta.
    
    A diferencia de versiones anteriores, NO multiplica arbitrariamente por 100
    a menos que se indique explícitamente con `multiplicar_por_100=True`.
    Por lo tanto, 0.056 se muestra como 0,056% y 5.56 se muestra como 5,56%.
    
    Ejemplos:
        formato_porcentaje(5.56)                     -> "5,56%"
        formato_porcentaje(0.056)                    -> "0,056%"
        formato_porcentaje("0.056%")                 -> "0,056%"
        formato_porcentaje(0.056, multiplicar_por_100=True)  -> "5,6%"
        formato_porcentaje(15.42, decimales=1)       -> "15,4%"
    """
    if pd.isna(val) or val is None or str(val).strip() == "":
        return ""

    s_orig = str(val).strip()
    tenia_simbolo_pct = "%" in s_orig

    num = limpiar_numero(val)
    if pd.isna(num) or not isinstance(num, (int, float)):
        return s_orig

    # Solo multiplicar por 100 si el usuario lo pide explícitamente y el valor NO traía ya el '%'
    if multiplicar_por_100 and not tenia_simbolo_pct:
        num = num * 100

    if decimales is not None:
        fmt_str = f"{{:.{decimales}f}}"
        txt_num = fmt_str.format(num)
    else:
        # Conservar los decimales significativos del número sin truncar
        if isinstance(num, int) or (isinstance(num, float) and num.is_integer()):
            txt_num = str(int(num))
        else:
            txt_num = f"{num:.6f}".rstrip('0').rstrip('.')

    if separador_decimal == ",":
        txt_num = txt_num.replace(".", ",")
    else:
        txt_num = txt_num.replace(",", ".")

    return f"{txt_num}%"


def formato_miles(val: Any, decimales: int = 0) -> str:
    """Formatea un número con separador de miles."""
    return formato_moneda(val, simbolo="", decimales=decimales)


def formatear_dataframe(
    df: pd.DataFrame,
    reglas: Dict[str, str]
) -> pd.DataFrame:
    """
    Aplica formatos a múltiples columnas de un DataFrame para presentación o exportación.
    
    Ejemplo:
        df_fmt = formatear_dataframe(df, {
            'Ventas': 'moneda',
            'Precio': 'moneda_2dec',
            'Margen': 'porcentaje',
            'Participacion': 'porcentaje_1dec',
            'Ratio': 'ratio_pct',
            'Cantidad': 'miles'
        })
    """
    df_out = df.copy()
    for col, tipo in reglas.items():
        if col not in df_out.columns:
            continue
        if tipo in ['moneda', 'clp', 'moneda_clp']:
            df_out[col] = df_out[col].apply(lambda x: formato_clp(x))
        elif tipo == 'moneda_2dec':
            df_out[col] = df_out[col].apply(lambda x: formato_moneda(x, decimales=2))
        elif tipo == 'usd':
            df_out[col] = df_out[col].apply(lambda x: formato_moneda(x, simbolo="USD $", decimales=2, separador_miles=","))
        elif tipo == 'uf':
            df_out[col] = df_out[col].apply(lambda x: formato_moneda(x, simbolo="UF", decimales=2, separador_miles="."))
        elif tipo == 'porcentaje':
            df_out[col] = df_out[col].apply(lambda x: formato_porcentaje(x))
        elif tipo == 'porcentaje_1dec':
            df_out[col] = df_out[col].apply(lambda x: formato_porcentaje(x, decimales=1))
        elif tipo == 'porcentaje_2dec':
            df_out[col] = df_out[col].apply(lambda x: formato_porcentaje(x, decimales=2))
        elif tipo == 'ratio_pct':
            df_out[col] = df_out[col].apply(lambda x: formato_porcentaje(x, multiplicar_por_100=True, decimales=1))
        elif tipo == 'miles':
            df_out[col] = df_out[col].apply(formato_miles)
    return df_out


# ==========================================
# 3. CÁLCULOS RÁPIDOS Y RESÚMENES EN DATAFRAME
# ==========================================

def agregar_fila_totales(
    df: pd.DataFrame,
    columnas_sumar: Optional[List[str]] = None,
    columnas_promedio: Optional[List[str]] = None,
    etiqueta: str = "TOTAL",
    columna_etiqueta: Optional[str] = None
) -> pd.DataFrame:
    """
    Agrega una fila final con la suma o promedio de las columnas especificadas.
    
    Ejemplo:
        df_con_total = agregar_fila_totales(df, columnas_sumar=['Subtotal', 'Cantidad'])
    """
    df_out = df.copy()
    if df_out.empty:
        return df_out

    # Determinar columnas numéricas si no se especificaron
    if columnas_sumar is None and columnas_promedio is None:
        columnas_sumar = [c for c in df_out.columns if pd.api.types.is_numeric_dtype(df_out[c])]

    fila_total = {}
    for col in df_out.columns:
        if columnas_sumar and col in columnas_sumar:
            fila_total[col] = df_out[col].sum()
        elif columnas_promedio and col in columnas_promedio:
            fila_total[col] = df_out[col].mean()
        else:
            fila_total[col] = ""

    # Asignar etiqueta 'TOTAL' en la primera columna o la indicada
    target_label_col = columna_etiqueta if columna_etiqueta else df_out.columns[0]
    fila_total[target_label_col] = etiqueta

    df_total = pd.DataFrame([fila_total])
    return pd.concat([df_out, df_total], ignore_index=True)


def calcular_participacion(
    df: pd.DataFrame,
    columna_valor: str,
    nombre_col: str = "% Participación",
    decimales: int = 1,
    mostrar_feedback: bool = True
) -> pd.DataFrame:
    """
    Calcula el porcentaje de participación de cada fila sobre el total de la columna.
    """
    df_out = df.copy()
    if df_out.empty:
        return df_out
    if columna_valor not in df_out.columns:
        if mostrar_feedback:
            print(f"⚠️ [calcular_participacion] La columna '{columna_valor}' no existe. Columnas disponibles: {list(df_out.columns)}")
        return df_out
    total = df_out[columna_valor].sum()
    if total != 0 and pd.notna(total):
        df_out[nombre_col] = ((df_out[columna_valor] / total) * 100).round(decimales)
    else:
        df_out[nombre_col] = 0.0
    return df_out


def calcular_variacion(
    df: pd.DataFrame,
    col_actual: str,
    col_anterior: str,
    nombre_col: str = "% Variación",
    decimales: int = 1,
    mostrar_feedback: bool = True
) -> pd.DataFrame:
    """
    Calcula la variación porcentual entre dos columnas: ((actual - anterior) / anterior) * 100.
    """
    df_out = df.copy()
    if df_out.empty:
        return df_out
    faltantes = [c for c in [col_actual, col_anterior] if c not in df_out.columns]
    if faltantes:
        if mostrar_feedback:
            print(f"⚠️ [calcular_variacion] Columnas no encontradas: {faltantes}. Columnas disponibles: {list(df_out.columns)}")
        return df_out
    ant = df_out[col_anterior]
    act = df_out[col_actual]
    var = np.where(ant != 0, ((act - ant) / ant.abs()) * 100, 0.0)
    df_out[nombre_col] = np.round(var, decimales)
    return df_out


def aplicar_impuesto(
    df: pd.DataFrame,
    col_neto: str,
    tasa: float = 0.19,
    col_iva: str = "IVA (19%)",
    col_total: str = "Total Bruto",
    mostrar_feedback: bool = True
) -> pd.DataFrame:
    """
    Calcula el IVA (o impuesto) y el valor bruto a partir de una columna neta.
    """
    df_out = df.copy()
    if df_out.empty:
        return df_out
    if col_neto not in df_out.columns:
        if mostrar_feedback:
            print(f"⚠️ [aplicar_impuesto] La columna '{col_neto}' no existe en el DataFrame. Columnas disponibles: {list(df_out.columns)}")
        return df_out
    df_out[col_iva] = (df_out[col_neto] * tasa).round(2)
    df_out[col_total] = (df_out[col_neto] + df_out[col_iva]).round(2)
    return df_out


def agrupar_y_resumir(
    df: pd.DataFrame,
    por: Union[str, List[str]],
    metricas: Dict[str, Union[str, List[str]]],
    mostrar_feedback: bool = True
) -> pd.DataFrame:
    """
    Agrupa un DataFrame y calcula sumas, promedios o conteos en una sola línea.
    
    Ejemplo:
        resumen = agrupar_y_resumir(df, por='Categoria', metricas={'Subtotal': 'sum', 'Cantidad': 'sum'})
    """
    if df.empty:
        return pd.DataFrame()
    cols_por = [por] if isinstance(por, str) else list(por)
    cols_faltantes = [c for c in cols_por if c not in df.columns]
    if cols_faltantes:
        if mostrar_feedback:
            print(f"⚠️ [agrupar_y_resumir] Columnas de agrupación no encontradas: {cols_faltantes}. Columnas disponibles: {list(df.columns)}")
        return df
    metricas_validas = {k: v for k, v in metricas.items() if k in df.columns}
    if not metricas_validas:
        if mostrar_feedback:
            print(f"⚠️ [agrupar_y_resumir] Ninguna de las columnas de métricas {list(metricas.keys())} existe en el DataFrame.")
        return df
    agrupado = df.groupby(por).agg(metricas_validas).reset_index()
    return agrupado


def obtener_celda(
    df: pd.DataFrame,
    fila: Any,
    columna: str,
    columna_identificador: Optional[str] = None,
    default: Any = None,
    lanzar_error: bool = False,
    mostrar_feedback: bool = True
) -> Any:
    """
    Obtiene el valor de una celda puntual indicando el nombre/etiqueta de la fila y el nombre de la columna.
    Si no se encuentra, muestra un mensaje de feedback amigable y devuelve un valor por defecto (None)
    en lugar de interrumpir el programa con un error.

    Parámetros:
    -----------
    df : pd.DataFrame
        El DataFrame a consultar.
    fila : Any (str, int)
        El nombre o valor que identifica la fila (ej: 'PROD-101', 'Distribuidora Los Andes', 'Enero').
    columna : str
        El nombre de la columna deseada (ej: 'Precio Unitario', 'monto_neto').
    columna_identificador : str, opcional
        Columna donde buscar el identificador. Si es None, busca automáticamente.
    default : Any (por defecto None)
        Valor a retornar si no se encuentra la fila o columna.
    lanzar_error : bool (por defecto False)
        Si es True, lanza KeyError en lugar de retornar el valor default.
    mostrar_feedback : bool (por defecto True)
        Si es True, imprime una advertencia informativa con sugerencias cuando no encuentra el dato.

    Ejemplos:
    ---------
    # Si existe:
    precio = obtener_celda(df, fila="PROD-101", columna="Precio Unitario")

    # Si NO existe, muestra feedback y devuelve None (sin dar error):
    precio = obtener_celda(df, fila="PROD-999", columna="Precio Unitario")
    # 👉 ⚠️ [obtener_celda] No se encontró la fila 'PROD-999' en el DataFrame. Retornando None.
    """
    if df is None or df.empty:
        if mostrar_feedback:
            print(f"⚠️ [obtener_celda] El DataFrame está vacío o es None. Retornando {default}.")
        if lanzar_error:
            raise ValueError("El DataFrame está vacío o es None.")
        return default

    # Verificar columna
    if columna not in df.columns and columna != df.index.name:
        if mostrar_feedback:
            print(f"⚠️ [obtener_celda] La columna '{columna}' no existe. Columnas disponibles: {list(df.columns)}. Retornando {default}.")
        if lanzar_error:
            raise KeyError(f"La columna '{columna}' no existe en el DataFrame. Columnas disponibles: {list(df.columns)}")
        return default

    # 1. Si la fila coincide directamente con el índice de pandas
    if fila in df.index:
        return df.at[fila, columna]

    # 2. Si se especificó una columna identificadora
    if columna_identificador and columna_identificador in df.columns:
        coincidencias = df[df[columna_identificador] == fila]
        if not coincidencias.empty:
            return coincidencias.iloc[0][columna]
        if mostrar_feedback:
            print(f"⚠️ [obtener_celda] No se encontró ninguna fila con {columna_identificador}='{fila}'. Retornando {default}.")
        if lanzar_error:
            raise KeyError(f"No se encontró ninguna fila con {columna_identificador}='{fila}'")
        return default

    # 3. Búsqueda automática en la primera columna o cualquier columna
    for col in df.columns:
        coincidencias = df[df[col].astype(str) == str(fila)]
        if not coincidencias.empty:
            return coincidencias.iloc[0][columna]

    if mostrar_feedback:
        print(f"⚠️ [obtener_celda] No se encontró la fila identificada con '{fila}' en el DataFrame. Retornando {default}.")
    if lanzar_error:
        raise KeyError(f"No se encontró la fila identificada con '{fila}' en el DataFrame.")
    
    return default


def modificar_celda(
    df: pd.DataFrame,
    fila: Any,
    columna: str,
    nuevo_valor: Any,
    columna_identificador: Optional[str] = None,
    lanzar_error: bool = False,
    mostrar_feedback: bool = True
) -> pd.DataFrame:
    """
    Modifica el valor de una celda puntual buscando por nombre de fila y nombre de columna.
    Si no se encuentra, muestra un mensaje de feedback y devuelve el DataFrame sin modificaciones.
    """
    df_out = df.copy()
    if df_out.empty:
        if mostrar_feedback:
            print("⚠️ [modificar_celda] El DataFrame está vacío.")
        if lanzar_error:
            raise ValueError("El DataFrame está vacío.")
        return df_out

    if columna not in df_out.columns:
        if mostrar_feedback:
            print(f"⚠️ [modificar_celda] La columna '{columna}' no existe. Columnas disponibles: {list(df_out.columns)}")
        if lanzar_error:
            raise KeyError(f"La columna '{columna}' no existe.")
        return df_out

    if fila in df_out.index:
        df_out.at[fila, columna] = nuevo_valor
        return df_out

    target_col = columna_identificador or df_out.columns[0]
    idx = df_out.index[df_out[target_col].astype(str) == str(fila)].tolist()
    if idx:
        df_out.at[idx[0], columna] = nuevo_valor
        return df_out

    if mostrar_feedback:
        print(f"⚠️ [modificar_celda] No se encontró la fila '{fila}' para modificar. El DataFrame no fue alterado.")
    if lanzar_error:
        raise KeyError(f"No se encontró la fila '{fila}' para modificar.")

    return df_out


def buscar_v(
    df_origen: pd.DataFrame,
    df_destino: pd.DataFrame,
    clave: str,
    columna_a_traer: str,
    clave_destino: Optional[str] = None,
    nombre_columna: Optional[str] = None,
    default: Any = np.nan,
    lanzar_error: bool = False,
    mostrar_feedback: bool = True
) -> Union[pd.Series, pd.DataFrame]:
    """
    Equivalente al BUSCARV / VLOOKUP de Excel para cruzar dos tablas en una sola línea.
    """
    target_key = clave_destino or clave
    errores = []
    if clave not in df_origen.columns:
        errores.append(f"La clave '{clave}' no existe en df_origen (disponibles: {list(df_origen.columns)})")
    if target_key not in df_destino.columns:
        errores.append(f"La clave '{target_key}' no existe en df_destino (disponibles: {list(df_destino.columns)})")
    if columna_a_traer not in df_destino.columns:
        errores.append(f"La columna a traer '{columna_a_traer}' no existe en df_destino (disponibles: {list(df_destino.columns)})")

    if errores:
        if mostrar_feedback:
            print(f"⚠️ [buscar_v] Error de configuración: {'; '.join(errores)}")
        if lanzar_error:
            raise KeyError("; ".join(errores))
        serie_vacia = pd.Series([default] * len(df_origen), index=df_origen.index)
        if nombre_columna:
            df_out = df_origen.copy()
            df_out[nombre_columna] = serie_vacia
            return df_out
        return serie_vacia

    # Mapeo rápido usando diccionario para máxima velocidad
    mapeo = df_destino.drop_duplicates(subset=[target_key]).set_index(target_key)[columna_a_traer].to_dict()
    serie_resultado = df_origen[clave].map(mapeo).fillna(default)

    if nombre_columna:
        df_out = df_origen.copy()
        df_out[nombre_columna] = serie_resultado
        return df_out

    return serie_resultado


def conciliar_tablas(
    df_a: pd.DataFrame,
    df_b: pd.DataFrame,
    clave: str,
    columnas_comparar: Optional[List[str]] = None,
    sufijo_a: str = "_A",
    sufijo_b: str = "_B",
    lanzar_error: bool = False,
    mostrar_feedback: bool = True
) -> Dict[str, pd.DataFrame]:
    """
    Concilia y audita dos tablas (ej: sistema vs extracto bancario, o inventario teórico vs físico).

    Retorna un diccionario con 4 DataFrames:
    - 'coincidentes': Filas idénticas en clave y valores comparados.
    - 'diferencias': Filas que existen en ambas tablas pero con valores distintos.
    - 'solo_en_A': Filas que solo existen en la primera tabla.
    - 'solo_en_B': Filas que solo existen en la segunda tabla.
    """
    if clave not in df_a.columns or clave not in df_b.columns:
        msg = f"La columna clave '{clave}' debe existir en ambas tablas. (df_a: {list(df_a.columns)}, df_b: {list(df_b.columns)})"
        if mostrar_feedback:
            print(f"⚠️ [conciliar_tablas] {msg}")
        if lanzar_error:
            raise KeyError(msg)
        return {
            "coincidentes": pd.DataFrame(),
            "diferencias": pd.DataFrame(),
            "solo_en_A": df_a.copy() if df_a is not None else pd.DataFrame(),
            "solo_en_B": df_b.copy() if df_b is not None else pd.DataFrame(),
        }

    # Unir ambas tablas con outer join
    merged = pd.merge(df_a, df_b, on=clave, how='outer', suffixes=(sufijo_a, sufijo_b), indicator=True)

    solo_en_a = merged[merged['_merge'] == 'left_only'].drop(columns=['_merge']).reset_index(drop=True)
    solo_en_b = merged[merged['_merge'] == 'right_only'].drop(columns=['_merge']).reset_index(drop=True)
    ambos = merged[merged['_merge'] == 'both'].drop(columns=['_merge']).reset_index(drop=True)

    if not columnas_comparar:
        # Detectar columnas comunes con sufijo
        columnas_comparar = [c for c in df_a.columns if c != clave and c in df_b.columns]

    if not columnas_comparar or ambos.empty:
        return {
            "coincidentes": ambos,
            "diferencias": pd.DataFrame(),
            "solo_en_A": solo_en_a,
            "solo_en_B": solo_en_b,
        }

    # Verificar discrepancias en las columnas a comparar
    mascara_coincide = pd.Series(True, index=ambos.index)
    for col in columnas_comparar:
        col_a = f"{col}{sufijo_a}" if f"{col}{sufijo_a}" in ambos.columns else col
        col_b = f"{col}{sufijo_b}" if f"{col}{sufijo_b}" in ambos.columns else col
        
        # Comparar numérico o texto
        coincide_col = (ambos[col_a] == ambos[col_b]) | (ambos[col_a].isna() & ambos[col_b].isna())
        mascara_coincide = mascara_coincide & coincide_col

    coincidentes = ambos[mascara_coincide].reset_index(drop=True)
    diferencias = ambos[~mascara_coincide].reset_index(drop=True)

    return {
        "coincidentes": coincidentes,
        "diferencias": diferencias,
        "solo_en_A": solo_en_a,
        "solo_en_B": solo_en_b,
    }


