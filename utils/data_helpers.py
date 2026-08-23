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

def limpiar_numero(val: Any) -> Union[float, int, Any]:
    """
    Convierte cualquier cadena con formato numérico, moneda o porcentaje a float o int.
    Maneja formatos latinoamericanos (1.234,56) y anglosajones (1,234.56),
    símbolos ($ / € / USD / CLP / %), y negativos entre paréntesis (100) -> -100.
    
    Ejemplos:
        limpiar_numero("$ 1.250.000") -> 1250000.0
        limpiar_numero("18,5 %")      -> 18.5
        limpiar_numero("(450.50)")    -> -450.50
    """
    if pd.isna(val) or val is None:
        return np.nan
    if isinstance(val, (int, float, np.number)):
        return val

    s = str(val).strip()
    if not s:
        return np.nan

    # Detectar negativos con paréntesis: (123.45)
    es_negativo = False
    if s.startswith("(") and s.endswith(")"):
        es_negativo = True
        s = s[1:-1].strip()
    elif s.startswith("-"):
        es_negativo = True
        s = s[1:].strip()

    # Remover símbolos de monedas y palabras comunes
    s = re.sub(r'[\$\€\£\¥\s]|USD|CLP|EUR|UF', '', s, flags=re.IGNORECASE)
    # Remover %
    s = s.replace('%', '').strip()

    if not s:
        return np.nan

    # Determinar si el separador decimal es coma o punto
    # Caso 1: Tiene comas y puntos (ej: 1.234.567,89 o 1,234,567.89)
    if '.' in s and ',' in s:
        if s.rfind(',') > s.rfind('.'):
            # Formato latino: 1.234,56 -> quitar puntos y coma a punto
            s = s.replace('.', '').replace(',', '.')
        else:
            # Formato anglo: 1,234.56 -> quitar comas
            s = s.replace(',', '')
    elif ',' in s:
        # Solo tiene comas. Si tiene 1 coma y max 2 decimales al final: 1234,56
        partes = s.split(',')
        if len(partes) == 2 and len(partes[1]) <= 2:
            s = s.replace(',', '.')
        else:
            # Es separador de miles: 1,000,000
            s = s.replace(',', '')
    elif '.' in s:
        # Solo tiene puntos. Si tiene múltiples puntos: 1.000.000
        if s.count('.') > 1:
            s = s.replace('.', '')
        # Si tiene 1 punto y exactamente 3 dígitos después, puede ser miles (ej: 1.500)
        # pero en float estándar suele ser decimal. Lo dejamos a float directo.

    try:
        num = float(s)
        if es_negativo:
            num = -num
        # Si no tiene decimales reales, retornar float o int
        return num
    except ValueError:
        return val


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


def formato_porcentaje(val: Any, decimales: int = 1) -> str:
    """
    Formatea un valor a porcentaje.
    Ejemplo:
        formato_porcentaje(15.42) -> "15.4%"
        formato_porcentaje(0.1542, es_ratio=True)
    """
    num = limpiar_numero(val)
    if pd.isna(num) or not isinstance(num, (int, float)):
        return "" if pd.isna(val) else str(val)
    
    # Si el valor está entre 0 y 1 (ratio), convertir a base 100
    if -1.0 <= num <= 1.0 and num != 0:
        num = num * 100

    return f"{num:.{decimales}f}%"


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
            'Cantidad': 'miles'
        })
    """
    df_out = df.copy()
    for col, tipo in reglas.items():
        if col not in df_out.columns:
            continue
        if tipo == 'moneda':
            df_out[col] = df_out[col].apply(lambda x: formato_moneda(x, decimales=0))
        elif tipo == 'moneda_2dec':
            df_out[col] = df_out[col].apply(lambda x: formato_moneda(x, decimales=2))
        elif tipo == 'porcentaje':
            df_out[col] = df_out[col].apply(formato_porcentaje)
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
    decimales: int = 1
) -> pd.DataFrame:
    """
    Calcula el porcentaje de participación de cada fila sobre el total de la columna.
    """
    df_out = df.copy()
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
    decimales: int = 1
) -> pd.DataFrame:
    """
    Calcula la variación porcentual entre dos columnas: ((actual - anterior) / anterior) * 100.
    """
    df_out = df.copy()
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
    col_total: str = "Total Bruto"
) -> pd.DataFrame:
    """
    Calcula el IVA (o impuesto) y el valor bruto a partir de una columna neta.
    """
    df_out = df.copy()
    df_out[col_iva] = (df_out[col_neto] * tasa).round(2)
    df_out[col_total] = (df_out[col_neto] + df_out[col_iva]).round(2)
    return df_out


def agrupar_y_resumir(
    df: pd.DataFrame,
    por: Union[str, List[str]],
    metricas: Dict[str, Union[str, List[str]]]
) -> pd.DataFrame:
    """
    Agrupa un DataFrame y calcula sumas, promedios o conteos en una sola línea.
    
    Ejemplo:
        resumen = agrupar_y_resumir(df, por='Categoria', metricas={'Subtotal': 'sum', 'Cantidad': 'sum'})
    """
    agrupado = df.groupby(por).agg(metricas).reset_index()
    return agrupado


def obtener_celda(
    df: pd.DataFrame,
    fila: Any,
    columna: str,
    columna_identificador: Optional[str] = None
) -> Any:
    """
    Obtiene el valor de una celda puntual indicando el nombre/etiqueta de la fila y el nombre de la columna.

    Parámetros:
    -----------
    df : pd.DataFrame
        El DataFrame a consultar.
    fila : Any (str, int)
        El nombre o valor que identifica la fila (ej: 'PROD-101', 'Distribuidora Los Andes', 'Enero').
    columna : str
        El nombre de la columna deseada (ej: 'Precio Unitario', 'monto_neto').
    columna_identificador : str, opcional
        Si el DataFrame no tiene como índice los nombres de filas, especifica qué columna
        contiene el nombre buscado (ej: 'Codigo' o 'Producto'). Si es None, busca automáticamente.

    Ejemplos:
    ---------
    # Caso 1: Con índice asignado
    precio = obtener_celda(df_con_indice, fila="PROD-101", columna="Precio Unitario")

    # Caso 2: Sin cambiar índice (busca en la columna 'Codigo')
    precio = obtener_celda(df, fila="PROD-101", columna="Precio Unitario", columna_identificador="Codigo")
    """
    if columna not in df.columns and columna != df.index.name:
        raise KeyError(f"La columna '{columna}' no existe en el DataFrame. Columnas disponibles: {list(df.columns)}")

    # 1. Si la fila coincide directamente con el índice de pandas
    if fila in df.index:
        return df.at[fila, columna]

    # 2. Si se especificó una columna identificadora
    if columna_identificador and columna_identificador in df.columns:
        coincidencias = df[df[columna_identificador] == fila]
        if not coincidencias.empty:
            return coincidencias.iloc[0][columna]
        raise KeyError(f"No se encontró ninguna fila con {columna_identificador}='{fila}'")

    # 3. Búsqueda automática en la primera columna o cualquier columna de texto
    for col in df.columns:
        coincidencias = df[df[col].astype(str) == str(fila)]
        if not coincidencias.empty:
            return coincidencias.iloc[0][columna]

    raise KeyError(f"No se encontró la fila identificada con '{fila}' en el DataFrame.")


def modificar_celda(
    df: pd.DataFrame,
    fila: Any,
    columna: str,
    nuevo_valor: Any,
    columna_identificador: Optional[str] = None
) -> pd.DataFrame:
    """
    Modifica el valor de una celda puntual buscando por nombre de fila y nombre de columna.
    """
    df_out = df.copy()
    if fila in df_out.index:
        df_out.at[fila, columna] = nuevo_valor
        return df_out

    target_col = columna_identificador or df_out.columns[0]
    idx = df_out.index[df_out[target_col].astype(str) == str(fila)].tolist()
    if idx:
        df_out.at[idx[0], columna] = nuevo_valor
        return df_out

    raise KeyError(f"No se encontró la fila '{fila}' para modificar.")


def buscar_v(
    df_origen: pd.DataFrame,
    df_destino: pd.DataFrame,
    clave: str,
    columna_a_traer: str,
    clave_destino: Optional[str] = None,
    nombre_columna: Optional[str] = None,
    default: Any = np.nan
) -> Union[pd.Series, pd.DataFrame]:
    """
    Equivalente al BUSCARV / VLOOKUP de Excel para cruzar dos tablas en una sola línea.

    Parámetros:
    -----------
    df_origen : pd.DataFrame
        El DataFrame donde quieres insertar el nuevo valor (ej: df_ventas).
    df_destino : pd.DataFrame
        El DataFrame que contiene la tabla maestra con el dato buscado (ej: df_clientes).
    clave : str
        Nombre de la columna común en df_origen (ej: 'id_cliente').
    columna_a_traer : str
        Nombre de la columna que deseas extraer de df_destino (ej: 'nombre').
    clave_destino : str, opcional
        Nombre de la columna clave en df_destino si se llama distinto a 'clave'.
    nombre_columna : str, opcional
        Si se especifica, agrega la columna a df_origen y devuelve el DataFrame completo.
        Si es None, devuelve una pd.Series lista para asignar.
    default : Any (por defecto np.nan)
        Valor a colocar si no se encuentra coincidencia.

    Ejemplos:
    ---------
    # Forma 1: Asignación directa a una nueva columna
    df_ventas["Nombre_Cliente"] = buscar_v(df_ventas, df_clientes, clave="id_cliente", columna_a_traer="nombre")

    # Forma 2: Retornar DataFrame actualizado
    df_resultado = buscar_v(df_ventas, df_clientes, clave="id_cliente", columna_a_traer="nombre", nombre_columna="Cliente")
    """
    target_key = clave_destino or clave
    
    if clave not in df_origen.columns:
        raise KeyError(f"La clave '{clave}' no existe en df_origen.")
    if target_key not in df_destino.columns:
        raise KeyError(f"La clave '{target_key}' no existe en df_destino.")
    if columna_a_traer not in df_destino.columns:
        raise KeyError(f"La columna '{columna_a_traer}' no existe en df_destino.")

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
    sufijo_b: str = "_B"
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
        raise KeyError(f"La columna clave '{clave}' debe existir en ambas tablas.")

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


