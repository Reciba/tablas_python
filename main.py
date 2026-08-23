"""
Script Principal Simple
========================
Demuestra cómo extraer tablas desde PDF, Excel y SQLite,
multiplicar celdas/columnas específicas y mostrar los resultados en pantalla.
Incluye todo el kit de herramientas base importado para su uso inmediato.
"""

# ==============================================================================
# 🧰 IMPORTACIÓN DEL KIT BASE COMPLETO (Disponible para usar en cualquier momento)
# ==============================================================================

# 1. Gestores principales de tablas y extracción
from helpers.table_manager import (
    TableManager,           # Clase orquestadora completa (PDF, Excel xlwings, SQLite)
    obtener_tabla,          # Función rápida de 1 línea para obtener un DataFrame
    inspeccionar_archivo,   # Muestra resumen y vista previa de todas las filas crudas
    exportar_archivo_a_csv, # Exporta todas las tablas de un archivo a CSVs independientes
)
from helpers.display_helper import DisplayHelper # Visualizador con tablas estilizadas en consola

# 2. Utilidades de exportación
from utils.exporter import (
    guardar_csv,            # Guarda un DataFrame en CSV (con sep=';' y utf-8-sig para Excel)
    guardar_excel,          # Guarda un DataFrame en Excel (.xlsx)
    TableExporter,          # Exportador avanzado
)

# 3. Cruces de datos y cálculos
from utils.data_helpers import (
    buscar_v,               # Equivalente al BUSCARV / VLOOKUP de Excel en 1 línea
    conciliar_tablas,       # Concilia 2 tablas y detecta diferencias, faltantes y coincidencias
    aplicar_impuesto,       # Calcula IVA o impuesto y columna Total Bruto
    calcular_participacion, # Calcula el % de participación de cada fila sobre el total
    calcular_variacion,     # Calcula el % de crecimiento o variación entre dos columnas
    agregar_fila_totales,   # Agrega fila 'TOTAL' al final con sumas o promedios
    agrupar_y_resumir,      # Agrupación y agregaciones rápidas (sum, mean, count)
)

# 4. Formateo, celdas y limpieza
from utils.data_helpers import (
    limpiar_numero,             # Convierte strings como '$ 1.250.000', '18,5%' o '(100)' a float
    limpiar_columnas_numericas, # Limpia y convierte columnas completas a numéricas
    normalizar_fechas,          # Estandariza fechas a formato YYYY-MM-DD
    formato_moneda,             # Formatea números a '$ 1.500.000'
    formato_porcentaje,         # Formatea números a '15.4%'
    formato_miles,              # Formatea números con separador de miles
    formatear_dataframe,        # Aplica formatos de moneda/porcentaje a un DataFrame completo
    obtener_celda,              # Obtiene una celda por nombre de fila y columna
    modificar_celda,            # Modifica una celda por nombre de fila y columna
)

# 5. Automatización avanzada: Escritura en vivo en Excel, unión de carpetas y validación
from utils.excel_writer import escribir_en_excel   # Escribe un DataFrame en una celda de Excel (en vivo o cerrado)
from utils.batch_processor import unir_archivos_carpeta # Une todos los archivos de una carpeta en 1 DataFrame
from utils.validator import (
    validar_dataframe,  # Valida esquema de columnas, tipos y nulos
    reporte_calidad,    # Diagnóstico completo de nulos y tipos
    detectar_duplicados # Extrae registros duplicados para auditoría
)


# ==============================================================================
# 🚀 LÓGICA PRINCIPAL DEL SCRIPT
# ==============================================================================

def main():
    print("\n" + "=" * 60)
    print("🚀 EXTRACCIÓN Y MULTIPLICACIÓN DE CELDAS (PDF, EXCEL Y SQLITE)")
    print("=" * 60)

    # ----------------------------------------------------
    # 1. PROCESAR ARCHIVO PDF (pdfplumber)
    # ----------------------------------------------------
    print("\n📄 [1] ARCHIVO PDF: 'samples/ejemplo_facturas.pdf'")
    # Extraer Tabla 1 indicando que el encabezado está en la Fila 4 (descarta filas 1 a 3)
    df_pdf = obtener_tabla("samples/ejemplo_facturas.pdf", tabla=1, fila_encabezado=4, skip_footer=1)

    # Multiplicación puntual de dos celdas de la primera fila:
    cant_0 = df_pdf.at[0, "Cantidad"]
    precio_0 = df_pdf.at[0, "Precio Unitario"]
    resultado_celda_pdf = float(cant_0) * float(precio_0)

    # Multiplicar las dos columnas completas:
    df_pdf["Total_Calculado"] = df_pdf["Cantidad"].astype(float) * df_pdf["Precio Unitario"].astype(float)

    print(f"👉 Multiplicación puntual de la Fila 1: Cantidad ({cant_0}) x Precio ({precio_0}) = {formato_moneda(resultado_celda_pdf)}")
    print("\nDataFrame PDF con la columna calculada:")
    print(df_pdf[["Codigo", "Descripcion Producto", "Cantidad", "Precio Unitario", "Total_Calculado"]])
    guardar_csv(df_pdf, "exports/pdf_multiplicado.csv")

    # ----------------------------------------------------
    # 2. PROCESAR ARCHIVO EXCEL (xlwings)
    # ----------------------------------------------------
    print("\n" + "-" * 60)
    print("📊 [2] ARCHIVO EXCEL: 'samples/ejemplo_inventario.xlsx'")
    # Extraer la tabla oficial 'TablaStock' (o por celda_inicio="C4")
    df_excel = obtener_tabla("samples/ejemplo_inventario.xlsx", tabla="TablaStock")

    # Multiplicación puntual de dos celdas de la primera fila:
    stock_0 = df_excel.at[0, "Stock"]
    costo_0 = df_excel.at[0, "Costo_Unitario"]
    resultado_celda_excel = float(stock_0) * float(costo_0)

    # Multiplicar las dos columnas completas:
    df_excel["Valor_Total_Calculado"] = df_excel["Stock"].astype(float) * df_excel["Costo_Unitario"].astype(float)

    print(f"👉 Multiplicación puntual de la Fila 1: Stock ({stock_0}) x Costo ({costo_0}) = {formato_moneda(resultado_celda_excel)}")
    print("\nDataFrame Excel con la columna calculada:")
    print(df_excel[["ID_Item", "Nombre_Articulo", "Stock", "Costo_Unitario", "Valor_Total_Calculado"]])
    guardar_csv(df_excel, "exports/excel_multiplicado.csv")

    # ----------------------------------------------------
    # 3. PROCESAR BASE DE DATOS SQLITE
    # ----------------------------------------------------
    print("\n" + "-" * 60)
    print("🗄️ [3] ARCHIVO SQLITE: 'samples/ejemplo_empresa.db'")
    # Extraer tabla 'ventas'
    df_sqlite = obtener_tabla("samples/ejemplo_empresa.db", tabla="ventas")

    # Multiplicación puntual (ej: aplicar IVA 19% al monto neto de la Fila 1):
    monto_0 = df_sqlite.at[0, "monto_neto"]
    iva_0 = float(monto_0) * 0.19

    # Multiplicar columna por tasa para calcular IVA:
    df_sqlite["IVA_19%"] = df_sqlite["monto_neto"].astype(float) * 0.19
    df_sqlite["Total_Bruto"] = df_sqlite["monto_neto"].astype(float) + df_sqlite["IVA_19%"]

    print(f"👉 Multiplicación puntual de la Fila 1: Monto ({monto_0}) x 0.19 (IVA) = {formato_moneda(iva_0)}")
    print("\nDataFrame SQLite con el cálculo aplicado:")
    print(df_sqlite[["id_venta", "fecha", "monto_neto", "IVA_19%", "Total_Bruto"]])
    guardar_csv(df_sqlite, "exports/sqlite_multiplicado.csv")

    print("\n" + "=" * 60)
    print("✅ Todos los cálculos realizados y exportados a la carpeta 'exports/'.")
    print("=" * 60 + "\n")


if __name__ == "__main__":
    main()
