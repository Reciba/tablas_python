# 📊 Tablas Python - Suite Integral de Procesamiento de Tablas

Sistema modular en Python diseñado para **reconocer, extraer, transformar, conciliar, validar y exportar tablas** provenientes de múltiples fuentes:
- 📄 **Documentos PDF** (usando `pdfplumber` con descarte de filas de basura iniciales)
- 📊 **Archivos Excel / CSV** (usando `xlwings`, compatible con libros **abiertos** o **cerrados**, tablas oficiales `ListObjects`, celdas de inicio y rangos)
- 🗄️ **Bases de Datos SQLite** (archivos `.db`, `.sqlite`, `.sqlite3`)

Además, incluye un conjunto completo de herramientas para **cálculos rápidos (IVA, márgenes, % participación), formateo de monedas/porcentajes, cruces tipo `BUSCARV`, conciliación automática de tablas, escritura en vivo en Excel, unión masiva de carpetas y reportes de calidad de datos**.

---

## 📁 Arquitectura del Proyecto

El código está estructurado de forma modular y desacoplada por capas:

```text
tablas_python/
│
├── main.py                     # Script principal con todas las utilidades importadas y listas para usar
├── requirements.txt            # Dependencias del proyecto
├── README.md                   # Documentación técnica completa
├── .gitignore                  # Exclusiones de Git
│
├── utils/                      # Capa de bajo nivel (lógica desacoplada y reutilizable)
│   ├── __init__.py             # Exportaciones unificadas del paquete utils
│   ├── file_utils.py           # Resolución de rutas (relativas/absolutas) y nombres seguros
│   ├── pdf_extractor.py        # Extracción multipágina en PDF con pdfplumber
│   ├── excel_extractor.py      # Extracción en Excel con xlwings (tablas oficiales, celdas y rangos)
│   ├── excel_writer.py         # Escritura en Excel en vivo o en segundo plano con xlwings
│   ├── sqlite_extractor.py     # Extracción y consultas SQL en SQLite
│   ├── table_cleaner.py        # Limpieza, descarte de basura superior, tipos y normalización
│   ├── batch_processor.py      # Unión masiva de carpetas con múltiples archivos a 1 DataFrame
│   ├── validator.py            # Validación de esquemas, diagnóstico de calidad y duplicados
│   ├── exporter.py             # Exportación individual y por lotes a CSV/Excel (utf-8-sig / sep=';')
│   └── data_helpers.py         # Cálculos: BUSCARV, conciliación, IVA, %, totales, celdas y formato
│
├── helpers/                    # Capa de fachada y presentación
│   ├── __init__.py             # Exportaciones unificadas del paquete helpers
│   ├── table_manager.py        # TableManager unificado, obtener_tabla y exportar_archivo_a_csv
│   └── display_helper.py       # Visualizador enriquecido en consola con Rich
│
├── samples/                    # Archivos y generador de prueba
│   ├── generate_samples.py     # Script generador de datos de prueba
│   ├── ejemplo_facturas.pdf    # PDF de muestra con tablas y encabezados desplazados
│   ├── ejemplo_inventario.xlsx # Excel con tablas oficiales, celdas específicas y múltiples hojas
│   └── ejemplo_empresa.db      # Base SQLite de muestra con tablas 'clientes', 'ventas' y vistas
│
└── exports/                    # Carpeta por defecto para salidas CSV y Excel
```

---

## 🚀 Instalación

Clona el repositorio e instala las dependencias:

```bash
git clone https://github.com/Reciba/tablas_python.git
cd tablas_python
pip install -r requirements.txt
```

---

## 💡 Guía de Uso Completa

### 1. Extracción desde PDF (`pdfplumber`)

Resuelve el problema donde las filas 1, 2 o 3 contienen títulos, membretes o metadatos irrelevantes y los encabezados reales comienzan más abajo (por ejemplo en la **fila 4**):

```python
from helpers.table_manager import obtener_tabla, TableManager

# Modo 1: En una sola línea indicando tabla 3 y fila 4 como encabezado:
df = obtener_tabla("samples/ejemplo_facturas.pdf", tabla=3, fila_encabezado=4)
print(df.head())

# Modo 2: Con inspección previa usando TableManager:
manager = TableManager("samples/ejemplo_facturas.pdf")
manager.resumen()             # Muestra cuántas tablas hay y en qué páginas
manager.ver_crudo(tabla=1)    # Muestra las primeras filas numeradas (Fila 1, Fila 2, Fila 3...)
df_limpio = manager.get_df(tabla=1, fila_encabezado=4, skip_footer=1)
```

---

### 2. Extracción desde Excel con `xlwings`

Soporta conectarse a archivos **abiertos en pantalla** (sin conflictos de bloqueo) o **cerrados en disco** (en segundo plano):

```python
from helpers.table_manager import obtener_tabla

# A) Por Nombre Oficial de Tabla de Excel (ListObject / Tabla con Formato):
df_stock = obtener_tabla("samples/ejemplo_inventario.xlsx", tabla="TablaStock")

# B) Por Celda de Inicio donde parte la tabla (ej. celda C4 en la hoja 'Despacho'):
df_despacho = obtener_tabla(
    "samples/ejemplo_inventario.xlsx",
    celda_inicio="C4",
    hoja="Despacho"
)

# C) Por Rango Exacto:
df_rango = obtener_tabla(
    "samples/ejemplo_inventario.xlsx",
    rango="C4:F8",
    hoja="Despacho"
)

# D) Control de archivo abierto/cerrado:
# archivo_abierto=None (autodetecta), True (fuerza conexión al Excel abierto), False (abre oculto)
df = obtener_tabla("reporte.xlsx", tabla=1, fila_encabezado=4, archivo_abierto=True)
```

---

### 3. Extracción y Consultas en SQLite (`.db`, `.sqlite`)

```python
from helpers.table_manager import TableManager, obtener_tabla

# A) Obtener tabla completa por nombre:
df_clientes = obtener_tabla("samples/ejemplo_empresa.db", tabla="clientes")

# B) Ejecutar consultas SQL personalizadas directamente a DataFrame:
db = TableManager("samples/ejemplo_empresa.db")
db.resumen() # Lista tablas y vistas

df_ventas = db.query("""
    SELECT c.nombre AS cliente, c.ciudad, v.monto_neto, v.fecha
    FROM clientes c
    JOIN ventas v ON c.id_cliente = v.id_cliente
    WHERE v.estado = 'Pagado'
""")
```

---

### 4. Búsqueda de Celdas por Nombre de Fila y Nombre de Columna

Puedes acceder a cualquier celda específica mediante nombres sin depender de posiciones numéricas:

```python
from utils.data_helpers import obtener_celda, modificar_celda

# A) Usando la función helper obtener_celda:
precio = obtener_celda(df, fila="PROD-101", columna="Precio Unitario")
print("Precio PROD-101:", precio)

# B) Usando set_index nativo de pandas con .at o .loc:
df_idx = df.set_index("Codigo")
precio = df_idx.at["PROD-101", "Precio Unitario"]
cantidad = df_idx.at["PROD-101", "Cantidad"]
total = float(precio) * float(cantidad)

# C) Modificar una celda por su nombre:
df_actualizado = modificar_celda(df, fila="PROD-101", columna="Precio Unitario", nuevo_valor=49.90)
```

---

### 5. Cruces de Datos con `buscar_v` (BUSCARV / VLOOKUP en 1 línea)

Cruza dos DataFrames asociando datos a partir de una clave en común:

```python
from utils.data_helpers import buscar_v

# Trae el nombre del cliente desde df_clientes a df_ventas usando 'id_cliente':
df_ventas["Nombre_Cliente"] = buscar_v(
    df_origen=df_ventas,
    df_destino=df_clientes,
    clave="id_cliente",
    columna_a_traer="nombre"
)
```

---

### 6. Conciliación y Auditoría entre 2 Tablas (`conciliar_tablas`)

Compara dos tablas (por ejemplo: extracto bancario vs. registro contable, o inventario físico vs. teórico):

```python
from utils.data_helpers import conciliar_tablas

resultado = conciliar_tablas(df_sistema, df_banco, clave="id_transaccion")

print("Coincidentes:", resultado["coincidentes"]) # Idénticas en ambas tablas
print("Diferencias:",  resultado["diferencias"])   # Existen en ambas pero con valores distintos
print("Solo en A:",    resultado["solo_en_A"])     # Registros que faltan en B
print("Solo en B:",    resultado["solo_en_B"])     # Registros que faltan en A
```

---

### 7. Escritura en Vivo en Excel con `escribir_en_excel`

Pega un DataFrame en una celda exacta de una plantilla Excel que tengas **abierta en pantalla** o **cerrada en disco**, conservando fórmulas y formatos:

```python
from utils.excel_writer import escribir_en_excel

escribir_en_excel(
    df=df_reporte,
    archivo_excel="plantilla_ventas.xlsx",
    hoja="Resumen",
    celda_inicio="B5",
    archivo_abierto=None # Detecta automáticamente si la ventana de Excel está abierta
)
```

---

### 8. Consolidación Masiva de Carpetas (`unir_archivos_carpeta`)

Lee y une decenas de archivos periódicos (facturas PDF, Excels mensuales, etc.) en un solo DataFrame maestro:

```python
from utils.batch_processor import unir_archivos_carpeta

df_consolidado = unir_archivos_carpeta(
    carpeta="facturas_2026/",
    extension="pdf",
    tabla=1,
    fila_encabezado=4
)
# Agrega automáticamente la columna 'Archivo_Origen' con el nombre de cada archivo
```

---

### 9. Calidad y Validación de Datos (`validar_dataframe` y `reporte_calidad`)

```python
from utils.validator import validar_dataframe, reporte_calidad, detectar_duplicados

# 1. Diagnóstico completo de nulos, completitud y tipos por columna:
print(reporte_calidad(df))

# 2. Validación de reglas de negocio antes de procesar:
check = validar_dataframe(
    df,
    columnas_requeridas=["id_venta", "monto_neto", "fecha"],
    no_nulos=["id_venta", "monto_neto"],
    tipos_esperados={"monto_neto": "numeric"}
)

if not check["es_valido"]:
    print("❌ Errores encontrados:", check["errores"])

# 3. Detección de registros duplicados:
duplicados = detectar_duplicados(df, columnas_clave=["id_venta"])
```

---

### 10. Cálculos Rápidos y Formato para Reportes

```python
from utils.data_helpers import (
    aplicar_impuesto,
    calcular_participacion,
    calcular_variacion,
    agregar_fila_totales,
    formatear_dataframe,
    formato_moneda,
    formato_porcentaje,
    limpiar_numero
)

# A) Calcular IVA (19%) y Total Bruto:
df = aplicar_impuesto(df, col_neto="monto_neto", tasa=0.19)

# B) Calcular % de participación sobre el total:
df = calcular_participacion(df, columna_valor="Total Bruto")

# C) Agregar fila final con totales:
df = agregar_fila_totales(df, columnas_sumar=["monto_neto", "IVA (19%)", "Total Bruto"])

# D) Formatear números a moneda ($) y porcentaje (%):
df_formateado = formatear_dataframe(df, {
    "monto_neto": "moneda",
    "IVA (19%)": "moneda",
    "Total Bruto": "moneda",
    "% Participación": "porcentaje"
})
```

---

### 11. Exportación Fácil a CSV y Excel

```python
from utils.exporter import guardar_csv, guardar_excel
from helpers.table_manager import exportar_archivo_a_csv

# A) Guardar DataFrame individual a CSV (optimizado con ';' y utf-8-sig para Excel):
guardar_csv(df, "exports/mi_tabla.csv")

# B) Guardar a Excel:
guardar_excel(df, "exports/mi_tabla.xlsx")

# C) Exportar TODAS las tablas encontradas en un archivo a CSVs independientes:
archivos = exportar_archivo_a_csv("samples/ejemplo_facturas.pdf", carpeta_salida="exports/pdf_tablas", fila_encabezado=4)
```

---

## 🖥️ Ejecución por Consola

```bash
# 1. Ejecutar script principal de demostración:
python main.py

# 2. Modo interactivo en terminal (asistente guiado):
python main.py -i

# 3. Exportar todas las tablas de cualquier archivo por línea de comandos:
python main.py -f "samples/ejemplo_facturas.pdf" -r 4 --export-all "exports/salida_csv"
```

---

## 📦 Tecnologías Utilizadas

- **[pandas](https://pandas.pydata.org/)**: Manipulación y estructuras tabulares en DataFrames.
- **[xlwings](https://docs.xlwings.org/)**: Integración avanzada con Microsoft Excel (archivos abiertos/cerrados, rangos y tablas).
- **[pdfplumber](https://github.com/jsvine/pdfplumber)**: Extracción precisa de texto y tablas multipágina en PDFs.
- **[openpyxl](https://openpyxl.readthedocs.io/)**: Soporte nativo y fallback para archivos `.xlsx`.
- **[reportlab](https://www.reportlab.com/)**: Generación de PDFs de prueba.
- **[rich](https://github.com/Textualize/rich)**: Visualización y formato de tablas en consola.

---

## 📄 Licencia

Distribuido bajo licencia MIT. Consulta `LICENSE` para más información.
