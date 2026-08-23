"""
Módulo para escribir y actualizar datos en archivos Excel usando xlwings.
Permite pegar DataFrames directamente en celdas específicas (ej: 'B5') de libros
abiertos en pantalla (en vivo) o cerrados en segundo plano, sin dañar formatos ni fórmulas existentes.
"""

from typing import Optional, Union, Any
import os
import pandas as pd

try:
    import xlwings as xw
    HAS_XLWINGS = True
except ImportError:
    HAS_XLWINGS = False

try:
    import openpyxl
    from openpyxl.utils.cell import coordinate_to_tuple
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

from .file_utils import resolve_file_path


def _is_workbook_open_in_excel(file_path: str, file_name: str) -> bool:
    """Comprueba si el libro ya está abierto en Excel activo."""
    if not HAS_XLWINGS:
        return False
    try:
        for book in xw.books:
            if book.name.lower() == file_name.lower() or book.fullname.lower() == file_path.lower():
                return True
    except Exception:
        return False
    return False


def escribir_en_excel(
    df: pd.DataFrame,
    archivo_excel: str,
    hoja: Union[str, int] = 1,
    celda_inicio: str = "A1",
    incluir_encabezados: bool = True,
    incluir_indice: bool = False,
    guardar: bool = True,
    archivo_abierto: Optional[bool] = None,
    crear_si_no_existe: bool = True
) -> str:
    """
    Escribe un DataFrame directamente en una celda específica de un libro de Excel.

    Parámetros:
    -----------
    df : pd.DataFrame
        El DataFrame con los datos a escribir.
    archivo_excel : str
        Ruta o nombre del archivo Excel.
    hoja : str o int (por defecto 1)
        Nombre o número de la hoja donde se escribirán los datos.
    celda_inicio : str (por defecto 'A1')
        Coordenada de la celda superior izquierda donde comenzará a escribirse (ej: 'B5', 'C4').
    incluir_encabezados : bool (por defecto True)
        Si escribe los nombres de las columnas en la primera fila.
    incluir_indice : bool (por defecto False)
        Si incluye la columna de índice de pandas.
    guardar : bool (por defecto True)
        Si guarda el archivo tras escribir.
    archivo_abierto : bool o None
        - True: Se conecta a la ventana activa de Excel.
        - False: Trabaja en segundo plano cerrado.
        - None: Detecta automáticamente si está abierto o cerrado.
    crear_si_no_existe : bool (por defecto True)
        Crea un nuevo archivo Excel si no existe en la ruta dada.

    Retorna:
    --------
    str
        Ruta absoluta del archivo modificado.
    """
    try:
        resolved_path = resolve_file_path(archivo_excel)
    except Exception:
        resolved_path = os.path.abspath(archivo_excel)

    if not os.path.exists(resolved_path):
        if crear_si_no_existe:
            os.makedirs(os.path.dirname(resolved_path), exist_ok=True)
            # Crear libro vacío con openpyxl o pandas
            df_init = pd.DataFrame()
            with pd.ExcelWriter(resolved_path, engine='openpyxl') as writer:
                sheet_title = hoja if isinstance(hoja, str) else "Hoja1"
                df_init.to_excel(writer, sheet_name=sheet_title)
        else:
            raise FileNotFoundError(f"No existe el archivo Excel: {resolved_path}")

    file_name = os.path.basename(resolved_path)

    # Intentar con xlwings si está disponible
    if HAS_XLWINGS:
        try:
            is_open = _is_workbook_open_in_excel(resolved_path, file_name)
            should_connect = archivo_abierto is True or (archivo_abierto is None and is_open)

            app = None
            book = None
            needs_close = False

            try:
                if should_connect:
                    try:
                        book = xw.books[file_name]
                    except Exception:
                        book = xw.Book(resolved_path)
                else:
                    app = xw.App(visible=False, add_book=False)
                    app.display_alerts = False
                    app.screen_updating = False
                    book = app.books.open(resolved_path)
                    needs_close = True

                # Seleccionar o crear hoja
                sheet_names = [s.name for s in book.sheets]
                if isinstance(hoja, int):
                    if 1 <= hoja <= len(book.sheets):
                        sht = book.sheets[hoja - 1]
                    else:
                        sht = book.sheets.add(f"Hoja{hoja}")
                else:
                    if str(hoja) in sheet_names:
                        sht = book.sheets[str(hoja)]
                    else:
                        sht = book.sheets.add(str(hoja))

                # Escribir DataFrame en la celda indicada
                clean_cell = celda_inicio.replace("$", "").upper()
                sht.range(clean_cell).options(
                    index=incluir_indice,
                    header=incluir_encabezados
                ).value = df

                if guardar:
                    book.save()

                return resolved_path

            finally:
                if needs_close:
                    if book:
                        try: book.close()
                        except Exception: pass
                    if app:
                        try: app.quit()
                        except Exception: pass

        except Exception as e:
            # Fallback a openpyxl si ocurre algún error COM
            pass

    # Fallback con openpyxl si xlwings no pudo ejecutarse
    if HAS_OPENPYXL:
        wb = openpyxl.load_workbook(resolved_path)
        sheet_title = hoja if isinstance(hoja, str) else (f"Hoja{hoja}" if isinstance(hoja, int) and hoja > len(wb.sheetnames) else wb.sheetnames[hoja-1])
        if sheet_title in wb.sheetnames:
            ws = wb[sheet_title]
        else:
            ws = wb.create_sheet(title=sheet_title)

        start_row, start_col = coordinate_to_tuple(celda_inicio.replace("$", "").upper())

        current_r = start_row
        if incluir_encabezados:
            for c_idx, col_name in enumerate(df.columns, start=start_col):
                ws.cell(row=current_r, column=c_idx, value=str(col_name))
            current_r += 1

        for _, row in df.iterrows():
            for c_idx, val in enumerate(row.values, start=start_col):
                cell_val = None if pd.isna(val) else val
                ws.cell(row=current_r, column=c_idx, value=cell_val)
            current_r += 1

        if guardar:
            wb.save(resolved_path)
        return resolved_path

    raise RuntimeError("Se requiere xlwings u openpyxl para escribir en Excel.")
