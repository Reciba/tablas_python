"""
Módulo helper para visualización limpia en consola de tablas detectadas y DataFrames.
"""
from typing import List, Any, Optional, Union
import sys
import pandas as pd

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

try:
    from rich.console import Console
    from rich.table import Table
    from rich.panel import Panel
    from rich.text import Text
    HAS_RICH = True
    console = Console(highlight=False)
except ImportError:
    HAS_RICH = False
    console = None


class DisplayHelper:
    """Helper para presentar resúmenes visuales de tablas y datos crudos."""

    @staticmethod
    def print_header(title: str):
        """Imprime un título estilizado."""
        if HAS_RICH:
            console.print(f"\n[bold cyan]=== {title} ===[/bold cyan]")
        else:
            print(f"\n=== {title} ===")

    @staticmethod
    def _get_location_str(t: Any) -> str:
        """Obtiene la cadena descriptiva de ubicación (Página, Hoja, Celda o Tabla SQLite)."""
        if hasattr(t, 'table_type'):
            return f"SQLite: {t.table_name} ({t.table_type})"
        elif hasattr(t, 'page_number'):
            return f"PDF Pág. {t.page_number}"
        elif hasattr(t, 'sheet_name'):
            t_name = getattr(t, 'table_name', None)
            range_addr = getattr(t, 'range_address', '') or getattr(t, 'start_cell', '')
            if t_name and t_name != t.sheet_name:
                return f"Excel Tabla '{t_name}' ({t.sheet_name} - {range_addr})"
            elif range_addr and range_addr != "A1":
                return f"Hoja: {t.sheet_name} ({range_addr})"
            return f"Hoja: {t.sheet_name}"
        return "N/A"

    @staticmethod
    def print_tables_summary(tables_info: List[Any], file_path: str):
        """Muestra un resumen de todas las tablas encontradas en el archivo."""
        if HAS_RICH:
            table = Table(title=f"📋 Tablas detectadas en: {file_path}", show_lines=True)
            table.add_column("ID Tabla", justify="center", style="bold green")
            table.add_column("Ubicación / Nombre", justify="center", style="cyan")
            table.add_column("Dimensiones (Filas x Cols)", justify="center")
            table.add_column("Vista previa fila 1 y 2", style="dim")

            for t in tables_info:
                loc = DisplayHelper._get_location_str(t)
                dims = f"{t.num_rows} x {t.num_cols}"
                preview_rows = t.get_preview(2)
                preview_str = " | ".join(str(r) for r in preview_rows)
                if len(preview_str) > 60:
                    preview_str = preview_str[:57] + "..."
                table.add_row(f"Tabla {t.table_id}", loc, dims, preview_str)
            
            console.print(table)
        else:
            print(f"\n--- Resumen de tablas en {file_path} ---")
            for t in tables_info:
                loc = DisplayHelper._get_location_str(t)
                print(f"  [Tabla {t.table_id}] -> {loc} | Dimensiones: {t.num_rows} filas x {t.num_cols} columnas")

    @staticmethod
    def print_raw_preview(table_info: Any, max_rows: int = 8):
        """
        Muestra una vista previa de las primeras filas con sus números de fila (1-indexed).
        Esto permite al usuario identificar fácilmente cuál fila contiene el encabezado.
        """
        raw_rows = table_info.get_preview(max_rows)
        loc = DisplayHelper._get_location_str(table_info)
        
        if HAS_RICH:
            table = Table(
                title=f"🔍 Vista Cruda: Tabla {table_info.table_id} ({loc}) - Primeras {len(raw_rows)} filas",
                show_lines=True
            )
            table.add_column("Fila #", justify="center", style="bold yellow")
            
            # Determinar máximo de columnas
            max_cols = max(len(r) for r in raw_rows) if raw_rows else 0
            for c in range(max_cols):
                table.add_column(f"Col {c+1}", style="white")

            for idx, row in enumerate(raw_rows, start=1):
                cells = [str(val) if val is not None else "" for val in row]
                # Rellenar columnas faltantes si la fila es corta
                while len(cells) < max_cols:
                    cells.append("")
                table.add_row(f"Fila {idx}", *cells)

            console.print(table)
            console.print("[italic dim]Consejo: Revisa los números de 'Fila #' para elegir tu 'fila_encabezado'.[/italic dim]\n")
        else:
            print(f"\n--- Vista Cruda: Tabla {table_info.table_id} ({loc}) ---")
            for idx, row in enumerate(raw_rows, start=1):
                row_str = " | ".join(str(val) if val is not None else "[vacío]" for val in row)
                print(f"  Fila {idx:2d}: {row_str}")
            print("Consejo: Usa el número de 'Fila' para tu parámetro fila_encabezado.\n")

    @staticmethod
    def print_dataframe(
        df: pd.DataFrame,
        title: str = "DataFrame Limpio",
        max_rows: Optional[Union[int, str]] = 15
    ):
        """
        Imprime un DataFrame formateado con estilo Rich.

        Parámetros:
        -----------
        df : pd.DataFrame
            El DataFrame a mostrar.
        title : str
            Título de la tabla.
        max_rows : int, None o 'all' (por defecto 15)
            Cantidad máxima de filas a renderizar. Usa None o 'all' para mostrar todas las filas.
        """
        if df is None or df.empty:
            if HAS_RICH:
                console.print(f"[yellow]⚠️ {title}: El DataFrame está vacío o es None.[/yellow]")
            else:
                print(f"⚠️ {title}: El DataFrame está vacío o es None.")
            return

        total_filas = len(df)
        if max_rows is None or max_rows == "all" or (isinstance(max_rows, int) and max_rows <= 0):
            df_mostrar = df
            limite = total_filas
        else:
            limite = int(max_rows)
            df_mostrar = df.head(limite)

        if HAS_RICH:
            table = Table(
                title=f"✨ {title} (Mostrando {len(df_mostrar)} de {total_filas} filas x {len(df.columns)} columnas)",
                show_lines=True
            )
            for col in df.columns:
                table.add_column(str(col), style="cyan", justify="left")

            for _, row in df_mostrar.iterrows():
                table.add_row(*[str(val) if pd.notna(val) else "" for val in row.values])

            console.print(table)
            if total_filas > limite:
                console.print(f"[dim]... y {total_filas - limite} filas más. (Pasa max_rows=None para ver todas)[/dim]\n")
        else:
            print(f"\n=== {title} (Mostrando {len(df_mostrar)} de {total_filas} filas) ===")
            print(df_mostrar)
            if total_filas > limite:
                print(f"... y {total_filas - limite} filas más.\n")


def mostrar_tabla(
    df: pd.DataFrame,
    title: str = "DataFrame",
    max_rows: Optional[Union[int, str]] = 15
):
    """
    Imprime un DataFrame formateado con bordes y colores en consola usando Rich.
    
    Parámetros:
    -----------
    df : pd.DataFrame
        El DataFrame a visualizar.
    title : str
        Título de la tabla.
    max_rows : int, None o 'all' (por defecto 15)
        Número de filas a mostrar. Usa max_rows=50, max_rows=None o max_rows='all' para mostrar todas.

    Ejemplos:
    ---------
    mostrar_tabla(df, "Mis Facturas", max_rows=30)     # Muestra hasta 30 filas
    mostrar_tabla(df, "Todas las Facturas", max_rows=None) # Muestra todas las filas sin límite
    """
    DisplayHelper.print_dataframe(df, title=title, max_rows=max_rows)


