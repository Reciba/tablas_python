"""
CLI entry point for tablas-python command line tool.
"""

import sys
import argparse
from helpers.table_manager import TableManager, obtener_tabla, exportar_archivo_a_csv
from helpers.display_helper import DisplayHelper
from utils.exporter import guardar_csv, guardar_excel


def main():
    parser = argparse.ArgumentParser(
        prog="tablas",
        description="tablas-python: Extractor de tablas (PDF, Excel xlwings, SQLite) para pandas."
    )
    parser.add_argument("-f", "--file", type=str, help="Ruta al archivo PDF, Excel o SQLite")
    parser.add_argument("-t", "--table", type=str, default="1", help="Número de tabla o nombre")
    parser.add_argument("-r", "--header-row", type=str, default="1", help="Fila de encabezado (ej: 4 o 'auto')")
    parser.add_argument("-o", "--output", type=str, help="Ruta de exportación (ej: salida.csv)")
    parser.add_argument("--export-all", type=str, help="Carpeta para exportar TODAS las tablas a CSV")
    
    args = parser.parse_args()

    if not args.file:
        parser.print_help()
        sys.exit(0)

    if args.export_all:
        h_row = int(args.header_row) if args.header_row.isdigit() else args.header_row
        guardados = exportar_archivo_a_csv(args.file, carpeta_salida=args.export_all, fila_encabezado=h_row)
        print(f"✅ Se exportaron {len(guardados)} tablas a '{args.export_all}'.")
        return

    tab_id = int(args.table) if args.table.isdigit() else args.table
    h_row = int(args.header_row) if args.header_row.isdigit() else args.header_row
    df = obtener_tabla(args.file, tabla=tab_id, fila_encabezado=h_row)
    DisplayHelper.print_dataframe(df, title=f"Tabla {args.table} de {args.file}")
    
    if args.output:
        if args.output.lower().endswith(".csv"):
            guardar_csv(df, args.output)
        else:
            guardar_excel(df, args.output)
        print(f"✅ Exportado a: {args.output}")


if __name__ == "__main__":
    main()
