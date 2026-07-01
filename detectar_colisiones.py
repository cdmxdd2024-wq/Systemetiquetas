"""
detectar_colisiones.py
----------------------
Compara 4 archivos de Excel para detectar colisiones de etiquetado:
códigos que están asignados a modelos DIFERENTES entre los archivos.

Uso:
    python detectar_colisiones.py

Requisitos:
    pip install pandas openpyxl
"""

import os
import sys
import pandas as pd


# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURACIÓN — ajusta aquí los nombres de archivos y columnas
# ─────────────────────────────────────────────────────────────────────────────

ARCHIVOS_EXCEL = [
    r"c:\Users\Aureleano\Desktop\BASE ELEVENTA ODAKI 1.xlsx",
    r"c:\Users\Aureleano\Desktop\BASE ELEVENTA ODAKI 2.xlsx",
    r"c:\Users\Aureleano\Desktop\BASE ELEVENTA ODAKI 3.xlsx",
    r"c:\Users\Aureleano\Desktop\BASE ELEVENTA ODAKI 4.xlsx",
]

COLUMNA_CODIGO = "Código"    # Nombre exacto de la columna de código en los Excel
COLUMNA_MODELO = "Producto"  # Nombre exacto de la columna de producto/modelo en los Excel

ARCHIVO_REPORTE = "reporte_colisiones_etiquetado.xlsx"

# ─────────────────────────────────────────────────────────────────────────────


def cargar_archivo(ruta: str, col_codigo: str, col_modelo: str) -> pd.DataFrame | None:
    """
    Carga un archivo Excel, valida que tenga las columnas requeridas,
    aplica limpieza de espacios y retorna solo las columnas relevantes.

    Retorna None si el archivo no existe o tiene un problema estructural.
    """
    if not os.path.isfile(ruta):
        print(f"  [ADVERTENCIA] Archivo no encontrado: '{ruta}' — se omitirá.")
        return None

    try:
        df = pd.read_excel(ruta, dtype=str)
    except Exception as e:
        print(f"  [ERROR] No se pudo leer '{ruta}': {e}")
        return None

    columnas_faltantes = [c for c in (col_codigo, col_modelo) if c not in df.columns]
    if columnas_faltantes:
        print(
            f"  [ERROR] El archivo '{ruta}' no tiene las columnas: "
            f"{columnas_faltantes}. Columnas disponibles: {list(df.columns)}"
        )
        return None

    df = df[[col_codigo, col_modelo]].copy()

    # Limpieza de espacios al inicio y al final para evitar falsos positivos
    df[col_codigo] = df[col_codigo].str.strip()
    df[col_modelo] = df[col_modelo].str.strip()

    # Eliminar filas donde código o modelo sean vacíos tras la limpieza
    df.dropna(subset=[col_codigo, col_modelo], inplace=True)
    df = df[df[col_codigo] != ""]
    df = df[df[col_modelo] != ""]

    print(f"  [OK] '{ruta}' - {len(df)} filas validas cargadas.")
    return df


def detectar_colisiones(df_total: pd.DataFrame, col_codigo: str, col_modelo: str) -> pd.DataFrame:
    """
    Agrupa por código y detecta aquellos asociados a más de un modelo único.
    Retorna un DataFrame con las colisiones encontradas.
    """
    # Para cada código, recopilar el conjunto de modelos únicos
    agrupado = (
        df_total
        .groupby(col_codigo)[col_modelo]
        .apply(lambda modelos: sorted(set(modelos)))
        .reset_index()
    )
    agrupado.columns = ["Codigo", "Modelos_unicos"]

    # Una colision ocurre cuando hay mas de un modelo diferente para el mismo codigo
    colisiones = agrupado[agrupado["Modelos_unicos"].apply(len) > 1].copy()

    # Convertir la lista de modelos en una cadena legible para el reporte
    colisiones["Modelos_en_conflicto"] = colisiones["Modelos_unicos"].apply(
        lambda lst: " | ".join(lst)
    )
    colisiones["Cantidad_de_modelos"] = colisiones["Modelos_unicos"].apply(len)

    colisiones = colisiones[["Codigo", "Cantidad_de_modelos", "Modelos_en_conflicto"]]
    colisiones.sort_values("Codigo", inplace=True)
    colisiones.reset_index(drop=True, inplace=True)

    return colisiones


def mostrar_reporte_en_terminal(colisiones: pd.DataFrame) -> None:
    """Imprime el reporte de colisiones en la terminal de forma legible."""
    separador = "-" * 70
    print(f"\n{separador}")
    print(f"  REPORTE DE COLISIONES DE ETIQUETADO")
    print(f"{separador}")

    if colisiones.empty:
        print("  No se encontraron colisiones. Todos los códigos son consistentes.")
    else:
        print(f"  Se encontraron {len(colisiones)} codigo(s) en conflicto:\n")
        for _, fila in colisiones.iterrows():
            print(f"  Codigo          : {fila['Codigo']}")
            print(f"  Productos unicos: {fila['Cantidad_de_modelos']}")
            print(f"  En conflicto    : {fila['Modelos_en_conflicto']}")
            print()

    print(separador)


def guardar_reporte_excel(colisiones: pd.DataFrame, ruta_salida: str) -> None:
    """Exporta el reporte de colisiones a un archivo Excel con formato básico."""
    try:
        with pd.ExcelWriter(ruta_salida, engine="openpyxl") as writer:
            colisiones.to_excel(writer, index=False, sheet_name="Colisiones")

            # Ajuste automático del ancho de columnas
            hoja = writer.sheets["Colisiones"]
            for col_cells in hoja.columns:
                max_len = max(
                    len(str(celda.value)) if celda.value is not None else 0
                    for celda in col_cells
                )
                hoja.column_dimensions[col_cells[0].column_letter].width = min(max_len + 4, 80)

        print(f"\n  [OK] Reporte guardado en: '{ruta_salida}'")
    except Exception as e:
        print(f"\n  [ERROR] No se pudo guardar el reporte: {e}")


def main() -> None:
    print("=" * 70)
    print("  DETECTOR DE COLISIONES DE ETIQUETADO - Inicio")
    print("=" * 70)
    print(f"\nColumna de codigo : '{COLUMNA_CODIGO}'")
    print(f"Columna de modelo : '{COLUMNA_MODELO}'")
    print(f"\nCargando {len(ARCHIVOS_EXCEL)} archivos...\n")
    sys.stdout.reconfigure(encoding="utf-8", errors="replace") if hasattr(sys.stdout, "reconfigure") else None

    dataframes = []
    for archivo in ARCHIVOS_EXCEL:
        df = cargar_archivo(archivo, COLUMNA_CODIGO, COLUMNA_MODELO)
        if df is not None:
            df["_origen"] = archivo  # columna auxiliar para trazabilidad (no se usa en reporte)
            dataframes.append(df)

    if not dataframes:
        print("\n[FATAL] No se pudo cargar ningún archivo. Verifica la configuración.")
        sys.exit(1)

    if len(dataframes) < len(ARCHIVOS_EXCEL):
        print(
            f"\n[ADVERTENCIA] Solo se cargaron {len(dataframes)} de "
            f"{len(ARCHIVOS_EXCEL)} archivos. El analisis continua con los disponibles."
        )

    # Unir todos los datos en un único DataFrame
    df_total = pd.concat(dataframes, ignore_index=True)

    # Descartar duplicados exactos (mismo código Y mismo modelo) antes de analizar
    df_total.drop_duplicates(subset=[COLUMNA_CODIGO, COLUMNA_MODELO], inplace=True)

    print(f"\nTotal de pares (codigo, modelo) unicos para analizar: {len(df_total)}")

    # Detectar colisiones
    colisiones = detectar_colisiones(df_total, COLUMNA_CODIGO, COLUMNA_MODELO)

    # Mostrar en terminal
    mostrar_reporte_en_terminal(colisiones)

    # Guardar reporte Excel
    guardar_reporte_excel(colisiones, ARCHIVO_REPORTE)

    print("\n  Proceso completado.\n")


if __name__ == "__main__":
    main()
