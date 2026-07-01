"""
app_colisiones.py
-----------------
Sistema ODAKI — dos herramientas integradas:
  1. Detector de Colisiones de Etiquetado
  2. Generador de Códigos de Barras Secuenciales

Ejecutar con:
    streamlit run app_colisiones.py
"""

import io
import os
import pandas as pd
import streamlit as st

# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURACION
# ─────────────────────────────────────────────────────────────────────────────

COLUMNA_CODIGO   = "Código"
COLUMNA_PRODUCTO = "Producto"

# ─────────────────────────────────────────────────────────────────────────────


def leer_excel(fuente, nombre: str, col_codigo: str, col_producto: str) -> pd.DataFrame | None:
    """
    Lee un Excel desde una ruta de disco o desde un objeto de archivo subido.
    Retorna None si falta alguna columna requerida.
    """
    try:
        df = pd.read_excel(fuente, dtype=str)
    except Exception as e:
        st.error(f"No se pudo leer '{nombre}': {e}")
        return None

    faltantes = [c for c in (col_codigo, col_producto) if c not in df.columns]
    if faltantes:
        st.error(f"'{nombre}' no tiene las columnas: {faltantes}. Disponibles: {list(df.columns)}")
        return None

    df[col_codigo]   = df[col_codigo].str.strip()
    df[col_producto] = df[col_producto].str.strip()
    df.dropna(subset=[col_codigo, col_producto], inplace=True)
    df = df[(df[col_codigo] != "") & (df[col_producto] != "")]
    df["_origen"] = nombre
    return df[[col_codigo, col_producto, "_origen"]]


def cargar_datos(fuentes: list, col_codigo: str, col_producto: str) -> pd.DataFrame | None:
    """
    Recibe lista de (fuente, nombre). Carga, limpia y une todos los archivos.
    """
    dfs = []
    for fuente, nombre in fuentes:
        df = leer_excel(fuente, nombre, col_codigo, col_producto)
        if df is not None:
            dfs.append(df)

    if not dfs:
        return None

    total = pd.concat(dfs, ignore_index=True)
    total.drop_duplicates(subset=[col_codigo, col_producto], inplace=True)
    return total


def detectar_compartidos(df: pd.DataFrame, col_codigo: str, col_producto: str) -> pd.DataFrame:
    """
    Detecta códigos que aparecen en más de un archivo con EXACTAMENTE el mismo
    producto — es decir, están correctamente unificados entre bases.
    """
    # Contar en cuántos archivos distintos aparece cada par (codigo, producto)
    agrupado = (
        df.groupby([col_codigo, col_producto])["_origen"]
        .apply(lambda x: sorted(set(x)))
        .reset_index()
    )
    agrupado.columns = ["Codigo", "Producto", "Archivos"]
    agrupado["Num_archivos"] = agrupado["Archivos"].apply(len)
    agrupado["Archivos_texto"] = agrupado["Archivos"].apply(lambda l: " | ".join(l))

    compartidos = agrupado[agrupado["Num_archivos"] > 1].copy()
    compartidos.sort_values(["Num_archivos", "Codigo"], ascending=[False, True], inplace=True)
    compartidos.reset_index(drop=True, inplace=True)
    return compartidos


def detectar_colisiones(df: pd.DataFrame, col_codigo: str, col_producto: str) -> pd.DataFrame:
    """Devuelve los códigos asociados a más de un producto único, con el archivo de origen de cada uno."""

    # Paso 1: para cada par (codigo, producto) obtener los archivos donde aparece
    origen_por_par = (
        df.groupby([col_codigo, col_producto])["_origen"]
        .apply(lambda x: ", ".join(sorted(set(x))))
        .reset_index()
    )
    origen_por_par.columns = ["Codigo", "Producto", "Archivos"]

    # Paso 2: iterar por código y construir la descripción solo si hay >1 producto
    filas = []
    for codigo, grupo in origen_por_par.groupby("Codigo"):
        if grupo["Producto"].nunique() <= 1:
            continue
        partes = [
            f"{row['Producto']}  [{row['Archivos']}]"
            for _, row in grupo.iterrows()
        ]
        filas.append({
            "Codigo":          codigo,
            "Cantidad":        grupo["Producto"].nunique(),
            "Productos_texto": " | ".join(partes),
        })

    if not filas:
        return pd.DataFrame(columns=["Codigo", "Cantidad", "Productos_texto"])

    resultado = pd.DataFrame(filas)
    resultado.sort_values("Codigo", inplace=True)
    resultado.reset_index(drop=True, inplace=True)
    return resultado


def generar_excel_compartidos(df_compartidos: pd.DataFrame) -> bytes:
    """Genera el reporte de compartidos en memoria."""
    df_export = df_compartidos[["Codigo", "Producto", "Num_archivos", "Archivos_texto"]].rename(columns={
        "Codigo":         "Código",
        "Producto":       "Producto",
        "Num_archivos":   "Cantidad de archivos",
        "Archivos_texto": "Presente en",
    })
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df_export.to_excel(writer, index=False, sheet_name="Compartidos")
        hoja = writer.sheets["Compartidos"]
        for col_cells in hoja.columns:
            ancho = max(len(str(c.value or "")) for c in col_cells)
            hoja.column_dimensions[col_cells[0].column_letter].width = min(ancho + 4, 80)
    buf.seek(0)
    return buf.read()


def generar_excel(df_colisiones: pd.DataFrame) -> bytes:
    """Genera el reporte Excel en memoria y devuelve los bytes."""
    df_export = df_colisiones[["Codigo", "Cantidad", "Productos_texto"]].rename(columns={
        "Codigo":          "Código",
        "Cantidad":        "Cantidad de productos distintos",
        "Productos_texto": "Productos en conflicto (con archivo de origen)",
    })
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df_export.to_excel(writer, index=False, sheet_name="Colisiones")
        hoja = writer.sheets["Colisiones"]
        for col_cells in hoja.columns:
            ancho = max(len(str(c.value or "")) for c in col_cells)
            hoja.column_dimensions[col_cells[0].column_letter].width = min(ancho + 4, 80)
    buf.seek(0)
    return buf.read()


# ─────────────────────────────────────────────────────────────────────────────
# LÓGICA — GENERADOR DE SKUs
# ─────────────────────────────────────────────────────────────────────────────

CATALOGO_TALLAS: dict[int, str] = {
    0:  "2",
    1:  "4",
    2:  "6",
    3:  "8",
    4:  "10",
    5:  "12",
    6:  "CH",
    7:  "M",
    8:  "G",
    9:  "XL",
    10: "XXL",
    11: "XXXL",
}

CATALOGO_COLORES: dict[int, str] = {
    1:  "NEGRO",        2:  "MARINO",       3:  "MORADO",
    4:  "AZUL PETROLEO",5:  "VINO",         6:  "V. MILITAR",
    7:  "GRIS",         8:  "TURQUESA",     9:  "ROJO",
    10: "REY",          11: "JADE",         12: "CARNE/NUDE",
    13: "FIUSHA",       14: "PALO DE ROSA", 15: "ANARANJADO",
    16: "OCRE",         17: "SALMON",       18: "V. AGUA",
    19: "AZUL CIELO",   20: "LILA",         21: "ROSA PASTEL",
    22: "BLANCO",       23: "VARIOS/X TONO",24: "AMARILLO",
}

def generar_skus(
    modelo: int,
    colores: list[int],
    variantes: list[int],
    tallas: list[int],
) -> list[dict]:
    """
    Genera permutaciones: Variante → Color (orden numérico) → Talla (orden numérico).
    Devuelve lista de dicts con 'codigo' y 'descripcion'.
    """
    filas = []
    for var in sorted(variantes):
        for col in sorted(colores):
            nombre_color = CATALOGO_COLORES.get(col, f"COLOR {col:02d}")
            for talla_cod in sorted(tallas):
                nombre_talla = CATALOGO_TALLAS.get(talla_cod, f"{talla_cod:02d}")
                codigo = f"{modelo:03d}{talla_cod:02d}{col:02d}{var:02d}"
                desc   = f"Modelo {modelo:03d} - Talla {nombre_talla} - {nombre_color} - Var {var:02d}"
                filas.append({"Código": codigo, "Descripción": desc})
    return filas


# ─────────────────────────────────────────────────────────────────────────────
# LÓGICA — GENERADOR SECUENCIAL (herramienta auxiliar)
# ─────────────────────────────────────────────────────────────────────────────

def _decodificar_sec(codigo: str) -> tuple[int, int]:
    if len(codigo) != 9 or not codigo.isdigit():
        raise ValueError("El código debe tener exactamente 9 dígitos numéricos.")
    if not codigo.startswith("121"):
        raise ValueError("El código debe comenzar con '121'.")
    if not codigo.endswith("01"):
        raise ValueError("El código debe terminar con '01'.")
    xx = int(codigo[3:5])
    yy = int(codigo[5:7])
    if not (1 <= xx <= 12):
        raise ValueError(f"XX debe estar entre 01 y 12. Encontrado: {xx:02d}.")
    return xx, yy

def _generar_sec(codigo_inicial: str, cantidad: int = 240) -> list[str]:
    xx, yy = _decodificar_sec(codigo_inicial)
    codigos = []
    for _ in range(cantidad):
        xx += 1
        if xx > 12:
            xx = 1
            yy += 1
        codigos.append(f"121{xx:02d}{yy:02d}01")
    return codigos


# ─────────────────────────────────────────────────────────────────────────────
# UI — PÁGINAS
# ─────────────────────────────────────────────────────────────────────────────

def pagina_generador():
    st.title("🔢 Generador de Códigos")
    st.caption("ODAKI — Genera códigos de 9 dígitos para tus productos")

    tab_sku, tab_sec = st.tabs(["🔢 Generador de Códigos por Modelo", "↗️ Generador Secuencial"])

    # ══════════════════════════════════════════════════════════════════════════
    # TAB 1 — GENERADOR DE SKUs
    # ══════════════════════════════════════════════════════════════════════════
    with tab_sku:
        st.markdown(
            "**Formato:** `MMM` + `TT` + `CC` + `VV`  \n"
            "`MMM` = Modelo (3 dígitos) · `TT` = Talla · `CC` = Color · `VV` = Variante/Escote"
        )
        st.divider()

        # ── Paso 1: Modelo ────────────────────────────────────────────────────
        st.subheader("Paso 1 — Modelo")
        modelo_str = st.text_input(
            "Número de modelo (3 dígitos)",
            placeholder="Ej: 121",
            max_chars=3,
            key="sku_modelo",
        )

        st.divider()

        # ── Paso 1b: Tallas ───────────────────────────────────────────────────
        st.subheader("Paso 2 — Tallas")
        st.caption("Selecciona las tallas en que se produce este modelo.")

        opciones_talla = [
            f"{cod:02d} — Talla {nombre}"
            for cod, nombre in CATALOGO_TALLAS.items()
        ]
        todas_tallas = st.checkbox("Seleccionar todas las tallas", value=True, key="sku_todas_tallas")
        seleccion_tallas = st.multiselect(
            "Tallas disponibles",
            options=opciones_talla,
            default=opciones_talla if todas_tallas else [],
            placeholder="Selecciona una o más tallas...",
            key="sku_tallas",
        )
        tallas_num = [int(t.split(" — ")[0]) for t in seleccion_tallas]

        st.divider()

        # ── Paso 3: Colores ───────────────────────────────────────────────────
        st.subheader("Paso 3 — Colores fabricados")
        st.caption("Selecciona solo los colores en que se produce este modelo.")

        opciones_color = [
            f"{num:02d} — {nombre}"
            for num, nombre in CATALOGO_COLORES.items()
        ]
        seleccion_colores = st.multiselect(
            "Colores disponibles",
            options=opciones_color,
            placeholder="Selecciona uno o más colores...",
            key="sku_colores",
        )
        colores_num = [int(c.split(" — ")[0]) for c in seleccion_colores]

        st.divider()

        # ── Paso 4: Variantes ─────────────────────────────────────────────────
        st.subheader("Paso 4 — Variantes / Escotes")
        variantes_str = st.text_input(
            "Variantes (separadas por coma)",
            value="01",
            placeholder="Ej: 01  ó  01, 02, 03",
            key="sku_variantes",
            help="Escribe los números de variante separados por coma. Ej: 01, 02",
        )

        st.divider()

        # ── Generación ────────────────────────────────────────────────────────
        if st.button("Generar códigos", type="primary", use_container_width=True, key="btn_sku"):
            errores = []
            if not modelo_str or not modelo_str.isdigit():
                errores.append("El modelo debe ser un número de hasta 3 dígitos.")
            if not tallas_num:
                errores.append("Selecciona al menos una talla.")
            if not colores_num:
                errores.append("Selecciona al menos un color.")

            variantes_num = []
            for v in variantes_str.split(","):
                v = v.strip()
                if v.isdigit():
                    variantes_num.append(int(v))
                else:
                    errores.append(f"Variante inválida: '{v}'. Solo números.")

            if errores:
                for e in errores:
                    st.error(e)
            else:
                modelo_int = int(modelo_str)
                filas = generar_skus(modelo_int, colores_num, variantes_num, tallas_num)
                df_result = pd.DataFrame(filas)

                st.success(f"Se generaron **{len(df_result)}** códigos.")

                # Métricas
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("Total códigos", len(df_result))
                c2.metric("Tallas",        len(tallas_num))
                c3.metric("Colores",       len(colores_num))
                c4.metric("Variantes",     len(variantes_num))

                st.divider()

                # Vista previa
                st.subheader("Vista previa")
                st.dataframe(df_result, use_container_width=True, height=380, hide_index=True)

                st.divider()

                # Descargas
                col_d1, col_d2, col_d3 = st.columns(3)

                # TSV — para pegar en Excel con 2 columnas
                tsv = "\n".join(
                    f"{r['Código']}\t{r['Descripción']}"
                    for _, r in df_result.iterrows()
                )
                with col_d1:
                    st.download_button(
                        label="📋 Descargar .txt / TSV\n(2 columnas en Excel)",
                        data=tsv.encode("utf-8"),
                        file_name=f"codigos_modelo_{modelo_str}.txt",
                        mime="text/plain",
                        use_container_width=True,
                    )

                # Excel
                buf_xl = io.BytesIO()
                with pd.ExcelWriter(buf_xl, engine="openpyxl") as w:
                    df_result.to_excel(w, index=False, sheet_name="Codigos")
                    hoja = w.sheets["Codigos"]
                    hoja.column_dimensions["A"].width = 14
                    hoja.column_dimensions["B"].width = 50
                buf_xl.seek(0)
                with col_d2:
                    st.download_button(
                        label="📊 Descargar .xlsx",
                        data=buf_xl.read(),
                        file_name=f"codigos_modelo_{modelo_str}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )

                # Solo códigos (para pegar en una sola columna)
                solo_codigos = "\n".join(df_result["Código"].tolist())
                with col_d3:
                    st.download_button(
                        label="🔢 Solo códigos (.txt)\n(1 columna en Excel)",
                        data=solo_codigos.encode("utf-8"),
                        file_name=f"codigos_modelo_{modelo_str}.txt",
                        mime="text/plain",
                        use_container_width=True,
                    )

                st.divider()
                st.subheader("Texto TSV — copia y pega en Excel")
                st.caption("Pega en Excel: la primera columna será el Código, la segunda la Descripción.")
                st.text_area("Selecciona todo y copia (Ctrl+A → Ctrl+C)", value=tsv, height=300, key="tsv_area")

    # ══════════════════════════════════════════════════════════════════════════
    # TAB 2 — GENERADOR SECUENCIAL
    # ══════════════════════════════════════════════════════════════════════════
    with tab_sec:
        st.markdown(
            "**Formato:** `121` + `XX` (01-12) + `YY` + `01`  \n"
            "Útil para continuar una numeración existente de forma secuencial."
        )
        st.divider()

        col_inp, col_cant = st.columns([2, 1])
        with col_inp:
            codigo_inicial = st.text_input(
                "Código inicial (9 dígitos)",
                placeholder="Ej: 121010201",
                max_chars=9,
                key="sec_inicial",
            )
        with col_cant:
            cantidad = st.number_input(
                "Cantidad a generar",
                min_value=1, max_value=1200, value=240, step=12,
                key="sec_cantidad",
            )

        if st.button("Generar secuencia", type="primary", use_container_width=True, key="btn_sec"):
            if not codigo_inicial:
                st.warning("Ingresa un código inicial.")
            else:
                try:
                    codigos = _generar_sec(codigo_inicial.strip(), int(cantidad))
                    df_sec = pd.DataFrame({"Código": codigos})

                    st.success(f"Se generaron **{len(codigos)}** códigos.")
                    c1, c2 = st.columns(2)
                    c1.metric("Primero", codigos[0])
                    c2.metric("Último",  codigos[-1])

                    st.divider()
                    st.dataframe(df_sec, use_container_width=True, height=360, hide_index=True)
                    st.divider()

                    col_s1, col_s2 = st.columns(2)
                    with col_s1:
                        st.download_button(
                            "📄 Descargar .txt",
                            data="\n".join(codigos).encode("utf-8"),
                            file_name=f"seq_desde_{codigo_inicial.strip()}.txt",
                            mime="text/plain",
                            use_container_width=True,
                        )
                    with col_s2:
                        buf_s = io.BytesIO()
                        with pd.ExcelWriter(buf_s, engine="openpyxl") as w:
                            df_sec.to_excel(w, index=False, sheet_name="Secuencia")
                        buf_s.seek(0)
                        st.download_button(
                            "📊 Descargar .xlsx",
                            data=buf_s.read(),
                            file_name=f"seq_desde_{codigo_inicial.strip()}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True,
                        )

                    st.text_area("Copia y pega en Excel", value="\n".join(codigos), height=250, key="sec_area")
                except ValueError as e:
                    st.error(f"Error: {e}")

    st.caption("Desarrollado para ODAKI · Sistema de Revisión de Bases ELEVENTA")


def pagina_detector():
    # ── Sidebar: carga de archivos ───────────────────────────────────────────
    with st.sidebar:
        st.header("📂 Archivos Excel")
        st.caption("Sube entre 1 y 4 archivos. Si no subes ninguno, se usan los archivos por defecto del sistema.")

        archivos_subidos = st.file_uploader(
            "Selecciona los archivos Excel",
            type=["xlsx", "xls"],
            accept_multiple_files=True,
            help="Puedes subir hasta 4 archivos a la vez.",
        )

        st.divider()

        if archivos_subidos:
            st.success(f"{len(archivos_subidos)} archivo(s) cargado(s)")
            for f in archivos_subidos:
                st.markdown(f"- `{f.name}`")
            fuentes = [(f, f.name) for f in archivos_subidos]
            modo = "subidos"
        else:
            fuentes = []
            modo = "ninguno"

        st.divider()
        col_cod  = st.text_input("Columna de código",   value=COLUMNA_CODIGO)
        col_prod = st.text_input("Columna de producto", value=COLUMNA_PRODUCTO)

    # ── Encabezado ───────────────────────────────────────────────────────────
    st.title("⚠️ Detector de Colisiones de Etiquetado")
    st.caption("ODAKI — Sistema de revisión de códigos duplicados entre bases ELEVENTA")

    # ── Sin archivos ─────────────────────────────────────────────────────────
    if modo == "ninguno":
        st.info("👈 Sube tus archivos Excel desde el panel izquierdo para comenzar el análisis.")
        st.stop()

    # Nombres de columna activos (pueden venir del sidebar)
    col_codigo   = col_cod.strip()
    col_producto = col_prod.strip()

    # ── Carga y análisis ─────────────────────────────────────────────────────
    with st.spinner("Analizando archivos..."):
        df_total = cargar_datos(fuentes, col_codigo, col_producto)

    if df_total is None or df_total.empty:
        st.error("No se pudo cargar ningún dato válido. Revisa los archivos y los nombres de columnas.")
        st.stop()

    df_colisiones  = detectar_colisiones(df_total, col_codigo, col_producto)
    df_compartidos = detectar_compartidos(df_total, col_codigo, col_producto)

    # ── Métricas ─────────────────────────────────────────────────────────────
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Archivos analizados",   len(fuentes))
    c2.metric("Pares únicos totales",  f"{len(df_total):,}")
    c3.metric("Códigos en conflicto",  f"{len(df_colisiones):,}",
              delta=f"-{len(df_colisiones)} a corregir", delta_color="inverse")
    c4.metric("Códigos sin conflicto",
              f"{df_total[col_codigo].nunique() - len(df_colisiones):,}")
    c5.metric("Compartidos (sin conflicto)", f"{len(df_compartidos):,}",
              delta="mismo codigo, mismo producto", delta_color="normal")

    st.divider()

    # ── Tabs principales ──────────────────────────────────────────────────────
    tab_conflicto, tab_compartidos = st.tabs([
        f"⚠️ Códigos en conflicto ({len(df_colisiones)})",
        f"✅ Códigos compartidos sin conflicto ({len(df_compartidos)})",
    ])

    # ══════════════════════════════════════════════════════════════════════════
    # TAB 1 — COLISIONES
    # ══════════════════════════════════════════════════════════════════════════
    with tab_conflicto:
        if df_colisiones.empty:
            st.success("No se encontraron colisiones. Todos los códigos son consistentes.")
        else:
            col_busq, col_filt = st.columns([2, 1])
            with col_busq:
                busqueda = st.text_input("🔍 Buscar código o producto",
                                         placeholder="Ej: 94000101  o  MALLA",
                                         key="busq_conflicto")
            with col_filt:
                min_conf = st.selectbox("Mínimo de conflictos", options=[2, 3, 4, 5], index=0)

            df_vista = df_colisiones[df_colisiones["Cantidad"] >= min_conf].copy()
            if busqueda:
                mask = (
                    df_vista["Codigo"].str.contains(busqueda, case=False, na=False)
                    | df_vista["Productos_texto"].str.contains(busqueda, case=False, na=False)
                )
                df_vista = df_vista[mask]

            st.caption(f"Mostrando **{len(df_vista)}** colisiones")
            st.dataframe(
                df_vista[["Codigo", "Cantidad", "Productos_texto"]].rename(columns={
                    "Codigo":          "Código",
                    "Cantidad":        "# Productos distintos",
                    "Productos_texto": "Productos en conflicto",
                }),
                use_container_width=True,
                height=420,
                hide_index=True,
            )

            st.subheader("🔎 Detalle de un código en conflicto")
            codigo_sel = st.selectbox(
                "Selecciona un código para ver en qué archivos aparece",
                options=[""] + df_colisiones["Codigo"].tolist(),
                key="sel_conflicto",
            )
            if codigo_sel:
                detalle = df_total[df_total[col_codigo] == codigo_sel][
                    [col_codigo, col_producto, "_origen"]
                ].rename(columns={
                    col_codigo:   "Código",
                    col_producto: "Producto",
                    "_origen":    "Archivo de origen",
                })
                st.dataframe(detalle, use_container_width=True, hide_index=True)

            st.divider()
            st.download_button(
                label="📥 Descargar reporte de colisiones (.xlsx)",
                data=generar_excel(df_colisiones),
                file_name="reporte_colisiones_etiquetado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

    # ══════════════════════════════════════════════════════════════════════════
    # TAB 2 — COMPARTIDOS SIN CONFLICTO
    # ══════════════════════════════════════════════════════════════════════════
    with tab_compartidos:
        st.caption(
            "Estos códigos tienen **exactamente el mismo producto** en más de un archivo. "
            "No representan un error, pero confirman que el código está unificado entre bases."
        )

        if df_compartidos.empty:
            st.info("No hay códigos compartidos entre archivos.")
        else:
            col_busq2, col_filt2 = st.columns([2, 1])
            with col_busq2:
                busqueda2 = st.text_input(
                    "🔍 Buscar código o producto",
                    placeholder="Ej: 71002201  o  CALCETA",
                    key="busq_compartidos",
                )
            with col_filt2:
                min_archivos = st.selectbox(
                    "Aparece en al menos N archivos",
                    options=list(range(2, len(fuentes) + 1)),
                    index=0,
                    key="filt_compartidos",
                )

            df_comp_vista = df_compartidos[df_compartidos["Num_archivos"] >= min_archivos].copy()
            if busqueda2:
                mask2 = (
                    df_comp_vista["Codigo"].str.contains(busqueda2, case=False, na=False)
                    | df_comp_vista["Producto"].str.contains(busqueda2, case=False, na=False)
                )
                df_comp_vista = df_comp_vista[mask2]

            st.caption(f"Mostrando **{len(df_comp_vista)}** códigos compartidos")
            st.dataframe(
                df_comp_vista[["Codigo", "Producto", "Num_archivos", "Archivos_texto"]].rename(columns={
                    "Codigo":         "Código",
                    "Producto":       "Producto",
                    "Num_archivos":   "# Archivos",
                    "Archivos_texto": "Presente en",
                }),
                use_container_width=True,
                height=420,
                hide_index=True,
            )

            st.divider()
            # Descarga de compartidos
            buf_comp = generar_excel_compartidos(df_compartidos)
            st.download_button(
                label="📥 Descargar reporte de compartidos (.xlsx)",
                data=buf_comp,
                file_name="reporte_compartidos.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

    st.caption("Desarrollado para ODAKI · Sistema de Revisión de Bases ELEVENTA")


# ─────────────────────────────────────────────────────────────────────────────
# MAIN — Navegación entre herramientas
# ─────────────────────────────────────────────────────────────────────────────

def main():
    st.set_page_config(
        page_title="Sistema ODAKI",
        page_icon="🏷️",
        layout="wide",
    )

    with st.sidebar:
        st.image("https://img.icons8.com/color/96/barcode.png", width=60)
        st.title("Sistema ODAKI")
        st.caption("Gestión de códigos de etiquetado")
        st.divider()
        pagina = st.radio(
            "Herramientas",
            options=["⚠️ Detector de Colisiones", "🔢 Generador de Códigos"],
            index=0,
        )
        st.divider()

    if pagina == "⚠️ Detector de Colisiones":
        pagina_detector()
    else:
        pagina_generador()


if __name__ == "__main__":
    main()
