"""
generar_codigos.py
------------------
Genera 240 códigos de barras secuenciales a partir de un código inicial de 9 dígitos.

Formato: 121 + XX + YY + 01
  - 121  : prefijo fijo
  - XX   : contador primario (01-12), se incrementa primero
  - YY   : contador secundario, sube cuando XX completa el ciclo de 12
  - 01   : sufijo fijo

Uso:
    python generar_codigos.py
"""

PREFIJO  = "121"
SUFIJO   = "01"
XX_MIN   = 1
XX_MAX   = 12
CANTIDAD = 240


def decodificar(codigo: str) -> tuple[int, int]:
    """Extrae XX e YY de un código de 9 dígitos."""
    if len(codigo) != 9:
        raise ValueError(f"El código debe tener exactamente 9 dígitos. Recibido: '{codigo}' ({len(codigo)} dígitos).")
    if not codigo.isdigit():
        raise ValueError(f"El código debe contener solo números. Recibido: '{codigo}'.")
    if not codigo.startswith(PREFIJO):
        raise ValueError(f"El código debe comenzar con '{PREFIJO}'. Recibido: '{codigo}'.")
    if not codigo.endswith(SUFIJO):
        raise ValueError(f"El código debe terminar con '{SUFIJO}'. Recibido: '{codigo}'.")

    xx = int(codigo[3:5])   # posiciones 3-4
    yy = int(codigo[5:7])   # posiciones 5-6

    if not (XX_MIN <= xx <= XX_MAX):
        raise ValueError(f"XX debe estar entre {XX_MIN:02d} y {XX_MAX:02d}. Valor encontrado: {xx:02d}.")
    if yy < 1:
        raise ValueError(f"YY debe ser mayor o igual a 01. Valor encontrado: {yy:02d}.")

    return xx, yy


def siguiente(xx: int, yy: int) -> tuple[int, int]:
    """Devuelve el siguiente par (XX, YY) respetando el patrón de ciclos."""
    xx += 1
    if xx > XX_MAX:
        xx = XX_MIN
        yy += 1
    return xx, yy


def construir_codigo(xx: int, yy: int) -> str:
    """Construye el código de 9 dígitos con el formato 121 + XX + YY + 01."""
    return f"{PREFIJO}{xx:02d}{yy:02d}{SUFIJO}"


def generar(codigo_inicial: str, cantidad: int = CANTIDAD) -> list[str]:
    """Genera `cantidad` códigos a partir del siguiente al código inicial."""
    xx, yy = decodificar(codigo_inicial)
    codigos = []
    for _ in range(cantidad):
        xx, yy = siguiente(xx, yy)
        codigos.append(construir_codigo(xx, yy))
    return codigos


def main():
    print("=" * 50)
    print("  GENERADOR DE CODIGOS DE BARRAS SECUENCIALES")
    print("=" * 50)
    print(f"  Formato: {PREFIJO} + XX(01-12) + YY + {SUFIJO}")
    print("=" * 50)

    while True:
        codigo_inicial = input("\nIngresa el codigo inicial (9 digitos): ").strip()
        try:
            codigos = generar(codigo_inicial, CANTIDAD)
            break
        except ValueError as e:
            print(f"\n  [ERROR] {e}")
            print("  Intenta de nuevo.\n")

    # Mostrar muestra en terminal
    print(f"\nSe generaron {len(codigos)} codigos.")
    print(f"Primeros 5  : {', '.join(codigos[:5])}")
    print(f"Ultimos 5   : {', '.join(codigos[-5:])}")

    # Guardar en archivo de texto (listo para pegar en Excel)
    nombre_salida = f"codigos_desde_{codigo_inicial}.txt"
    ruta_salida   = f"c:\\Users\\Aureleano\\Desktop\\SistemasODAKIDM\\ArreglaCodigos\\{nombre_salida}"

    with open(ruta_salida, "w", encoding="utf-8") as f:
        f.write("\n".join(codigos))

    print(f"\nArchivo guardado en:\n  {ruta_salida}")
    print("\nAbre el archivo .txt, selecciona todo (Ctrl+A),")
    print("copia (Ctrl+C) y pega directo en la columna de Excel.")
    print("\n" + "=" * 50)


if __name__ == "__main__":
    main()
