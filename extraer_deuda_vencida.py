"""Consolida deuda vencida de reportes DIAN por NIT.

Instalar: py -m pip install pymupdf openpyxl
Ejecutar: py extraer_deuda_vencida.py
"""
from __future__ import annotations

import csv
import re
from collections import defaultdict
from datetime import datetime
from pathlib import Path

import fitz  # PyMuPDF
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

CARPETA_RAIZ = Path(r"D:\CONFIRMACION OBL.FIN ENDOSOS")
SOLO_VALORES_POSITIVOS = True
PATRON_CARPETA = re.compile(r"^Obligaciones Financieras NIT\s*-\s*(\d+)\s*$", re.I)
PATRON_ANIO = re.compile(r"(?:19|20)\d{2}$")
PATRON_NUMERO = re.compile(r"-?\d[\d.,]*$")
CABECERA = ["NIT", "Año", "Obligación", "Estado", "Valor", "Número de obligaciones", "Archivo PDF", "Página"]


def numero(texto: str) -> int:
    """Importes enteros de los reportes (ej. 184.537.000)."""
    return int(re.sub(r"[^\d-]", "", texto))


def lineas_pagina(pagina):
    """Une fragmentos de texto pertenecientes a una misma línea."""
    resultado = []
    for bloque in pagina.get_text("dict")["blocks"]:
        for linea in bloque.get("lines", []):
            partes = linea["spans"]
            if not partes:
                continue
            texto = " ".join(p["text"].strip() for p in partes).strip()
            x0, y0, x1, y1 = linea["bbox"]
            resultado.append((x0, y0, x1, y1, texto))
    return resultado


def extraer_pdf(ruta: Path, nit: str):
    registros = []
    advertencias = []
    with fitz.open(ruta) as documento:
        if documento.needs_pass:
            raise ValueError("PDF protegido con contraseña")
        for pagina_num, pagina in enumerate(documento, 1):
            palabras = pagina.get_text("words")
            lineas = lineas_pagina(pagina)
            if not palabras:
                advertencias.append(f"Página {pagina_num} sin texto extraíble (PDF posiblemente escaneado)")
                continue

            # Encabezados de cada bloque de tabla: tres años en columnas.
            encabezados = defaultdict(list)
            for x0, y0, x1, y1, palabra, *_ in palabras:
                if PATRON_ANIO.fullmatch(palabra) and x0 > 320:
                    encabezados[round(y0, 1)].append((x0, int(palabra)))
            bloques = []
            for y, columnas in sorted(encabezados.items()):
                if len(columnas) >= 2:
                    bloques.append((y, sorted(columnas)))

            # Nombres de obligación: primera columna de la tabla.
            obligaciones = sorted(
                (y0, txt) for x0, y0, _x1, _y1, txt in lineas
                if 25 <= x0 <= 60 and txt and txt.lower() != "obligación"
            )

            estados = sorted(
                (y0, y1) for x0, y0, _x1, y1, txt in lineas
                if x0 > 210 and re.sub(r"\s+", " ", txt).upper() == "DEUDA VENCIDA"
            )
            for estado_y0, estado_y1 in estados:
                # Año del encabezado que aparece inmediatamente antes de la fila.
                bloque = next((cols for by, cols in reversed(bloques) if by <= estado_y0 + 2), None)
                obligacion = next((t for oy, t in reversed(obligaciones) if oy <= estado_y0 + 3), None)
                if not bloque or not obligacion:
                    advertencias.append(f"Página {pagina_num}, y={estado_y0:.1f}: no se pudo identificar obligación o años")
                    continue

                # Números situados horizontalmente en la fila; las cifras suelen
                # estar 1-3 puntos más arriba que el texto del estado.
                tokens = sorted(
                    (x0, palabra) for x0, y0, x1, y1, palabra, *_ in palabras
                    if x0 > 325 and abs((y0 + y1)/2 - (estado_y0 + estado_y1)/2) <= 5
                    and PATRON_NUMERO.fullmatch(palabra)
                )
                if not tokens:
                    advertencias.append(f"Página {pagina_num}, y={estado_y0:.1f}: fila sin cifras")
                    continue

                for indice, (x_anio, anio) in enumerate(bloque):
                    # Dos celdas por año: valor a la izquierda, cantidad a la derecha.
                    # Partimos el espacio entre años por sus puntos medios.
                    limite_izq = (bloque[indice - 1][0] + x_anio) / 2 if indice else x_anio - 70
                    limite_der = (x_anio + bloque[indice + 1][0]) / 2 if indice + 1 < len(bloque) else x_anio + 95
                    celdas = [v for x, v in tokens if limite_izq <= x < limite_der]
                    if len(celdas) != 2:
                        advertencias.append(
                            f"Página {pagina_num}, {obligacion}, {anio}: esperadas 2 cifras; obtenidas {len(celdas)}"
                        )
                        continue
                    valor, cantidad = map(numero, celdas)
                    if SOLO_VALORES_POSITIVOS and valor <= 0:
                        continue
                    registros.append({
                        "NIT": nit,
                        "Año": anio,
                        "Obligación": obligacion,
                        "Estado": "DEUDA VENCIDA",
                        "Valor": valor,
                        "Número de obligaciones": cantidad,
                        "Archivo PDF": str(ruta),
                        "Página": pagina_num,
                    })
    return registros, advertencias


def guardar_csv(ruta: Path, filas: list[dict], columnas: list[str]):
    with ruta.open("w", newline="", encoding="utf-8-sig") as archivo:
        escritor = csv.DictWriter(archivo, fieldnames=columnas, delimiter=";")
        escritor.writeheader()
        escritor.writerows(filas)


def guardar_excel(ruta: Path, hojas):
    libro = Workbook()
    libro.remove(libro.active)
    for nombre, columnas, filas in hojas:
        hoja = libro.create_sheet(nombre)
        hoja.append(columnas)
        for fila in filas:
            hoja.append([fila.get(col, "") for col in columnas])
        hoja.freeze_panes = "A2"
        hoja.auto_filter.ref = hoja.dimensions
        for celda in hoja[1]:
            celda.font = Font(color="FFFFFF", bold=True)
            celda.fill = PatternFill("solid", fgColor="17365D")
        for idx, col in enumerate(columnas, 1):
            ancho = min(75, max(12, len(col) + 3))
            if col == "Archivo PDF":
                ancho = 75
            elif col == "Obligación":
                ancho = 48
            hoja.column_dimensions[get_column_letter(idx)].width = ancho
        for fila in hoja.iter_rows(min_row=2):
            for celda in fila:
                if columnas[celda.column - 1] in ("Valor", "Valor total"):
                    celda.number_format = '#,##0'
    libro.save(ruta)


def main():
    if not CARPETA_RAIZ.is_dir():
        raise SystemExit(f"No existe la carpeta: {CARPETA_RAIZ}")

    encontrados = []
    control = []
    errores = []
    carpetas = sorted(p for p in CARPETA_RAIZ.iterdir() if p.is_dir() and PATRON_CARPETA.fullmatch(p.name))
    print(f"Carpetas NIT identificadas: {len(carpetas)}")
    for carpeta in carpetas:
        nit = PATRON_CARPETA.fullmatch(carpeta.name).group(1)
        ruta = carpeta / f"{nit}_REPORTE_CONSOLIDADO.pdf"
        if not ruta.is_file():
            control.append({"NIT": nit, "Estado procesamiento": "PDF no encontrado", "Registros": 0, "Valor total": 0, "Detalle": str(ruta)})
            continue
        try:
            filas, avisos = extraer_pdf(ruta, nit)
            encontrados.extend(filas)
            total = sum(f["Valor"] for f in filas)
            estado = "Procesado con advertencias" if avisos else ("Con deuda vencida" if filas else "Sin deuda vencida positiva")
            control.append({"NIT": nit, "Estado procesamiento": estado, "Registros": len(filas), "Valor total": total, "Detalle": " | ".join(avisos)})
            for aviso in avisos:
                errores.append({"NIT": nit, "Archivo PDF": str(ruta), "Advertencia": aviso})
            print(f"{nit}: {len(filas)} registros; valor total {total:,}; advertencias {len(avisos)}")
        except Exception as exc:
            mensaje = f"{type(exc).__name__}: {exc}"
            control.append({"NIT": nit, "Estado procesamiento": "Error de lectura", "Registros": 0, "Valor total": 0, "Detalle": mensaje})
            errores.append({"NIT": nit, "Archivo PDF": str(ruta), "Advertencia": mensaje})
            print(f"{nit}: ERROR {mensaje}")

    encontrados.sort(key=lambda x: (x["NIT"], x["Año"], x["Obligación"]))
    columnas_control = ["NIT", "Estado procesamiento", "Registros", "Valor total", "Detalle"]
    columnas_errores = ["NIT", "Archivo PDF", "Advertencia"]
    guardar_csv(CARPETA_RAIZ / "DEUDA_VENCIDA_CONSOLIDADO.csv", encontrados, CABECERA)
    guardar_csv(CARPETA_RAIZ / "DEUDA_VENCIDA_CONTROL.csv", control, columnas_control)
    guardar_excel(CARPETA_RAIZ / "DEUDA_VENCIDA_CONSOLIDADO.xlsx", [
        ("Deuda vencida", CABECERA, encontrados),
        ("Control NIT", columnas_control, control),
        ("Advertencias", columnas_errores, errores),
    ])
    print(f"\nTERMINADO: {len(encontrados)} registros; {sum(r['Valor'] for r in encontrados):,} COP")
    print(f"Resultados guardados en {CARPETA_RAIZ}")


if __name__ == "__main__":
    main()