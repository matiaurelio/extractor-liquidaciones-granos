"""
Extrae datos de "Liquidación Primaria de Granos" (PDF de ARCA) y los vuelca a un Excel.

USO:
    python extraer_liquidaciones.py carpeta_con_pdfs/ salida.xlsx
    python extraer_liquidaciones.py archivo1.pdf archivo2.pdf salida.xlsx

Datos que extrae por cada PDF:
    - FECHA            -> fecha de la operación (campo "Fecha:" de CONDICIONES DE LA OPERACIÓN)
    - COE               -> número de C.O.E.
    - Retención LA PAMPA / BUENOS AIRES (u otra provincia que aparezca)
    - Otras retenciones (IVA, Ganancias, Sellado, etc. - se suman aparte)
    - Total retenciones
    - Importe Neto a Pagar (para liquidaciones de "Ajuste Unificado" toma el
      "Pago según condiciones" final, que es el neto real a cobrar/pagar)
    - Archivo de origen (para poder rastrear cualquier dato)

Si tu PDF tiene un formato distinto y algún dato no se detecta, va a quedar
vacío esa celda -- no se inventa nada.
"""

import re
import sys
from pathlib import Path

import pdfplumber
import pandas as pd

# Provincias que reconocemos como columna propia (podés agregar más acá)
PROVINCIAS_CONOCIDAS = ["LA PAMPA", "BUENOS AIRES", "SANTA FE", "CORDOBA", "CÓRDOBA"]


def normalizar_monto(texto):
    """Convierte '$ 1.234,56' / '1,234.56' / '$1234.56' a float."""
    if texto is None:
        return None
    t = texto.replace("$", "").replace("\n", " ").strip()
    t = t.replace(" ", "")
    if not t:
        return None
    # Formato con coma de miles y punto decimal (el que usan estos PDF)
    t = t.replace(",", "")
    try:
        return float(t)
    except ValueError:
        return None


def extraer_fecha(texto_completo):
    # Preferimos la fecha del encabezado (fecha de emisión del comprobante),
    # que aparece como "DD/MM/AAAA, LOCALIDAD" arriba de cada página.
    m = re.search(r"(\d{2}/\d{2}/\d{4}),\s*[A-ZÁÉÍÓÚÑ ]+", texto_completo)
    if m:
        return m.group(1)
    # fallback: campo "Fecha:" de la tabla de condiciones de la operación
    m = re.search(r"Fecha:\s*(\d{2}/\d{2}/\d{4})", texto_completo)
    if m:
        return m.group(1)
    return None


def extraer_coe(texto_completo):
    m = re.search(r"C\.O\.E\.:\s*(\d+)", texto_completo)
    return m.group(1) if m else None


def extraer_importe_neto(texto_completo):
    """
    Toma el valor del campo "Importe Neto a Pagar" tal cual aparece en el PDF.
    En los documentos de "Ajuste Unificado" este campo puede aparecer dos
    veces (Débito y luego Crédito): se usa la PRIMERA ocurrencia (la de
    Débito), que es la que corresponde al ajuste principal.
    """
    m = re.search(r"Importe Neto a Pagar:?\s*\$?\s*([\d.,]+)", texto_completo)
    if m:
        return normalizar_monto(m.group(1))
    return None


def extraer_retenciones(pdf):
    """
    Recorre todas las páginas buscando la tabla RETENCIONES y devuelve una lista
    de dicts: {concepto, detalle, importe}
    """
    retenciones = []
    for page in pdf.pages:
        for tabla in page.extract_tables():
            # Buscamos, en CUALQUIER fila de la tabla, el encabezado de la
            # sección RETENCIONES (a veces pdfplumber junta varias secciones
            # visuales en un solo "table" lógico).
            header_row_idx = None
            for i, fila in enumerate(tabla):
                celdas = [str(c).strip() for c in fila if c]
                if any(c == "Retenciones" for c in celdas) and any("Concepto" == c for c in celdas):
                    header_row_idx = i
                    break
            if header_row_idx is None:
                continue

            for fila in tabla[header_row_idx + 1:]:
                if not fila or not fila[0]:
                    continue
                concepto = str(fila[0]).replace("\n", " ").strip()
                # Si llegamos a otra sección (p.ej. "IMPORTES TOTALES..."), cortamos
                if concepto.upper().startswith("IMPORTES TOTALES") or concepto == "":
                    break
                detalle = str(fila[1]).replace("\n", " ").strip() if len(fila) > 1 and fila[1] else ""
                importe_txt = fila[-1]
                importe = normalizar_monto(str(importe_txt)) if importe_txt else None
                if importe is None:
                    continue
                retenciones.append({"concepto": concepto, "detalle": detalle, "importe": importe})
    return retenciones


def procesar_pdf(path):
    with pdfplumber.open(path) as pdf:
        texto_completo = "\n".join(page.extract_text() or "" for page in pdf.pages)
        retenciones = extraer_retenciones(pdf)

    fila = {
        "Archivo": Path(path).name,
        "FECHA": extraer_fecha(texto_completo),
        "COE": extraer_coe(texto_completo),
    }

    for prov in PROVINCIAS_CONOCIDAS:
        fila[f"Retención {prov.title()}"] = None

    for r in retenciones:
        prov_encontrada = None
        for prov in PROVINCIAS_CONOCIDAS:
            if prov in r["detalle"].upper():
                prov_encontrada = prov
                break
        if prov_encontrada:
            col = f"Retención {prov_encontrada.title()}"
            fila[col] = (fila.get(col) or 0) + r["importe"]

    fila["Importe Neto a Pagar"] = extraer_importe_neto(texto_completo)

    return fila


def main():
    # Si se corre SIN argumentos (por ejemplo con el botón ▶ Play de VS Code,
    # o escribiendo solo "python extraer_liquidaciones.py"), usamos valores
    # por defecto: buscar los PDF en una carpeta "pdfs" al lado del script,
    # y generar el Excel como "salida.xlsx" en esa misma carpeta.
    if len(sys.argv) == 1:
        carpeta_script = Path(__file__).resolve().parent
        entradas = [str(carpeta_script / "pdfs")]
        salida = str(carpeta_script / "salida.xlsx")
        print(f"Sin argumentos: uso la carpeta '{entradas[0]}' y genero '{salida}'")
    elif len(sys.argv) >= 3:
        salida = sys.argv[-1]
        entradas = sys.argv[1:-1]
    else:
        print("Uso: python extraer_liquidaciones.py <pdf1 pdf2 ... | carpeta> <salida.xlsx>")
        print("     (o corré el script sin argumentos para usar la carpeta 'pdfs' por defecto)")
        sys.exit(1)

    pdfs = []
    for entrada in entradas:
        p = Path(entrada)
        if p.is_dir():
            encontrados = list(p.glob("*.pdf")) + list(p.glob("*.PDF"))
            # En Windows/Mac el filesystem no distingue mayúsculas, así que
            # "*.pdf" y "*.PDF" pueden devolver los mismos archivos dos veces.
            # Deduplicamos por ruta resuelta en minúsculas.
            vistos = set()
            for f in sorted(encontrados):
                clave = str(f.resolve()).lower()
                if clave not in vistos:
                    vistos.add(clave)
                    pdfs.append(f)
        elif p.suffix.lower() == ".pdf":
            pdfs.append(p)

    if not pdfs:
        print("No se encontraron PDFs.")
        sys.exit(1)

    filas = []
    for pdf_path in pdfs:
        try:
            filas.append(procesar_pdf(pdf_path))
        except Exception as e:
            filas.append({"Archivo": pdf_path.name, "ERROR": str(e)})

    df = pd.DataFrame(filas)

    # Orden de columnas prolijo
    columnas_base = ["Archivo", "FECHA", "COE"]
    columnas_prov = [f"Retención {p.title()}" for p in PROVINCIAS_CONOCIDAS]
    columnas_finales = ["Importe Neto a Pagar"]
    orden = [c for c in columnas_base + columnas_prov + columnas_finales if c in df.columns]
    otras_cols = [c for c in df.columns if c not in orden]
    df = df[orden + otras_cols]

    # Sacamos columnas de provincias que quedaron 100% vacías (no aparecieron en ningún PDF)
    for c in columnas_prov:
        if c in df.columns and df[c].isna().all():
            df = df.drop(columns=[c])

    df.to_excel(salida, index=False)
    print(f"Listo. {len(df)} liquidaciones procesadas -> {salida}")


if __name__ == "__main__":
    main()
