"""
================================================================================
 EXTRACTOR DE LIQUIDACIONES DE GRANOS - interfaz gráfica
================================================================================

Este archivo está pensado para servir como PLANTILLA / MOLDE de futuros
programas parecidos: "elegir uno o varios PDF -> sacar ciertos datos -> volcar
todo a un Excel". Por eso está dividido en secciones bien marcadas, y comentado
en detalle para que sea fácil encontrar qué tocar cuando quieras adaptarlo a
otro tipo de documento.

Para reutilizar esta plantilla en un proyecto nuevo, básicamente hay que tocar
SOLO la SECCIÓN 2 (la lógica de extracción). Las secciones 1, 3 y 4 son
genéricas y sirven casi sin cambios para cualquier "PDFs -> Excel".

--------------------------------------------------------------------------------

"""

import os
import re
import sys
import threading
import traceback
from pathlib import Path

import pdfplumber
import pandas as pd

import tkinter as tk
from tkinter import filedialog, messagebox, ttk


# ================================================================================
# SECCIÓN 0: CONFIGURACIÓN GENERAL
# --------------------------------------------------------------------------------
# Constantes que definen "look & feel" y textos de la ventana. Cambiar esto para
# adaptar la plantilla a otro programa (otro título, otro logo, otros colores).
# ================================================================================

TITULO_VENTANA = "Extractor de Liquidaciones de Granos"
SUBTITULO_VENTANA = "Elegí tus PDF, apretá Generar Excel, listo."

# Nombre del archivo de imagen que se muestra centrado arriba de la ventana.
# Tiene que ser un PNG con fondo transparente (para que se vea "flotando"
# sobre el fondo de la ventana en vez de tener un cuadrado blanco alrededor).
# Si el archivo no existe, la ventana igual funciona (simplemente no muestra
# el logo) -- así no se rompe nada si te olvidás de ponerlo.
NOMBRE_ARCHIVO_LOGO = "logo.png"

# Ancho máximo (en píxeles) al que se reduce el logo si es muy grande, para
# que no ocupe toda la ventana. Ajustá este número a gusto.
ANCHO_MAXIMO_LOGO = 180

COLOR_BOTON_PRINCIPAL = "#2e7d32"   # verde
COLOR_TEXTO_BOTON_PRINCIPAL = "white"


def ruta_recurso(nombre_archivo):
    """
    Devuelve la ruta correcta a un archivo "extra" (como el logo) tanto si:
      a) estás corriendo este archivo directamente con "python archivo.py", o
      b) ya lo empaquetaste con PyInstaller en un .exe.

    ESTO ES IMPORTANTE PARA LA PLANTILLA: cuando PyInstaller arma el .exe con
    la opción --onefile, en tiempo de ejecución descomprime los archivos
    extra (como logo.png) en una carpeta temporal, cuya ruta guarda en
    "sys._MEIPASS". Si no usás esta función y simplemente escribís
    "logo.png" a secas, funciona al correr el .py pero el .exe empaquetado
    NO va a encontrar el archivo.

    Usalo así en cualquier proyecto nuevo: ruta_recurso("mi_imagen.png")
    """
    base_path = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, nombre_archivo)


# ================================================================================
# SECCIÓN 1: HELPERS GENÉRICOS (montos, fechas, etc.)
# --------------------------------------------------------------------------------
# Funciones chiquitas de propósito general. Suelen servir tal cual para
# cualquier otro extractor de PDF que trabaje con montos/fechas en español.
# ================================================================================

def normalizar_monto(texto):
    """Convierte un texto tipo '$ 1.234,56' / '1,234.56' / '$1234.56' a float."""
    if texto is None:
        return None
    t = texto.replace("$", "").replace("\n", " ").strip()
    t = t.replace(" ", "")
    if not t:
        return None
    t = t.replace(",", "")  # formato con coma de miles y punto decimal
    try:
        return float(t)
    except ValueError:
        return None


# ================================================================================
# SECCIÓN 2: LÓGICA DE EXTRACCIÓN (¡ACÁ ES DONDE SE ADAPTA A CADA PDF DISTINTO!)
# --------------------------------------------------------------------------------
# Esta es la ÚNICA sección que hay que reescribir si querés usar esta plantilla
# para extraer datos de OTRO tipo de documento (una factura, un remito, etc.).
# La receta es siempre la misma:
#
#   1. Definir funciones "extraer_XXX(texto_completo)" que, con una expresión
#      regular (re.search), encuentran un dato puntual dentro del texto plano
#      del PDF.
#   2. Una función "procesar_pdf(path)" que abre UN pdf, saca todo el texto,
#      llama a las funciones de extracción, y arma un diccionario con los
#      resultados (una fila del futuro Excel).
#   3. Una función "procesar_todos(rutas_pdf)" que llama a procesar_pdf() para
#      cada archivo y junta todo en una tabla (DataFrame de pandas).
#
# El resto del programa (la ventana, los botones, etc.) NO necesita saber nada
# de liquidaciones de granos: solo le pide a esta sección "dame una tabla a
# partir de esta lista de PDFs", así que podés reemplazar todo este bloque por
# la lógica de otro documento sin tocar el resto del código.
# ================================================================================

# Provincias que reconocemos como columna propia en el Excel final.
# (Ejemplo de "constante de configuración" típica de este tipo de proyectos:
# para otro documento, acá irían las categorías/campos que quieras separar en
# columnas propias.)
PROVINCIAS_CONOCIDAS = ["LA PAMPA", "BUENOS AIRES", "SANTA FE", "CORDOBA", "CÓRDOBA"]


def extraer_fecha(texto_completo):
    """Fecha de emisión del comprobante: 'DD/MM/AAAA, LOCALIDAD' (encabezado)."""
    m = re.search(r"(\d{2}/\d{2}/\d{4}),\s*[A-ZÁÉÍÓÚÑ ]+", texto_completo)
    if m:
        return m.group(1)
    # Alternativa (fallback) por si el encabezado no matchea: campo "Fecha:"
    m = re.search(r"Fecha:\s*(\d{2}/\d{2}/\d{4})", texto_completo)
    if m:
        return m.group(1)
    return None


def extraer_coe(texto_completo):
    """Número de C.O.E. (identificador único de la liquidación en AFIP)."""
    m = re.search(r"C\.O\.E\.:\s*(\d+)", texto_completo)
    return m.group(1) if m else None


def extraer_importe_neto(texto_completo):
    """
    Toma el valor del campo "Importe Neto a Pagar" tal cual aparece en el PDF.
    En los documentos de "Ajuste Unificado" este campo puede aparecer dos
    veces (Débito y luego Crédito): se usa SIEMPRE la PRIMERA ocurrencia (la
    de Débito), sea cual sea su valor -- incluso si es $0.
    """
    m = re.search(r"Importe Neto a Pagar:?\s*\$?\s*([\d.,]+)", texto_completo)
    if m:
        return normalizar_monto(m.group(1))
    return None


def extraer_retenciones(pdf):
    """
    Recorre TODAS las páginas del PDF buscando la tabla "RETENCIONES" y
    devuelve una lista de diccionarios: [{concepto, detalle, importe}, ...]

    Nota técnica para la plantilla: acá usamos pdf.extract_tables() de
    pdfplumber en vez de expresiones regulares sobre texto plano, porque el
    dato que buscamos está organizado en una tabla real dentro del PDF (con
    columnas). Para datos "sueltos" dentro de un párrafo, las expresiones
    regulares (como en las funciones de arriba) suelen ser más simples.
    """
    retenciones = []
    for page in pdf.pages:
        for tabla in page.extract_tables():
            # Buscamos, en CUALQUIER fila de la tabla, el encabezado de la
            # sección RETENCIONES (a veces pdfplumber junta varias secciones
            # visuales del PDF en un solo "table" lógico).
            header_row_idx = None
            for i, fila in enumerate(tabla):
                celdas = [str(c).strip() for c in fila if c]
                if any(c == "Retenciones" for c in celdas) and any("Concepto" == c for c in celdas):
                    header_row_idx = i
                    break
            if header_row_idx is None:
                continue  # esta tabla no es la de retenciones, seguimos buscando

            for fila in tabla[header_row_idx + 1:]:
                if not fila or not fila[0]:
                    continue
                concepto = str(fila[0]).replace("\n", " ").strip()
                # Si llegamos a otra sección del PDF, cortamos el recorrido
                if concepto.upper().startswith("IMPORTES TOTALES") or concepto == "":
                    break
                detalle = str(fila[1]).replace("\n", " ").strip() if len(fila) > 1 and fila[1] else ""
                importe_txt = fila[-1]  # el importe siempre es la última columna
                importe = normalizar_monto(str(importe_txt)) if importe_txt else None
                if importe is None:
                    continue
                retenciones.append({"concepto": concepto, "detalle": detalle, "importe": importe})
    return retenciones


def procesar_pdf(path):
    """
    Procesa UN solo PDF y devuelve un diccionario con los datos extraídos
    (una futura fila del Excel). Esta es la función "orquestadora": abre el
    PDF, saca el texto, y llama a cada función extraer_XXX de arriba.
    """
    with pdfplumber.open(path) as pdf:
        texto_completo = "\n".join(page.extract_text() or "" for page in pdf.pages)
        retenciones = extraer_retenciones(pdf)

    fila = {
        "Archivo": Path(path).name,
        "FECHA": extraer_fecha(texto_completo),
        "COE": extraer_coe(texto_completo),
    }

    # Creamos una columna por cada provincia conocida, arrancando en None
    for prov in PROVINCIAS_CONOCIDAS:
        fila[f"Retención {prov.title()}"] = None

    # Sumamos cada retención encontrada a la columna de su provincia
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


def procesar_todos(rutas_pdf):
    """
    Procesa una LISTA de PDFs y devuelve una tabla (DataFrame de pandas) lista
    para volcar a Excel, con las columnas ordenadas de forma prolija.

    Esta función es la que llama la interfaz gráfica (Sección 3) -- es el
    "puente" entre la lógica de extracción y la ventana.
    """
    filas = []
    for p in rutas_pdf:
        try:
            filas.append(procesar_pdf(p))
        except Exception as e:
            # Si UN PDF falla (por ejemplo, tiene un formato raro), no
            # frenamos todo el proceso: anotamos el error en esa fila y
            # seguimos con el resto.
            filas.append({"Archivo": Path(p).name, "ERROR": str(e)})

    df = pd.DataFrame(filas)

    # --- Orden prolijo de columnas ---
    columnas_base = ["Archivo", "FECHA", "COE"]
    columnas_prov = [f"Retención {p.title()}" for p in PROVINCIAS_CONOCIDAS]
    columnas_finales = ["Importe Neto a Pagar"]
    orden = [c for c in columnas_base + columnas_prov + columnas_finales if c in df.columns]
    otras_cols = [c for c in df.columns if c not in orden]
    df = df[orden + otras_cols]

    # Sacamos columnas de provincias que quedaron 100% vacías (no aparecieron
    # en ningún PDF de esta tanda)
    for c in columnas_prov:
        if c in df.columns and df[c].isna().all():
            df = df.drop(columns=[c])

    return df


# ================================================================================
# SECCIÓN 3: INTERFAZ GRÁFICA (la ventanita)
# --------------------------------------------------------------------------------
# Esta sección es GENÉRICA: no sabe nada de liquidaciones de granos. Solo
# sabe pedirle a "procesar_todos()" (Sección 2) que convierta una lista de
# PDFs en una tabla, y mostrar el resultado. Para reutilizar la plantilla en
# otro proyecto, en general esta sección se puede dejar CASI IGUAL.
# ================================================================================

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(TITULO_VENTANA)
        self.geometry("640x560")
        self.resizable(True, True)

        self.rutas_pdf = []  # lista de rutas a los PDF que el usuario eligió

        # ---------------------------------------------------------------
        # LOGO (imagen centrada arriba de la ventana, con fondo transparente)
        # ---------------------------------------------------------------
        # tk.PhotoImage soporta PNG con transparencia de forma nativa desde
        # Tcl/Tk 8.6 (que es lo que trae Python moderno), así que NO hace
        # falta instalar Pillow solo para esto.
        #
        # Guardamos la imagen en self._logo_img (con guión bajo) para que
        # Python no la "tire a la basura" apenas termina el __init__ -- es un
        # detalle de tkinter: si no guardás una referencia a la imagen en
        # algún lado que siga vivo, la imagen desaparece de la pantalla sola.
        self._logo_img = self._cargar_logo()
        if self._logo_img is not None:
            logo_label = tk.Label(self, image=self._logo_img)
            logo_label.pack(pady=(15, 5))

        # --- Título y subtítulo ---
        titulo = tk.Label(self, text=TITULO_VENTANA, font=("Segoe UI", 14, "bold"))
        titulo.pack(pady=(5 if self._logo_img else 15, 0))
        subtitulo = tk.Label(self, text=SUBTITULO_VENTANA, font=("Segoe UI", 10), fg="#555555")
        subtitulo.pack(pady=(0, 10))

        # --- Botones para elegir archivos ---
        frame_botones = tk.Frame(self)
        frame_botones.pack(pady=5)

        btn_archivos = tk.Button(frame_botones, text="📄 Elegir PDF sueltos...",
                                  width=22, height=2, command=self.elegir_archivos)
        btn_archivos.grid(row=0, column=0, padx=8)

        btn_carpeta = tk.Button(frame_botones, text="📁 Elegir una carpeta...",
                                 width=22, height=2, command=self.elegir_carpeta)
        btn_carpeta.grid(row=0, column=1, padx=8)

        # --- Lista de archivos elegidos ---
        frame_lista = tk.Frame(self)
        frame_lista.pack(pady=10, fill="both", expand=True, padx=20)

        tk.Label(frame_lista, text="PDF seleccionados:", anchor="w").pack(fill="x")

        scrollbar = tk.Scrollbar(frame_lista)
        scrollbar.pack(side="right", fill="y")

        self.lista = tk.Listbox(frame_lista, yscrollcommand=scrollbar.set, height=8)
        self.lista.pack(fill="both", expand=True)
        scrollbar.config(command=self.lista.yview)

        btn_limpiar = tk.Button(self, text="Limpiar selección", command=self.limpiar)
        btn_limpiar.pack(pady=(0, 5))

        # --- Botón principal ---
        self.btn_generar = tk.Button(
            self, text="⚙️  Generar Excel", font=("Segoe UI", 12, "bold"),
            bg=COLOR_BOTON_PRINCIPAL, fg=COLOR_TEXTO_BOTON_PRINCIPAL, height=2,
            command=self.generar_excel
        )
        self.btn_generar.pack(pady=10, fill="x", padx=40)

        # --- Barra de progreso y estado ---
        self.progreso = ttk.Progressbar(self, mode="indeterminate")
        self.progreso.pack(fill="x", padx=40, pady=(0, 5))

        self.estado = tk.Label(self, text="", fg="#333333")
        self.estado.pack(pady=(0, 10))

    # -----------------------------------------------------------------
    # Carga del logo. Separado en su propio método para que sea fácil de
    # copiar/pegar a otro proyecto tal cual.
    # -----------------------------------------------------------------
    def _cargar_logo(self):
        """
        Intenta cargar NOMBRE_ARCHIVO_LOGO como imagen para mostrar en la
        ventana. Si el archivo no existe o falla la carga, devuelve None
        (la ventana sigue funcionando igual, simplemente sin logo).
        """
        ruta = ruta_recurso(NOMBRE_ARCHIVO_LOGO)
        if not os.path.exists(ruta):
            return None
        try:
            img = tk.PhotoImage(file=ruta)
            # Si la imagen es más ancha que ANCHO_MAXIMO_LOGO, la reducimos.
            # tk.PhotoImage solo permite reducir en números enteros (subsample),
            # así que calculamos el factor entero más cercano.
            if img.width() > ANCHO_MAXIMO_LOGO:
                factor = max(1, img.width() // ANCHO_MAXIMO_LOGO)
                img = img.subsample(factor, factor)
            return img
        except tk.TclError:
            # Archivo corrupto, formato no soportado, etc. -- seguimos sin logo.
            return None

    # -----------------------------------------------------------------
    # Manejo de selección de archivos
    # -----------------------------------------------------------------
    def elegir_archivos(self):
        archivos = filedialog.askopenfilenames(
            title="Elegí uno o más PDF",
            filetypes=[("Archivos PDF", "*.pdf *.PDF")]
        )
        if archivos:
            for a in archivos:
                if a not in self.rutas_pdf:
                    self.rutas_pdf.append(a)
            self.actualizar_lista()

    def elegir_carpeta(self):
        carpeta = filedialog.askdirectory(title="Elegí la carpeta con tus PDF")
        if carpeta:
            encontrados = list(Path(carpeta).glob("*.pdf")) + list(Path(carpeta).glob("*.PDF"))
            vistos = {str(Path(p).resolve()).lower() for p in self.rutas_pdf}
            for f in sorted(encontrados):
                clave = str(f.resolve()).lower()
                if clave not in vistos:
                    self.rutas_pdf.append(str(f))
                    vistos.add(clave)
            self.actualizar_lista()

    def actualizar_lista(self):
        self.lista.delete(0, tk.END)
        for p in self.rutas_pdf:
            self.lista.insert(tk.END, Path(p).name)
        self.estado.config(text=f"{len(self.rutas_pdf)} PDF listo(s) para procesar.")

    def limpiar(self):
        self.rutas_pdf = []
        self.actualizar_lista()
        self.estado.config(text="")

    # -----------------------------------------------------------------
    # Generación del Excel
    # -----------------------------------------------------------------
    def generar_excel(self):
        if not self.rutas_pdf:
            messagebox.showwarning("Sin PDF", "Primero elegí al menos un PDF (o una carpeta).")
            return

        destino = filedialog.asksaveasfilename(
            title="Guardar Excel como...",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile="liquidaciones.xlsx"
        )
        if not destino:
            return

        self.btn_generar.config(state="disabled")
        self.estado.config(text="Procesando, un momento...")
        self.progreso.start(12)

        # Procesamos en un hilo (thread) separado para que la ventana NO se
        # quede "congelada" (sin responder) mientras se procesan los PDF.
        # Esto es importante en cualquier programa con interfaz gráfica que
        # haga un trabajo que tarde más de una fracción de segundo.
        hilo = threading.Thread(target=self._procesar_en_segundo_plano, args=(destino,))
        hilo.start()

    def _procesar_en_segundo_plano(self, destino):
        """
        Corre en un hilo aparte. IMPORTANTE: nunca hay que tocar elementos de
        la ventana (labels, botones, etc.) directamente desde acá -- hay que
        usar self.after(0, funcion, argumentos) para pedirle a tkinter que
        actualice la interfaz de forma segura desde el hilo principal.
        """
        try:
            df = procesar_todos(self.rutas_pdf)
            df.to_excel(destino, index=False)
            self.after(0, self._terminar_ok, destino, len(df))
        except Exception:
            error_texto = traceback.format_exc()
            self.after(0, self._terminar_error, error_texto)

    def _terminar_ok(self, destino, cantidad):
        self.progreso.stop()
        self.btn_generar.config(state="normal")
        self.estado.config(text=f"Listo: {cantidad} liquidaciones procesadas.")
        respuesta = messagebox.askyesno(
            "Listo",
            f"Se generó el Excel con {cantidad} liquidaciones:\n{destino}\n\n¿Querés abrirlo ahora?"
        )
        if respuesta:
            os.startfile(destino)  # abre el archivo con el programa por defecto (Windows)

    def _terminar_error(self, error_texto):
        self.progreso.stop()
        self.btn_generar.config(state="normal")
        self.estado.config(text="Ocurrió un error.")
        messagebox.showerror("Error", f"Algo falló al procesar:\n\n{error_texto}")


# ================================================================================
# SECCIÓN 4: PUNTO DE ENTRADA
# --------------------------------------------------------------------------------
# Esto es lo que se ejecuta cuando corrés "python extraer_liquidaciones_gui.py"
# o hacés doble clic en el .exe. Prácticamente nunca hace falta tocar esto.
# ================================================================================

if __name__ == "__main__":
    app = App()
    app.mainloop()
