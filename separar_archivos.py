from PyPDF2 import PdfReader, PdfWriter
import os
import re
import sys
import shutil

def contiene_lista_asegurados(texto):
    frases_lista = ["listado de asegurados", "relación de asegurados"]
    return any(frase in texto.lower() for frase in frases_lista)

def contiene_paginacion(texto):
    patrones = [r'\b1\s*/\s*\d+\b', r'\bp[áa]gina\s*1\s*de\s*\d+\b', r'\b[Pp][áa]gina\s*1\s*/\s*\d+\b']
    return any(re.search(patron, texto.lower()) for patron in patrones)

def extraer_paginas_pdf(ruta_archivo, carpeta_salida):
    try:
        reader = PdfReader(ruta_archivo)
    except Exception as e:
        print(f" No se pudo abrir el PDF '{ruta_archivo}': {e}")
        return

    total_paginas = len(reader.pages)
    nombre_archivo = os.path.splitext(os.path.basename(ruta_archivo))[0]
    os.makedirs(carpeta_salida, exist_ok=True)

    grupos = []
    grupo_actual = []
    pagina_actual = 0

    while pagina_actual < total_paginas:
        try:
            page = reader.pages[pagina_actual]
            texto = page.extract_text() or ""
            texto = texto.strip()
        except Exception as e:
            print(f" Error al leer página {pagina_actual + 1} de '{ruta_archivo}': {e}")
            pagina_actual += 1
            continue

        if not texto or len(texto) < 10:
            if grupo_actual:
                grupo_actual.append(pagina_actual)
            pagina_actual += 1
            continue

        if contiene_lista_asegurados(texto):
            grupo_actual.append(pagina_actual)
            pagina_actual += 1
            while pagina_actual < total_paginas:
                try:
                    texto_extra = reader.pages[pagina_actual].extract_text() or ""
                    texto_extra = texto_extra.strip()
                except:
                    break

                if not texto_extra or len(texto_extra) < 10 or contiene_lista_asegurados(texto_extra):
                    grupo_actual.append(pagina_actual)
                    pagina_actual += 1
                else:
                    break
            grupos.append(grupo_actual)
            grupo_actual = []
            continue

        match = re.search(r'p[áa]gina\s*1\s*(de|/)\s*(\d+)', texto, re.IGNORECASE)
        if match:
            try:
                numero_final = int(match.group(2))
                if grupo_actual:
                    grupos.append(grupo_actual)
                    grupo_actual = []
                rango = list(range(pagina_actual, pagina_actual + numero_final))
                grupo_actual.extend(rango)
                pagina_actual += numero_final
                grupos.append(grupo_actual)
                grupo_actual = []
                continue
            except ValueError:
                pass

        if pagina_actual + 1 < total_paginas:
            try:
                texto_siguiente = reader.pages[pagina_actual + 1].extract_text() or ""
                if contiene_paginacion(texto_siguiente) and not contiene_paginacion(texto):
                    grupos.append([pagina_actual])
                    pagina_actual += 1
                    continue
            except:
                pass

        grupo_actual.append(pagina_actual)
        pagina_actual += 1

    if grupo_actual:
        grupos.append(grupo_actual)

    if len(grupos) == 1:
        idx = grupos[0][0]
        try:
            texto = reader.pages[idx].extract_text() or ""
        except:
            texto = ""
        if not texto.strip() or len(texto.strip()) < 10:
            return
        destino = os.path.join(carpeta_salida, f"{nombre_archivo}.pdf")
        shutil.copy2(ruta_archivo, destino)
        print(f" PDF completo copiado: {destino}")
    else:
        for i, grupo in enumerate(grupos):
            contenido_util = any((reader.pages[idx].extract_text() or "").strip() for idx in grupo)
            if not contenido_util:
                continue
            writer = PdfWriter()
            for idx in grupo:
                writer.add_page(reader.pages[idx])
            output_filename = os.path.join(carpeta_salida, f"{nombre_archivo}_parte_{i + 1}.pdf")
            with open(output_filename, "wb") as f_out:
                writer.write(f_out)
            print(f" PDF separado guardado: {output_filename}")

def procesar_carpeta(carpeta):
    carpeta_salida = os.path.join(carpeta, "separados")
    os.makedirs(carpeta_salida, exist_ok=True)

    for archivo in os.listdir(carpeta):
        ruta_archivo = os.path.join(carpeta, archivo)

        if os.path.isfile(ruta_archivo):
            if archivo.lower().endswith(".pdf"):
                print(f" Procesando PDF: {archivo}")
                extraer_paginas_pdf(ruta_archivo, carpeta_salida)
            else:
                try:
                    shutil.copy2(ruta_archivo, os.path.join(carpeta_salida, archivo))
                    print(f" Archivo no-PDF copiado: {archivo}")
                except Exception as e:
                    print(f" No se pudo copiar '{archivo}': {e}")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Uso: python separar_pdfs.py <ruta_carpeta>")
        sys.exit(1)

    carpeta_pdf = sys.argv[1]
    if not os.path.isdir(carpeta_pdf):
        print(" La ruta proporcionada no es una carpeta válida.")
        sys.exit(1)

    procesar_carpeta(carpeta_pdf)
    print(" Proceso de separación finalizado.")
