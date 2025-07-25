import os
import re
import shutil
import zipfile
import rarfile
import py7zr
from pathlib import Path
import extract_msg
from bs4 import BeautifulSoup
import pandas as pd

# Verifica si es archivo comprimido
def es_comprimido(nombre):
    return nombre.lower().endswith(('.zip', '.rar', '.7z'))

# Verifica si es archivo de correo Outlook
def es_correo(nombre):
    return nombre.lower().endswith('.msg')

# Extrae cualquier archivo comprimido
def extraer_archivo(archivo, destino):
    ext = archivo.lower()
    os.makedirs(destino, exist_ok=True)

    try:
        if ext.endswith('.zip'):
            with zipfile.ZipFile(archivo, 'r') as z:
                z.extractall(destino)
        elif ext.endswith('.rar'):
            with rarfile.RarFile(archivo) as rf:
                rf.extractall(destino)
        elif ext.endswith('.7z'):
            with py7zr.SevenZipFile(archivo, mode='r') as sz:
                sz.extractall(destino)
        else:
            print(f" Tipo no soportado: {archivo}")
            return False

        descomprimir_recursivamente(destino)  # Recursivo
        print(f" Descomprimido: {archivo} → {destino}")
        return True

    except Exception as e:
        print(f" Error al descomprimir {archivo}: {e}")
        return False

# Descomprime recursivamente todo
def descomprimir_recursivamente(ruta):
    for root, _, files in os.walk(ruta):
        for file in files:
            path_archivo = os.path.join(root, file)
            if es_comprimido(file):
                carpeta_destino = os.path.join(root, Path(file).stem)
                extraer_archivo(path_archivo, carpeta_destino)
            elif es_correo(file):
                procesar_correo(path_archivo, os.path.join(ruta, "correos"))

# Procesa .msg y guarda adjuntos + tabla como CSV
def procesar_correo(ruta_msg, carpeta_correos):
    os.makedirs(carpeta_correos, exist_ok=True)
    try:
        msg = extract_msg.Message(ruta_msg)
        msg_subject = Path(ruta_msg).stem.replace(" ", "_")
        correo_dir = os.path.join(carpeta_correos, msg_subject)
        os.makedirs(correo_dir, exist_ok=True)

        # Extraer adjuntos
        msg.extract_attachments(correo_dir)
        print(f" Adjuntos extraídos de: {ruta_msg}")

        # Analizar cuerpo del correo
        html = msg.htmlBody or ""
        if "<table" in html:
            soup = BeautifulSoup(html, "html.parser")
            tabla = soup.find("table")
            if tabla:
                df = pd.read_html(str(tabla))[0]
                # Buscar "comisión del xx%"
                cuerpo = soup.get_text()
                comision = buscar_comision(cuerpo)
                df["comision"] = comision
                df.to_csv(os.path.join(correo_dir, "contenido_correo.csv"), index=False)
                print(f"📝 Tabla y comisión extraídas de: {ruta_msg}")
    except Exception as e:
        print(f" Error al procesar correo {ruta_msg}: {e}")

# Extrae "comisión del x%" como número
def buscar_comision(texto):
    match = re.search(r"comisi[oó]n del (\d{1,2}(?:[\.,]\d{1,2})?)\s*%", texto.lower())
    if match:
        return match.group(1).replace(",", ".")
    return None

# Copia todos los archivos planos a entrega/
def mover_archivos_a_entrega(ruta_base):
    carpeta_entrega = os.path.join(ruta_base, "entrega")
    os.makedirs(carpeta_entrega, exist_ok=True)

    for root, _, files in os.walk(ruta_base):
        if carpeta_entrega in root:
            continue
        for file in files:
            if es_comprimido(file) or es_correo(file):
                continue
            origen = os.path.join(root, file)
            destino = os.path.join(carpeta_entrega, file)
            if not os.path.exists(destino):
                shutil.copy2(origen, destino)
                print(f" Copiado a entrega: {origen}")
            else:
                print(f" Ya existe (omitido): {destino}")

# ======================
# 🚀 Ejecución principal
# ======================
if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print(" Uso: python descomprimir_y_procesar.py /ruta/carpeta")
        sys.exit(1)

    ruta = sys.argv[1]
    print(ruta)
    if not os.path.isdir(ruta):
        print("Ruta inválida.", ruta)
        sys.exit(1)

    print(" Descomprimiendo y procesando...")
    descomprimir_recursivamente(ruta)

    print(" Copiando archivos a carpeta 'entrega'...")
    mover_archivos_a_entrega(ruta)

    print(" Proceso completo.")

