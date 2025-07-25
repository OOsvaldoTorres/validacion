import os
import pandas as pd
import argparse
import PyPDF2
import zipfile
import shutil
import xml.etree.ElementTree as ET

EXCEL_ESTRUCTURA = "C:/Users/otorres/Downloads/Check HS-CRS 1.xlsx"  # Usa el nuevo archivo

# Sinónimos de carpeta
SINONIMOS = {
    "caratula": ["caratula", "portada", "polizas", "poliza","caratulas"],
    "certificados": ["certificados", "certificado asegurado", "certificados asegurados"],
    "list. de aseg en excel": ["listado excel", "lista excel", "registro de asegurados excel"],
    "list. de aseg en pdf": ["listado pdf", "lista pdf", "registro de asegurados"],
    "detalle de coberturas": ["coberturas", "detalle de coberturas", "resumen de coberturas"],
    "endosos": ["endoso", "documento de endoso"],
    "recibo": ["recibo", "comprobante de pago"],
    "factura": ["factura", "facturación"],
    "comision en recibo": ["comisión recibo"],
    "comision en car": ["comisión carátula"],
    "comision por correo": ["comisión por email"],
    "observacion facturas": ["notas factura", "observaciones"]
}

# Claves por tipo de carpeta
CLAVES_VALIDAS = {
    "caratula": ["póliza", "aseguradora"],
    "certificados": ["certificado individual", "asegurado"],
    "detalle de coberturas": ["cobertura", "riesgo cubierto"],
    "endosos": ["endoso", "modificación"],
    "recibo": ["recibo", "importe", "pago"],
    "factura": ["factura", "rfc", "total"],
    "comision en recibo": ["comisión", "recibo"],
    "comision en car": ["comisión", "carátula"],
    "comision por correo": ["comisión", "correo"],
    "observacion facturas": ["observaciones", "facturas"],
    "list. de aseg en excel": ["excel", "listado"],
    "list. de aseg en pdf": [ "listado", "registro de asegurados"]
}

def descomprimir_zip(ruta_zip, ruta_destino):
    try:
        # Verificar si la ruta de destino existe, si no, crearla
        if not os.path.exists(ruta_destino):
            os.makedirs(ruta_destino)

        # Descomprimir el archivo zip
        with zipfile.ZipFile(ruta_zip, 'r') as archivo_zip:
            archivo_zip.extractall(ruta_destino)
        print(f"✅ Archivos descomprimidos en {ruta_destino}")
    except Exception as e:
        print(f"⚠️ Error al descomprimir el archivo: {e}")

def mover_archivo_a_carpeta(ruta_archivo, ruta_destino, nombre_carpeta):
    # Crear la carpeta si no existe
    carpeta_destino = os.path.join(ruta_destino, nombre_carpeta)
    if not os.path.exists(carpeta_destino):
        os.makedirs(carpeta_destino)
        print(f"✅ Carpeta '{nombre_carpeta}' creada en {ruta_destino} ")

    # Mover el archivo a la carpeta correspondiente
    nombre_archivo = os.path.basename(ruta_archivo)
    destino_archivo = os.path.join(carpeta_destino, nombre_archivo)

    # Mover el archivo a la nueva carpeta
    shutil.move(ruta_archivo, destino_archivo)
    print(f"✅ Archivo '{nombre_archivo}' movido a '{carpeta_destino}'")

def leer_archivos_xml(ruta_carpeta):
    xml_files = [f for f in os.listdir(ruta_carpeta) if f.lower().endswith('.xml')]
    if not xml_files:
        print("⚠️ No se encontraron archivos XML.")
        return

    for archivo_xml in xml_files:
        try:
            tree = ET.parse(os.path.join(ruta_carpeta, archivo_xml))
            root = tree.getroot()

            # Aquí puedes procesar el contenido del XML
            print(f"📄 Leyendo archivo XML: {archivo_xml}")
            for elem in root.iter():
                print(f"   {elem.tag}: {elem.text}")
        except Exception as e:
            print(f"⚠️ Error al leer el archivo XML '{archivo_xml}': {e}")

def cargar_estructura_nueva(ramo_buscado):
    try:
        df = pd.read_excel(EXCEL_ESTRUCTURA, header=1)  # fila 2 es el header real

        carpetas = list(df.columns[1:])
        estructura = {}
        estructura_restante = set()
        estructura_obligatoria = set()

        for _, row in df.iterrows():
            ramo = str(row[df.columns[0]]).strip().lower()
            if ramo != ramo_buscado.strip().lower():
                continue

            for carpeta in carpetas:
                valor = str(row[carpeta]).strip().lower() if pd.notna(row[carpeta]) else ""

                if valor == "si":
                    estructura_obligatoria.add(carpeta.strip().lower())
                else:
                    aseguradoras = [a.strip() for a in valor.split(",") if a.strip()]
                    for aseguradora in aseguradoras:
                        if aseguradora == "aseguradoras restantes":
                            estructura_restante.add(carpeta.strip().lower())
                        else:
                            if aseguradora not in estructura:
                                estructura[aseguradora] = set()
                            estructura[aseguradora].add(carpeta.strip().lower())

        if estructura_restante:
            estructura["__restantes__"] = estructura_restante
        if estructura_obligatoria:
            estructura["__obligatorias__"] = estructura_obligatoria

        return estructura
    except Exception as e:
        print(f"❌ Error al leer estructura del Excel: {e}")
        return {}

def encontrar_nombre_estandar(nombre_real):
    nombre_real = nombre_real.strip().lower()
    for clave, sinonimos in SINONIMOS.items():
        if nombre_real == clave or nombre_real in sinonimos:
            return clave
    return None

def validar_archivo_excel(archivo_path, claves):
    try:
        # Cargar archivo Excel
        df = pd.read_excel(archivo_path)
        for clave in claves:
            if df.apply(lambda row: row.astype(str).str.contains(clave, case=False).any(), axis=1).any():
                print(f"   ✅ '{os.path.basename(archivo_path)}' contiene clave válida.")
                return True
        print(f"   ⚠️ '{os.path.basename(archivo_path)}' NO contiene clave esperada.")
        return False
    except Exception as e:
        print(f"⚠️ Error al procesar Excel '{archivo_path}': {e}")
        return False

def validar_estructura(ramo, aseguradora, ruta_base):
    estructura = cargar_estructura_nueva(ramo)
    aseguradora = aseguradora.strip().lower()

    carpetas_esperadas = set()

    # Agregar carpetas obligatorias
    if "__obligatorias__" in estructura:
        carpetas_esperadas.update(estructura["__obligatorias__"])

    # Agregar carpetas específicas de la aseguradora o las de "restantes"
    if aseguradora in estructura:
        carpetas_esperadas.update(estructura[aseguradora])
    elif "__restantes__" in estructura:
        print(f"⚠️ Aseguradora '{aseguradora}' no encontrada. Usando estructura de 'aseguradoras restantes'.")
        carpetas_esperadas.update(estructura["__restantes__"])
    else:
        print(f"⚠️ No se encontró '{aseguradora}' ni estructura para aseguradoras restantes en el Excel.")
        return

    print("\n📋 Carpetas esperadas según el Excel:")
    for carpeta in sorted(carpetas_esperadas):
        print(f"   - {carpeta}")

    carpetas_reales = {
        carpeta.lower(): carpeta
        for carpeta in os.listdir(ruta_base)
        if os.path.isdir(os.path.join(ruta_base, carpeta))
    }

    carpetas_detectadas = set()
    carpetas_no_validas = []

    for carpeta_real in carpetas_reales:
        nombre_estandar = encontrar_nombre_estandar(carpeta_real)
        if nombre_estandar:
            carpetas_detectadas.add(nombre_estandar)
            validar_archivos(os.path.join(ruta_base, carpetas_reales[carpeta_real]), nombre_estandar)
        else:
            carpetas_no_validas.append(carpeta_real)

    faltantes = carpetas_esperadas - carpetas_detectadas

    if faltantes:
        print("\n❌ Carpetas faltantes:")
        for carpeta in sorted(faltantes):
            print(f"   ❌ {carpeta}")
        print("\n❌ ❌ La estructura NO es correcta.")
    else:
        print("\n✅ La estructura es correcta. No faltan carpetas obligatorias.")

    if carpetas_no_validas:
        print("\n⚠️ Carpetas que no coinciden con ninguna esperada:")
        for c in carpetas_no_validas:
            print(f"   ⚠️ {c}")

def extraer_texto_pdf(pdf_path):
    try:
        with open(pdf_path, "rb") as archivo:
            reader = PyPDF2.PdfReader(archivo)
            texto = ""
            for pagina in reader.pages:
                texto += pagina.extract_text() or ""
        return texto.lower()
    except Exception as e:
        print(f"⚠️ Error al leer PDF '{pdf_path}': {e}")
        return ""

def validar_archivos(ruta_carpeta, tipo_carpeta):
    print(f"\n📂 Validando archivos en carpeta: {os.path.basename(ruta_carpeta)}")
    claves = CLAVES_VALIDAS.get(tipo_carpeta, [])
    encontrados = 0

    for archivo in os.listdir(ruta_carpeta):
        archivo_path = os.path.join(ruta_carpeta, archivo)

        if archivo.lower().endswith(".pdf"):
            texto = extraer_texto_pdf(archivo_path)
            if any(clave in texto for clave in claves):
                print(f"   ✅ '{archivo}' contiene clave válida.")
                encontrados += 1
            else:
                print(f"   ⚠️ '{archivo}' NO contiene clave esperada.")
        
        elif archivo.lower().endswith(".xlsx") or archivo.lower().endswith(".xls"):
            if validar_archivo_excel(archivo_path, claves):
                encontrados += 1

        elif archivo.lower().endswith(".xml"):
            leer_archivos_xml(ruta_carpeta)

    if encontrados == 0:
        print("   ⚠️ Ningún archivo válido detectado.")


# MAIN
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Valida estructura y contenido por ramo y aseguradora.")
    parser.add_argument("ramo", type=str, help="Nombre del ramo (área)")
    parser.add_argument("aseguradora", type=str, help="Nombre de la aseguradora")
    parser.add_argument("ruta_base", type=str, help="Ruta base de las carpetas a validar")
    args = parser.parse_args()

    validar_estructura(args.ramo, args.aseguradora, args.ruta_base)
