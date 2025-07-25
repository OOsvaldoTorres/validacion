import os
import pandas as pd
import argparse
import PyPDF2

EXCEL_ESTRUCTURA = "C:/Users/otorres/Downloads/checklist_junio.xlsx"

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

def cargar_requisitos_excel(ramo, aseguradora):
    try:
        df = pd.read_excel(EXCEL_ESTRUCTURA)

        df_filtrado = df[
            (df["ramo"].str.lower() == ramo.lower()) &
            (df["aseguradora"].str.lower() == aseguradora.lower())
        ]

        requisitos = []
        for _, row in df_filtrado.iterrows():
            nombre = str(row["nombre"]).strip()
            clave = str(row["clave"]).strip().lower() if pd.notna(row["clave"]) else ""
            requisitos.append({"nombre": nombre, "clave": clave})

        return requisitos
    except Exception as e:
        print(f"❌ Error leyendo el Excel: {e}")
        return []

def validar_archivos(ruta_base, requisitos):
    archivos_en_carpeta = os.listdir(ruta_base)
    archivos_en_carpeta = [
        f for f in archivos_en_carpeta if os.path.isfile(os.path.join(ruta_base, f))
    ]

    encontrados = []
    faltantes = []

    for req in requisitos:
        nombre_req = req["nombre"]
        clave_req = req["clave"]

        encontrado = False

        for archivo in archivos_en_carpeta:
            nombre_archivo = archivo.lower()
            ruta_archivo = os.path.join(ruta_base, archivo)

            # Solo PDF por ahora
            if not nombre_archivo.endswith(".pdf"):
                continue

            texto_pdf = extraer_texto_pdf(ruta_archivo)

            # Casos según nombre_req
            if nombre_req == "**" or "*" in nombre_req:
                # Solo validar contenido
                if clave_req in texto_pdf:
                    encontrado = True
                    break

            else:
                # Validar que nombre contenga texto Y contenido tenga clave
                if nombre_req.lower() in nombre_archivo and clave_req in texto_pdf:
                    encontrado = True
                    break

        if encontrado:
            encontrados.append(nombre_req)
        else:
            faltantes.append(nombre_req)

    print("\n✅ Archivos encontrados:")
    for e in encontrados:
        print(f"   ✅ {e}")

    if faltantes:
        print("\n❌ Archivos faltantes:")
        for f in faltantes:
            print(f"   ❌ {f}")
    else:
        print("\n🎉 Todos los archivos requeridos fueron encontrados correctamente.")

# MAIN
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Valida archivos requeridos por ramo y aseguradora.")
    parser.add_argument("ramo", type=str, help="Nombre del ramo")
    parser.add_argument("aseguradora", type=str, help="Nombre de la aseguradora")
    parser.add_argument("ruta_base", type=str, help="Ruta de la carpeta a validar")
    args = parser.parse_args()

    requisitos = cargar_requisitos_excel(args.ramo, args.aseguradora)
    if not requisitos:
        print("⚠️ No se encontraron requisitos para esos parámetros.")
    else:
        validar_archivos(args.ruta_base, requisitos)
