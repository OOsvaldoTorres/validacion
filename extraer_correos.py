import extract_msg
from bs4 import BeautifulSoup
import pandas as pd
import os
import re
from io import StringIO


ruta_msg = "C:/Users/Oscar Torres/Downloads/aseguradoras/ABRIL POLIZAS/ASOCIACION NECROLOGICA MEXICANA SA DE CV/Nueva carpeta/Archivos - 2025-02-28T092509.231/CFDI DE ANEMEX Factura  31359.msg"  # Cambia esta ruta
carpeta_correos= "C:/Users/Oscar Torres/Downloads/aseguradoras/ABRIL POLIZAS/ASOCIACION NECROLOGICA MEXICANA SA DE CV/Nueva carpeta"
msg = extract_msg.Message(ruta_msg)
html = msg.htmlBody or msg.body or ""

soup = BeautifulSoup(html, "html.parser")
tabla = soup.find("table")



def procesar_correo(ruta_msg, carpeta_correos):
    os.makedirs(carpeta_correos, exist_ok=True)

    try:
        msg = extract_msg.Message(ruta_msg)
        html = msg.htmlBody or msg.body or ""

        # Crear subcarpeta para guardar adjuntos y CSV
        nombre_base = os.path.splitext(os.path.basename(ruta_msg))[0].replace(" ", "_")
        carpeta_correo = os.path.join(carpeta_correos, nombre_base)
        os.makedirs(carpeta_correo, exist_ok=True)

        # 💾 Guardar adjuntos manualmente si no existe extract_attachments
        if hasattr(msg, "extract_attachments"):
            msg.extract_attachments(carpeta_correo)
        else:
            for att in msg.attachments:
                filename = att.longFilename or att.shortFilename or "adjunto.bin"
                with open(os.path.join(carpeta_correo, filename), "wb") as f:
                    f.write(att.data)
        print(f"📬 Adjuntos extraídos de: {ruta_msg}")

        # Extraer tabla HTML si la hay
        soup = BeautifulSoup(html, "html.parser")
        tabla = soup.find("table")
        texto_completo = soup.get_text()

        # Buscar comisión del correo
        match = re.search(r"comisi[oó]n del (\d{1,2}(?:[\.,]\d{1,2})?)\s*%", texto_completo.lower())
        comision = match.group(1).replace(",", ".") if match else None

        if tabla:
            try:
                df = pd.read_html(StringIO(str(tabla)))[0]
                df["comision"] = comision
                ruta_csv = os.path.join(carpeta_correo, "contenido_correo.csv")
                df.to_csv(ruta_csv, index=False)
                print(f"✅ Tabla y comisión guardadas en: {ruta_csv}")
            except Exception as e:
                print(f"⚠️ No se pudo procesar la tabla del correo {ruta_msg}: {e}")
        else:
            print(f"⚠️ No se encontró tabla HTML en el correo: {ruta_msg}")

    except Exception as e:
        print(f"❌ Error procesando {ruta_msg}: {e}")

ejecutar = procesar_correo(ruta_msg, carpeta_correos)