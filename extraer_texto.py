# from PyPDF2 import PdfReader
# reader = PdfReader("C:/Users/Oscar Torres/Downloads/aseguradoras/23 ENERO/VANESSA met, zurich grupomexicanode seguros/MET/MET/FUNERARIOS/FUNERARIOS/01001_G0024669_0560644/01001_G0024669_0560644/Recibos/00083231761/AC/AVISO_TTI151104L63_83231761_0024669__A.pdf")
# print(reader.pages[0].extract_text())
import pdfplumber

ruta_pdf = "C:/Users/Oscar Torres/Downloads/aseguradoras/3er ETAPA 2da Parte/MAPFRE MEXICO SA BENEFICIOS/Archivos - 2025-02-28T091346.168/1602400002598/1602400002598/Consentimientos 1602400002598.pdf"

with pdfplumber.open(ruta_pdf) as pdf:
    primera_pagina = pdf.pages[0]
    texto = primera_pagina.extract_text()
    print(texto)


    