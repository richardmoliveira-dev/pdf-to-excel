#!/usr/bin/env python
# coding: utf-8

# In[ ]:


from flask import Flask, request, send_file
import fitz  # PyMuPDF
import re
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
import os

app = Flask(__name__)

# Função para extrair informações do PDF
def extract_info_from_pdf(pdf_path):
    try:
        pdf_document = fitz.open(pdf_path)
        all_beneficiarios = []

        for page_number in range(len(pdf_document)):
            page = pdf_document.load_page(page_number)
            text = page.get_text("text")
            beneficiarios_pagina = extract_info_from_text(text)
            all_beneficiarios.extend(beneficiarios_pagina)

        pdf_document.close()
        return all_beneficiarios
    except Exception as e:
        print(f"Erro ao extrair informações do PDF: {str(e)}")
        return []

# Função para extrair informações do texto
def extract_info_from_text(text):
    padrao_numero_guia = r"71 - Nome Social do Beneficiário\n(\d{9})"
    padrao_nome_beneficiario = r"21 - Nome do Beneficiário\n(.+)"
    padrao_valor_guia = r"40 - Valor Total Liberado Guia \(R\$\)\n([\d,]+)"

    matches = re.finditer(f"{padrao_numero_guia}|{padrao_nome_beneficiario}|{padrao_valor_guia}", text)
    beneficiarios_pagina = []
    info_beneficiario = {}

    for match in matches:
        if match.group(1):
            if info_beneficiario:
                beneficiarios_pagina.append(info_beneficiario.copy())
                info_beneficiario.clear()
            info_beneficiario['numero_guia'] = int(match.group(1))
        elif match.group(2):
            info_beneficiario['nome_beneficiario'] = match.group(2)
        elif match.group(3):
            info_beneficiario['valor_guia'] = match.group(3)

    if info_beneficiario:
        beneficiarios_pagina.append(info_beneficiario)

    return beneficiarios_pagina

# Rota da API para processar o PDF
@app.route('/process_pdf', methods=['POST'])
def process_pdf():
    if 'file' not in request.files:
        return "Nenhum arquivo enviado", 400

    uploaded_file = request.files['file']
    pdf_path = f"uploads/{uploaded_file.filename}"

    os.makedirs("uploads", exist_ok=True)
    uploaded_file.save(pdf_path)

    # Processar o PDF e criar o Excel
    beneficiarios = extract_info_from_pdf(pdf_path)
    if not beneficiarios:
        return "Erro ao processar o PDF ou PDF vazio", 500

    excel_bytes = BytesIO()
    wb = Workbook()
    ws = wb.active

    ws.append(["Número da Guia", "Nome do Beneficiário", "Valor da Guia"])
    for beneficiario in beneficiarios:
        ws.append([
            beneficiario.get('numero_guia', ''),
            beneficiario.get('nome_beneficiario', ''),
            beneficiario.get('valor_guia', '')
        ])

    wb.save(excel_bytes)
    excel_bytes.seek(0)

    return send_file(
        excel_bytes,
        as_attachment=True,
        download_name=uploaded_file.filename.replace('.pdf', '_informacoes_beneficiarios.xlsx'),
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

if __name__ == '__main__':
    app.run(debug=True)

