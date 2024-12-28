#!/usr/bin/env python
# coding: utf-8




import fitz  # PyMuPDF
import re
import pandas as pd
from io import BytesIO
import os
from flask import Flask, request, send_file, jsonify

app = Flask(__name__)

# Diretório para uploads temporários
UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def extract_info_from_pdf(pdf_path):
    try:
        pdf_document = fitz.open(pdf_path)
        all_text = ""
        for page_number in range(len(pdf_document)):
            page = pdf_document.load_page(page_number)
            all_text += page.get_text("text") + "\n"

        pdf_document.close()

        lines = all_text.splitlines()
        lines_sem_branco = [l for l in lines if l.strip() != '']
        all_text = "\n".join(lines_sem_branco)

        all_beneficiarios = extract_info_from_text(all_text)
        return all_beneficiarios

    except Exception as e:
        print(f"Erro ao extrair informações do PDF: {str(e)}")
        return []

def extract_info_from_text(text):
    pattern = re.compile(
        r"71 - Nome Social do Beneficiário\n(?P<numero_guia>\d{9})|"
        r"9\d{8}-00\d\r?\n(?P<dentista>.+)|"
        r"21 - Nome do Beneficiário\n(?P<nome_beneficiario>.+)|"
        r"40 - Valor Total Liberado Guia \(R\$\)\n(?P<valor_guia>[\d.,]+)"
    )

    beneficiarios = []
    info_beneficiario = {}

    for match in pattern.finditer(text):
        if match.lastgroup == "numero_guia":
            if info_beneficiario:
                beneficiarios.append(info_beneficiario.copy())
                info_beneficiario.clear()
            info_beneficiario['numero_guia'] = match.group("numero_guia")

        elif match.lastgroup == "dentista":
            info_beneficiario['dentista'] = match.group("dentista")

        elif match.lastgroup == "nome_beneficiario":
            info_beneficiario['nome_beneficiario'] = match.group("nome_beneficiario")

        elif match.lastgroup == "valor_guia":
            info_beneficiario['valor_guia'] = match.group("valor_guia")

    if info_beneficiario:
        beneficiarios.append(info_beneficiario)

    return beneficiarios

@app.route('/process_pdf', methods=['POST'])
def process_pdf():
    if 'file' not in request.files:
        return jsonify({"error": "Nenhum arquivo enviado"}), 400

    file = request.files['file']
    if not file.filename.endswith('.pdf'):
        return jsonify({"error": "Arquivo inválido. Apenas PDFs são aceitos"}), 400

    temp_pdf_path = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(temp_pdf_path)

    try:
        todos_beneficiarios = extract_info_from_pdf(temp_pdf_path)

        # Processa os dados e gera o Excel
        linhas_detalhadas = []
        for benef in todos_beneficiarios:
            numero = benef.get('numero_guia', '')
            dentista = benef.get('dentista', '')
            nome_ben = benef.get('nome_beneficiario', '')
            valor_str = benef.get('valor_guia', '').strip()

            valor_num = None
            if valor_str:
                valor_sem_pontos = valor_str.replace('.', '')
                valor_convertido = valor_sem_pontos.replace(',', '.')
                try:
                    valor_num = float(valor_convertido)
                except ValueError:
                    valor_num = None

            linhas_detalhadas.append({
                "Número da Guia": numero,
                "Dentista": dentista,
                "Nome do Beneficiário": nome_ben,
                "Valor da Guia": valor_num
            })

        df_detalhado = pd.DataFrame(linhas_detalhadas)

        df_detalhado_agrupado = df_detalhado.groupby("Número da Guia", as_index=False).agg({
            "Dentista": "first",
            "Nome do Beneficiário": "first",
            "Valor da Guia": "sum"
        })

        df_resumo = df_detalhado_agrupado.groupby('Dentista', as_index=False, dropna=False).agg(
            quant_guias=('Número da Guia', 'count'),
            total_guias=('Valor da Guia', 'sum')
        )

        qtd_guias_geral = df_resumo['quant_guias'].sum()
        total_guias_geral = df_resumo['total_guias'].sum()

        df_total = pd.DataFrame({
            'Dentista': ['TOTAL GERAL'],
            'quant_guias': [qtd_guias_geral],
            'total_guias': [total_guias_geral]
        })
        df_resumo = pd.concat([df_resumo, df_total], ignore_index=True)

        output = BytesIO()
        with pd.ExcelWriter(output) as writer:
            df_detalhado_agrupado.to_excel(writer, sheet_name='Detalhado', index=False)
            df_resumo.to_excel(writer, sheet_name='Resumo por Dentista', index=False)

        output.seek(0)
        os.remove(temp_pdf_path)  # Remove o PDF temporário
        return send_file(output, as_attachment=True, download_name='resultado.xlsx', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    except Exception as e:
        os.remove(temp_pdf_path)  # Remove o PDF temporário em caso de erro
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    app.run(debug=True)
