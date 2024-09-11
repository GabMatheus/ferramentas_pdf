import os
import py7zr
from flask import Flask, request, jsonify, render_template
from werkzeug.utils import secure_filename
from datetime import datetime
from docx2pdf import convert
import logging
from PyPDF2 import PdfMerger
import pythoncom
from pdf2docx import Converter

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'

# Configure logging
logging.basicConfig(level=logging.INFO)

# Certifique-se de que a pasta de upload existe
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

def create_7z(files, archive_file):
    try:
        with py7zr.SevenZipFile(archive_file, 'w') as archive:
            for file in files:
                archive.write(file, os.path.basename(file))
    except Exception as e:
        logging.error(f'Falha ao criar o arquivo 7z: {e}')
        raise RuntimeError(f'Falha ao criar o arquivo 7z: {e}')

def convert_docx_to_pdf(input_file, output_file):
    try:
        pythoncom.CoInitialize()  # Inicializa o COM
        convert(input_file, output_file)
    except Exception as e:
        logging.error(f'Falha ao converter DOCX para PDF: {e}')
        raise RuntimeError(f'Falha ao converter DOCX para PDF: {e}')
    finally:
        pythoncom.CoUninitialize()  
def convert_pdf_to_docx(input_file, output_file):
    try:
        cv = Converter(input_file)
        cv.convert(output_file, start=0, end=None)
        cv.close()
    except Exception as e:
        logging.error(f'Falha ao converter PDF para DOCX: {e}')
        raise RuntimeError(f'Falha ao converter PDF para DOCX: {e}')

def merge_pdfs(file_paths, output_file):
    try:
        pdf_merger = PdfMerger()
        for file_path in file_paths:
            pdf_merger.append(file_path)
        pdf_merger.write(output_file)
        pdf_merger.close()
    except Exception as e:
        logging.error(f'Falha ao mesclar PDFs: {e}')
        raise RuntimeError(f'Falha ao mesclar PDFs: {e}')

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/convert_docx', methods=['POST'])
def convert_docx_route():
    if 'file' not in request.files:
        return jsonify({'Erro': 'Nenhum arquivo fornecido'}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({'Erro': 'Nenhum arquivo selecionado'}), 400

    filename = secure_filename(file.filename)
    if not filename.lower().endswith('.docx'):
        return jsonify({'Erro': 'Tipo de arquivo inválido. Por favor selecione um arquivo DOCX'}), 400

    input_file = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(input_file)

    output_file = os.path.join(app.config['UPLOAD_FOLDER'], f'{os.path.splitext(filename)[0]}.pdf')

    try:
        convert_docx_to_pdf(input_file, output_file)
    except RuntimeError as e:
        os.remove(input_file)
        return jsonify({'Erro': str(e)}), 500

    os.remove(input_file)
    return jsonify({'Aviso': 'Arquivo DOCX convertido para PDF com sucesso!', 'output_file': output_file})

@app.route('/merge_pdfs', methods=['POST'])
def merge_pdfs_route():
    files = request.files.getlist('files')
    if not files:
        return jsonify({'Erro': 'Nenhum arquivo fornecido'}), 400

    saved_files = []
    date_str = datetime.now().strftime('%d-%m-%Y_%H-%M-%S')

    try:
        for file in files:
            if file.filename == '':
                continue

            filename = secure_filename(file.filename)
            if not filename.lower().endswith('.pdf'):
                return jsonify({'Erro': 'Tipo inválido, apenas PDFs são permitidos'}), 400

            file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(file_path)
            saved_files.append(file_path)

        merged_pdf_file = os.path.join(app.config['UPLOAD_FOLDER'], f'arquivo_mesclado_{date_str}.pdf')
        merge_pdfs(saved_files, merged_pdf_file)

        return jsonify({'Aviso': 'PDFs mesclados com sucesso!', 'output': merged_pdf_file})

    except RuntimeError as e:
        return jsonify({'Erro': str(e)}), 500

    finally:
        # Limpeza dos arquivos temporários
        for file_path in saved_files:
            try:
                os.remove(file_path)
            except Exception as e:
                logging.error(f'Falha ao remover o arquivo {file_path}: {e}')

@app.route('/convert_pdf', methods=['POST'])
def convert_pdf_route():
    if 'file' not in request.files:
        return jsonify({'Erro': 'Nenhum arquivo fornecido'}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({'Erro': 'Nenhum arquivo selecionado'}), 400

    filename = secure_filename(file.filename)
    if not filename.lower().endswith('.pdf'):
        return jsonify({'Erro': 'Tipo de arquivo inválido. Por favor selecione um arquivo PDF'}), 400

    input_file = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(input_file)

    output_file = os.path.join(app.config['UPLOAD_FOLDER'], f'{os.path.splitext(filename)[0]}.docx')

    try:
        convert_pdf_to_docx(input_file, output_file)
    except RuntimeError as e:
        os.remove(input_file)
        return jsonify({'Erro': str(e)}), 500

    os.remove(input_file)
    return jsonify({'Aviso': 'Arquivo PDF convertido para DOCX com sucesso!', 'output_file': output_file})

@app.route('/create_archive', methods=['POST'])
def create_archive_route():
    files = request.files.getlist('files')
    if not files:
        return jsonify({'Erro': 'Nenhum arquivo fornecido'}), 400

    saved_files = []
    date_str = datetime.now().strftime('%d-%m-%Y_%H-%M-%S')

    try:
        for file in files:
            if file.filename == '':
                continue

            filename = secure_filename(file.filename)
            if not filename:
                continue

            file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(file_path)
            saved_files.append(file_path)

        archive_file = os.path.join(app.config['UPLOAD_FOLDER'], f'arquivo_{date_str}.7z')
        create_7z(saved_files, archive_file)

        return jsonify({'Aviso': 'Arquivos comprimidos com sucesso!', 'output': archive_file})

    except RuntimeError as e:
        return jsonify({'Erro': str(e)}), 500

    finally:
        # Limpeza dos arquivos temporários
        for file_path in saved_files:
            try:
                os.remove(file_path)
            except Exception as e:
                logging.error(f'Falha ao remover o arquivo {file_path}: {e}')

if __name__ == '__main__':
    app.run(debug=True)
