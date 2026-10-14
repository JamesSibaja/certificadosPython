import io
import os
import re
import subprocess
import tempfile
import zipfile
from docx import Document
from flask import Flask, jsonify, request, send_file
from flask_cors import CORS
from mailmerge import MailMerge  # <-- IMPORTANTE: Se importa como mailmerge
import openpyxl
import pandas as pd
import requests

app = Flask(__name__)
CORS(app)


def convertir_enlace_drive(url):
  """Convierte enlaces de Google Drive / Docs para descarga directa."""
  if not url:
    return url
  match_docs = re.search(r'/document/d/([a-zA-Z0-9_-]+)', url)
  if match_docs:
    return (
        f'https://docs.google.com/document/d/{match_docs.group(1)}/export?format=docx'
    )
  match_drive = re.search(r'/file/d/([a-zA-Z0-9_-]+)', url)
  if match_drive:
    return (
        f'https://drive.google.com/uc?export=download&id={match_drive.group(1)}'
    )
  return url


def descargar_machote_docx(url):
  url_directa = convertir_enlace_drive(url)
  response = requests.get(url_directa)
  if response.status_code != 200:
    raise Exception(
        'No se pudo descargar el machote desde Google Drive. Asegúrate de que'
        ' sea público.'
    )
  return io.BytesIO(response.content)


def extraer_etiquetas_docx(doc_path):
  """Extrae etiquetas conservando la compatibilidad con MailMerge y llaves dobles/triples."""
  with MailMerge(doc_path) as mm:
    campos = list(mm.get_merge_fields())

  if not campos:
    doc = Document(doc_path)
    etiquetas = set()
    for p in doc.paragraphs:
      texto = ''.join([run.text for run in p.runs])
      matches = re.findall(r'<<(.*?)>>|\{\{(.*?)\}\}', texto)
      for m in matches:
        tag = m[0] if m[0] else m[1]
        if tag.strip():
          etiquetas.add(tag.strip())
    campos = list(etiquetas)

  return campos


@app.route('/api/plantilla-datos', methods=['POST'])
def plantilla_datos():
  try:
    data = request.get_json()
    url_machote = data.get('url')
    if not url_machote:
      return jsonify({'error': 'Falta la URL del machote'}), 400

    machote_io = descargar_machote_docx(url_machote)

    with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp:
      tmp.write(machote_io.getvalue())
      tmp_path = tmp.name

    tags = extraer_etiquetas_docx(tmp_path)
    os.unlink(tmp_path)

    if not tags:
      tags = ['nombre', 'cedula', 'curso', 'fecha']

    columnas = ['nombre_archivo_pdf'] + [
        t for t in tags if t != 'nombre_archivo_pdf'
    ]

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Plantilla Datos'
    ws.append(columnas)

    ejemplo = ['certificado_ejemplo'] + [f'Dato {t}' for t in columnas[1:]]
    ws.append(ejemplo)

    excel_io = io.BytesIO()
    wb.save(excel_io)
    excel_io.seek(0)

    return send_file(
        excel_io,
        mimetype=(
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        ),
        as_attachment=True,
        download_name='plantilla_datos_certificado.xlsx',
    )
  except Exception as e:
    return jsonify({'error': str(e)}), 500


@app.route('/api/generar-certificados', methods=['POST'])
def generar_certificados():
  temp_dir = tempfile.mkdtemp()
  try:
    url_machote = request.form.get('url_machote')
    file_datos = request.files.get('archivo_datos')

    if not url_machote or not file_datos:
      return jsonify({'error': 'Faltan parámetros o archivos'}), 400

    machote_bytes = descargar_machote_docx(url_machote).getvalue()

    machote_path = os.path.join(temp_dir, 'machote.docx')
    with open(machote_path, 'wb') as f:
      f.write(machote_bytes)

    if file_datos.filename.endswith('.xlsx'):
      df = pd.read_excel(file_datos)
    else:
      df = pd.read_csv(file_datos)

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w') as zip_file:
      for index, row in df.iterrows():
        row_dict = {}
        for k, v in row.items():
          row_dict[str(k)] = '' if pd.isna(v) else str(v)

        nombre_pdf = row_dict.get('nombre_archivo_pdf', '').strip()
        if not nombre_pdf:
          nombre_pdf = (
              str(row.iloc[1]).replace(' ', '_')
              if len(row) > 1
              else f'certificado_{index}'
          )
        nombre_pdf = ''.join(
            c for c in nombre_pdf if c.isalnum() or c in ('_', '-', '.')
        )

        with MailMerge(machote_path) as document:
          available_fields = document.get_merge_fields()
          filtered_dict = {
              k: v for k, v in row_dict.items() if k in available_fields
          }
          if filtered_dict:
            document.merge(**filtered_dict)

          doc_out_path = os.path.join(temp_dir, f'{nombre_pdf}.docx')
          document.write(doc_out_path)

        pdf_out_path = os.path.join(temp_dir, f'{nombre_pdf}.pdf')
        subprocess.run(
            [
                'libreoffice',
                '--headless',
                '--convert-to',
                'pdf',
                '--outdir',
                temp_dir,
                doc_out_path,
            ],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )

        if os.path.exists(pdf_out_path):
          with open(pdf_out_path, 'rb') as pf:
            zip_file.writestr(f'{nombre_pdf}.pdf', pf.read())
        else:
          with open(doc_out_path, 'rb') as wf:
            zip_file.writestr(f'{nombre_pdf}.docx', wf.read())

    zip_buffer.seek(0)
    return send_file(
        zip_buffer,
        mimetype='application/zip',
        as_attachment=True,
        download_name='certificados_pdf.zip',
    )

  except Exception as e:
    return jsonify({'error': str(e)}), 500
  finally:
    for root, dirs, files in os.walk(temp_dir, topdown=False):
      for name in files:
        try:
          os.remove(os.path.join(root, name))
        except:
          pass
      try:
        os.rmdir(root)
      except:
        pass


if __name__ == '__main__':
  # IMPORTANTE: Tomar el puerto que asigna Render dinámicamente
  port = int(os.environ.get('PORT', 5000))
  app.run(host='0.0.0.0', port=port)
