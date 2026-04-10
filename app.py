from flask import Flask, request, jsonify
from flask_cors import CORS
import base64, os, subprocess, shutil, tempfile, json
from generators.robo import generar_robo
from generators.abuso import generar_abuso
from generators.credito import generar_credito

app = Flask(__name__)
CORS(app)  # Permite requests desde cualquier origen

@app.route('/health', methods=['GET'])
def health():
    return jsonify({'ok': True})

@app.route('/generar-pdf', methods=['POST'])
def generar_pdf():
    try:
        data = request.get_json()
        tipo  = data.get('tipoEvento', 'ROBO')
        datos = data.get('datos', {})

        with tempfile.TemporaryDirectory() as tmpdir:
            if tipo == 'ROBO':
                pdf_path = generar_robo(datos, tmpdir)
            elif tipo == 'ABUSO DE CONFIANZA':
                pdf_path = generar_abuso(datos, tmpdir)
            elif tipo == 'RECUPERACIÓN DE CRÉDITO':
                pdf_path = generar_credito(datos, tmpdir)
            else:
                return jsonify({'error': f'Tipo de evento no soportado: {tipo}'}), 400

            with open(pdf_path, 'rb') as f:
                pdf_b64 = base64.b64encode(f.read()).decode('utf-8')

            # Buscar también el docx o pptx generado
            folio = datos.get('folio', 'REPORTE')
            doc_b64 = None
            doc_nombre = None
            for ext in ['_final.pdf', '.docx', '.pptx']:
                candidate = os.path.join(tmpdir, f'{folio}{ext}')
                if ext not in ('_final.pdf',) and os.path.exists(candidate):
                    with open(candidate, 'rb') as f:
                        doc_b64 = base64.b64encode(f.read()).decode('utf-8')
                    doc_nombre = f"{folio} {datos.get('cliente','CLIENTE')}{ext}"
                    break

        nombre = f"{datos.get('folio','REPORTE')} {datos.get('cliente','CLIENTE') or datos.get('nombre','CLIENTE')}.pdf"
        resp = {'ok': True, 'pdf': pdf_b64, 'nombre': nombre}
        if doc_b64:
            resp['doc'] = doc_b64
            resp['doc_nombre'] = doc_nombre
        return jsonify(resp)

    except Exception as e:
        import traceback
        return jsonify({'ok': False, 'error': str(e), 'trace': traceback.format_exc()}), 500

if __name__ == '__main__':
    app.run(debug=True, port=5000)
