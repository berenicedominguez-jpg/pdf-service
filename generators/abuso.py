import os, base64, io
from PIL import Image as PILImage
from .utils import *

CX_FIJO    = 8029440
CY_MAX_MAPA = 5500000
CY_MAX_EVID = 3800000

def generar_abuso(d, tmpdir):
    F = d.get('folio','—')
    N = d.get('nick','—')
    M = d.get('monitorista','—')

    img_lugar_b64     = d.get('img_lugar')
    img_evidencia_raw = d.get('img_evidencia', [])
    if isinstance(img_evidencia_raw, str):
        img_evidencia_raw = [img_evidencia_raw] if img_evidencia_raw else []

    # Todas las imágenes van por ReportLab
    evidencias_pdf = []
    if img_lugar_b64:
        evidencias_pdf.append({'titulo': 'LUGAR DEL EVENTO', 'b64': img_lugar_b64, 'tipo': 'mapa'})
    if len(img_evidencia_raw) > 0:
        evidencias_pdf.append({'titulo': 'EVIDENCIA',    'b64': img_evidencia_raw[0], 'tipo': 'evidencia'})
    if len(img_evidencia_raw) > 1:
        evidencias_pdf.append({'titulo': 'EVIDENCIA (2)', 'b64': img_evidencia_raw[1], 'tipo': 'evidencia'})

    body = '<w:body>'

    # ══ HOJA 1 ══
    body += enc_pagina(F, N, M, titulo='REPORTE DE ABUSO DE CONFIANZA')
    body += tabla(
        enc_sec('RECEPCIÓN DEL REPORTE') +
        fila('Fecha y hora del reporte', d.get('fecha_atencion','—')) +
        fila('Cliente',                  d.get('cliente','—'), True) +
        fila('No. Póliza',               d.get('poliza','—')) +
        fila('Denunciante',              d.get('nombre_reportante','—'), True) +
        fila('Relación con la unidad',   d.get('relacion_unidad','—')) +
        fila('Medio de contacto',        d.get('medio_contacto','—'), True) +
        fila('Teléfono',                 d.get('tel_reportante','—')) +
        fila('Correo electrónico',       d.get('correo_reportante','—'), True)
    )
    body += sep()
    body += tabla(
        enc_sec('DATOS DE LA UNIDAD') +
        fila('Marca',                d.get('marca','—')) +
        fila('Submarca',             d.get('modelo','—'), True) +
        fila('Año',                  d.get('anio','—')) +
        fila('Color',                d.get('color','—'), True) +
        fila('Placas',               d.get('placas','—')) +
        fila('Número de serie (VIN)', d.get('serie','—'), True) +
        fila('Número económico',     d.get('nick','—')) +
        fila('Eventos registrados',  d.get('alertas','N/A'), True) +
        fila('Paro de motor',        d.get('paro_motor','N/A')) +
        fila('Última ubicación GPS', d.get('lugar','—'), True) +
        fila('Coordenadas',          d.get('coordenadas','—')) +
        fila('Última señal GPS',     d.get('ultima_senal','—'), True)
    )

    # ══ HOJA 2 ══ — Descripción + placeholder para mapa (imagen va por ReportLab)
    body += salto()
    body += enc_pagina(F, N, M, titulo='REPORTE DE ABUSO DE CONFIANZA')
    body += tabla(
        enc_sec('DESCRIPCIÓN DEL EVENTO') +
        fila('Tipo de evento',               'ABUSO DE CONFIANZA') +
        fila('Lugar del evento',             d.get('estado_rep','—'), True) +
        fila('Fecha aproximada del robo',    d.get('fecha_robo','—')) +
        fila('Hora aproximada del robo',     d.get('hora_robo','—'), True) +
        fila('Folio de predenuncia',         d.get('folio_predenuncia','—')) +
        fila('Número de carpeta',            d.get('num_carpeta','—'), True) +
        fila('Se levanta denuncia en el MP', d.get('denuncia_mp','—')) +
        fila('Autoridad que apoya',          d.get('autoridades','—'), True)
    )
    body += sep()
    # Barra "LUGAR DEL EVENTO" — la imagen la superpone ReportLab
    body += tabla(enc_sec('LUGAR DEL EVENTO'))

    # ══ HOJA 3 ══
    body += salto()
    body += enc_pagina(F, N, M, titulo='REPORTE DE ABUSO DE CONFIANZA')
    body += tabla_texto('ACCIONES REALIZADAS', d.get('acciones','—'))
    body += sep()
    es_rec = 'RECUPERADA' in d.get('resultado','') and 'NO' not in d.get('resultado','')
    body += tabla(
        enc_sec('ESTATUS ACTUAL') +
        fila('Resultado',                d.get('resultado','—'), verde=es_rec, rojo=not es_rec) +
        fila('Lugar de la recuperación', d.get('lugar_recuperacion','N/A'), True) +
        fila('Vehículo remitido',        d.get('vehiculo_remitido','NO')) +
        fila('Lugar de remisión',        d.get('lugar_remision','N/A'), True) +
        fila('Entregado al cliente',     d.get('entregado_cliente','NO'))
    )

    # ══ HOJA 4 ══ — página con membrete, las imágenes las superpone ReportLab
    body += salto()
    body += enc_pagina(F, N, M, titulo='REPORTE DE ABUSO DE CONFIANZA')

    body += '''<w:sectPr>
  <w:headerReference r:id="rId6" w:type="default"/>
  <w:footerReference r:id="rId7" w:type="default"/>
  <w:pgSz w:h="16834" w:w="11909" w:orient="portrait"/>
  <w:pgMar w:bottom="1440" w:top="1440" w:left="1440" w:right="1440" w:header="720" w:footer="720"/>
</w:sectPr></w:body>'''

    return build_docx_pdf(body, tmpdir, F,
                          evidencias_pdf=evidencias_pdf if evidencias_pdf else None)
