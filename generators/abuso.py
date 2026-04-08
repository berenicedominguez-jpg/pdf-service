import os, base64, io
from PIL import Image as PILImage
from .utils import *

# Ancho fijo para todas las imágenes (EMU) = ancho de tabla completa
CX_FIJO = 8029440
CY_MAX_MAPA = 7199985
CY_MAX_EVID = 4514392

def cx_cy_proporcional(b64_str, cx_max, cy_max):
    """Calcula cx y cy manteniendo aspect ratio, sin exceder los máximos."""
    try:
        img_data = base64.b64decode(b64_str)
        img = PILImage.open(io.BytesIO(img_data))
        w, h = img.size
        cx = cx_max
        cy = int(cx_max * h / w)
        if cy > cy_max:
            cy = cy_max
            cx = int(cy_max * w / h)
        return cx, cy
    except Exception:
        return cx_max, cy_max

def generar_abuso(d, tmpdir):
    F = d.get('folio','—')
    N = d.get('nick','—')
    M = d.get('monitorista','—')

    # Preparar imágenes
    imagenes = []
    img_lugar_b64 = d.get('img_lugar')
    img_evidencia_raw = d.get('img_evidencia', [])
    if isinstance(img_evidencia_raw, str):
        img_evidencia_raw = [img_evidencia_raw] if img_evidencia_raw else []

    rId_lugar      = 'rId10'
    rId_evidencia1 = 'rId11'
    rId_evidencia2 = 'rId12'

    body = '<w:body>'

    # Calcular dimensiones de imágenes ANTES de construir el body
    cx_mapa, cy_mapa = cx_cy_proporcional(img_lugar_b64, CX_FIJO, CY_MAX_MAPA) if img_lugar_b64 else (CX_FIJO, CY_MAX_MAPA)
    cx_ev1, cy_ev1 = cx_cy_proporcional(img_evidencia_raw[0], CX_FIJO, CY_MAX_EVID) if len(img_evidencia_raw) > 0 else (CX_FIJO, CY_MAX_EVID)
    cx_ev2, cy_ev2 = cx_cy_proporcional(img_evidencia_raw[1], CX_FIJO, CY_MAX_EVID) if len(img_evidencia_raw) > 1 else (CX_FIJO, CY_MAX_EVID)

    # Actualizar imagenes con cx/cy para redimensionado físico
    imagenes = []
    if img_lugar_b64:
        imagenes.append({'b64': img_lugar_b64, 'nombre': 'img_lugar', 'rId': rId_lugar, 'cx': cx_mapa, 'cy': cy_mapa})
    if len(img_evidencia_raw) > 0:
        imagenes.append({'b64': img_evidencia_raw[0], 'nombre': 'img_evidencia1', 'rId': rId_evidencia1, 'cx': cx_ev1, 'cy': cy_ev1})
    if len(img_evidencia_raw) > 1:
        imagenes.append({'b64': img_evidencia_raw[1], 'nombre': 'img_evidencia2', 'rId': rId_evidencia2, 'cx': cx_ev2, 'cy': cy_ev2})

    # ══ HOJA 1 ══
    body += enc_pagina(F, N, M, titulo='REPORTE ABUSO DE CONFIANZA')
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
        fila('Marca',               d.get('marca','—')) +
        fila('Submarca',            d.get('modelo','—'), True) +
        fila('Año',                 d.get('anio','—')) +
        fila('Color',               d.get('color','—'), True) +
        fila('Placas',              d.get('placas','—')) +
        fila('Número de serie (VIN)',d.get('serie','—'), True) +
        fila('Número económico',    d.get('nick','—')) +
        fila('Eventos registrados', d.get('alertas','N/A'), True) +
        fila('Paro de motor',       d.get('paro_motor','N/A')) +
        fila('Última ubicación GPS',d.get('lugar','—'), True) +
        fila('Coordenadas',         d.get('coordenadas','—')) +
        fila('Última señal GPS',    d.get('ultima_senal','—'), True)
    )

    # ══ HOJA 2 ══
    body += salto()
    body += enc_pagina(F, N, M, titulo='REPORTE DE ABUSO DE CONFIANZA')
    body += tabla(
        enc_sec('DESCRIPCIÓN DEL EVENTO') +
        fila('Tipo de evento',              'ABUSO DE CONFIANZA') +
        fila('Lugar del evento',            d.get('estado_rep','—'), True) +
        fila('Fecha aproximada del robo',   d.get('fecha_robo','—')) +
        fila('Hora aproximada del robo',    d.get('hora_robo','—'), True) +
        fila('Folio de predenuncia',        d.get('folio_predenuncia','—')) +
        fila('Número de carpeta',           d.get('num_carpeta','—'), True) +
        fila('Se levanta denuncia en el MP',d.get('denuncia_mp','—')) +
        fila('Autoridad que apoya',         d.get('autoridades','—'), True)
    )
    body += sep()

    # Imagen lugar del robo — ocupa todo el espacio disponible de la hoja
    if img_lugar_b64:
        body += imagen_real('LUGAR DEL EVENTO', rId_lugar, cx=cx_mapa, cy=cy_mapa)
    else:
        body += imagen_placeholder('LUGAR DEL EVENTO', alto=2800)

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
    body += sep()
    body += tabla_texto('OBSERVACIONES ADICIONALES', d.get('observaciones','—'))

    # ══ HOJA 4 ══ — ambas evidencias calculadas para caber exactamente en la hoja
    body += salto()
    body += enc_pagina(F, N, M, titulo='REPORTE DE ABUSO DE CONFIANZA')
    if len(img_evidencia_raw) > 0:
        body += imagen_real('EVIDENCIA', rId_evidencia1, cx=cx_ev1, cy=cy_ev1)
        if len(img_evidencia_raw) > 1:
            body += sep()
            body += imagen_real('EVIDENCIA (2)', rId_evidencia2, cx=cx_ev2, cy=cy_ev2)
    else:
        body += imagen_placeholder('EVIDENCIA', alto=3500)

    body += '''<w:sectPr>
  <w:headerReference r:id="rId6" w:type="default"/>
  <w:footerReference r:id="rId7" w:type="default"/>
  <w:pgSz w:h="16834" w:w="11909" w:orient="portrait"/>
  <w:pgMar w:bottom="1440" w:top="1440" w:left="1440" w:right="1440" w:header="720" w:footer="720"/>
</w:sectPr></w:body>'''

    return build_docx_pdf(body, tmpdir, F, imagenes=imagenes if imagenes else None)
