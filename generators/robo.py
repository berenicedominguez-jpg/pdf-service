import os
from .utils import *

def generar_robo(d, tmpdir):
    F = d.get('folio','—')
    N = d.get('nick','—')
    M = d.get('monitorista','—')

    es_inbursa_compact = d.get('plataforma','').upper() == 'INBURSA' and any(
        v not in ('N/A','','—') for v in [d.get('caja_tipo','N/A'), d.get('caja_marca','N/A')]
    )

    caja_vals = [d.get('caja_tipo','N/A'),d.get('caja_marca','N/A'),d.get('caja_color','N/A'),
                 d.get('caja_anio','N/A'),d.get('caja_placas','N/A'),d.get('caja_serie','N/A')]
    tiene_caja = any(v not in ('N/A','','—') for v in caja_vals)

    body = '<w:body>'

    # ══ HOJA 1 ══
    body += enc_pagina(F, N, M)

    # Recepción
    recep = (enc_sec('RECEPCIÓN DEL REPORTE') +
        fila('Fecha y hora del reporte', d.get('fecha_atencion','—')) +
        fila('Cliente',                  d.get('cliente','—'), True) +
        fila('Nombre',                   d.get('nombre_reportante','—')) +
        fila('Relación con la unidad',   d.get('relacion_unidad','—'), True) +
        fila('Medio de contacto',        d.get('medio_contacto','—')) +
        fila('Teléfono',                 d.get('tel_reportante','—'), True) +
        fila('Correo electrónico',       d.get('correo_reportante','—'))
    )
    if d.get('plataforma','').upper() == 'INBURSA':
        recep += (
            fila('No. Póliza',   d.get('inbursa_poliza','—'), True) +
            fila('No. CIS',      d.get('inbursa_cis','—')) +
            fila('No. Siniestro',d.get('inbursa_siniestro','—'), True) +
            fila('Alarmado por', d.get('inbursa_alarmado','—'))
        )
    body += tabla(recep)
    body += sep()

    # Datos unidad
    body += tabla(
        enc_sec('DATOS DE LA UNIDAD') +
        fila('Marca',               d.get('marca','—')) +
        fila('Modelo',              d.get('modelo','—'), True) +
        fila('Año',                 d.get('anio','—')) +
        fila('Color',               d.get('color','—'), True) +
        fila('Placas',              d.get('placas','—')) +
        fila('Número de serie (VIN)',d.get('serie','—'), True) +
        fila('Número económico',    d.get('nick','—')) +
        fila('Alerta(s) recibidas', d.get('alertas','N/A'), True) +
        fila('Paro de motor',       d.get('paro_motor','N/A')) +
        fila('Última ubicación GPS',d.get('lugar','—'), True) +
        fila('Coordenadas',         d.get('coordenadas','—')) +
        fila('Última señal GPS',    d.get('ultima_senal','—'), True)
    )

    # Acoplado
    if tiene_caja and not es_inbursa_compact:
        body += sep()
        body += tabla(
            enc_sec('DATOS DEL ACOPLADO') +
            fila('Tipo de caja', d.get('caja_tipo','N/A')) +
            fila('Marca',        d.get('caja_marca','N/A'), True) +
            fila('Color',        d.get('caja_color','N/A')) +
            fila('Año',          d.get('caja_anio','N/A'), True) +
            fila('Placas',       d.get('caja_placas','N/A')) +
            fila('Serie',        d.get('caja_serie','N/A'), True)
        )

    # Compact hoja 1 si INBURSA + acoplado
    if es_inbursa_compact:
        hoja1_start = body.rfind('<w:body>') + len('<w:body>')
        hoja1 = body[hoja1_start:]
        hoja1 = (hoja1.replace('w:w="55"','w:w="38"')
                       .replace('w:w="85"','w:w="65"')
                       .replace('w:line="255"','w:line="220"')
                       .replace('w:val="17"','w:val="16"'))
        body = body[:hoja1_start] + hoja1

    # ══ HOJA 2 ══
    body += salto()
    body += enc_pagina(F, N, M)

    if tiene_caja and es_inbursa_compact:
        body += tabla(
            enc_sec('DATOS DEL ACOPLADO') +
            fila('Tipo de caja', d.get('caja_tipo','N/A')) +
            fila('Marca',        d.get('caja_marca','N/A'), True) +
            fila('Color',        d.get('caja_color','N/A')) +
            fila('Año',          d.get('caja_anio','N/A'), True) +
            fila('Placas',       d.get('caja_placas','N/A')) +
            fila('Serie',        d.get('caja_serie','N/A'), True)
        )
        body += sep()

    body += tabla(
        enc_sec('DESCRIPCIÓN DEL EVENTO') +
        fila('Lugar del evento',          d.get('estado_rep','—'), True) +
        fila('Fecha aproximada del evento',d.get('fecha_robo','—')) +
        fila('Hora aproximada del evento', d.get('hora_robo','—'), True) +
        fila('Folio de predenuncia',       d.get('folio_predenuncia','—'), True) +
        fila('Situación',                  d.get('situacion','—'))
    )
    body += sep()
    body += tabla_texto('ACCIONES REALIZADAS', d.get('acciones','—'))
    body += sep()
    body += tabla(
        enc_sec('ESTATUS ACTUAL') +
        fila('Resultado',              d.get('resultado','—'),
             verde=('RECUPERADA' in d.get('resultado','') and 'NO' not in d.get('resultado',''))) +
        fila('Observaciones adicionales', d.get('observaciones','—'))
    )

    body += '''<w:sectPr>
  <w:headerReference r:id="rId6" w:type="default"/>
  <w:footerReference r:id="rId7" w:type="default"/>
  <w:pgSz w:h="16834" w:w="11909" w:orient="portrait"/>
  <w:pgMar w:bottom="1440" w:top="1440" w:left="1440" w:right="1440" w:header="720" w:footer="720"/>
</w:sectPr></w:body>'''

    return build_docx_pdf(body, tmpdir, F)
