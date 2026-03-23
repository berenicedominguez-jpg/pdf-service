import os, base64
from .utils import *

def imagen_placeholder(titulo, alto=3000):
    AZUL='1B2A4A'; BORDE='C5D0E4'; BLANCO='FFFFFF'
    return f'''<w:tbl>
  <w:tblPr><w:tblW w:w="8709" w:type="dxa"/><w:tblLayout w:type="fixed"/>
    <w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar>
  </w:tblPr>
  <w:tblGrid><w:gridCol w:w="8709"/></w:tblGrid>
  <w:tr><w:trPr><w:tblHeader/></w:trPr>
    <w:tc><w:tcPr><w:tcW w:w="8709" w:type="dxa"/>
      <w:shd w:val="clear" w:color="auto" w:fill="{AZUL}"/>
      <w:tcBorders><w:top w:val="single" w:sz="2" w:color="{AZUL}"/><w:left w:val="single" w:sz="2" w:color="{AZUL}"/><w:bottom w:val="single" w:sz="2" w:color="{AZUL}"/><w:right w:val="single" w:sz="2" w:color="{AZUL}"/></w:tcBorders>
      <w:tcMar><w:top w:w="85" w:type="dxa"/><w:left w:w="100" w:type="dxa"/><w:bottom w:w="85" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tcMar>
    </w:tcPr>
    <w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr>
      <w:r><w:rPr><w:b/><w:bCs/><w:color w:val="FFFFFF"/><w:sz w:val="18"/><w:szCs w:val="18"/><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/></w:rPr>
      <w:t>{titulo}</w:t></w:r></w:p>
    </w:tc>
  </w:tr>
  <w:tr>
    <w:trPr><w:trHeight w:val="{alto}" w:hRule="exact"/></w:trPr>
    <w:tc><w:tcPr><w:tcW w:w="8709" w:type="dxa"/>
      <w:shd w:val="clear" w:color="auto" w:fill="F5F5F5"/>
      <w:tcBorders><w:top w:val="single" w:sz="2" w:color="{BORDE}"/><w:left w:val="single" w:sz="2" w:color="{BORDE}"/><w:bottom w:val="single" w:sz="2" w:color="{BORDE}"/><w:right w:val="single" w:sz="2" w:color="{BORDE}"/></w:tcBorders>
    </w:tcPr>
    <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="200" w:after="0"/></w:pPr>
      <w:r><w:rPr><w:color w:val="AAAAAA"/><w:sz w:val="16"/><w:szCs w:val="16"/><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/></w:rPr>
      <w:t>[ Insertar captura de pantalla ]</w:t></w:r></w:p>
    </w:tc>
  </w:tr>
</w:tbl>'''

def generar_abuso(d, tmpdir):
    F = d.get('folio','—')
    N = d.get('nick','—')
    M = d.get('monitorista','—')

    body = '<w:body>'

    # ══ HOJA 1 ══
    body += enc_pagina(F, N, M, titulo='REPORTE DE ROBO')
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
        fila('Modelo',              d.get('anio','—')) +
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
    body += enc_pagina(F, N, M, titulo='REPORTE DE ROBO')
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
    body += imagen_placeholder('LUGAR DEL ROBO', alto=2800)
    body += sep()
    body += tabla_texto('ACCIONES REALIZADAS', d.get('acciones','—'))

    # ══ HOJA 3 ══
    body += salto()
    body += enc_pagina(F, N, M, titulo='REPORTE DE ROBO')
    es_rec = 'RECUPERADA' in d.get('resultado','') and 'NO' not in d.get('resultado','')
    body += tabla(
        enc_sec('ESTATUS ACTUAL') +
        fila('Resultado',            d.get('resultado','—'), verde=es_rec, rojo=not es_rec) +
        fila('Lugar de la recuperación', d.get('lugar_recuperacion','N/A'), True) +
        fila('Vehículo remitido',    d.get('vehiculo_remitido','NO')) +
        fila('Lugar de remisión',    d.get('lugar_remision','N/A'), True) +
        fila('Entregado al cliente', d.get('entregado_cliente','NO'))
    )
    body += sep()
    body += imagen_placeholder('EVIDENCIA', alto=3500)

    body += '''<w:sectPr>
  <w:headerReference r:id="rId6" w:type="default"/>
  <w:footerReference r:id="rId7" w:type="default"/>
  <w:pgSz w:h="16834" w:w="11909" w:orient="portrait"/>
  <w:pgMar w:bottom="1440" w:top="1440" w:left="1440" w:right="1440" w:header="720" w:footer="720"/>
</w:sectPr></w:body>'''

    return build_docx_pdf(body, tmpdir, F)
