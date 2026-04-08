import os, subprocess, shutil, zipfile, base64

BASE_DIR     = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MEMBRETE_ZIP = os.path.join(BASE_DIR, 'membrete.zip')
MEMBRETE     = os.path.join(BASE_DIR, 'membrete_unpacked')

# Unzip membrete on first use
if not os.path.exists(MEMBRETE) and os.path.exists(MEMBRETE_ZIP):
    with zipfile.ZipFile(MEMBRETE_ZIP, 'r') as z:
        z.extractall(BASE_DIR)

LOGO_TC      = os.path.join(BASE_DIR, 'assets', 'logo_tc_transparent.png')
LOGO_NUMARIS = os.path.join(BASE_DIR, 'assets', 'numaris_logo.png')

AZUL='1B2A4A'; AZUL_MED='EEF2F8'; VERDE='C6EFCE'; VERDE_T='276221'
BLANCO='FFFFFF'; BORDE='C5D0E4'; ROJO='C00000'

def lbl(txt, compact=False):
    p = '38' if compact else '55'
    l = '220' if compact else '255'
    f = '16' if compact else '17'
    return f'''<w:tc>
  <w:tcPr><w:tcW w:w="2600" w:type="dxa"/>
    <w:tcBorders><w:top w:val="single" w:sz="2" w:color="{BORDE}"/><w:left w:val="single" w:sz="2" w:color="{BORDE}"/><w:bottom w:val="single" w:sz="2" w:color="{BORDE}"/><w:right w:val="single" w:sz="2" w:color="{BORDE}"/></w:tcBorders>
    <w:tcMar><w:top w:w="{p}" w:type="dxa"/><w:left w:w="100" w:type="dxa"/><w:bottom w:w="{p}" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tcMar>
  </w:tcPr>
  <w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="{l}" w:lineRule="exact"/></w:pPr>
    <w:r><w:rPr><w:b/><w:bCs/><w:color w:val="{AZUL}"/><w:sz w:val="{f}"/><w:szCs w:val="{f}"/><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/></w:rPr>
    <w:t xml:space="preserve">{txt}</w:t></w:r></w:p></w:tc>'''

def val(txt, shade=False, verde=False, rojo=False, compact=False):
    fill  = VERDE if verde else (AZUL_MED if shade else BLANCO)
    color = VERDE_T if verde else (ROJO if rojo else '222222')
    bold  = '<w:b/><w:bCs/>' if (verde or rojo) else ''
    p = '38' if compact else '55'
    l = '220' if compact else '255'
    f = '16' if compact else '17'
    return f'''<w:tc>
  <w:tcPr><w:tcW w:w="6109" w:type="dxa"/>
    <w:shd w:val="clear" w:color="auto" w:fill="{fill}"/>
    <w:tcBorders><w:top w:val="single" w:sz="2" w:color="{BORDE}"/><w:left w:val="single" w:sz="2" w:color="{BORDE}"/><w:bottom w:val="single" w:sz="2" w:color="{BORDE}"/><w:right w:val="single" w:sz="2" w:color="{BORDE}"/></w:tcBorders>
    <w:tcMar><w:top w:w="{p}" w:type="dxa"/><w:left w:w="100" w:type="dxa"/><w:bottom w:w="{p}" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tcMar>
  </w:tcPr>
  <w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="{l}" w:lineRule="exact"/></w:pPr>
    <w:r><w:rPr>{bold}<w:color w:val="{color}"/><w:sz w:val="{f}"/><w:szCs w:val="{f}"/><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/></w:rPr>
    <w:t xml:space="preserve">{txt}</w:t></w:r></w:p></w:tc>'''

def fila(l_txt, v_txt, shade=False, verde=False, rojo=False, compact=False):
    return f'<w:tr><w:trPr><w:cantSplit/></w:trPr>{lbl(l_txt,compact)}{val(v_txt,shade,verde,rojo,compact)}</w:tr>'

def enc_sec(titulo, compact=False):
    p = '65' if compact else '85'
    return f'''<w:tr><w:trPr><w:cantSplit/><w:tblHeader/></w:trPr>
  <w:tc><w:tcPr><w:tcW w:w="8709" w:type="dxa"/><w:gridSpan w:val="2"/>
    <w:shd w:val="clear" w:color="auto" w:fill="{AZUL}"/>
    <w:tcBorders><w:top w:val="single" w:sz="2" w:color="{AZUL}"/><w:left w:val="single" w:sz="2" w:color="{AZUL}"/><w:bottom w:val="single" w:sz="2" w:color="{AZUL}"/><w:right w:val="single" w:sz="2" w:color="{AZUL}"/></w:tcBorders>
    <w:tcMar><w:top w:w="{p}" w:type="dxa"/><w:left w:w="100" w:type="dxa"/><w:bottom w:w="{p}" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tcMar>
  </w:tcPr>
  <w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr>
    <w:r><w:rPr><w:b/><w:bCs/><w:color w:val="FFFFFF"/><w:sz w:val="18"/><w:szCs w:val="18"/><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/></w:rPr>
    <w:t>{titulo}</w:t></w:r></w:p>
  </w:tc></w:tr>'''

def tabla(filas_xml):
    return f'''<w:tbl>
  <w:tblPr><w:tblW w:w="8709" w:type="dxa"/><w:tblLayout w:type="fixed"/>
    <w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar>
  </w:tblPr>
  <w:tblGrid><w:gridCol w:w="2600"/><w:gridCol w:w="6109"/></w:tblGrid>
  {filas_xml}</w:tbl>'''

def tabla_texto(titulo, texto):
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
  <w:tr><w:trPr><w:cantSplit/></w:trPr>
    <w:tc><w:tcPr><w:tcW w:w="8709" w:type="dxa"/>
      <w:shd w:val="clear" w:color="auto" w:fill="{BLANCO}"/>
      <w:tcBorders><w:top w:val="single" w:sz="2" w:color="{BORDE}"/><w:left w:val="single" w:sz="2" w:color="{BORDE}"/><w:bottom w:val="single" w:sz="2" w:color="{BORDE}"/><w:right w:val="single" w:sz="2" w:color="{BORDE}"/></w:tcBorders>
      <w:tcMar><w:top w:w="100" w:type="dxa"/><w:left w:w="100" w:type="dxa"/><w:bottom w:w="100" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tcMar>
    </w:tcPr>
    <w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="300" w:lineRule="auto"/></w:pPr>
      <w:r><w:rPr><w:color w:val="222222"/><w:sz w:val="17"/><w:szCs w:val="17"/><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/></w:rPr>
      <w:t xml:space="preserve">{texto}</w:t></w:r></w:p>
    </w:tc>
  </w:tr>
</w:tbl>'''

def sep():
    return '<w:p><w:pPr><w:spacing w:before="0" w:after="0"/><w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr></w:pPr></w:p>'

def salto():
    return '<w:p><w:pPr><w:pageBreakBefore/><w:spacing w:before="0" w:after="0"/><w:rPr><w:sz w:val="2"/></w:rPr></w:pPr></w:p>'

def enc_pagina(folio, nick, monitorista, titulo='INFORME DE CASO'):
    return f'''<w:p>
  <w:pPr><w:spacing w:before="0" w:after="30"/><w:keepNext/>
    <w:tabs><w:tab w:val="right" w:pos="8709"/></w:tabs>
  </w:pPr>
  <w:r><w:rPr><w:b/><w:bCs/><w:color w:val="{AZUL}"/><w:sz w:val="42"/><w:szCs w:val="42"/>
    <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/></w:rPr>
    <w:t>{titulo}</w:t></w:r>
  <w:r><w:rPr><w:b/><w:bCs/><w:color w:val="{AZUL}"/><w:sz w:val="42"/><w:szCs w:val="42"/>
    <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/></w:rPr>
    <w:tab/><w:t>FOLIO: {folio}</w:t></w:r>
</w:p>
<w:p>
  <w:pPr><w:spacing w:before="0" w:after="120"/><w:keepNext/>
    <w:tabs><w:tab w:val="right" w:pos="8709"/></w:tabs>
    <w:pBdr><w:bottom w:val="single" w:sz="6" w:space="4" w:color="C5D0E4"/></w:pBdr>
  </w:pPr>
  <w:r><w:rPr><w:b/><w:bCs/><w:color w:val="{AZUL}"/><w:sz w:val="18"/><w:szCs w:val="18"/>
    <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/></w:rPr>
    <w:t>UNIDAD: {nick}</w:t></w:r>
  <w:r><w:rPr><w:b/><w:bCs/><w:color w:val="{AZUL}"/><w:sz w:val="18"/><w:szCs w:val="18"/>
    <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/></w:rPr>
    <w:tab/><w:t>ATIENDE: {monitorista}</w:t></w:r>
</w:p>'''

def imagen_placeholder(titulo, alto=3000):
    AZUL='1B2A4A'; BORDE='C5D0E4'
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

def imagen_real(titulo, rId, cx=8029440, cy=4500000):
    """Genera XML para insertar imagen real en el documento"""
    AZUL='1B2A4A'; BORDE='C5D0E4'
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
    <w:tc><w:tcPr><w:tcW w:w="8709" w:type="dxa"/>
      <w:tcBorders><w:top w:val="single" w:sz="2" w:color="{BORDE}"/><w:left w:val="single" w:sz="2" w:color="{BORDE}"/><w:bottom w:val="single" w:sz="2" w:color="{BORDE}"/><w:right w:val="single" w:sz="2" w:color="{BORDE}"/></w:tcBorders>
      <w:tcMar><w:top w:w="100" w:type="dxa"/><w:left w:w="100" w:type="dxa"/><w:bottom w:w="100" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tcMar>
    </w:tcPr>
    <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="0" w:after="0"/></w:pPr>
      <w:r><w:rPr><w:noProof/></w:rPr>
        <w:drawing>
          <wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
            <wp:extent cx="{cx}" cy="{cy}"/>
            <wp:effectExtent l="0" t="0" r="0" b="0"/>
            <wp:docPr id="1" name="imagen"/>
            <wp:cNvGraphicFramePr/>
            <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                  <pic:nvPicPr>
                    <pic:cNvPr id="0" name="imagen"/>
                    <pic:cNvPicPr/>
                  </pic:nvPicPr>
                  <pic:blipFill>
                    <a:blip r:embed="{rId}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
                    <a:stretch><a:fillRect/></a:stretch>
                  </pic:blipFill>
                  <pic:spPr>
                    <a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>
                    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                  </pic:spPr>
                </pic:pic>
              </a:graphicData>
            </a:graphic>
          </wp:inline>
        </w:drawing>
      </w:r>
    </w:p>
    </w:tc>
  </w:tr>
</w:tbl>'''

def agregar_imagen_al_docx(dst, img_b64, nombre_archivo, rId):
    """Agrega imagen al docx y registra la relación"""
    if not img_b64:
        return False
    
    # Si es lista, tomar el primer elemento
    if isinstance(img_b64, list):
        img_b64 = img_b64[0] if img_b64 else None
    if not img_b64:
        return False
    
    # Detectar tipo de imagen
    if img_b64.startswith('/9j/') or img_b64.startswith('iVBOR'):
        ext = 'png' if img_b64.startswith('iVBOR') else 'jpg'
    else:
        ext = 'png'
    
    content_type = 'image/png' if ext == 'png' else 'image/jpeg'
    
    # Guardar imagen en word/media/
    media_dir = os.path.join(dst, 'word', 'media')
    os.makedirs(media_dir, exist_ok=True)
    img_path = os.path.join(media_dir, f'{nombre_archivo}.{ext}')
    with open(img_path, 'wb') as f:
        f.write(base64.b64decode(img_b64))
    
    # Agregar relación en word/_rels/document.xml.rels
    rels_path = os.path.join(dst, 'word', '_rels', 'document.xml.rels')
    with open(rels_path, 'r', encoding='utf-8') as f:
        rels = f.read()
    
    nueva_rel = f'<Relationship Id="{rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/{nombre_archivo}.{ext}"/>'
    rels = rels.replace('</Relationships>', f'{nueva_rel}</Relationships>')
    
    with open(rels_path, 'w', encoding='utf-8') as f:
        f.write(rels)
    
    # Agregar content type en [Content_Types].xml
    ct_path = os.path.join(dst, '[Content_Types].xml')
    with open(ct_path, 'r', encoding='utf-8') as f:
        ct = f.read()
    
    ext_type = f'<Default Extension="{ext}" ContentType="{content_type}"/>'
    if f'Extension="{ext}"' not in ct:
        ct = ct.replace('</Types>', f'{ext_type}</Types>')
        with open(ct_path, 'w', encoding='utf-8') as f:
            f.write(ct)
    
    return True

def build_docx_pdf(body_xml, tmpdir, nombre, imagenes=None):
    """Inserta body en el membrete, empaqueta como .docx y convierte a PDF con LibreOffice"""

    # 1. Copiar membrete desempaquetado a carpeta temporal
    dst = os.path.join(tmpdir, 'doc')
    shutil.copytree(MEMBRETE, dst)

    # 2. Agregar imágenes si se proporcionaron
    if imagenes:
        for img_info in imagenes:
            agregar_imagen_al_docx(dst, img_info['b64'], img_info['nombre'], img_info['rId'])

    # 3. Leer document.xml del membrete ORIGINAL y sustituir el body
    with open(os.path.join(MEMBRETE, 'word', 'document.xml'), 'r', encoding='utf-8') as f:
        orig = f.read()
    new_doc = orig[:orig.find('<w:background')] + body_xml + '</w:document>'
    with open(os.path.join(dst, 'word', 'document.xml'), 'w', encoding='utf-8') as f:
        f.write(new_doc)

    # 4. Empaquetar carpeta dst como ZIP con extensión .docx
    docx_path = os.path.join(tmpdir, f'{nombre}.docx')
    with zipfile.ZipFile(docx_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(dst):
            for file in files:
                abs_path = os.path.join(root, file)
                arc_name = os.path.relpath(abs_path, dst)
                zf.write(abs_path, arc_name)

    # 5. Convertir .docx a PDF con LibreOffice
    subprocess.run(
        ['soffice', '--headless', '--convert-to', 'pdf', '--outdir', tmpdir, docx_path],
        check=True,
        env={**os.environ, 'HOME': tmpdir}
    )

    return os.path.join(tmpdir, f'{nombre}.pdf')
