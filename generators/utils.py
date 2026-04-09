import os, subprocess, shutil, zipfile, base64

BASE_DIR     = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MEMBRETE_ZIP = os.path.join(BASE_DIR, 'membrete.zip')
MEMBRETE     = os.path.join(BASE_DIR, 'membrete_unpacked')

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

def agregar_imagen_al_docx(dst, img_b64, nombre_archivo, rId, cx_emu=None, cy_emu=None):
    if not img_b64:
        return False
    if isinstance(img_b64, list):
        img_b64 = img_b64[0] if img_b64 else None
    if not img_b64:
        return False
    img_bytes = base64.b64decode(img_b64)
    ext = 'png' if img_b64.startswith('iVBOR') else 'jpg'
    content_type = 'image/png' if ext == 'png' else 'image/jpeg'
    media_dir = os.path.join(dst, 'word', 'media')
    os.makedirs(media_dir, exist_ok=True)
    img_path = os.path.join(media_dir, f'{nombre_archivo}.{ext}')
    with open(img_path, 'wb') as f:
        f.write(img_bytes)
    rels_path = os.path.join(dst, 'word', '_rels', 'document.xml.rels')
    with open(rels_path, 'r', encoding='utf-8') as f:
        rels = f.read()
    nueva_rel = f'<Relationship Id="{rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/{nombre_archivo}.{ext}"/>'
    rels = rels.replace('</Relationships>', f'{nueva_rel}</Relationships>')
    with open(rels_path, 'w', encoding='utf-8') as f:
        f.write(rels)
    ct_path = os.path.join(dst, '[Content_Types].xml')
    with open(ct_path, 'r', encoding='utf-8') as f:
        ct = f.read()
    ext_type = f'<Default Extension="{ext}" ContentType="{content_type}"/>'
    if f'Extension="{ext}"' not in ct:
        ct = ct.replace('</Types>', f'{ext_type}</Types>')
        with open(ct_path, 'w', encoding='utf-8') as f:
            f.write(ct)
    return True


def _imagen_pdf(b64_str, titulo, tmpdir, idx, max_h_pts):
    """
    Genera un PDF de una sola página con la imagen completa escalada.
    max_h_pts: alto máximo disponible en puntos para la imagen.
    """
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas as rl_canvas
    from reportlab.lib.units import cm
    from PIL import Image as PILImage
    import io as _io

    W, H = A4
    MARGIN = 1.8 * cm
    TITLE_H = 18  # pts para barra de título
    AVAIL_W = W - 2 * MARGIN
    AVAIL_H = max_h_pts - TITLE_H - 6  # espacio para imagen

    img_bytes = base64.b64decode(b64_str)
    img = PILImage.open(_io.BytesIO(img_bytes)).convert('RGB')
    iw, ih = img.size

    # Escalar para caber sin recorte
    scale = min(AVAIL_W / iw, AVAIL_H / ih)
    dw = iw * scale
    dh = ih * scale

    img_tmp = os.path.join(tmpdir, f'_img_{idx}.jpg')
    img.save(img_tmp, 'JPEG', quality=92)

    out = os.path.join(tmpdir, f'_img_page_{idx}.pdf')
    c = rl_canvas.Canvas(out, pagesize=(W, max_h_pts))

    # Barra azul de título
    c.setFillColorRGB(0.106, 0.165, 0.290)
    c.rect(MARGIN, max_h_pts - TITLE_H, AVAIL_W, TITLE_H, fill=1, stroke=0)
    c.setFillColorRGB(1, 1, 1)
    c.setFont('Helvetica-Bold', 9)
    c.drawString(MARGIN + 6, max_h_pts - TITLE_H + 5, titulo)

    # Imagen centrada
    x = MARGIN + (AVAIL_W - dw) / 2
    y = (AVAIL_H - dh) / 2
    c.drawImage(img_tmp, x, y, width=dw, height=dh)
    c.save()
    return out


def build_docx_pdf(body_xml, tmpdir, nombre, imagenes=None, evidencias_pdf=None):
    """
    Genera PDF de 4 hojas:
    - Hoja 1: Recepción + Datos (LibreOffice)
    - Hoja 2: Descripción + Mapa (LibreOffice texto + ReportLab imagen superpuesta)
    - Hoja 3: Acciones + Estatus + Observaciones (LibreOffice)
    - Hoja 4: Evidencia 1 + Evidencia 2 (ReportLab, ambas en una página)

    evidencias_pdf: lista de dicts con claves:
      - 'titulo': str
      - 'b64': str (base64)
      - 'tipo': 'mapa' | 'evidencia'
    """
    from pypdf import PdfWriter, PdfReader
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas as rl_canvas
    from reportlab.lib.units import cm
    from PIL import Image as PILImage
    import io as _io

    W, H = A4
    MARGIN = 1.8 * cm

    # 1. Copiar membrete
    dst = os.path.join(tmpdir, 'doc')
    shutil.copytree(MEMBRETE, dst)

    # 3. Sustituir body
    with open(os.path.join(MEMBRETE, 'word', 'document.xml'), 'r', encoding='utf-8') as f:
        orig = f.read()
    new_doc = orig[:orig.find('<w:background')] + body_xml + '</w:document>'
    with open(os.path.join(dst, 'word', 'document.xml'), 'w', encoding='utf-8') as f:
        f.write(new_doc)

    # 4. Empaquetar .docx
    docx_path = os.path.join(tmpdir, f'{nombre}.docx')
    with zipfile.ZipFile(docx_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(dst):
            for file in files:
                abs_path = os.path.join(root, file)
                arc_name = os.path.relpath(abs_path, dst)
                zf.write(abs_path, arc_name)

    # 5. Convertir a PDF con LibreOffice
    subprocess.run(
        ['soffice', '--headless', '--convert-to', 'pdf', '--outdir', tmpdir, docx_path],
        check=True, env={**os.environ, 'HOME': tmpdir}
    )
    main_pdf = os.path.join(tmpdir, f'{nombre}.pdf')

    if not evidencias_pdf:
        return main_pdf

    # Separar mapa y evidencias
    mapa = next((e for e in evidencias_pdf if e.get('tipo') == 'mapa'), None)
    evids = [e for e in evidencias_pdf if e.get('tipo') == 'evidencia']

    reader_main = PdfReader(main_pdf)
    writer = PdfWriter()

    for i, page in enumerate(reader_main.pages):
        page_num = i + 1  # 1-based

        if page_num == 2 and mapa:
            # ── Hoja 2: superponer imagen del mapa sobre la página de LibreOffice ──
            # La página de LO tiene texto hasta ~300pts desde arriba (A4=842)
            # La imagen va desde ~300pts hacia abajo

            # Calcular posición Y donde termina el contenido de texto en hoja 2
            # Encabezado ~100pts + tabla descripción ~170pts + barra "LUGAR" ~20pts = ~290pts
            # desde arriba → en coords PDF (0=abajo): 842 - 290 = 552
            TEXT_BOTTOM_Y = 290  # pts desde arriba donde termina el texto
            IMG_TOP_Y = H - TEXT_BOTTOM_Y  # en coords reportlab (0=abajo)
            AVAIL_H_MAPA = IMG_TOP_Y - MARGIN  # espacio disponible para imagen

            # Generar overlay con solo la imagen del mapa
            overlay_path = os.path.join(tmpdir, '_mapa_overlay.pdf')
            c = rl_canvas.Canvas(overlay_path, pagesize=A4)

            img_bytes = base64.b64decode(mapa['b64'])
            img = PILImage.open(_io.BytesIO(img_bytes)).convert('RGB')
            iw, ih = img.size
            AVAIL_W = W - 2 * MARGIN
            scale = min(AVAIL_W / iw, AVAIL_H_MAPA / ih)
            dw = iw * scale
            dh = ih * scale
            x = MARGIN + (AVAIL_W - dw) / 2
            y = MARGIN + (AVAIL_H_MAPA - dh) / 2

            img_tmp = os.path.join(tmpdir, '_mapa.jpg')
            img.save(img_tmp, 'JPEG', quality=92)
            c.drawImage(img_tmp, x, y, width=dw, height=dh)
            c.save()

            # Fusionar página LO + overlay mapa
            overlay_reader = PdfReader(overlay_path)
            page.merge_page(overlay_reader.pages[0])
            writer.add_page(page)

        else:
            writer.add_page(page)

    # ── Hoja 4: ambas evidencias en una sola página con ReportLab ──
    if evids:
        ev_path = os.path.join(tmpdir, '_evidencias.pdf')
        c = rl_canvas.Canvas(ev_path, pagesize=A4)
        AVAIL_W = W - 2 * MARGIN

        # Dividir la página en 2 mitades iguales
        n = len(evids)
        slot_h = (H - 2 * MARGIN) / n
        TITLE_H = 18

        for idx, ev in enumerate(evids):
            # Y base de este slot (de arriba hacia abajo)
            slot_top = H - MARGIN - idx * slot_h
            slot_bottom = slot_top - slot_h

            # Barra título
            c.setFillColorRGB(0.106, 0.165, 0.290)
            c.rect(MARGIN, slot_top - TITLE_H, AVAIL_W, TITLE_H, fill=1, stroke=0)
            c.setFillColorRGB(1, 1, 1)
            c.setFont('Helvetica-Bold', 9)
            c.drawString(MARGIN + 6, slot_top - TITLE_H + 5, ev.get('titulo', 'EVIDENCIA'))

            # Imagen
            img_bytes = base64.b64decode(ev['b64'])
            img = PILImage.open(_io.BytesIO(img_bytes)).convert('RGB')
            iw, ih = img.size
            avail_img_h = slot_h - TITLE_H - 8
            scale = min(AVAIL_W / iw, avail_img_h / ih)
            dw = iw * scale
            dh = ih * scale
            x = MARGIN + (AVAIL_W - dw) / 2
            y = slot_bottom + (avail_img_h - dh) / 2 + 4

            img_tmp = os.path.join(tmpdir, f'_ev_{idx}.jpg')
            img.save(img_tmp, 'JPEG', quality=92)
            c.drawImage(img_tmp, x, y, width=dw, height=dh)

        c.showPage()
        c.save()

        ev_reader = PdfReader(ev_path)
        for pg in ev_reader.pages:
            writer.add_page(pg)

    final_pdf = os.path.join(tmpdir, f'{nombre}_final.pdf')
    with open(final_pdf, 'wb') as f:
        writer.write(f)
    return final_pdf
