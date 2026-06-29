import os, re, io, zipfile, datetime, tempfile, subprocess
from flask import Flask, render_template, request, send_file, jsonify
import openpyxl
import pytesseract
import fitz
from PIL import Image

app = Flask(__name__, template_folder='.')
app.config['MAX_CONTENT_LENGTH'] = 30 * 1024 * 1024

BASE_DIR  = os.path.dirname(os.path.abspath(__file__))
PLANT_DIR = BASE_DIR

@app.route('/')
def index():
    return render_template('index.html')


@app.route('/generar', methods=['POST'])
def generar():
    try:
        piqueo_f    = request.files.get('piqueo')
        reporte_f   = request.files.get('reporte')
        sanitario_f = request.files.get('sanitario')
        remito_f    = request.files.get('remito')
        shipment    = request.form.get('shipment_no', '').strip()
        tipo_via    = request.form.get('tipo_via', '').strip()

        errores = []
        if not piqueo_f:    errores.append('Falta el Piqueo (.xlsx)')
        if not reporte_f:   errores.append('Falta el Reporte DOC (.xlsx)')
        if not sanitario_f: errores.append('Falta el Sanitario Provisorio (PDF)')
        if not remito_f:    errores.append('Falta el Remito (PDF)')
        if not shipment:    errores.append('Falta el numero de Shipment')
        if not tipo_via:    errores.append('Selecciona la via (Aereo o Maritimo)')
        if errores:
            return jsonify({'ok': False, 'errores': errores}), 400

        datos_piqueo = leer_piqueo(piqueo_f)
        reporte      = leer_reporte(reporte_f, shipment)
        datos_remito = leer_remito(remito_f.read())
        datos_prov   = leer_sanitario_provisorio(sanitario_f.read())

        # Combinar: remito base, fechas del piqueo, pallets del provisorio
        datos = {**datos_remito, **datos_piqueo, **datos_prov}
        for prod in datos.get('productos', []):
            cod = prod.get('codigo', '')
            if cod in reporte.get('descripciones', {}):
                prod['nombre_en'] = reporte['descripciones'][cod]
            else:
                prod['nombre_en'] = buscar_nombre_en(prod.get('nombre_es', ''))

        import glob
        patron_via = 'aereo' if tipo_via == 'aereo' else 'mar'
        candidatos = glob.glob(os.path.join(PLANT_DIR, 'Malasia*' + patron_via + '*.docx'))
        if not candidatos:
            return jsonify({'ok': False, 'errores': [
                'Plantilla para via "' + tipo_via + '" no encontrada. Archivos: ' + str(os.listdir(PLANT_DIR))
            ]}), 500
        plantilla = candidatos[0]

        with open(plantilla, 'rb') as f:
            docx_bytes = f.read()

        resultado, alertas = generar_sanitario(docx_bytes, datos, tipo_via)

        nombre_archivo = 'Sanitario_Malasia_' + tipo_via + '_' + shipment + '.docx'
        resp = send_file(
            io.BytesIO(resultado), as_attachment=True,
            download_name=nombre_archivo,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
        if alertas:
            resp.headers['X-Alertas'] = ' | '.join(alertas)
        return resp

    except Exception as e:
        import traceback
        return jsonify({'ok': False, 'errores': [str(e), traceback.format_exc()]}), 500


# ── PIQUEO (xlsx) - fechas ────────────────────────────────────────────────────

def leer_piqueo(file):
    wb = openpyxl.load_workbook(file)
    ws = wb.active

    faena_min = faena_max = None
    prod_min  = prod_max  = None
    venc_min  = venc_max  = None

    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or not row[1]:
            continue
        fecha_f   = row[7]   # col H = Fecha F (faena)
        fecha_p   = row[6]   # col G = Fecha P (produccion)
        fecha_ven = row[10]  # col K = Fecha Ven (vencimiento)

        if isinstance(fecha_f, datetime.datetime):
            faena_min = min(faena_min, fecha_f) if faena_min else fecha_f
            faena_max = max(faena_max, fecha_f) if faena_max else fecha_f
        if isinstance(fecha_p, datetime.datetime):
            prod_min = min(prod_min, fecha_p) if prod_min else fecha_p
            prod_max = max(prod_max, fecha_p) if prod_max else fecha_p
        if isinstance(fecha_ven, datetime.datetime):
            venc_min = min(venc_min, fecha_ven) if venc_min else fecha_ven
            venc_max = max(venc_max, fecha_ven) if venc_max else fecha_ven

    def fmt_rango(d_min, d_max):
        if not d_min:
            return None
        s = d_min.strftime('%d/%m/%Y')
        if d_max and d_max != d_min:
            s += ' al ' + d_max.strftime('%d/%m/%Y')
        return s

    return {
        'fecha_faena':       fmt_rango(faena_min, faena_max),
        'fecha_produccion':  fmt_rango(prod_min,  prod_max),
        'fecha_vencimiento': fmt_rango(venc_min,  venc_max),
    }


# ── REPORTE DOC (xlsx) ────────────────────────────────────────────────────────

def leer_reporte(file, shipment):
    wb = openpyxl.load_workbook(file)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return {}
    hdr = None
    hdr_idx = 0
    for i, row in enumerate(rows[:5]):
        if row and any(str(v or '').strip() in ('Shipment No', 'Shipment') for v in row):
            hdr = row
            hdr_idx = i
            break
    if hdr is None:
        hdr = rows[0]
    def col(keys):
        for i, h in enumerate(hdr):
            if h and any(k.lower() in str(h).lower() for k in keys):
                return i
        return None
    c_ship = col(['Shipment No', 'Shipment'])
    c_code = col(['Code'])
    c_desc = col(['Description'])
    d = {'descripciones': {}}
    ship_base = shipment.split('-')[0]
    for row in rows[hdr_idx + 1:]:
        if not row or c_ship is None:
            continue
        sv = str(row[c_ship] or '').strip()
        if sv != shipment and not sv.startswith(ship_base):
            continue
        if c_code is not None and c_desc is not None and row[c_code] and row[c_desc]:
            d['descripciones'][str(row[c_code]).strip()] = str(row[c_desc]).strip()
    return d


# ── UTILIDADES NUMÉRICAS ──────────────────────────────────────────────────────

def limpiar_num(s):
    if not s:
        return None
    s = str(s).strip().replace(' ', '')
    if not s:
        return None
    # Primero: N.3decimales remito Devesa: 320.000->320, 343.890->343.89
    if re.match(r'^\d+\.\d{3}$', s):
        entero, dec = s.split('.')
        dec_limpio = dec.rstrip('0')
        s = entero + '.' + dec_limpio if dec_limpio else entero
    # Segundo: europeo miles+coma: 10.020,000
    elif ',' in s and re.match(r'^\d{1,3}(\.\d{3})+(,\d+)?$', s):
        s = s.replace('.', '').replace(',', '.')
    else:
        s = s.replace(',', '.')
    try:
        return '{:.2f}'.format(float(s))
    except Exception:
        return s


# ── REMITO (fitz) ────────────────────────────────────────────────────────────

def leer_remito(pdf_bytes):
    doc = fitz.open(stream=pdf_bytes, filetype='pdf')
    texto = ''
    for page in doc:
        texto += page.get_text()
    doc.close()
    datos = {}
    m_vuelo = re.search(r'Buque/Aerol[ii]nea[:\s]+([A-Z][A-Z\-\d]+)', texto, re.IGNORECASE)
    if m_vuelo:
        transport = m_vuelo.group(1).strip()
        if '-' in transport:
            partes = transport.split('-')
            if len(partes) == 2 and re.match(r'[A-Z]{2}\d+', partes[1]):
                datos['transporte'] = 'VUELO / FLIGHT: ' + partes[1]
            else:
                datos['transporte'] = 'VAPOR / VESSEL: ' + transport
        else:
            datos['transporte'] = 'VAPOR / VESSEL: ' + transport
    m_cont = re.search(r'CONTAINER[:\s]*([A-Z]{4}\d{6,7}-?\d?)', texto, re.IGNORECASE)
    datos['contenedor'] = m_cont.group(1).strip() if m_cont else None
    m_ps = re.search(r'P\.S\.[:\s]+([A-Z0-9/]+)', texto)
    m_pa = re.search(r'P\.A\.[:\s]+([A-Z]{2,3}\d{4,8})', texto)
    datos['precinto_senasa'] = m_ps.group(1).strip() if m_ps else None
    datos['precinto_afip']   = m_pa.group(1).strip() if m_pa else None
    m_pallets = re.search(r'EN\s+(\d+)\s+PALLETS?', texto, re.IGNORECASE)
    datos['pallets'] = m_pallets.group(1) if m_pallets else None
    m_tot_cajas = re.search(r'Total General\s+(\d[\d\.]*)', texto)
    m_tot_neto  = re.search(r'PESO NETO TOTAL[:\s]+([\d\.,]+)', texto)
    m_tot_bruto = re.search(r'PESO BRUTO TOTAL[:\s]+([\d\.,]+)', texto)
    datos['total_cajas'] = m_tot_cajas.group(1).replace('.', '') if m_tot_cajas else None
    datos['total_neto']  = limpiar_num(m_tot_neto.group(1)) if m_tot_neto else None
    datos['total_bruto'] = limpiar_num(m_tot_bruto.group(1)) if m_tot_bruto else None
    productos = []
    lineas = texto.split('\n')
    i = 0
    while i < len(lineas):
        linea = lineas[i].strip()
        if re.match(r'^CD\d+$', linea):
            codigo    = linea
            desc      = lineas[i+1].strip() if i+1 < len(lineas) else ''
            cajas     = lineas[i+2].strip() if i+2 < len(lineas) else ''
            neto_raw  = lineas[i+4].strip() if i+4 < len(lineas) else ''
            bruto_raw = lineas[i+5].strip() if i+5 < len(lineas) else ''
            nombre_es = buscar_nombre_es_remito(desc)
            productos.append({
                'codigo':    codigo,
                'nombre_es': nombre_es,
                'nombre_en': '',
                'cajas':     cajas,
                'neto':      limpiar_num(neto_raw),
                'bruto':     limpiar_num(bruto_raw),
            })
            i += 6
        else:
            i += 1
    datos['productos'] = productos
    return datos


# ── SANITARIO PROVISORIO (OCR - solo pallets y KGS) ──────────────────────────

def ocr_pdf(pdf_bytes):
    texto = ''
    with tempfile.TemporaryDirectory() as tmpdir:
        pdf_path = os.path.join(tmpdir, 'input.pdf')
        with open(pdf_path, 'wb') as f:
            f.write(pdf_bytes)
        out_prefix = os.path.join(tmpdir, 'page')
        subprocess.run(
            ['pdftoppm', '-r', '150', '-l', '1', '-jpeg', pdf_path, out_prefix],
            check=True, capture_output=True
        )
        archivos = sorted([f for f in os.listdir(tmpdir) if f.startswith('page') and f.endswith('.jpg')])
        for nombre in archivos:
            img_path = os.path.join(tmpdir, nombre)
            img = Image.open(img_path)
            texto += pytesseract.image_to_string(img, lang='spa') + '\n'
            img.close()
    return texto


def leer_sanitario_provisorio(pdf_bytes):
    # Solo extraemos pallets y KGS del provisorio via OCR
    # Las fechas vienen del piqueo
    texto = ocr_pdf(pdf_bytes)
    datos = {}
    # Pallets: 'ACONDICIONADA EN 14 PALLETS: 614.02KGS'
    m_kg = re.search(r'EN\s+(\d+)\s+PALLETS?[:\s]+(\d[\d\.]*)', texto, re.IGNORECASE)
    if m_kg:
        datos['pallets_prov']  = m_kg.group(1)   # solo como respaldo
        datos['kg_pallets']    = limpiar_num(m_kg.group(2))
    else:
        datos['kg_pallets'] = None
    datos['fecha_emision'] = datetime.datetime.now().strftime('%d/%m/%Y')
    return datos


# ── MAPAS NOMBRE ─────────────────────────────────────────────────────────────

MAPA_EN = {
    'BOLA DE LOMO':              'KNUCKLE',
    'CUADRADA':                  'OUTSIDE FLAT',
    'LOMO SIN CORDON':           'TENDERLOIN CHAIN OFF',
    'LOMO S/ CORDON':            'TENDERLOIN CHAIN OFF',
    'LOMO SC':                   'TENDERLOIN CHAIN OFF',
    'NALGA DE ADENTRO CON TAPA': 'TOPSIDE CAP ON',
    'NALGA CON TAPA':            'TOPSIDE CAP ON',
    'NALGA SIN TAPA':            'TOPSIDE CAP OFF',
    'NALGA':                     'TOPSIDE',
    'CARNAZA DE PALETA':         'BOLAR BLADE',
    'BIFE ANGOSTO':              'STRIPLOIN',
    'BIFE ANCHO SIN TAPA':       'RIBEYE',
    'BIFE ANCHO ST':             'RIBEYE',
    'BIFE ANCHO':                'RIB EYE',
    'COLITA DE CUADRIL':         'TRI-TIP',
    'CORAZON DE CUADRIL':        'HEART OF RUMP',
    'MARUCHA':                   'OYSTER BLADE',
    'ASADO SIN HUESO':           'SHORT RIB MEAT',
    'PECHO':                     'BRISKET POINT END',
    'TAPA DE BIFE ANCHO':        'RIB CAP',
    'TAPA DE CUADRIL':           'RUMP CAP',
    'BIFE DE VACIO GRANDE':      'FLAP MEAT',
    'BIFE DE VACIO':             'FLANK',
    'PECETO':                    'EYE ROUND',
    'AGUJA':                     'CHUCK',
    'CHINGOLO':                  'CHUCK TENDER',
    'CORAZON DE PALETA':         'SHOULDER CLOD HEART',
}

NOMBRES_PRODUCTO = sorted(MAPA_EN.keys(), key=len, reverse=True)


def buscar_nombre_es_remito(desc):
    d = desc.upper()
    for nombre in NOMBRES_PRODUCTO:
        if nombre in d:
            lbs_m = re.search(r'(\d/\d\s*LBS|\+\s*\d\s*LBS|\+5\s*LBS)', desc, re.IGNORECASE)
            if lbs_m and 'LOMO' in nombre:
                return nombre + ' ' + lbs_m.group(1).strip()
            return nombre
    return desc


def buscar_nombre_en(nombre_es):
    n = nombre_es.upper()
    for clave, en in sorted(MAPA_EN.items(), key=lambda x: len(x[0]), reverse=True):
        if clave in n:
            return en
    return ''


def armar_nombre_bilingue(nombre_es, nombre_en):
    es = nombre_es.strip().upper()
    en = (nombre_en or '').strip().upper()
    if en and en != es:
        return es + '/ ' + en
    return es


# ── GENERACION DOCX ──────────────────────────────────────────────────────────

def get_trs(xml):
    return list(re.finditer(r'<w:tr[ >]', xml))


def get_fila_xml(xml, trs, idx):
    ini = trs[idx].start()
    fin = trs[idx + 1].start() if idx + 1 < len(trs) else len(xml)
    return xml[ini:fin], ini, fin


def _reemplazar_celda(xml_fila, celda_idx, nuevo_texto):
    celda_starts = [m.start() for m in re.finditer(r'<w:tc>', xml_fila)]
    celda_ends   = [m.start() for m in re.finditer(r'</w:tc>', xml_fila)]
    if celda_idx >= len(celda_starts):
        return xml_fila
    bloque = xml_fila[celda_starts[celda_idx]:celda_ends[celda_idx]]
    textos = re.findall(r'<w:t[^>]*>[^<]*</w:t>', bloque)
    if not textos:
        return xml_fila
    primer   = textos[0]
    tag_open = re.match(r'<w:t[^>]*>', primer).group()
    nuevo_bloque = bloque.replace(primer, tag_open + nuevo_texto + '</w:t>', 1)
    for t in textos[1:]:
        tag2 = re.match(r'<w:t[^>]*>', t).group()
        nuevo_bloque = nuevo_bloque.replace(t, tag2 + '</w:t>', 1)
    return xml_fila[:celda_starts[celda_idx]] + nuevo_bloque + xml_fila[celda_ends[celda_idx]:]


def _construir_fila(fila_modelo, cajas, nombre_bi, neto, bruto):
    nueva = fila_modelo
    nueva = _reemplazar_celda(nueva, 0, str(cajas))
    nueva = _reemplazar_celda(nueva, 1, nombre_bi)
    nueva = _reemplazar_celda(nueva, 6, str(neto))
    nueva = _reemplazar_celda(nueva, 7, str(bruto))
    return nueva


def _get_fila_por_contenido(xml, trs, texto_clave):
    for i, m in enumerate(trs):
        ini = m.start()
        fin = trs[i+1].start() if i+1 < len(trs) else len(xml)
        if texto_clave in xml[ini:fin]:
            return xml[ini:fin], ini, fin, i
    return None, None, None, None


def _reemplazar_pallets_en_fila(fila_xml, pallets, kg_pallets):
    # Reemplazar numero en texto largo: 'ACONDICIONADO EN 1 PALLET'
    fila_xml = re.sub(r'(ACONDICIONADO EN )\d+( PALLET)', r'\g<1>' + str(pallets) + r'\2', fila_xml)
    fila_xml = re.sub(r'(ACONDITIONED IN )\d+( PALLET)', r'\g<1>' + str(pallets) + r'\2', fila_xml)
    # Reemplazar el <w:t>1</w:t> separado (sin atributos, es el numero de pallets)
    fila_xml = fila_xml.replace('<w:t>1</w:t>', '<w:t>' + str(pallets) + '</w:t>', 1)
    # Reemplazar kg pallets (32.44)
    if kg_pallets:
        fila_xml = fila_xml.replace('<w:t>32.44</w:t>', '<w:t>' + str(kg_pallets) + '</w:t>')
    return fila_xml

def _reemplazar_bloque_productos(xml, trs, primera_idx, total_idx_fallback,
                                  productos, total_cajas, total_neto, total_bruto,
                                  pallets, kg_pallets):
    fila_pal, ini_pal, fin_pal, idx_pal = _get_fila_por_contenido(xml, trs, 'ACONDICIONADO EN')
    total_idx = (idx_pal + 1) if idx_pal is not None else total_idx_fallback

    fila_modelo, ini_mod, _ = get_fila_xml(xml, trs, primera_idx)
    fila_total, ini_tot, fin_tot = get_fila_xml(xml, trs, total_idx)

    nuevas_filas = ''
    for prod in productos:
        nombre_bi = armar_nombre_bilingue(prod.get('nombre_es', ''), prod.get('nombre_en', ''))
        nuevas_filas += _construir_fila(
            fila_modelo, prod.get('cajas', ''), nombre_bi,
            prod.get('neto', ''), prod.get('bruto', '')
        )

    nueva_pal = _reemplazar_pallets_en_fila(fila_pal, pallets, kg_pallets) if fila_pal else ''

    # Calcular total bruto incluyendo kg pallets
    try:
        total_bruto_final = '{:.2f}'.format(float(total_bruto) + float(kg_pallets or 0))
    except Exception:
        total_bruto_final = total_bruto

    nums_tot = re.findall(r'<w:t[^>]*>(\d[\d\.]*)</w:t>', fila_total)
    nueva_total = fila_total
    if len(nums_tot) >= 3:
        nueva_total = nueva_total.replace('>' + nums_tot[0] + '<', '>' + str(total_cajas) + '<', 1)
        nueva_total = nueva_total.replace('>' + nums_tot[1] + '<', '>' + str(total_neto) + '<', 1)
        nueva_total = nueva_total.replace('>' + nums_tot[2] + '<', '>' + str(total_bruto_final) + '<', 1)

    return xml[:ini_mod] + nuevas_filas + nueva_pal + nueva_total + xml[fin_tot:]


def fmt_fecha_aereo(f):
    if f and ' al ' in f.lower():
        partes = re.split(r'\s+al\s+', f, flags=re.IGNORECASE)
        return partes[0] + ' AL/TO ' + partes[1]
    return f or ''


def fmt_fecha_maritimo(f):
    if f and ' al ' in f.lower():
        partes = re.split(r'\s+al\s+', f, flags=re.IGNORECASE)
        return partes[0] + ' AL ' + partes[1]
    return f or ''


def generar_sanitario(docx_bytes, datos, tipo_via):
    alertas = []
    with zipfile.ZipFile(io.BytesIO(docx_bytes), 'r') as z:
        archivos = {n: z.read(n) for n in z.namelist()}
    xml = archivos['word/document.xml'].decode('utf-8')
    if tipo_via == 'aereo':
        xml, al = _gen_aereo(xml, datos)
    else:
        xml, al = _gen_maritimo(xml, datos)
    alertas.extend(al)
    archivos['word/document.xml'] = xml.encode('utf-8')
    out = io.BytesIO()
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
        for n, d in archivos.items():
            z.writestr(n, d)
    out.seek(0)
    return out.read(), alertas


def _gen_aereo(xml, datos):
    alertas = []
    productos   = datos.get('productos', [])
    total_cajas = datos.get('total_cajas', '')
    total_neto  = datos.get('total_neto', '')
    total_bruto = datos.get('total_bruto', '')
    pallets     = datos.get('pallets', '1')
    kg_pallets  = datos.get('kg_pallets', '')

    trs = get_trs(xml)
    xml = _reemplazar_bloque_productos(
        xml, trs, primera_idx=5, total_idx_fallback=19,
        productos=productos,
        total_cajas=total_cajas, total_neto=total_neto, total_bruto=total_bruto,
        pallets=pallets, kg_pallets=kg_pallets
    )

    f_faena = datos.get('fecha_faena', '')
    f_prod  = datos.get('fecha_produccion', '')
    f_venc  = datos.get('fecha_vencimiento', '')
    if f_faena: xml = xml.replace('>16/12/2025<', '>' + fmt_fecha_aereo(f_faena) + '<')
    if f_prod:  xml = xml.replace('>20/12/2025 AL/TO 22/12/2025<', '>' + fmt_fecha_aereo(f_prod) + '<')
    if f_venc:  xml = xml.replace('>19/04/2026 AL/TO 21/04/2026<', '>' + fmt_fecha_aereo(f_venc) + '<')

    transporte = datos.get('transporte', '')
    if transporte:
        xml = xml.replace('>VUELO / FLIGHT: EK248<', '>' + transporte + '<')

    fecha_emi = datos.get('fecha_emision') or datetime.datetime.now().strftime('%d/%m/%Y')
    try:
        dia, mes, anio = fecha_emi.split('/')
    except Exception:
        dia = mes = anio = ''
        alertas.append('Fecha de emision no parseada - completar manualmente')
    xml = xml.replace('>2026<',  '>' + anio + '<',  1)
    xml = xml.replace('>01 <',   '>' + mes + ' <',  1)
    xml = xml.replace('>23<',    '>' + dia + '<',   1)

    return xml, alertas


def _gen_maritimo(xml, datos):
    alertas = []
    productos   = datos.get('productos', [])
    total_cajas = datos.get('total_cajas', '')
    total_neto  = datos.get('total_neto', '')
    total_bruto = datos.get('total_bruto', '')
    pallets     = datos.get('pallets', '1')
    kg_pallets  = datos.get('kg_pallets', '')

    trs = get_trs(xml)
    xml = _reemplazar_bloque_productos(
        xml, trs, primera_idx=5, total_idx_fallback=15,
        productos=productos,
        total_cajas=total_cajas, total_neto=total_neto, total_bruto=total_bruto,
        pallets=pallets, kg_pallets=kg_pallets
    )

    f_faena = datos.get('fecha_faena', '')
    f_prod  = datos.get('fecha_produccion', '')
    f_venc  = datos.get('fecha_vencimiento', '')
    if f_faena: xml = xml.replace('>28/05/2026 AL 05/06/2026<', '>' + fmt_fecha_maritimo(f_faena) + '<')
    if f_prod:  xml = xml.replace('>03/06/2026 AL 10/06/2026<', '>' + fmt_fecha_maritimo(f_prod) + '<')
    if f_venc:  xml = xml.replace('>01/10/2026 AL 08/10/2026<', '>' + fmt_fecha_maritimo(f_venc) + '<')

    transporte = datos.get('transporte', '')
    if transporte:
        xml = xml.replace('>VAPOR / VESSEL: TIGER PLATA<', '>' + transporte + '<')

    contenedor    = datos.get('contenedor', '')
    precinto_afip = datos.get('precinto_afip', '')
    if contenedor:    xml = xml.replace('>TCLU129408-4<', '>' + contenedor + '<')
    if precinto_afip: xml = xml.replace('>BAH79585<',    '>' + precinto_afip + '<')
    if not contenedor:    alertas.append('Contenedor no encontrado - completar manualmente')
    if not precinto_afip: alertas.append('Precinto A.F.I.P. no encontrado - completar manualmente')

    fecha_emi = datos.get('fecha_emision') or datetime.datetime.now().strftime('%d/%m/%Y')
    try:
        dia, mes, anio = fecha_emi.split('/')
    except Exception:
        dia = mes = anio = ''
        alertas.append('Fecha de emision no parseada - completar manualmente')
    xml = xml.replace('>:      2026<', '>:      ' + anio + '<')
    xml = re.sub(r'>\)\s+\d{2}\s+<', ')         ' + mes + ' <', xml, count=1)
    xml = xml.replace('>14<', '>' + dia + '<', 1)

    return xml, alertas


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
