import os, re, io, zipfile, datetime
from flask import Flask, render_template, request, send_file, jsonify
import openpyxl
import pytesseract
from pdf2image import convert_from_bytes
from PIL import Image

app = Flask(__name__, template_folder=".")
app.config['MAX_CONTENT_LENGTH'] = 30 * 1024 * 1024

BASE_DIR  = os.path.dirname(os.path.abspath(__file__))
PLANT_DIR = BASE_DIR

# ─────────────────────────────────────────────
# RUTAS
# ─────────────────────────────────────────────

@app.route('/')
def index():
    return render_template('index.html')


@app.route('/generar', methods=['POST'])
def generar():
    try:
        reporte_f   = request.files.get('reporte')
        sanitario_f = request.files.get('sanitario')
        shipment    = request.form.get('shipment_no', '').strip()
        tipo_via    = request.form.get('tipo_via', '').strip()

        errores = []
        if not reporte_f:   errores.append("Falta el Reporte DOC (.xlsx)")
        if not sanitario_f: errores.append("Falta el Sanitario Provisorio (PDF)")
        if not shipment:    errores.append("Falta el número de Shipment")
        if not tipo_via:    errores.append("Seleccioná la vía (Aéreo o Marítimo)")
        if errores:
            return jsonify({'ok': False, 'errores': errores}), 400

        reporte     = leer_reporte(reporte_f, shipment)
        pdf_bytes   = sanitario_f.read()
        datos_prov  = leer_sanitario_provisorio(pdf_bytes, reporte)

        nombre_plantilla = f'Malasia_{"aereo" if tipo_via == "aereo" else "maritimo"}.docx'
        plantilla = os.path.join(PLANT_DIR, nombre_plantilla)
        if not os.path.exists(plantilla):
            return jsonify({'ok': False, 'errores': [
                f'Plantilla "{nombre_plantilla}" no encontrada en el servidor.'
            ]}), 500

        with open(plantilla, 'rb') as f:
            docx_bytes = f.read()

        resultado, alertas = generar_sanitario(docx_bytes, datos_prov, tipo_via)

        nombre_archivo = f"Sanitario_Malasia_{tipo_via}_{shipment}.docx"
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


# ─────────────────────────────────────────────
# LEER REPORTE DOC
# ─────────────────────────────────────────────

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


# ─────────────────────────────────────────────
# OCR DEL SANITARIO PROVISORIO
# ─────────────────────────────────────────────

def ocr_pdf(pdf_bytes):
    """Convierte PDF a texto via pytesseract."""
    images = convert_from_bytes(pdf_bytes, dpi=150)
    texto = ''
    for img in images[:2]:  # solo primeras 2 páginas
        texto += pytesseract.image_to_string(img, lang='spa') + '\n'
    return texto


def limpiar_num(s):
    """Normaliza número: '1.438,00' o '1438,00' → '1438.00'"""
    if not s:
        return None
    s = s.strip().replace(' ', '')
    # formato europeo: punto como miles, coma como decimal
    if re.match(r'^\d{1,3}(\.\d{3})+(,\d+)?$', s):
        s = s.replace('.', '').replace(',', '.')
    else:
        s = s.replace(',', '.')
    try:
        return f'{float(s):.2f}'
    except Exception:
        return s


def buscar(patron, texto, grupo=1, flags=re.IGNORECASE):
    m = re.search(patron, texto, flags)
    return m.group(grupo).strip() if m else None


def leer_sanitario_provisorio(pdf_bytes, reporte):
    texto = ocr_pdf(pdf_bytes)
    datos = {}

    # ── Fechas ───────────────────────────────────────────────────
    # Formato: "Faena: 28/05/2026 al 05/06/2026"
    # o "Faena: 16/12/2025" (sin rango)
    def extraer_fecha_campo(etiqueta):
        m = re.search(
            rf'{etiqueta}[:\s\-—]+(\d{{2}}/\d{{2}}/\d{{4}}(?:\s+al?\s+\d{{2}}/\d{{2}}/\d{{4}})?)',
            texto, re.IGNORECASE
        )
        if m:
            return m.group(1).strip()
        return None

    datos['fecha_faena']      = extraer_fecha_campo(r'Faena')
    datos['fecha_produccion'] = extraer_fecha_campo(r'Producci[oó]n')
    datos['fecha_vencimiento']= extraer_fecha_campo(r'Vencimiento')

    # Limpiar ruido OCR en fechas
    for k in ('fecha_faena', 'fecha_produccion', 'fecha_vencimiento'):
        if datos.get(k):
            # Eliminar caracteres basura al final
            datos[k] = re.sub(r'[^0-9/\s].*$', '', datos[k]).strip()
            datos[k] = re.sub(r'\s+al?\s+', ' al ', datos[k])

    # ── Transporte ───────────────────────────────────────────────
    # Aéreo: "Vuelo de Aerolínea: EMIRATES N°: EK248 de Fecha: 24/01/2026"
    # Marítimo: "Barco: TIGER PLATA 985"
    m_vuelo = re.search(
        r'(?:Vuelo de Aerol[ií]nea|Aerol[ií]nea)[:\s]+([A-Z][A-Z\s]+?)\s*N[°o"*]*[:\s]*([A-Z]{2}\d+)',
        texto, re.IGNORECASE
    )
    m_barco = re.search(
        r'Barco[:\s]+([A-Z][A-Z\s\d]+?)(?:\n|IV\.|$)',
        texto, re.IGNORECASE
    )

    if m_vuelo:
        aerolinea = m_vuelo.group(1).strip()
        numero    = m_vuelo.group(2).strip()
        datos['transporte'] = f'VUELO / FLIGHT: {numero}'
        datos['tipo_detectado'] = 'aereo'
    elif m_barco:
        barco = m_barco.group(1).strip().rstrip('0123456789').strip()
        datos['transporte'] = f'VAPOR / VESSEL: {barco}'
        datos['tipo_detectado'] = 'maritimo'
    else:
        datos['transporte'] = ''
        datos['tipo_detectado'] = 'desconocido'

    # ── Contenedor ───────────────────────────────────────────────
    m_cont = re.search(
        r'Contenedor N[^:]*:\s*([A-Z]{4}[\+\-]?\d{6,7}[\-\d]*)',
        texto, re.IGNORECASE
    )
    if m_cont:
        cont = m_cont.group(1).replace('+', '')
        datos['contenedor'] = cont
    else:
        datos['contenedor'] = None

    # ── Precinto A.F.I.P. ────────────────────────────────────────
    m_afip = re.search(
        r'A\.F[\w\.]+P[^\:]*[:\"\*\s]+([A-Z]{2,3}\d{4,8})',
        texto, re.IGNORECASE
    )
    datos['precinto_afip'] = m_afip.group(1).strip() if m_afip else None

    # ── Pallets y KGS ────────────────────────────────────────────
    m_pallets = re.search(
        r'EN\s+(\d+)\s+PALLETS?[:\s]+([\d\.,]+)\s*KGS?',
        texto, re.IGNORECASE
    )
    if m_pallets:
        datos['pallets']    = m_pallets.group(1)
        datos['kg_pallets'] = limpiar_num(m_pallets.group(2))
    else:
        datos['pallets']    = None
        datos['kg_pallets'] = None

    # ── Totales (cajas, neto, bruto) ─────────────────────────────
    # Los totales suelen aparecer como dos números grandes juntos en el texto
    # cerca de "Totales" o justo antes de "I.10 Tipo de embalaje"
    m_totales = re.search(
        r'(?:Totales?|I\.9)[^\n]*\n[^\n]*?(\d{2,4})\s+([\d\.,]+)\s+([\d\.,]+)',
        texto, re.IGNORECASE | re.DOTALL
    )
    if not m_totales:
        # Buscar el par neto/bruto total: números grandes consecutivos
        # que aparecen solos en una zona sin descripción de producto
        m_totales2 = re.search(
            r'\n\s*(\d{3,6}[,\.]\d{2})\s+(\d{3,6}[,\.]\d{2})\s*\n',
            texto
        )
        if m_totales2:
            datos['total_neto']  = limpiar_num(m_totales2.group(1))
            datos['total_bruto'] = limpiar_num(m_totales2.group(2))
            datos['total_cajas'] = None  # se calculará desde productos
        else:
            datos['total_neto']  = None
            datos['total_bruto'] = None
            datos['total_cajas'] = None
    else:
        datos['total_cajas'] = m_totales.group(1)
        datos['total_neto']  = limpiar_num(m_totales.group(2))
        datos['total_bruto'] = limpiar_num(m_totales.group(3))

    # ── Fecha de emisión ─────────────────────────────────────────
    datos['fecha_emision'] = datetime.datetime.now().strftime('%d/%m/%Y')

    # ── Productos ────────────────────────────────────────────────
    datos['productos'] = extraer_productos(texto, reporte)

    # Si no hay totales, calcular desde productos
    if not datos['total_cajas'] and datos['productos']:
        datos['total_cajas'] = str(sum(
            int(p.get('cajas') or 0) for p in datos['productos']
        ))
    if not datos['total_neto'] and datos['productos']:
        total = sum(float(p.get('neto') or 0) for p in datos['productos'])
        datos['total_neto'] = f'{total:.2f}'
    if not datos['total_bruto'] and datos['productos']:
        total = sum(float(p.get('bruto') or 0) for p in datos['productos'])
        datos['total_bruto'] = f'{total:.2f}'

    return datos


# ─────────────────────────────────────────────
# EXTRACCIÓN DE PRODUCTOS DEL TEXTO OCR
# ─────────────────────────────────────────────

# Nombres clave en español que identifican un producto
NOMBRES_PRODUCTO = [
    'BOLA DE LOMO', 'CUADRADA', 'NALGA DE ADENTRO CON TAPA', 'NALGA CON TAPA',
    'NALGA SIN TAPA', 'NALGA', 'CARNAZA DE PALETA', 'BIFE ANGOSTO',
    'BIFE ANCHO SIN TAPA', 'BIFE ANCHO', 'BIFE DE VACIO GRANDE', 'BIFE DE VACIO',
    'COLITA DE CUADRIL', 'CORAZON DE CUADRIL', 'CORAZON DE PALETA',
    'LOMO SIN CORDON', 'LOMO S/ CORDON',
    'MARUCHA', 'ASADO SIN HUESO', 'PECHO', 'TAPA DE BIFE ANCHO',
    'TAPA DE CUADRIL', 'PECETO', 'AGUJA', 'CHINGOLO', 'PALETA',
    'CARNAZA', 'VACIO', 'ENTRAÑA',
]

# Mapa español → inglés para Malasia
MAPA_EN = {
    'BOLA DE LOMO':              'KNUCKLE',
    'CUADRADA':                  'OUTSIDE FLAT',
    'LOMO SIN CORDON':           'TENDERLOIN CHAIN OFF',
    'LOMO S/ CORDON':            'TENDERLOIN CHAIN OFF',
    'NALGA DE ADENTRO CON TAPA': 'TOPSIDE CAP ON',
    'NALGA CON TAPA':            'TOPSIDE CAP ON',
    'NALGA SIN TAPA':            'TOPSIDE CAP OFF',
    'NALGA':                     'TOPSIDE',
    'CARNAZA DE PALETA':         'BOLAR BLADE',
    'BIFE ANGOSTO':              'STRIPLOIN',
    'BIFE ANCHO SIN TAPA':       'RIBEYE',
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


def buscar_nombre_es(texto_linea):
    """Extrae el nombre corto en español de una línea de descripción SENASA."""
    t = texto_linea.upper()
    # Ordenar por longitud desc para matchear primero los más específicos
    for nombre in sorted(NOMBRES_PRODUCTO, key=len, reverse=True):
        if nombre in t:
            # Si tiene info de LBS, agregarla
            lbs_m = re.search(r'(\d/\d\s*LBS|\+\s*\d\s*LBS)', texto_linea, re.IGNORECASE)
            if lbs_m and 'LOMO' in nombre:
                return f'{nombre} {lbs_m.group(1).strip()}'
            return nombre
    return None


def buscar_nombre_en(nombre_es, descripciones_reporte=None):
    """Busca la traducción al inglés."""
    n = nombre_es.upper()

    # Intentar desde reporte DOC primero
    if descripciones_reporte:
        for code, desc in descripciones_reporte.items():
            for clave in MAPA_EN:
                if clave in n and clave.split()[0] in desc.upper():
                    return desc

    # Fallback al mapa interno (más específico primero)
    for clave, en in sorted(MAPA_EN.items(), key=lambda x: len(x[0]), reverse=True):
        if clave in n:
            return en
    return ''


def armar_nombre_bilingue(nombre_es, nombre_en):
    es = nombre_es.strip().upper()
    en = (nombre_en or '').strip().upper()
    if en and en != es:
        return f'{es}/ {en}'
    return es


def extraer_productos(texto, reporte):
    """
    Extrae los productos del texto OCR del sanitario provisorio.
    Estrategia: buscar líneas que contengan nombres conocidos de productos
    y los números de neto/bruto al final o en la misma línea.
    """
    descripciones = reporte.get('descripciones', {})
    productos = []
    lineas = texto.split('\n')

    i = 0
    while i < len(lineas):
        linea = lineas[i]
        nombre_es = buscar_nombre_es(linea)
        if nombre_es:
            # Buscar cajas: número al inicio de la línea anterior o esta
            cajas = None
            # Buscar número de cajas en la línea anterior (suele estar ahí)
            if i > 0:
                m_cajas = re.match(r'^\s*(\d{1,4})\s*$', lineas[i-1].strip())
                if m_cajas:
                    cajas = m_cajas.group(1)
            # Si no, buscar al inicio de la línea actual
            if not cajas:
                m_cajas = re.match(r'^\s*(\d{1,4})\s+', linea)
                if m_cajas:
                    cajas = m_cajas.group(1)

            # Buscar neto y bruto: dos números con decimales en la línea o siguientes
            neto = bruto = None
            # Buscar en la línea actual
            nums = re.findall(r'\b(\d{1,6}[,\.]\d{2})\b', linea)
            if len(nums) >= 2:
                neto  = limpiar_num(nums[-2])
                bruto = limpiar_num(nums[-1])
            elif len(nums) == 1:
                # Un número en esta línea, buscar el otro en la siguiente
                neto = limpiar_num(nums[0])
                if i + 1 < len(lineas):
                    nums2 = re.findall(r'\b(\d{1,6}[,\.]\d{2})\b', lineas[i+1])
                    if nums2:
                        bruto = limpiar_num(nums2[0])
            else:
                # Buscar en línea siguiente
                if i + 1 < len(lineas):
                    nums2 = re.findall(r'\b(\d{1,6}[,\.]\d{2})\b', lineas[i+1])
                    if len(nums2) >= 2:
                        neto  = limpiar_num(nums2[0])
                        bruto = limpiar_num(nums2[1])

            nombre_en = buscar_nombre_en(nombre_es, descripciones)

            if neto:  # solo agregar si tiene al menos el neto
                productos.append({
                    'cajas':    cajas or '1',
                    'nombre_es': nombre_es,
                    'nombre_en': nombre_en,
                    'neto':     neto,
                    'bruto':    bruto or neto,
                })
        i += 1

    return productos


# ─────────────────────────────────────────────
# GENERACIÓN DEL DOCX
# ─────────────────────────────────────────────

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
    nuevo_bloque = bloque.replace(primer, f'{tag_open}{nuevo_texto}</w:t>', 1)
    for t in textos[1:]:
        tag2 = re.match(r'<w:t[^>]*>', t).group()
        nuevo_bloque = nuevo_bloque.replace(t, f'{tag2}</w:t>', 1)
    return xml_fila[:celda_starts[celda_idx]] + nuevo_bloque + xml_fila[celda_ends[celda_idx]:]


def _construir_fila_aereo(fila_modelo, cajas, nombre_bi, neto, bruto):
    nueva = fila_modelo
    nueva = _reemplazar_celda(nueva, 0, str(cajas))
    nueva = _reemplazar_celda(nueva, 1, nombre_bi)
    nueva = _reemplazar_celda(nueva, 5, str(neto))
    nueva = _reemplazar_celda(nueva, 6, str(bruto))
    return nueva


def _construir_fila_maritimo(fila_modelo, cajas, nombre_bi, neto, bruto):
    nueva = fila_modelo
    nueva = _reemplazar_celda(nueva, 0, str(cajas))
    nueva = _reemplazar_celda(nueva, 1, nombre_bi)
    nueva = _reemplazar_celda(nueva, 6, str(neto))
    nueva = _reemplazar_celda(nueva, 7, str(bruto))
    return nueva


def _reemplazar_bloque_productos(xml, trs, primera_idx, total_idx, productos,
                                  total_cajas, total_neto, total_bruto, tipo_via):
    fila_modelo, ini_mod, _   = get_fila_xml(xml, trs, primera_idx)
    fila_total,  ini_tot, fin_tot = get_fila_xml(xml, trs, total_idx)

    nuevas_filas = ''
    for prod in productos:
        nombre_bi = armar_nombre_bilingue(prod.get('nombre_es',''), prod.get('nombre_en',''))
        if tipo_via == 'aereo':
            nuevas_filas += _construir_fila_aereo(
                fila_modelo, prod.get('cajas',''), nombre_bi,
                prod.get('neto',''), prod.get('bruto','')
            )
        else:
            nuevas_filas += _construir_fila_maritimo(
                fila_modelo, prod.get('cajas',''), nombre_bi,
                prod.get('neto',''), prod.get('bruto','')
            )

    # Actualizar totales
    nums_tot = re.findall(r'<w:t[^>]*>(\d[\d\.]*)</w:t>', fila_total)
    nueva_total = fila_total
    if len(nums_tot) >= 3:
        nueva_total = nueva_total.replace(f'>{nums_tot[0]}<', f'>{total_cajas}<', 1)
        nueva_total = nueva_total.replace(f'>{nums_tot[1]}<', f'>{total_neto}<', 1)
        nueva_total = nueva_total.replace(f'>{nums_tot[2]}<', f'>{total_bruto}<', 1)

    return xml[:ini_mod] + nuevas_filas + nueva_total + xml[fin_tot:]


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
    productos    = datos.get('productos', [])
    total_cajas  = datos.get('total_cajas', '18')
    total_neto   = datos.get('total_neto',  '285.00')
    total_bruto  = datos.get('total_bruto', '340.00')

    trs = get_trs(xml)
    xml = _reemplazar_bloque_productos(
        xml, trs, primera_idx=5, total_idx=19,
        productos=productos,
        total_cajas=total_cajas, total_neto=total_neto, total_bruto=total_bruto,
        tipo_via='aereo'
    )

    # Pallets
    pallets    = datos.get('pallets', '')
    kg_pallets = datos.get('kg_pallets', '')
    if pallets:
        xml = re.sub(r'(ACONDICIONADO EN )\d+( PALLET)', rf'\g<1>{pallets}\2', xml)
        xml = re.sub(r'(ACONDITIONED IN )\d+( PALLET)', rf'\g<1>{pallets}\2', xml)
    if kg_pallets:
        xml = re.sub(r'>32\.44<', f'>{kg_pallets}<', xml)

    # Fechas
    f_faena = datos.get('fecha_faena', '')
    f_prod  = datos.get('fecha_produccion', '')
    f_venc  = datos.get('fecha_vencimiento', '')
    if f_faena: xml = xml.replace('>16/12/2025<', f'>{f_faena}<')
    if f_prod:  xml = xml.replace('>20/12/2025 AL/TO 22/12/2025<', f'>{f_prod}<')
    if f_venc:  xml = xml.replace('>19/04/2026 AL/TO 21/04/2026<', f'>{f_venc}<')

    # Transporte
    transporte = datos.get('transporte', '')
    if transporte:
        xml = xml.replace('>VUELO / FLIGHT: EK248<', f'>{transporte}<')

    # Fecha de emisión: año | mes | día
    fecha_emi = datos.get('fecha_emision') or datetime.datetime.now().strftime('%d/%m/%Y')
    try:
        dia, mes, anio = fecha_emi.split('/')
    except Exception:
        dia = mes = anio = ''
        alertas.append('Fecha de emisión no parseada — completar manualmente')

    xml = xml.replace('>2026<',  f'>{anio}<',  1)
    xml = xml.replace('>01 <',   f'>{mes} <',  1)
    xml = xml.replace('>23<',    f'>{dia}<',   1)

    return xml, alertas


def _gen_maritimo(xml, datos):
    alertas = []
    productos    = datos.get('productos', [])
    total_cajas  = datos.get('total_cajas', '613')
    total_neto   = datos.get('total_neto',  '11895.00')
    total_bruto  = datos.get('total_bruto', '12655.34')

    trs = get_trs(xml)
    xml = _reemplazar_bloque_productos(
        xml, trs, primera_idx=5, total_idx=15,
        productos=productos,
        total_cajas=total_cajas, total_neto=total_neto, total_bruto=total_bruto,
        tipo_via='maritimo'
    )

    # Pallets
    pallets    = datos.get('pallets', '')
    kg_pallets = datos.get('kg_pallets', '')
    if pallets:
        xml = re.sub(r'(ACONDICIONADO EN )\d+( PALLETS)', rf'\g<1>{pallets}\2', xml)
        xml = re.sub(r'(ACONDITIONED IN )\d+( PALLETS)', rf'\g<1>{pallets}\2', xml)
    if kg_pallets:
        xml = re.sub(r'>\([\d\.]+\s*KGS?', f'>({kg_pallets} KGS', xml)

    # Fechas
    f_faena = datos.get('fecha_faena', '')
    f_prod  = datos.get('fecha_produccion', '')
    f_venc  = datos.get('fecha_vencimiento', '')
    if f_faena: xml = xml.replace('>28/05/2026 AL 05/06/2026<', f'>{f_faena}<')
    if f_prod:  xml = xml.replace('>03/06/2026 AL 10/06/2026<', f'>{f_prod}<')
    if f_venc:  xml = xml.replace('>01/10/2026 AL 08/10/2026<', f'>{f_venc}<')

    # Transporte
    transporte = datos.get('transporte', '')
    if transporte:
        xml = xml.replace('>VAPOR / VESSEL: TIGER PLATA<', f'>{transporte}<')

    # Contenedor y precinto
    contenedor    = datos.get('contenedor', '')
    precinto_afip = datos.get('precinto_afip', '')
    if contenedor:    xml = xml.replace('>TCLU129408-4<', f'>{contenedor}<')
    if precinto_afip: xml = xml.replace('>BAH79585<',    f'>{precinto_afip}<')
    if not contenedor:    alertas.append('Contenedor no encontrado — completar manualmente')
    if not precinto_afip: alertas.append('Precinto A.F.I.P. no encontrado — completar manualmente')

    # Fecha de emisión
    fecha_emi = datos.get('fecha_emision') or datetime.datetime.now().strftime('%d/%m/%Y')
    try:
        dia, mes, anio = fecha_emi.split('/')
    except Exception:
        dia = mes = anio = ''
        alertas.append('Fecha de emisión no parseada — completar manualmente')

    xml = xml.replace('>:      2026<', f'>:      {anio}<')
    xml = re.sub(r'>\)\s+\d{2}\s+<', f')         {mes} <', xml, count=1)
    xml = xml.replace('>14<', f'>{dia}<', 1)

    return xml, alertas


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
