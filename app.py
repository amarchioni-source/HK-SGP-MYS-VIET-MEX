import os, re, io, zipfile, datetime, tempfile, subprocess, unicodedata
from flask import Flask, render_template, request, send_file, jsonify
import openpyxl
import pytesseract
import fitz
from PIL import Image

app = Flask(__name__, template_folder='.')
app.config['MAX_CONTENT_LENGTH'] = 30 * 1024 * 1024

BASE_DIR  = os.path.dirname(os.path.abspath(__file__))
PLANT_DIR = BASE_DIR


def _normalizar(s):
    """minusculas + sin tildes, para matchear nombres de archivo de plantilla
    sin importar mayusculas/minusculas o acentos (ej. 'SINGAPUR Aéreo.docx')."""
    s = s.lower()
    s = unicodedata.normalize('NFKD', s)
    return ''.join(c for c in s if not unicodedata.combining(c))

@app.route('/')
def index():
    return render_template('index.html')


@app.route('/generar', methods=['POST'])
def generar():
    try:
        piqueo_f     = request.files.get('piqueo')
        reporte_f    = request.files.get('reporte')
        sanitario_f  = request.files.get('sanitario')
        sanitario2_f = request.files.get('sanitario2')  # opcional - algunos destinos (ej. Ecuador) reparten carne/menudencias en 2 certificados separados que van juntos en un solo sanitario final
        remito_f     = request.files.get('remito')
        shipment     = request.form.get('shipment_no', '').strip()
        tipo_via     = request.form.get('tipo_via', '').strip()
        destino      = request.form.get('destino', 'malasia').strip()

        errores = []
        if not piqueo_f:    errores.append('Falta el Piqueo (.xlsx)')
        if not reporte_f:   errores.append('Falta el Reporte DOC (.xlsx)')
        if not sanitario_f: errores.append('Falta el Sanitario Provisorio (PDF)')
        if not remito_f:    errores.append('Falta el Remito (PDF)')
        if not shipment:    errores.append('Falta el numero de Shipment')
        if not tipo_via:    errores.append('Selecciona la via (Aereo o Maritimo)')
        if destino == 'perucongeladomenudencias' and not sanitario2_f:
            errores.append('Perú Congelado + Menudencias necesita el Sanitario Provisorio 2 (de menudencias)')
        if errores:
            return jsonify({'ok': False, 'errores': errores}), 400

        datos_piqueo = leer_piqueo(piqueo_f)
        reporte      = leer_reporte(reporte_f, shipment)
        datos_remito = leer_remito(remito_f.read())

        if destino == 'perucongeladomenudencias':
            # Caso combinado: un solo remito trae carne + menudencia
            # (identificadas por CODIGOS_MENUDENCIA_PERU), pero SENASA exige 2
            # certificados separados, cada uno con su propio provisorio (no se
            # suman como en Ecuador - cada documento lleva SU porcion de
            # pallets/fechas, no el total combinado).
            datos_prov_carne = leer_sanitario_provisorio(sanitario_f.read())
            datos_prov_menud = leer_sanitario_provisorio(sanitario2_f.read())

            todos_productos = datos_remito.get('productos', [])
            productos_carne = [p for p in todos_productos if p.get('codigo') not in CODIGOS_MENUDENCIA_PERU]
            productos_menud = [p for p in todos_productos if p.get('codigo') in CODIGOS_MENUDENCIA_PERU]

            def _totales_subset(productos):
                cajas = sum(float(p.get('cajas') or 0) for p in productos)
                neto  = sum(float(p.get('neto') or 0) for p in productos)
                bruto = sum(float(p.get('bruto') or 0) for p in productos)
                return str(int(cajas)), '{:.2f}'.format(neto), '{:.2f}'.format(bruto)

            def _armar_datos_subset(productos, datos_prov):
                d = dict(datos_remito)
                d['productos'] = productos
                d['total_cajas'], d['total_neto'], d['total_bruto'] = _totales_subset(productos)
                d['fecha_faena']       = datos_prov.get('fecha_faena_prov')
                d['fecha_produccion']  = datos_prov.get('fecha_produccion_prov')
                d['fecha_vencimiento'] = datos_prov.get('fecha_vencimiento_prov')
                d['pallets']     = datos_prov.get('pallets_prov')
                d['kg_pallets']  = datos_prov.get('kg_pallets')
                if datos_remito.get('es_congelado') is not None:
                    d['es_congelado'] = datos_remito['es_congelado']
                d['fecha_emision'] = datos_prov.get('fecha_emision')
                return d

            datos_carne = _armar_datos_subset(productos_carne, datos_prov_carne)
            datos_menud = _armar_datos_subset(productos_menud, datos_prov_menud)

            todos_docx = [f for f in os.listdir(PLANT_DIR) if f.lower().endswith('.docx')]
            cand_congelado = [f for f in todos_docx if 'peru' in _normalizar(f) and 'congelado' in _normalizar(f)]
            cand_menud     = [f for f in todos_docx if 'peru' in _normalizar(f) and 'menudencia' in _normalizar(f)]
            if not cand_congelado or not cand_menud:
                return jsonify({'ok': False, 'errores': [
                    'Faltan plantillas de Peru Congelado y/o Menudencias. Archivos: ' + str(todos_docx)
                ]}), 500

            with open(os.path.join(PLANT_DIR, cand_congelado[0]), 'rb') as f:
                docx_congelado = f.read()
            with open(os.path.join(PLANT_DIR, cand_menud[0]), 'rb') as f:
                docx_menud = f.read()

            res_carne, al_carne = generar_sanitario(docx_congelado, datos_carne, 'maritimo', 'perucongelado')
            res_menud, al_menud = generar_sanitario(docx_menud, datos_menud, 'maritimo', 'perumenudencias')

            buf = io.BytesIO()
            with zipfile.ZipFile(buf, 'w') as zf:
                zf.writestr('Sanitario_Peru_Congelado_' + shipment + '.docx', res_carne)
                zf.writestr('Sanitario_Peru_Menudencias_' + shipment + '.docx', res_menud)
            buf.seek(0)

            resp = send_file(
                buf, as_attachment=True,
                download_name='Sanitario_Peru_CongeladoMenudencias_' + shipment + '.zip',
                mimetype='application/zip'
            )
            todas_alertas = al_carne + al_menud
            if todas_alertas:
                resp.headers['X-Alertas'] = ' | '.join(todas_alertas)
            return resp

        datos_prov   = leer_sanitario_provisorio(sanitario_f.read())

        if sanitario2_f:
            # Se subieron 2 provisorios (ej. Ecuador: uno de carne y otro de
            # menudencias) - se combinan sumando pallets y kg, porque el
            # certificado final los declara juntos.
            datos_prov2 = leer_sanitario_provisorio(sanitario2_f.read())
            try:
                p1 = int(datos_prov.get('pallets_prov') or 0)
                p2 = int(datos_prov2.get('pallets_prov') or 0)
                datos_prov['pallets_prov'] = str(p1 + p2)
            except (TypeError, ValueError):
                pass
            try:
                k1 = float(datos_prov.get('kg_pallets') or 0)
                k2 = float(datos_prov2.get('kg_pallets') or 0)
                datos_prov['kg_pallets'] = '{:.2f}'.format(k1 + k2)
            except (TypeError, ValueError):
                pass
            if datos_prov2.get('es_congelado'):
                datos_prov['es_congelado'] = True

        datos = {**datos_remito, **datos_piqueo, **datos_prov}
        datos['destino'] = destino
        # Pallets: el conteo declarado en el provisorio es mas confiable que el
        # del piqueo (que cuenta pallets de trabajo, no el conteo final del
        # embarque) - se usa como respaldo solo si el remito no lo trae.
        if not datos.get('pallets'):
            datos['pallets'] = datos_prov.get('pallets_prov') or datos_piqueo.get('pallets_piqueo')
        # Congelado: remito tiene prioridad sobre provisorio
        if datos_remito.get('es_congelado') is not None:
            datos['es_congelado'] = datos_remito['es_congelado']
        lotes_map = datos_piqueo.get('lotes_por_producto', {})
        fecha_prod_map  = datos_piqueo.get('fecha_produccion_por_producto', {})
        fecha_faena_map = datos_piqueo.get('fecha_faena_por_producto', {})
        productos_list = datos.get('productos', [])
        for prod in productos_list:
            cod = prod.get('codigo', '')
            if cod in reporte.get('descripciones', {}):
                prod['nombre_en'] = reporte['descripciones'][cod]
            else:
                prod['nombre_en'] = buscar_nombre_en(prod.get('nombre_es', ''))
            prod['lotes'] = lotes_map.get(cod, '')
            prod['fecha_produccion_prod'] = fecha_prod_map.get(cod, '')
            prod['fecha_faena_prod']      = fecha_faena_map.get(cod, '')

        # Cruce de Contramarca (formato USA con anexo) en 3 niveles, porque el
        # OCR del sanitario provisorio viene con ruido (numeros perdidos o mal
        # leidos): 1) por cantidad de cajas cuando matchea unico; 2) por nombre
        # del producto dentro del texto de la linea, para las que no tienen
        # cajas; 3) por eliminacion, si al final queda exactamente 1 producto
        # y 1 codigo valido sin asignar. Los codigos se validan contra la
        # lista real del remito (CONTRAMARCA:C164/65/66/...) para descartar
        # lecturas OCR invalidas (ej. C185 cuando el unico codigo posible es C165).
        lineas_usa = datos_prov.get('lineas_usa', [])
        codigos_validos = set(expandir_contramarcas(datos_remito.get('contramarca') or ''))
        if codigos_validos and len(codigos_validos) < len(productos_list):
            # El campo CONTRAMARCA del remito vino truncado/incompleto (la capa de
            # texto del PDF a veces corta el campo aunque visualmente se vea completo).
            # Como las contramarcas de USA son siempre un rango consecutivo, se infiere
            # el rango completo a partir del primer numero confirmado + la cantidad de
            # productos del envio.
            numeros = sorted(set(int(re.sub(r'\D', '', c)) for c in codigos_validos))
            base = numeros[0]
            codigos_validos = set('C' + str(base + i) for i in range(len(productos_list)))
        if codigos_validos:
            lineas_usa = [l for l in lineas_usa if l['contramarca'] in codigos_validos]

        usados_cod = set()
        usados_idx = set()

        por_cajas = {}
        for l in lineas_usa:
            if l['cajas']:
                por_cajas.setdefault(l['cajas'], []).append(l)
        for i, prod in enumerate(productos_list):
            cand = por_cajas.get(str(prod.get('cajas', '')), [])
            if len(cand) == 1 and cand[0]['contramarca'] not in usados_cod:
                prod['contramarca'] = cand[0]['contramarca']
                usados_cod.add(cand[0]['contramarca'])
                usados_idx.add(i)

        # Cruce por peso neto (nivel 2): el peso de cada linea del provisorio
        # es casi siempre unico por producto, incluso cuando dos productos
        # comparten la misma descripcion (ej. mismo "LOMO SC GF AA MB2+" pero
        # distinta cantidad de kilos). El neto del OCR viene con coma decimal
        # (ej. "816,00"); se normaliza a punto para comparar contra el neto
        # ya parseado del remito.
        def _normalizar_neto(valor):
            if not valor:
                return None
            try:
                return '{:.2f}'.format(float(str(valor).replace('.', '').replace(',', '.')))
            except (TypeError, ValueError):
                return None

        por_neto = {}
        for l in lineas_usa:
            neto_norm = _normalizar_neto(l.get('neto'))
            if neto_norm:
                por_neto.setdefault(neto_norm, []).append(l)
        for i, prod in enumerate(productos_list):
            if i in usados_idx:
                continue
            try:
                neto_prod_norm = '{:.2f}'.format(float(prod.get('neto') or 0))
            except (TypeError, ValueError):
                continue
            cand = [l for l in por_neto.get(neto_prod_norm, []) if l['contramarca'] not in usados_cod]
            if len(cand) == 1:
                prod['contramarca'] = cand[0]['contramarca']
                usados_cod.add(cand[0]['contramarca'])
                usados_idx.add(i)

        for l in lineas_usa:
            if l['contramarca'] in usados_cod:
                continue
            texto_l = l.get('texto', '')
            # Contar cuantos productos AUN SIN ASIGNAR comparten el mismo
            # nombre_es (ej. varios "LOMO SC" identicos por no tener rango de
            # peso que los distinga) - un nombre duplicado no puede resolverse
            # por texto, y ademas NO debe bloquear el cruce de otro producto
            # que si tiene un nombre especifico y distinguible.
            nombres_restantes = [(productos_list[i].get('nombre_es') or '').upper()
                                  for i in range(len(productos_list)) if i not in usados_idx]
            conteo_nombres = {}
            for n in nombres_restantes:
                conteo_nombres[n] = conteo_nombres.get(n, 0) + 1
            cand_idx = [i for i, p in enumerate(productos_list)
                        if i not in usados_idx
                        and (p.get('nombre_es') or '').upper() in texto_l
                        and conteo_nombres.get((p.get('nombre_es') or '').upper(), 0) == 1]
            if len(cand_idx) == 1:
                i = cand_idx[0]
                productos_list[i]['contramarca'] = l['contramarca']
                usados_cod.add(l['contramarca'])
                usados_idx.add(i)

        restantes_idx = [i for i in range(len(productos_list)) if i not in usados_idx]
        restantes_cod = [c for c in codigos_validos if c not in usados_cod]
        if len(restantes_idx) == 1 and len(restantes_cod) == 1:
            productos_list[restantes_idx[0]]['contramarca'] = restantes_cod[0]

        PATRONES_DESTINO = {
            'malasia':          'alasia',
            'singapur':         'ingapur',
            'mexico':           'exico',
            'usawclass':        'class',
            'usaorleans':       'orleans',
            'hongkongcongelado': 'congelado',
            'hongkongenfriado':  'enfriado',
            'filipinas':         'filipinas',
            'ecuador':           'ecuador',
            'egipto':            'egipto',
            'dubai':             'dubai',
            'brasil':            'brasil',
            'peruenfriado':      ('peru', 'enfriado'),
            'perumenudencias':   ('peru', 'menudencia'),
            'perucongelado':     ('peru', 'congelado'),
            'usaallecondimentada': 'condimentada',
            'usaallenatural':      'natural',
        }
        patron_via = 'aereo' if tipo_via == 'aereo' else 'mar'
        patron_dest = PATRONES_DESTINO.get(destino, PATRONES_DESTINO['malasia'])
        todos_docx = [f for f in os.listdir(PLANT_DIR) if f.lower().endswith('.docx')]

        def _coincide(nombre_normalizado, patron):
            # patron puede ser un string simple, o una tupla de palabras que
            # deben estar TODAS presentes (para evitar colisiones entre
            # destinos que comparten una sola palabra, ej. "enfriado" en
            # Hong Kong y en Peru)
            if isinstance(patron, tuple):
                return all(p in nombre_normalizado for p in patron)
            return patron in nombre_normalizado

        candidatos = [f for f in todos_docx if _coincide(_normalizar(f), patron_dest) and patron_via in _normalizar(f)]
        if not candidatos:
            candidatos = [f for f in todos_docx if _coincide(_normalizar(f), patron_dest)]
        if not candidatos:
            return jsonify({'ok': False, 'errores': [
                'Plantilla no encontrada para destino=' + destino + ' via=' + tipo_via + '. Archivos: ' + str(os.listdir(PLANT_DIR))
            ]}), 500
        plantilla = os.path.join(PLANT_DIR, candidatos[0])

        with open(plantilla, 'rb') as f:
            docx_bytes = f.read()

        resultado, alertas = generar_sanitario(docx_bytes, datos, tipo_via, destino)

        nombre_archivo = 'Sanitario_' + destino.capitalize() + '_' + tipo_via + '_' + shipment + '.docx'
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


# ── PIQUEO ───────────────────────────────────────────────────────────────────

def leer_piqueo(file):
    wb = openpyxl.load_workbook(file, data_only=True)

    # Buscar entre TODAS las hojas la que tenga headers de producto (Cod Prod/Codigo).
    # No asumir que la hoja activa (wb.active) es la correcta - puede haber hojas
    # sueltas/borrador (ej. "Hoja1") marcadas como activas por accidente.
    ws = None
    rows = None
    hdr_idx = None
    for sheet in wb.worksheets:
        rows_tmp = list(sheet.iter_rows(values_only=True))
        for i, row in enumerate(rows_tmp[:5]):
            if not row: continue
            valores = [str(v or '').strip() for v in row]
            tiene_cod   = any(v in ('Cod Prod', 'Codigo') for v in valores)
            tiene_fecha = any(v in ('Fecha F', 'Fecha P', 'Fecha Ven') for v in valores)
            if tiene_cod and tiene_fecha:
                ws, rows, hdr_idx = sheet, rows_tmp, i
                break
        if ws is not None:
            break

    if ws is None:
        # Fallback: comportamiento anterior (hoja activa) por si ninguna calza
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
        hdr_idx = 0
        for i, row in enumerate(rows[:5]):
            if row and any(str(v or '').strip() in ('Cod Prod', 'Producto', 'Fecha F', 'Fecha P') for v in row):
                hdr_idx = i
                break

    hdr = rows[hdr_idx]

    def col_idx(nombres):
        for i, h in enumerate(hdr):
            if h and any(n.lower() in str(h).lower() for n in nombres):
                return i
        return None

    c_cod     = col_idx(['Cod Prod', 'Codigo'])
    c_fecha_f = col_idx(['Fecha F'])
    c_fecha_p = col_idx(['Fecha P'])
    c_fecha_v = col_idx(['Fecha Ven', 'Vencimiento'])
    c_pallet  = col_idx(['Pallet'])

    faena_min = faena_max = None
    prod_min  = prod_max  = None
    venc_min  = venc_max  = None
    pallets_set = set()
    lotes_por_cod = {}        # Cod Prod -> set de fechas 'Fecha P' (YYYYMMDD), para Nº de lotes (ej. Mexico)
    prod_fechas_por_cod = {}   # Cod Prod -> [fecha_p_min, fecha_p_max] (rango de produccion por producto, ej. Orleans)
    faena_fechas_por_cod = {}  # Cod Prod -> [fecha_f_min, fecha_f_max] (rango de faena por producto, ej. Orleans)

    for row in rows[hdr_idx + 1:]:
        if not row: continue
        cod = row[c_cod] if c_cod is not None and c_cod < len(row) else None
        if not cod: continue
        cod = str(cod).strip()

        if c_pallet is not None and c_pallet < len(row) and row[c_pallet]:
            pallets_set.add(str(row[c_pallet]))

        fecha_f = row[c_fecha_f] if c_fecha_f is not None and c_fecha_f < len(row) else None
        fecha_p = row[c_fecha_p] if c_fecha_p is not None and c_fecha_p < len(row) else None
        fecha_v = row[c_fecha_v] if c_fecha_v is not None and c_fecha_v < len(row) else None

        if isinstance(fecha_f, datetime.datetime):
            faena_min = min(faena_min, fecha_f) if faena_min else fecha_f
            faena_max = max(faena_max, fecha_f) if faena_max else fecha_f
            fmin, fmax = faena_fechas_por_cod.get(cod, (fecha_f, fecha_f))
            faena_fechas_por_cod[cod] = (min(fmin, fecha_f), max(fmax, fecha_f))
        if isinstance(fecha_p, datetime.datetime):
            prod_min = min(prod_min, fecha_p) if prod_min else fecha_p
            prod_max = max(prod_max, fecha_p) if prod_max else fecha_p
            lotes_por_cod.setdefault(cod, set()).add(fecha_p.strftime('%Y%m%d'))
            pmin, pmax = prod_fechas_por_cod.get(cod, (fecha_p, fecha_p))
            prod_fechas_por_cod[cod] = (min(pmin, fecha_p), max(pmax, fecha_p))
        if isinstance(fecha_v, datetime.datetime):
            venc_min = min(venc_min, fecha_v) if venc_min else fecha_v
            venc_max = max(venc_max, fecha_v) if venc_max else fecha_v

    def fmt_rango(d_min, d_max):
        if not d_min: return None
        s = d_min.strftime('%d/%m/%Y')
        if d_max and isinstance(d_max, datetime.datetime) and d_max != d_min:
            s += ' al ' + d_max.strftime('%d/%m/%Y')
        return s

    lotes_por_producto = {
        cod: ' - '.join(sorted(fechas)) for cod, fechas in lotes_por_cod.items()
    }
    fecha_produccion_por_producto = {
        cod: fmt_rango(pmin, pmax) for cod, (pmin, pmax) in prod_fechas_por_cod.items()
    }
    fecha_faena_por_producto = {
        cod: fmt_rango(fmin, fmax) for cod, (fmin, fmax) in faena_fechas_por_cod.items()
    }

    return {
        'fecha_faena':       fmt_rango(faena_min, faena_max),
        'fecha_produccion':  fmt_rango(prod_min,  prod_max),
        'fecha_vencimiento': fmt_rango(venc_min,  venc_max),
        'pallets_piqueo':    str(len(pallets_set)) if pallets_set else None,
        'lotes_por_producto': lotes_por_producto,
        'fecha_produccion_por_producto': fecha_produccion_por_producto,
        'fecha_faena_por_producto': fecha_faena_por_producto,
    }


# ── REPORTE DOC ──────────────────────────────────────────────────────────────

def leer_reporte(file, shipment):
    wb = openpyxl.load_workbook(file)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows: return {}
    hdr = None
    hdr_idx = 0
    for i, row in enumerate(rows[:5]):
        if row and any(str(v or '').strip() in ('Shipment No', 'Shipment') for v in row):
            hdr = row; hdr_idx = i; break
    if hdr is None: hdr = rows[0]
    def col(keys):
        # Preferir coincidencia EXACTA del nombre de columna antes que por
        # substring - la planilla real tiene columnas como "Shipment
        # Description" que tambien contienen la palabra "Description" y
        # confundian la busqueda con la columna real de descripcion del
        # producto.
        for i, h in enumerate(hdr):
            if h and str(h).strip().lower() in [k.lower() for k in keys]:
                return i
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
        if not row or c_ship is None: continue
        sv = str(row[c_ship] or '').strip()
        if sv != shipment and not sv.startswith(ship_base): continue
        if c_code is not None and c_desc is not None and row[c_code] and row[c_desc]:
            d['descripciones'][str(row[c_code]).strip()] = str(row[c_desc]).strip()
    return d


# ── UTILIDADES NUMÉRICAS ─────────────────────────────────────────────────────

def expandir_contramarcas(campo):
    """Expande el formato abreviado de contramarcas del remito, ej.
    'C164/65/66/67/68/69/70/71/72/73/74' -> ['C164','C165',...,'C174'].
    Sirve como lista de codigos validos para validar/corregir lo leido por OCR
    en el sanitario provisorio (que a veces confunde digitos, ej. C165->C185)."""
    if not campo:
        return []
    partes = [p.strip() for p in re.split(r'[/,]', campo) if p.strip()]
    if not partes:
        return []
    primero = re.sub(r'\D', '', partes[0])
    if not primero:
        return []
    base_len = len(primero)
    numeros = [int(primero)]
    for p in partes[1:]:
        p_digits = re.sub(r'\D', '', p)
        if not p_digits:
            continue
        if len(p_digits) < base_len:
            prefijo = primero[:base_len - len(p_digits)]
            numeros.append(int(prefijo + p_digits))
        else:
            numeros.append(int(p_digits))
    return ['C' + str(n) for n in numeros]


def _es_numero(s):
    """True si s es un numero parseable (ej. '2504.00'), usado para validar
    filas de producto del remito antes de aceptarlas."""
    if not s:
        return False
    try:
        float(s)
        return True
    except (TypeError, ValueError):
        return False


def limpiar_num(s):
    if not s: return None
    s = str(s).strip().replace(' ', '')
    if not s: return None
    if re.match(r'^\d+\.\d{3}$', s):
        entero, dec = s.split('.')
        dec_limpio = dec.rstrip('0')
        s = entero + '.' + dec_limpio if dec_limpio else entero
    elif ',' in s and re.match(r'^\d{1,3}(\.\d{3})+(,\d+)?$', s):
        s = s.replace('.', '').replace(',', '.')
    else:
        s = s.replace(',', '.')
    try: return '{:.2f}'.format(float(s))
    except Exception: return s


def formatear_miles(valor):
    """Convierte un numero (string con punto decimal, ej '21692.00') al formato
    con separador de miles por punto y decimales con coma (ej '21.692,00'),
    usado en el certificado de Mexico."""
    try:
        f = float(valor)
    except (TypeError, ValueError):
        return valor
    entero, dec = '{:.2f}'.format(f).split('.')
    signo = ''
    if entero.startswith('-'):
        signo = '-'
        entero = entero[1:]
    entero_fmt = '{:,}'.format(int(entero)).replace(',', '.')
    return signo + entero_fmt + ',' + dec


def formatear_miles_en(valor):
    """Convierte un numero al formato ingles/EEUU: separador de miles por
    coma y decimales con punto (ej '24,833.00'), usado en Filipinas."""
    try:
        f = float(valor)
    except (TypeError, ValueError):
        return valor
    return '{:,.2f}'.format(f)


# ── REMITO (fitz) ────────────────────────────────────────────────────────────

def leer_remito(pdf_bytes):
    doc = fitz.open(stream=pdf_bytes, filetype='pdf')
    texto = ''
    for page in doc: texto += page.get_text()
    doc.close()
    datos = {}
    m_vuelo = re.search(r'Buque/Aerol[íi]nea[:\s]+([^\n\r]+)', texto, re.IGNORECASE)
    if m_vuelo:
        transport = m_vuelo.group(1).strip()
        if '-' in transport:
            partes = transport.split('-')
            if len(partes) == 2 and re.match(r'[A-Z]{2}\d+', partes[1]):
                datos['transporte'] = partes[1]
                datos['tipo_transporte'] = 'aereo'
            else:
                datos['transporte'] = transport
                datos['tipo_transporte'] = 'maritimo'
        else:
            datos['transporte'] = transport
            datos['tipo_transporte'] = 'maritimo'
    m_cont = re.search(r'CONTAINER[:\s]*([A-Z]{4}\d{6,7}-?\d?)', texto, re.IGNORECASE)
    contenedor = m_cont.group(1).strip() if m_cont else None
    if not contenedor:
        # Algunos PDF separan la etiqueta "CONTAINER:" de su valor al extraer
        # el texto (el numero queda flotando en otra parte del documento) -
        # como respaldo, se busca un codigo con forma de contenedor suelto.
        m_cont_suelto = re.search(r'\b([A-Z]{4}\d{7})\b', texto)
        if m_cont_suelto:
            contenedor = m_cont_suelto.group(1)
    if contenedor and '-' not in contenedor:
        # Formato estandar ISO 6346: 4 letras + 6 digitos de serie + guion + 1 digito verificador
        m_iso = re.match(r'^([A-Z]{4}\d{6})(\d)$', contenedor)
        if m_iso:
            contenedor = m_iso.group(1) + '-' + m_iso.group(2)
    datos['contenedor'] = contenedor
    m_ps = re.search(r'P\.S\.[:\s]+([A-Z0-9/]+)', texto)
    m_pa = re.search(r'P\.A\.[:\s]+([A-Z]{2,3}\s?\d{4,8}(?:/\d+)*)', texto)
    datos['precinto_senasa'] = m_ps.group(1).strip() if m_ps else None
    datos['precinto_afip']   = m_pa.group(1).strip() if m_pa else None
    m_camion = re.search(r'CAMION/ACOPLADO[:\s]*([A-Z0-9]+)\s*/\s*([A-Z0-9]*)[ \t]*$', texto, re.IGNORECASE | re.MULTILINE)
    if m_camion:
        camion   = m_camion.group(1).strip()
        acoplado = m_camion.group(2).strip()
        datos['camion']   = camion
        datos['acoplado'] = acoplado or None
        datos['camion_acoplado'] = (camion + '/ ' + acoplado) if acoplado else camion
    else:
        datos['camion'] = datos['acoplado'] = datos['camion_acoplado'] = None
    m_contra = re.search(r'CONTRAMARCA[:\s]+([^\n\r]+)', texto, re.IGNORECASE)
    contra = m_contra.group(1).strip() if m_contra else ''
    datos['contramarca'] = contra if contra else None
    m_marca = re.search(r'(?<!CONTRA)MARCA[:\s]+([^\n\r]+)', texto, re.IGNORECASE)
    datos['marca'] = m_marca.group(1).strip() if m_marca else None
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
        if re.match(r'^[A-Z]{1,4}\d+$', linea):
            codigo    = linea
            desc      = lineas[i+1].strip() if i+1 < len(lineas) else ''
            cajas     = lineas[i+2].strip() if i+2 < len(lineas) else ''
            neto_raw  = lineas[i+4].strip() if i+4 < len(lineas) else ''
            bruto_raw = lineas[i+5].strip() if i+5 < len(lineas) else ''
            neto  = limpiar_num(neto_raw)
            bruto = limpiar_num(bruto_raw)
            # Validar que realmente sea una fila de producto (cantidad de cajas
            # y pesos numericos) antes de aceptarla - evita falsos positivos
            # como un numero de contenedor suelto al final del PDF que por
            # casualidad matchea el patron de codigo (ej. "SEKU9270994").
            es_fila_valida = bool(re.match(r'^\d+$', cajas)) and _es_numero(neto) and _es_numero(bruto)
            if es_fila_valida:
                nombre_es = buscar_nombre_es_remito(desc)
                productos.append({
                    'codigo': codigo, 'nombre_es': nombre_es, 'nombre_en': '',
                    'desc_original': desc,
                    'cajas': cajas, 'neto': neto, 'bruto': bruto,
                })
                i += 6
            else:
                i += 1
        else:
            i += 1
    datos['productos'] = productos

    # Detectar congelado desde observaciones del remito
    if re.search(r'CONGELAD', texto, re.IGNORECASE):
        datos['es_congelado'] = True
    else:
        datos['es_congelado'] = False

    return datos


# ── SANITARIO PROVISORIO (OCR) ───────────────────────────────────────────────

def ocr_pdf(pdf_bytes):
    texto = ''
    with tempfile.TemporaryDirectory() as tmpdir:
        pdf_path = os.path.join(tmpdir, 'input.pdf')
        with open(pdf_path, 'wb') as f: f.write(pdf_bytes)
        out_prefix = os.path.join(tmpdir, 'page')
        subprocess.run(['pdftoppm', '-r', '300', '-l', '1', '-jpeg', pdf_path, out_prefix],
            check=True, capture_output=True)
        archivos = sorted([f for f in os.listdir(tmpdir) if f.startswith('page') and f.endswith('.jpg')])
        for nombre in archivos:
            img = Image.open(os.path.join(tmpdir, nombre))
            texto += pytesseract.image_to_string(img, lang='spa') + '\n'
            img.close()
    return texto


def leer_sanitario_provisorio(pdf_bytes):
    texto = ocr_pdf(pdf_bytes)
    datos = {}
    m_kg = re.search(r'EN\s+(\d+)\s+PALLETS?[:\s]+([\d\.,]+)', texto, re.IGNORECASE)
    if m_kg:
        datos['pallets_prov'] = m_kg.group(1)
        datos['kg_pallets']   = limpiar_num(m_kg.group(2))
    else:
        datos['kg_pallets'] = None
    # Detectar congelado vs enfriado
    if re.search(r'CONGELAD', texto, re.IGNORECASE):
        datos['es_congelado'] = True
    else:
        datos['es_congelado'] = False
    datos['fecha_emision'] = datetime.datetime.now().strftime('%d/%m/%Y')

    # Fechas resumen del provisorio (faena/produccion/vencimiento), ancladas al
    # inicio de linea para no confundirlas con las fechas de faena POR PRODUCTO
    # que aparecen embebidas en cada renglon como "(F. faena: ...)". Se usan
    # cuando hay que separar un envio en 2 documentos (ej. Peru
    # congelado+menudencias) y cada documento necesita el rango de SU propio
    # provisorio, no el agregado de todo el piqueo.
    m_faena_p = re.search(r'^\s*Faena\s*:\s*(\d{2}/\d{2}/\d{4})\s*al\s*(\d{2}/\d{2}/\d{4})', texto, re.IGNORECASE | re.MULTILINE)
    if m_faena_p:
        datos['fecha_faena_prov'] = m_faena_p.group(1) + ' al ' + m_faena_p.group(2)
    m_prod_p = re.search(r'^\s*(?:1\.11\s*Fecha\s*)?Producci[oó]n\s*:\s*(\d{2}/\d{2}/\d{4})\s*al\s*(\d{2}/\d{2}/\d{4})', texto, re.IGNORECASE | re.MULTILINE)
    if m_prod_p:
        datos['fecha_produccion_prov'] = m_prod_p.group(1) + ' al ' + m_prod_p.group(2)
    m_venc_p = re.search(r'^\s*Vencimiento\s*:\s*(\d{2}/\d{2}/\d{4})\s*al\s*(\d{2}/\d{2}/\d{4})', texto, re.IGNORECASE | re.MULTILINE)
    if m_venc_p:
        datos['fecha_vencimiento_prov'] = m_venc_p.group(1) + ' al ' + m_venc_p.group(2)

    # Patente de transporte (a veces con 2 chapas, ej. Peru: "Patente Transporte:
    # ADK931 / Z1V990") - el remito a veces solo trae la primera. El texto
    # entre "patente" y los valores tolera basura de OCR (ej. "patente S NE
    # ADLSSS /BAG9TS" en vez de "patente N°: ADL886 /BAG975") - au ncuando el
    # OCR lea mal algun caracter de la patente en si, al menos la ESTRUCTURA
    # (dos patentes separadas por "/") se reconoce.
    m_patente = re.search(
        r'(?:Patente\s*Transporte|Cami[oó]n\s*patente).{0,20}?([A-Z0-9]{5,8})\s*/\s*([A-Z0-9]{5,8})',
        texto, re.IGNORECASE
    )
    if m_patente:
        datos['patente1'] = m_patente.group(1).strip()
        datos['patente2'] = m_patente.group(2).strip()
    else:
        datos['patente1'] = datos['patente2'] = None

    # Contramarca por linea, en el orden en que aparecen (formato USA con anexo,
    # ej. "55 ... - C208 ( Fecha de Faena: ... )"). El OCR a veces confunde la
    # 'C' del codigo con otro caracter (0, 9, etc.) - se ignora ese caracter y
    # se reconstruye siempre con 'C' + los digitos. El numero de cajas al
    # principio de la linea es opcional porque el OCR a veces lo pierde del
    # todo (ej. una linea entera sin numero visible). La fecha de faena de esta
    # linea NO se usa (viene poco confiable del OCR) - se usa la del piqueo por
    # producto en su lugar; esto solo ancla el match a filas de producto reales
    # (evita matchear el resumen "CONTRAMARCA:C208/C209/...").
    # Tambien captura el peso neto/bruto que sigue despues del parentesis de
    # fecha - son casi siempre unicos por producto y sirven para desambiguar
    # cuando el nombre y la cantidad de cajas no alcanzan (ver el cruce en 3+1
    # niveles en el route /generar). El label de fecha varia entre "Fecha de
    # Faena" y la forma abreviada "F. faena" (ej. para menudencias).
    lineas_usa = []
    patron_linea = re.compile(
        r'^[\s|_\[\]]*(\d+)?([^\n]*?)-\s*[A-Za-z0-9](\d{2,4})\s*\(\s*(?:Fecha\s*de\s*Faena|F\.\s*faena)\s*:\s*'
        r'\d{2}/\d{2}/\d{4}(?:\s*al\s*\d{2}/\d{2}/\d{4})?\s*\)?'
        r'(?:[ \t]*([\d.,]+)(?:[ \t]*\|?[ \t]*([\d.,]+))?)?',
        re.IGNORECASE | re.MULTILINE
    )
    for m in patron_linea.finditer(texto):
        lineas_usa.append({
            'cajas': m.group(1), 'texto': (m.group(2) or '').upper(), 'contramarca': 'C' + m.group(3),
            'neto': m.group(4), 'bruto': m.group(5),
        })
    datos['lineas_usa'] = lineas_usa

    return datos


# ── MAPAS NOMBRE ─────────────────────────────────────────────────────────────

MAPA_EN = {
    'BOLA DE LOMO':              'KNUCKLE',
    'CUADRADA':                  'OUTSIDE FLAT',
    'LOMO SIN CORDON':           'TENDERLOIN CHAIN OFF',
    'LOMO S/ CORDON':            'TENDERLOIN CHAIN OFF',
    'LOMO SC':                   'TENDERLOIN CHAIN OFF',
    'LOMO CON CORDON':           'BEEF TENDERLOIN',
    'LOMO C/ CORDON':            'BEEF TENDERLOIN',
    'LOMO CC':                   'BEEF TENDERLOIN',
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
    'BRAZUELO':                  'CONICAL MUSCLE',
    'COGOTE':                    'NECK',
    'PECHO PEDO':                'BRISKET POINT END',
    'CABEZA DE LOMO':            'TENDERLOIN BUTT',
    'BIFE ANGOSTO CON HUESO':    'SHORTLOIN',
    'CARNAZA':                   'BOLAR BLADE',
}

NOMBRES_PRODUCTO = sorted(MAPA_EN.keys(), key=len, reverse=True)


def extraer_calibre(desc):
    """Extrae el calibre/talle de una descripcion de remito (ej. '3/4 LBS',
    '+5 LBS', '-1,3 KG', '+1,3 KG'), para que no se pierda al simplificar el
    nombre del producto. Aplica a cualquier corte, no solo LOMO - antes solo
    se agregaba el calibre para LOMO, perdiendose en cortes como TAPA DE
    CUADRIL (que puede venir en variantes -1,3 KG / +1,3 KG)."""
    m = re.search(r'([+\-]\s*\d+(?:[/,.]\d+)?\s*(?:KG|LBS)|\d+/\d+\s*(?:KG|LBS))', desc, re.IGNORECASE)
    return m.group(1).strip() if m else ''


def buscar_nombre_es_remito(desc):
    d = desc.upper()
    for nombre in NOMBRES_PRODUCTO:
        if nombre in d:
            calibre = extraer_calibre(desc)
            if calibre:
                return nombre + ' ' + calibre
            return nombre
    return desc


def buscar_nombre_en(nombre_es):
    n = nombre_es.upper()
    calibre_m = re.search(r'([+\-]\s*\d+(?:[/,.]\d+)?\s*(?:KG|LBS)|\d+/\d+\s*(?:KG|LBS))', n, re.IGNORECASE)
    calibre = calibre_m.group(1).strip() if calibre_m else ''
    for clave, en in sorted(MAPA_EN.items(), key=lambda x: len(x[0]), reverse=True):
        if clave in n:
            return (en + ' ' + calibre).strip() if calibre else en
    return ''


def armar_nombre_bilingue(nombre_es, nombre_en):
    es = nombre_es.strip().upper()
    en = (nombre_en or '').strip().upper()
    if en and en != es: return es + '/ ' + en
    return es


# ── NOMBRES ESPECIFICOS MÉXICO (corte / "pulpa" / ingles) ────────────────────
# Tabla fija por tipo de corte, requerida por la certificacion mexicana.
# No todos los cortes llevan calificador de "pulpa" (ver TAPA DE CUADRIL,
# BIFE ANGOSTO, BIFE ANCHO en los ejemplos - no llevan).
MAPA_MEXICO = {
    'NALGA DE AFUERA':  {'es': 'NALGA DE AFUERA CT', 'pulpa': 'PULPA BLANCA', 'en': 'BEEF GOOSENECK'},
    'NALGA CON TAPA':   {'es': 'NALGA CON TAPA',      'pulpa': 'PULPA NEGRA', 'en': 'BEEF TOP (INSIDE) ROUND'},
    'BOLA DE LOMO':     {'es': 'BOLA DE LOMO',        'pulpa': 'PULPA BOLA',  'en': 'BONELESS BEEF KNUCKLE'},
    'TAPA DE CUADRIL':  {'es': 'TAPA DE CUADRIL',     'pulpa': None,          'en': 'BONELESS BEEF RUMP CAP'},
    'BIFE ANGOSTO':     {'es': 'BIFE ANGOSTO CC',     'pulpa': None,          'en': 'BONELESS BEEF NEW YORK'},
    'BIFE ANCHO':       {'es': 'BIFE ANCHO',          'pulpa': None,          'en': 'BONELESS BEEF RIB EYE'},
}
CLAVES_MEXICO = sorted(MAPA_MEXICO.keys(), key=len, reverse=True)


def limpiar_desc_mexico(desc_original):
    """Limpia la descripcion cruda del remito para usarla como nombre de
    respaldo cuando el corte no esta en MAPA_MEXICO: se queda con todo lo que
    esta antes de '(MEX)' (los calificadores como GF/MC/AA van despues)."""
    d = (desc_original or '').upper()
    if '(MEX)' in d:
        d = d.split('(MEX)')[0]
    return d.strip()


def buscar_info_mexico(desc_original):
    """Busca el corte dentro de la descripcion cruda del remito (ej.
    'NALGA DE AFUERA C/TORTGUITA (MEX) GF') y devuelve su info de Mexico,
    o None si no esta en la tabla (corte nuevo, no mapeado todavia)."""
    d = (desc_original or '').upper()
    for clave in CLAVES_MEXICO:
        if clave in d:
            return MAPA_MEXICO[clave]
    return None


def armar_nombre_mexico(prod):
    """Arma el nombre de 3 partes 'ES/ PULPA / EN' (o 2 partes 'ES / EN' si
    el corte no lleva pulpa) para el certificado de Mexico. Si el corte no
    esta en MAPA_MEXICO (todavia no se le agrego el calificador de pulpa,
    si le corresponde), se arma un nombre de 2 partes igual de valido usando
    la descripcion completa del remito - no queda nunca en blanco ni con un
    nombre generico de una sola palabra, sea cual sea el corte."""
    info = buscar_info_mexico(prod.get('desc_original', ''))
    if info is not None:
        es, pulpa, en = info['es'], info['pulpa'], info['en']
        if pulpa:
            return es + '/ ' + pulpa + ' / ' + en
        return es + ' / ' + en

    es_generico = limpiar_desc_mexico(prod.get('desc_original', '')) or (prod.get('nombre_es', '') or '').strip().upper()
    en_generico = (prod.get('nombre_en', '') or buscar_nombre_en(es_generico) or '').strip().upper()
    if en_generico and en_generico != es_generico:
        return es_generico + ' / ' + en_generico
    return es_generico


# ── NOMBRES ESPECIFICOS FILIPINAS (nombre bilingue de una sola linea) ────────
# A diferencia de Malasia/Singapur (ES y EN en filas separadas), Filipinas pide
# el nombre bilingue combinado en UNA sola linea "ES / EN". La tabla se va
# completando con cada envio nuevo que Angie confirma. IMPORTANTE: el texto
# congelado y el enfriado de un mismo corte son independientes entre si (no se
# derivan uno del otro - cada uno viene confirmado por su propio envio real,
# y pueden diferir en fraseo, orden de palabras, o incluir "ANGUS"/grado de
# marmoleo segun el lote). Por eso cada corte guarda un dict separado por
# estado ('congelado'/'enfriado'), completado solo con lo confirmado.
MAPA_FILIPINAS = {
    'BIFE ANGOSTO CON HUESO Y LOMO': {
        'congelado': {'es': 'BIFE ANGOSTO CON HUESO Y LOMO',  'en': 'FROZEN BONE IN BEEF SHORTLOIN'},
        'enfriado':  {'es': 'BIFE ANGOSTO CON HUESO CON LOMO', 'en': 'CHILLED BEEF ANGUS SHORTLOIN BONE IN GF MB4+'},
    },
    'BIFE ANCHO CON HUESO': {
        'congelado': {'es': 'BIFE ANCHO CON HUESO', 'en': 'FROZEN BONE IN BEEF OP RIBS'},
        'enfriado':  {'es': 'BIFE ANCHO CON HUESO', 'en': 'CHILLED BEEF ANGUS BONE IN OP RIBS GF AA MB4+'},
    },
    'CABEZA DE LOMO': {
        'congelado': {'es': 'CABEZA DE LOMO', 'en': 'FROZEN BONELESS BEEF TENDERLOIN BUTT GF AA MB2+'},
        'enfriado':  {'es': 'CABEZA DE LOMO', 'en': 'CHILLED BEEF ANGUS BONELESS TENDERLOIN BUTT GF MB4+'},
    },
    # Los siguientes solo estan confirmados para congelado por ahora - se
    # agrega la variante enfriado cuando Angie confirme un envio con ese corte
    'CENTRO DE ENTRAÑA':          {'congelado': {'es': 'CENTRO DE ENTRAÑA',          'en': 'FROZEN BEEF HANGING TENDER GF'}},
    'RABO':                       {'congelado': {'es': 'RABO',                       'en': 'FROZEN BEEF TAILS'}},
    'LOMO SIN CORDON':            {'congelado': {'es': 'LOMO SIN CORDON',            'en': 'FROZEN BEEF TENDERLOIN GF AA MB2+ AGED'}},
    'LOMO SC':                    {'congelado': {'es': 'LOMO SIN CORDON',            'en': 'FROZEN BEEF TENDERLOIN GF AA MB2+ AGED'}},
    'MARUCHA':                    {'congelado': {'es': 'MARUCHA',                    'en': 'FROZEN BONELESS BEEF OYSTER BLADE'}},
    'GRASA VACUNA DE DESPOSTADA': {'congelado': {'es': 'GRASA VACUNA DE DESPOSTADA', 'en': 'FROZEN BEEF BODY FAT'}},
    'GRASA VACUNA R':             {'congelado': {'es': 'GRASA VACUNA R',             'en': 'FROZEN BONELESS BEEF FAT'}},
    'ENTRAÑA FINA':               {'congelado': {'es': 'ENTRAÑA FINA',               'en': 'FROZEN BONELESS BEEF OUTSIDE SKIRT GF'}},
    'BIFE ANCHO SIN TAPA':        {'congelado': {'es': 'BIFE ANCHO SIN TAPA',        'en': 'FROZEN BEEF BONELESS RIB EYE GF AA MB2+ AGED'}},
    'TAPA DE CUADRIL':            {'congelado': {'es': 'TAPA DE CUADRIL',            'en': 'FROZEN BEEF BONELESS RUMP CAP GF AA MB2+ AGED'}},
    'CORAZON DE CUADRIL':         {'congelado': {'es': 'CORAZON DE CUADRIL',         'en': 'FROZEN BONELESS BEEF SIRLOIN CC'}},
}
CLAVES_FILIPINAS = sorted(MAPA_FILIPINAS.keys(), key=len, reverse=True)


def buscar_info_filipinas(desc_original, es_congelado):
    d = (desc_original or '').upper()
    estado = 'congelado' if es_congelado else 'enfriado'
    for clave in CLAVES_FILIPINAS:
        if clave in d:
            variantes = MAPA_FILIPINAS[clave]
            if estado in variantes:
                return variantes[estado]
            # Todavia no se confirmo esta variante puntual (congelado/enfriado) -
            # mejor usar la que si esta confirmada que dejar el corte sin mapear.
            for v in variantes.values():
                return v
    return None


# ── NOMBRES Y FUSION DE FILAS ECUADOR ────────────────────────────────────────
# Ecuador pide un nombre en UNA sola columna (sin bilingue) y ademas fusiona en
# una sola fila los productos que son el mismo corte pero difieren solo en un
# RANGO DE PESO (ej. "TAPA DE CUADRIL -1,6 KG" y "+1,6 KG" -> una fila con la
# suma de cajas/neto/bruto). Los que difieren en grado de calidad (AA) o son
# "en trozos" (e/tzos) NO se fusionan, aunque el nombre final se vea igual.

def clave_fusion_ecuador(desc_original):
    """Clave para agrupar filas que se fusionan: la descripcion cruda del
    remito sin el rango de peso (ej. '-1,6 KG'), pero conservando todo lo
    demas (grado de calidad, "en trozos", etc.) para no fusionar productos
    que en realidad son distintos."""
    d = (desc_original or '').upper()
    d = re.sub(r'[+\-]\s*\d+[,.]\d+\s*KG', '', d)
    return re.sub(r'\s+', ' ', d).strip()


def limpiar_nombre_ecuador(desc_original):
    """Nombre simplificado en español para el certificado de Ecuador (sin
    bilingue). Quita el destino '(EC)' y los codigos de calidad, el rango de
    peso (ej. '-1,6 KG'), y los calificadores que Ecuador no usa en el nombre
    final (FINA/ST/CC/PE), conservando "EN TROZOS" si corresponde."""
    d = (desc_original or '').upper()
    d = re.sub(r'\s*E/TZOS\s*\(\d+\)', ' EN TROZOS', d)
    d = re.sub(r'\s*[+\-]\s*\d+[,.]\d+\s*KG', '', d)
    d = d.split('(')[0].strip()
    en_trozos = d.endswith('EN TROZOS')
    base = d[:-len('EN TROZOS')].strip() if en_trozos else d
    base = re.sub(r'\s+(FINA|ST|CC|PE)$', '', base).strip()
    return (base + ' EN TROZOS') if en_trozos else base


def fusionar_productos_ecuador(productos):
    """Agrupa los productos del remito por clave_fusion_ecuador, sumando
    cajas/neto/bruto de los que comparten clave (mismo corte, solo distinto
    rango de peso). Mantiene el orden de primera aparicion."""
    grupos = {}
    orden = []
    for prod in productos:
        clave = clave_fusion_ecuador(prod.get('desc_original', ''))
        if clave not in grupos:
            grupos[clave] = {
                'desc_original': prod.get('desc_original', ''),
                'cajas': 0.0, 'neto': 0.0, 'bruto': 0.0,
            }
            orden.append(clave)
        g = grupos[clave]
        try: g['cajas'] += float(prod.get('cajas', 0) or 0)
        except (TypeError, ValueError): pass
        try: g['neto'] += float(prod.get('neto', 0) or 0)
        except (TypeError, ValueError): pass
        try: g['bruto'] += float(prod.get('bruto', 0) or 0)
        except (TypeError, ValueError): pass
    fusionados = []
    for clave in orden:
        g = grupos[clave]
        fusionados.append({
            'desc_original': g['desc_original'],
            'cajas': str(int(g['cajas'])) if float(g['cajas']).is_integer() else str(g['cajas']),
            'neto': '{:.2f}'.format(g['neto']),
            'bruto': '{:.2f}'.format(g['bruto']),
        })
    return fusionados


# ── NOMBRES ESPECIFICOS EGIPTO (nombre bilingue de una sola linea) ──────────
MAPA_EGIPTO = {
    'HIGADO': 'HIGADOS BOVINOS CONGELADOS / FROZEN BEEF OFFALS LIVERS',
    'RIÑON':  'RIÑONES BOVINOS CONGELADOS / FROZEN BEEF OFFALS KIDNEYS',
}
CLAVES_EGIPTO = sorted(MAPA_EGIPTO.keys(), key=len, reverse=True)


# ── NOMBRES ESPECIFICOS DUBAI (por codigo, nombre bilingue de una sola linea) ──
MAPA_DUBAI = {
    'CD160208': {'es': 'BIFE ANCHO CON HUESO', 'en': 'TOMAHAWK'},
    'CD155309': {'es': 'BIFE ANCHO CON HUESO CON LOMO', 'en': 'SHORTLOINS'},
    'CD219023': {'es': 'ENTRAÑA FINA', 'en': 'ANGUS BONELESS BEEF OUTSIDE SKIRT'},
    'CD209167': {'es': 'TAPA DE CUADRIL', 'en': 'RUMP CAP'},
    'CD220018': {'es': 'VACIO', 'en': 'ANGUS BONELESS BEEF WHOLE FLANK'},
    'CD220625': {'es': 'BIFE DE VACIO GRANDE', 'en': 'FLAP MEAT'},
    'CD214093': {'es': 'BIFE ANGOSTO', 'en': 'STRIPLOIN'},
    'CD216150': {'es': 'LOMO S/ CORDON', 'en': 'TENDERLOIN CHAIN OFF'},
    'CD216151': {'es': 'LOMO S/ CORDON', 'en': 'TENDERLOIN CHAIN OFF'},
    'CD216152': {'es': 'LOMO S/ CORDON', 'en': 'TENDERLOIN CHAIN OFF'},
    'CD224889': {'es': 'BIFE ANCHO SIN TAPA', 'en': 'RIBEYE'},
    'CD217005': {'es': 'CABEZA DE LOMO', 'en': 'TENDERLOIN BUTT'},
}


def armar_nombre_dubai(prod):
    """Arma el nombre bilingue de una sola linea 'ES/ EN' para Dubai, buscando
    por codigo en MAPA_DUBAI. Si el codigo no esta todavia mapeado, cae a la
    descripcion limpia del remito + traduccion generica en vez de dejar la
    celda vacia."""
    codigo = (prod.get('codigo', '') or '').strip().upper()
    info = MAPA_DUBAI.get(codigo)
    if info:
        return info['es'] + '/ ' + info['en']
    desc_original = prod.get('desc_original', '')
    es = (desc_original.split('(')[0].strip().upper() if desc_original else '') or (prod.get('nombre_es', '') or '').strip().upper()
    en = (buscar_nombre_en(es) or '').strip().upper()
    if en:
        return es + '/ ' + en
    return es


def armar_nombre_egipto(prod):
    """Arma el nombre bilingue de una sola linea 'ES / EN' para Egipto. Si el
    corte no esta todavia en MAPA_EGIPTO, cae a la descripcion completa del
    remito + la tabla general de traducciones en vez de dejar la celda
    vacia o con un nombre generico de una sola palabra."""
    desc_original = prod.get('desc_original', '')
    d = (desc_original or '').upper()
    for clave in CLAVES_EGIPTO:
        if clave in d:
            return MAPA_EGIPTO[clave]

    es = (desc_original.split('(')[0].strip().upper() if desc_original else '') or (prod.get('nombre_es', '') or '').strip().upper()
    en = (buscar_nombre_en(es) or '').strip().upper()
    if en:
        return es + ' / ' + en
    return es


# ── NOMBRES ESPECIFICOS BRASIL (nombre bilingue de una sola linea) ──────────
# A diferencia de Ecuador, Brasil NO fusiona filas (cada linea del remito
# queda como su propia fila, aunque el nombre simplificado coincida entre
# variantes de peso).
MAPA_BRASIL = {
    'TAPA DE CUADRIL':      'PICANHA',
    'COLITA DE CUADRIL':    'MAMINHA',
    'ENTRAÑA FINA':         'FRALDINHA',
    'BIFE DE VACIO GRANDE': 'FRALDA',
    'BIFE ANCHO':           'FILE DE COSTELA',
}
CLAVES_BRASIL = sorted(MAPA_BRASIL.keys(), key=len, reverse=True)


def limpiar_nombre_es_brasil(desc_original):
    """Nombre en español simplificado para Brasil: quita el rango de peso
    (ej. 'A -1.6'), el pais destino '(BR)' y todo lo que sigue, y el
    calificativo suelto 'BR' si quedo afuera del parentesis."""
    d = (desc_original or '').upper()
    d = re.sub(r'\s+A\s*[+\-]\s*\d+[,.]\d+', '', d)
    d = d.split('(')[0].strip()
    d = re.sub(r'\bBR\b', '', d).strip()
    return re.sub(r'\s+', ' ', d)


# ── NOMBRES ESPECIFICOS PERU (por codigo de producto, tabla propia de Angie) ──
# A diferencia de otros destinos, Peru tiene una tabla CODIGO -> DESCRIPCION ya
# armada y confirmada (viene de "DESCRIPCION_DE_CAMION_Y_MULTI.xlsx"), separada
# por Enfriado/Congelado. Se usa esa tabla como fuente principal por codigo
# exacto, con una regla de limpieza de texto como respaldo para codigos que
# todavia no esten en la tabla (quita el destino "PE"/"(PE)" y los
# calificativos sueltos "CC"/"SC", pero conserva el grado de calidad GF/AA/MB2+).
MAPA_PERU_ENFRIADO = {
    'CD214022': 'BIFE ANGOSTO GF AA',
    'CD220501': 'FALSA ENTRAÑA',
    'CD220626': 'VACIO GF AA',
    'CD214897': 'BIFE ANGOSTO GF AA MB2+',
    'CD219000': 'ENTRAÑA FINA GF AA',
    'CD213100': 'COLITA DE CUADRIL GF AA',
    'CD214836': 'BIFE ANGOSTO GF',
    'CV224316': 'BIFE ANCHO ST GF AA MB2+',
    'CD264567': 'ASADO SIN HUESO GF AA',
    'CD224850': 'BIFE ANCHO ST GF',
    'CD216000': 'LOMO SC 3/4 LBS GF AA',
    'CD209144': 'TAPA DE CUADRIL GF AA',
    'CD209145': 'TAPA DE CUADRIL GF AA',
    'CD224365': 'BIFE ANCHO ST GF AA',
    'CV209300': 'CORAZON DE CUADRIL GF AA',
    'CD221011': 'BIFE DE VACIO GF AA',
    'CD264565': 'ASADO SIN HUESO GF AA',
    'CD216143': 'LOMO SC 4/5 LBS GF AA',
}


MAPA_PERU_MENUDENCIAS = {
    'FD610001': 'CORAZON',
    'FD608001': 'HIGADO TP',
    'FD608018': 'HIGADO PE',
    'FD615001': 'MONDONGO SEMICOCIDO CON BONETE TP',
    'FD615004': 'MONDONGO SEMICOCIDO CON BONETE B',
}


MAPA_PERU_CONGELADO = {
    'FD610001': 'CORAZON',
    'FD611008': 'MOLLEJAS',
    'FD256054': 'RECORTE DE CARNE VACUNA SIN HUESO PARA USO INDUSTRIAL RECORTE GR 80/20 AA',
    'FD256005': 'RECORTE DE CARNE VACUNA SIN HUESO PARA USO INDUSTRIAL RECORTE GR 80/20 AA',
    'FD219025': 'ENTRAÑA FINA GF AA',
    'FD219010': 'ENTRAÑA FINA GF',
    'FD216094': 'LOMO FINO GF AA',
    'FD216003': 'LOMO FINO S/C 3/4 LBS',
    'FD216019': 'LOMO FINO S/C 4/5 LBS',
    'FD209337': 'CORAZON DE CUADRIL GF',
    'FD214323': 'BIFE ANGOSTO',
    'FD214326': 'BIFE ANGOSTO',
    'FD209236': 'TAPA DE CUADRIL GF AA',
    'FD209237': 'TAPA DE CUADRIL GF AA',
    'FD214358': 'BIFE ANGOSTO GF AA',
    'FD224342': 'BIFE ANCHO GF',
    'FD224346': 'BIFE ANCHO GF AA',
    'FD216093': 'LOMO FINO GF AA',
    'FD221011': 'BIFE DE VACIO GF AA',
    'FD220629': 'VACIO GF AA',
}

# Codigos confirmados de menudencia (organo) para Peru - se usan para separar
# un envio combinado en 2 documentos (congelado + menudencias). Se va
# ampliando con cada envio nuevo que Angie confirme, igual que las tablas de
# nombres - mejor pecar de lista corta y pedir confirmacion para lo que no
# este, que asumir mal.
CODIGOS_MENUDENCIA_PERU = {
    'FD610001',  # CORAZON (organo, no "CORAZON DE CUADRIL" que es corte de carne)
    'FD608001', 'FD608018',  # HIGADO
    'FD615001', 'FD615004',  # MONDONGO
    'FD611008',  # MOLLEJAS
}


def limpiar_nombre_peru(desc_original):
    """Respaldo para codigos que no estan en MAPA_PERU_ENFRIADO/CONGELADO: quita
    el destino 'PE'/'(PE)' y los calificativos sueltos CC/SC, conservando el
    grado de calidad (GF, AA, MB2+)."""
    d = (desc_original or '').upper()
    d = d.replace('(PE)', ' ')
    d = re.sub(r'\bPE\b', ' ', d)
    d = re.sub(r'\b(CC|SC)\b', ' ', d)
    return re.sub(r'\s+', ' ', d).strip()


def armar_nombre_peru(prod, mapa_codigos):
    codigo = (prod.get('codigo', '') or '').strip().upper()
    if codigo in mapa_codigos:
        return mapa_codigos[codigo]
    return limpiar_nombre_peru(prod.get('desc_original', ''))


def armar_nombre_brasil(prod):
    """Arma el nombre bilingue de una sola linea 'ES / PT' para Brasil. Si el
    corte no esta todavia en MAPA_BRASIL, cae al nombre en español solo (sin
    version en portugues) en vez de dejar la celda vacia."""
    desc_original = prod.get('desc_original', '')
    es = limpiar_nombre_es_brasil(desc_original)
    d = (desc_original or '').upper()
    for clave in CLAVES_BRASIL:
        if clave in d:
            return es + ' / ' + MAPA_BRASIL[clave]
    return es


def armar_nombre_filipinas(prod, es_congelado):
    """Arma el nombre bilingue de una sola linea 'ES / EN' para Filipinas. Si el
    corte no esta todavia en MAPA_FILIPINAS, cae a la descripcion completa del
    remito + la tabla general de traducciones (igual que Mexico) en vez de
    dejar la celda vacia o con un nombre generico de una sola palabra."""
    desc_original = prod.get('desc_original', '')

    # Caso especial: "ASADO CON HUESO" lleva la cantidad de costillas, que
    # varia por envio (viene codificada en el remito como ej. "5C")
    m_costillas = re.search(r'ASADO CON HUESO\s*(\d+)\s*C\b', desc_original, re.IGNORECASE)
    if m_costillas:
        n = m_costillas.group(1)
        prefijo_en = 'FROZEN' if es_congelado else 'CHILLED'
        return 'ASADO CON HUESO ' + n + ' COSTILLAS / ' + prefijo_en + ' BEEF BONE IN TOP PLATE ' + n + ' RIBS'

    info = buscar_info_filipinas(desc_original, es_congelado)
    if info is not None:
        return info['es'] + ' / ' + info['en']

    es = (desc_original.split('(')[0].strip().upper() if desc_original else '') or (prod.get('nombre_es', '') or '').strip().upper()
    en = (buscar_nombre_en(es) or '').strip().upper()
    if en:
        return es + ' / ' + en
    return es


# ── NOMBRES ESPECIFICOS HONG KONG (corte / ingles, con prefijo BEEF/FROZEN BEEF) ──
# A diferencia de Malasia/Singapur, Hong Kong pide el nombre en ingles con el
# prefijo "BEEF " (y "FROZEN BEEF " si el envio es congelado). Ademas hay
# cortes con el mismo nombre corto en el remito que son en realidad productos
# distintos (ej. "BIFE ANCHO CON HUESO" vs "BIFE ANCHO CON COSTILLA TOMAHAWK"),
# por eso se buscan las claves mas especificas primero.
MAPA_HONGKONG = {
    'BIFE ANCHO CON COSTILLA TOMAHAWK': {'es': 'BIFE ANCHO CON HUESO',    'en': 'TOMAHAWK'},
    'BIFE ANCHO CON HUESO':             {'es': 'BIFE ANCHO CON HUESO',    'en': 'OP RIBS'},
    'BIFE ANCHO S/T CC':                {'es': 'BIFE ANCHO SIN TAPA CC', 'en': 'RIBEYE LIP ON'},
    'BIFE ANCHO SIN TAPA CC':           {'es': 'BIFE ANCHO SIN TAPA CC', 'en': 'RIBEYE LIP ON'},
    'BIFE ANCHO ST':                    {'es': 'BIFE ANCHO SIN TAPA',    'en': 'RIBEYE'},
    'BIFE ANCHO SIN TAPA':              {'es': 'BIFE ANCHO SIN TAPA',    'en': 'RIBEYE'},
    'BIFE ANGOSTO':                     {'es': 'BIFE ANGOSTO',           'en': 'STRIPLOIN'},
    'AGUJA':                            {'es': 'AGUJA',                  'en': 'CHUCK ROLL'},
    'CENTRO DE ENTRAÑA':                {'es': 'CENTRO DE ENTRAÑA',      'en': 'THICK SKIRT'},
    'ENTRAÑA FINA':                     {'es': 'ENTRAÑA FINA',           'en': 'THIN SKIRT'},
    'ENTRAÑA':                          {'es': 'CENTRO DE ENTRAÑA',      'en': 'THICK SKIRT'},  # sin calificativo = CENTRO DE ENTRAÑA (confirmado con CA218000)
    'CORAZON DE LENGUA':                {'es': 'CORAZON DE LENGUA',      'en': 'PEELED CENTRE CUT TONGUE'},
    'MOLLEJA':                          {'es': 'MOLLEJA',                'en': 'SWEETBREAD'},
    'HUESO DE PIERNA EN TROZOS':        {'es': 'HUESO DE PIERNA EN TROZOS', 'en': 'CENTRE CUT MARROW BONE'},
    'NUEZ DE QUIJADA':                  {'es': 'NUEZ DE QUIJADA',        'en': 'CHEEK MEAT'},
}
CLAVES_HONGKONG = sorted(MAPA_HONGKONG.keys(), key=len, reverse=True)


def buscar_info_hongkong(desc_original):
    d = (desc_original or '').upper()
    for clave in CLAVES_HONGKONG:
        if clave in d:
            return MAPA_HONGKONG[clave]
    return None


def armar_nombre_hongkong(prod, es_congelado):
    """Arma el nombre 'ES/ (FROZEN )BEEF EN' de una sola linea para Hong Kong.
    Si el corte no esta en MAPA_HONGKONG, usa la descripcion completa del
    remito como respaldo (igual que Mexico) en vez de quedar en blanco."""
    info = buscar_info_hongkong(prod.get('desc_original', ''))
    prefijo = 'FROZEN BEEF ' if es_congelado else 'BEEF '
    if info is not None:
        return info['es'] + '/ ' + prefijo + info['en']

    es_generico = limpiar_desc_mexico(prod.get('desc_original', '')) or (prod.get('nombre_es', '') or '').strip().upper()
    en_generico = (buscar_nombre_en(es_generico) or '').strip().upper()
    if en_generico:
        return es_generico + '/ ' + prefijo + en_generico
    return es_generico


# ── XML HELPERS ──────────────────────────────────────────────────────────────

def get_trs(xml):
    return list(re.finditer(r'<w:tr[ >]', xml))


def get_fila_xml(xml, trs, idx):
    ini = trs[idx].start()
    fin = trs[idx + 1].start() if idx + 1 < len(trs) else len(xml)
    return xml[ini:fin], ini, fin


def _reemplazar_tras_label(xml, patron_label, valor_nuevo, count=1, flags=re.IGNORECASE):
    """Reemplaza el valor que sigue a un label de texto fijo (ej. 'Contenedor N°) /
    Container Identification:'), tolerando que haya tags XML y/o runs vacios/con
    solo espacios de por medio (formato tipico de Word al fragmentar en runs).
    El valor debe empezar con un caracter alfanumerico - esto es clave para NO
    conformarse con un run intermedio que solo tiene un espacio, y seguir hasta
    encontrar el valor real."""
    patron = re.compile(r'(' + patron_label + r'(?:\s|<[^>]+>)*?)([A-Za-z0-9][^<]*)', flags)
    return patron.sub(lambda m: m.group(1) + valor_nuevo, xml, count=count)


def _reemplazar_ocurrencias_por_indice(xml, patron_label, valores, flags=re.IGNORECASE):
    """Como _reemplazar_tras_label, pero para cuando el MISMO label aparece
    varias veces con valores distintos (ej. 'Número de precinto:' en español
    y de nuevo en portugues, con el precinto SENASA y el AFIP respectivamente).
    'valores' es la lista de reemplazos en el orden en que aparecen las
    ocurrencias (None = no tocar esa ocurrencia). Se procesa de atras para
    adelante para no invalidar las posiciones ya encontradas."""
    patron = re.compile(r'(' + patron_label + r'(?:\s|<[^>]+>)*?)([A-Za-z0-9][^<]*)', flags)
    matches = list(patron.finditer(xml))
    for i in reversed(range(min(len(matches), len(valores)))):
        if not valores[i]:
            continue
        m = matches[i]
        xml = xml[:m.start()] + m.group(1) + valores[i] + xml[m.end():]
    return xml


def _reemplazar_celda(xml_fila, celda_idx, nuevo_texto):
    celda_starts = [m.start() for m in re.finditer(r'<w:tc>', xml_fila)]
    celda_ends   = [m.start() for m in re.finditer(r'</w:tc>', xml_fila)]
    if celda_idx >= len(celda_starts): return xml_fila
    bloque = xml_fila[celda_starts[celda_idx]:celda_ends[celda_idx]]
    textos = re.findall(r'<w:t[^>]*>[^<]*</w:t>', bloque)
    if not textos:
        # Celda sin ningun <w:t> (vacia de verdad, sin ni siquiera un run vacio):
        # insertar un run nuevo antes de cerrar el ultimo parrafo, en vez de
        # descartar el valor en silencio (eso escondia errores de indice de celda).
        insert_pos = bloque.rfind('</w:p>')
        if insert_pos < 0:
            return xml_fila
        nuevo_bloque = bloque[:insert_pos] + '<w:r><w:t>' + nuevo_texto + '</w:t></w:r>' + bloque[insert_pos:]
        return xml_fila[:celda_starts[celda_idx]] + nuevo_bloque + xml_fila[celda_ends[celda_idx]:]
    primer   = textos[0]
    tag_open = re.match(r'<w:t[^>]*>', primer).group()
    nuevo_bloque = bloque.replace(primer, tag_open + nuevo_texto + '</w:t>', 1)
    for t in textos[1:]:
        tag2 = re.match(r'<w:t[^>]*>', t).group()
        nuevo_bloque = nuevo_bloque.replace(t, tag2 + '</w:t>', 1)
    return xml_fila[:celda_starts[celda_idx]] + nuevo_bloque + xml_fila[celda_ends[celda_idx]:]


def _construir_fila(fila_modelo, cajas, nombre_bi, neto, bruto, neto_celda, bruto_celda,
                     lotes=None, lotes_celda=None):
    nueva = fila_modelo
    nueva = _reemplazar_celda(nueva, 0, str(cajas))
    nueva = _reemplazar_celda(nueva, 1, nombre_bi)
    if lotes_celda is not None:
        nueva = _reemplazar_celda(nueva, lotes_celda, str(lotes or ''))
    nueva = _reemplazar_celda(nueva, neto_celda, str(neto))
    nueva = _reemplazar_celda(nueva, bruto_celda, str(bruto))
    return nueva


def _get_fila_por_contenido(xml, trs, texto_clave):
    for i, m in enumerate(trs):
        ini = m.start()
        fin = trs[i+1].start() if i+1 < len(trs) else len(xml)
        if texto_clave in xml[ini:fin]:
            return xml[ini:fin], ini, fin, i
    return None, None, None, None


def _reemplazar_pallets_en_fila(fila_xml, pallets, kg_pallets):
    # El conteo de pallets puede aparecer 2 veces en la misma fila bilingue (ES y
    # EN). Segun como haya fusionado _merge_runs_xml los runs, el numero puede
    # quedar: (a) embebido en el mismo run que el texto "ACONDICIONADO EN"/
    # "ACONDITIONED IN" (fusion exitosa), o (b) en su propio run aislado (no se
    # fusiono, tenia formato distinto). Se cubren ambos casos.
    fila_xml = re.sub(r'(ACONDICIONADO EN\s*)\d+(\s*PALLET)', r'\g<1>' + str(pallets) + r'\2', fila_xml)
    fila_xml = re.sub(r'(ACONDITIONED IN\s*)\d+(\s*PALLET)', r'\g<1>' + str(pallets) + r'\2', fila_xml)

    def _reemplazar_numero_run(m):
        contenido = m.group(1)
        if contenido.strip().isdigit():
            sufijo = ' ' if contenido.endswith(' ') and not contenido.endswith('  ') else ''
            return m.group(0).replace(contenido, str(pallets) + sufijo, 1)
        return m.group(0)
    fila_xml = re.sub(r'<w:t[^>]*>([^<]*)</w:t>', _reemplazar_numero_run, fila_xml)

    if kg_pallets:
        # Caso generico: el numero de KGS esta en el mismo run que el texto "KGS)" (ej. Mexico)
        nueva_fila, n = re.subn(r'[\d\.,]+(\s*KGS\))', str(kg_pallets) + r'\1', fila_xml, count=1)
        if n:
            fila_xml = nueva_fila
        else:
            # Fallback generico: run con formato de numero decimal aislado (plantillas viejas,
            # ej. Malasia/Singapur con el numero de KGS en su propio run separado del texto)
            def _reemplazar_decimal_run(m):
                contenido = m.group(1)
                if re.match(r'^\d+[.,]\d+$', contenido.strip()):
                    return m.group(0).replace(contenido, str(kg_pallets), 1)
                return m.group(0)
            fila_xml = re.sub(r'<w:t[^>]*>([^<]*)</w:t>', _reemplazar_decimal_run, fila_xml)
    return fila_xml


def _reemplazar_bloque_productos(xml, trs, primera_idx, total_idx_fallback,
                                  productos, total_cajas, total_neto, total_bruto,
                                  pallets, kg_pallets, neto_celda=6, bruto_celda=7,
                                  lotes_celda=None, sumar_pallet_a_bruto=True,
                                  total_replacer=None, armar_nombre_func=None):
    fila_pal, ini_pal, fin_pal, idx_pal = _get_fila_por_contenido(xml, trs, 'ACONDICIONADO EN')
    total_idx = None
    if idx_pal is not None:
        # Buscar la primera fila con contenido real despues de la fila de pallets -
        # puede haber filas vacias "espaciadoras" de por medio antes de la fila de
        # totales real, y asumir siempre "+1" rompe en esos casos.
        for j in range(idx_pal + 1, len(trs)):
            fila_candidata, _, _ = get_fila_xml(xml, trs, j)
            if re.search(r'<w:t[^>]*>[^<]+</w:t>', fila_candidata):
                total_idx = j
                break
    if total_idx is None:
        total_idx = total_idx_fallback

    fila_modelo, ini_mod, _ = get_fila_xml(xml, trs, primera_idx)
    fila_total, ini_tot, fin_tot = get_fila_xml(xml, trs, total_idx)

    if armar_nombre_func is None:
        armar_nombre_func = lambda prod: armar_nombre_bilingue(prod.get('nombre_es', ''), prod.get('nombre_en', ''))

    nuevas_filas = ''
    for prod in productos:
        nombre_bi = armar_nombre_func(prod)
        nuevas_filas += _construir_fila(
            fila_modelo, prod.get('cajas', ''), nombre_bi,
            prod.get('neto', ''), prod.get('bruto', ''), neto_celda, bruto_celda,
            lotes=prod.get('lotes', ''), lotes_celda=lotes_celda
        )

    nueva_pal = _reemplazar_pallets_en_fila(fila_pal, pallets, kg_pallets) if (fila_pal and pallets) else fila_pal

    if sumar_pallet_a_bruto:
        try:
            total_bruto_final = '{:.2f}'.format(float(total_bruto) + float(kg_pallets or 0))
        except Exception:
            total_bruto_final = total_bruto
    else:
        total_bruto_final = total_bruto

    if total_replacer is not None:
        nueva_total = total_replacer(fila_total, total_cajas, total_neto, total_bruto_final)
    else:
        nums_tot = re.findall(r'<w:t[^>]*>(\d[\d\.]*)</w:t>', fila_total)
        nueva_total = fila_total
        if len(nums_tot) >= 3:
            nueva_total = nueva_total.replace('>' + nums_tot[0] + '<', '>' + str(total_cajas) + '<', 1)
            nueva_total = nueva_total.replace('>' + nums_tot[1] + '<', '>' + str(total_neto) + '<', 1)
            nueva_total = nueva_total.replace('>' + nums_tot[2] + '<', '>' + str(total_bruto_final) + '<', 1)

    xml_nuevo = xml[:ini_mod] + nuevas_filas + nueva_pal + nueva_total + xml[fin_tot:]
    return xml_nuevo


def fmt_fecha_al_to(f):
    if f and ' al ' in f.lower():
        partes = re.split(r'\s+al\s+', f, flags=re.IGNORECASE)
        return partes[0] + ' AL/TO ' + partes[1]
    return f or ''


def fmt_fecha_to_hongkong(f):
    """'dd/mm/yyyy al dd/mm/yyyy' -> 'dd/mm/yyyy to dd/mm/yyyy' (I.11 de Hong Kong, 'to' en minuscula sin 'AL')."""
    if f and ' al ' in f.lower():
        partes = re.split(r'\s+al\s+', f, flags=re.IGNORECASE)
        return partes[0].strip() + ' to ' + partes[1].strip()
    return f or ''


def fmt_fecha_al(f):
    if f and ' al ' in f.lower():
        partes = re.split(r'\s+al\s+', f, flags=re.IGNORECASE)
        return partes[0] + ' AL ' + partes[1]
    return f or ''


def fmt_fecha_al_minuscula(f):
    """'dd/mm/yyyy al dd/mm/yyyy' con 'al' en minuscula (formato de Ecuador)."""
    if f and ' al ' in f.lower():
        partes = re.split(r'\s+al\s+', f, flags=re.IGNORECASE)
        return partes[0].strip() + ' al ' + partes[1].strip()
    return f or ''


def fmt_fecha_al_to_usa(f):
    """'dd/mm/yyyy al dd/mm/yyyy' -> 'dd/mm/yyyy al/to dd/mm/yyyy' (formato USA, minuscula)."""
    if f and ' al ' in f.lower():
        partes = re.split(r'\s+al\s+', f, flags=re.IGNORECASE)
        return partes[0].strip() + ' al/to ' + partes[1].strip()
    return f or ''


def fecha_a_lote_usa(f):
    """'dd/mm/yyyy al dd/mm/yyyy' -> 'YYYYMMDD al/to YYYYMMDD' (mismo rango que fecha de produccion, Lote de USA)."""
    if not f or ' al ' not in f.lower():
        return f or ''
    partes = re.split(r'\s+al\s+', f, flags=re.IGNORECASE)
    salida = []
    for p in partes:
        m = re.match(r'(\d{2})/(\d{2})/(\d{4})', p.strip())
        salida.append(m.group(3) + m.group(2) + m.group(1) if m else p.strip())
    return ' al/to '.join(salida)


def kg_a_lbs(kg):
    try:
        return '{:.2f}'.format(float(str(kg).replace(',', '.')) * 2.20462)
    except (TypeError, ValueError):
        return ''


def _merge_runs_xml(xml):
    """Fusiona <w:r> adyacentes con el mismo <w:rPr> dentro de cada parrafo,
    concatenando sus <w:t>. Word fragmenta el texto en runs distintos (marcas
    de revision, corrector ortografico), lo que rompe los reemplazos simples
    basados en substring (ej. 'VAPOR/VESSEL:  NOMBRE' guardado en 3 runs
    separados). Solo fusiona runs de texto simple (un unico <w:t>, sin tabs,
    saltos de linea u otros elementos) para no arriesgar contenido complejo."""
    xml = re.sub(r'<w:proofErr[^/]*/>', '', xml)  # el corrector ortografico bloquea la fusion de runs adyacentes

    def _procesar_parrafo(m):
        parrafo = m.group(0)
        cambiado = True
        while cambiado:
            cambiado = False
            runs = list(re.finditer(r'<w:r(?:\s[^>]*)?>(.*?)</w:r>', parrafo, re.DOTALL))
            for i in range(len(runs) - 1):
                r1, r2 = runs[i], runs[i + 1]
                if parrafo[r1.end():r2.start()]:
                    continue  # no son estrictamente adyacentes
                rpr1 = re.search(r'<w:rPr>.*?</w:rPr>', r1.group(1), re.DOTALL)
                rpr2 = re.search(r'<w:rPr>.*?</w:rPr>', r2.group(1), re.DOTALL)
                if (rpr1.group(0) if rpr1 else '') != (rpr2.group(0) if rpr2 else ''):
                    continue
                t1 = re.findall(r'<w:t[^>]*>([^<]*)</w:t>', r1.group(1))
                t2 = re.findall(r'<w:t[^>]*>([^<]*)</w:t>', r2.group(1))
                # Solo fusionar runs de un unico <w:t> simple (evita tabs/breaks/drawings)
                if len(t1) != 1 or len(t2) != 1:
                    continue
                if r1.group(1).count('<w:t') != 1 or r2.group(1).count('<w:t') != 1:
                    continue
                texto_unido = t1[0] + t2[0]
                rpr_txt = rpr1.group(0) if rpr1 else ''
                nuevo_run = '<w:r>' + rpr_txt + '<w:t xml:space="preserve">' + texto_unido + '</w:t></w:r>'
                parrafo = parrafo[:r1.start()] + nuevo_run + parrafo[r2.end():]
                cambiado = True
                break
        return parrafo

    return re.sub(r'<w:p(?:\s[^>]*)?>.*?</w:p>', _procesar_parrafo, xml, flags=re.DOTALL)


def generar_sanitario(docx_bytes, datos, tipo_via, destino):
    alertas = []
    with zipfile.ZipFile(io.BytesIO(docx_bytes), 'r') as z:
        archivos = {n: z.read(n) for n in z.namelist()}
    xml = archivos['word/document.xml'].decode('utf-8')
    xml = _merge_runs_xml(xml)
    if destino == 'singapur':
        if tipo_via == 'aereo':
            xml, al = _gen_singapur_aereo(xml, datos)
        else:
            xml, al = _gen_singapur_maritimo(xml, datos)
    elif destino == 'mexico':
        xml, al = _gen_mexico_maritimo(xml, datos)
    elif destino == 'usawclass':
        xml, al = _gen_usa_wclass(xml, datos)
    elif destino == 'usaorleans':
        xml, al = _gen_usa_orleans(xml, datos)
    elif destino == 'hongkongcongelado':
        xml, al = _gen_hongkong(xml, datos, es_congelado=True, tipo_via=tipo_via)
    elif destino == 'hongkongenfriado':
        xml, al = _gen_hongkong(xml, datos, es_congelado=False, tipo_via=tipo_via)
    elif destino in ('usaallecondimentada', 'usaallenatural'):
        xml, al = _gen_alle_processing(xml, datos)
    elif destino == 'filipinas':
        if tipo_via == 'aereo':
            xml, al = _gen_filipinas_aereo(xml, datos)
        else:
            xml, al = _gen_filipinas_maritimo(xml, datos)
    elif destino == 'ecuador':
        xml, al = _gen_ecuador(xml, datos)
    elif destino == 'egipto':
        xml, al = _gen_egipto(xml, datos)
    elif destino == 'dubai':
        xml, al = _gen_dubai(xml, datos)
    elif destino == 'brasil':
        xml, al = _gen_brasil(xml, datos)
    elif destino == 'peruenfriado':
        xml, al = _gen_peru_enfriado(xml, datos)
    elif destino == 'perumenudencias':
        xml, al = _gen_peru_menudencias(xml, datos)
    elif destino == 'perucongelado':
        xml, al = _gen_peru_congelado(xml, datos)
    else:
        if tipo_via == 'aereo':
            xml, al = _gen_malasia_aereo(xml, datos)
        else:
            xml, al = _gen_malasia_maritimo(xml, datos)
    alertas.extend(al)

    # Red de seguridad: si por algun bug el XML quedo mal formado, no entregar
    # un .docx roto (Word no lo puede ni abrir) - mejor fallar con un error claro.
    try:
        from xml.etree import ElementTree as ET
        ET.fromstring(xml)
    except ET.ParseError as e:
        raise ValueError(
            'El documento generado quedo con XML invalido (' + str(e) + '). '
            'No se genero el archivo para evitar entregar un .docx corrupto - '
            'avisar para revisar el generador de este destino.'
        )

    archivos['word/document.xml'] = xml.encode('utf-8')
    out = io.BytesIO()
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
        for n, d in archivos.items(): z.writestr(n, d)
    out.seek(0)
    return out.read(), alertas



def _reemplazar_fechas(xml, trs, f_faena, f_prod, f_venc, fmt_func):
    """Busca filas con I.13/I.14/I.15 y reemplaza solo la parte de fecha,
    preservando los titulos (I.13.Fecha de faena/, Date of slaughter, etc.)"""

    def _get_celdas(fila):
        starts = [m.start() for m in re.finditer(r'<w:tc>', fila)]
        ends   = [m.start() for m in re.finditer(r'</w:tc>', fila)]
        return [(s, e) for s, e in zip(starts, ends)]

    def _reemplazar_fecha_en_celda(fila, cs, ce, nuevo):
        bloque = fila[cs:ce]
        # Encontrar todos los w:t con su posicion
        wts = list(re.finditer(r'<w:t[^>]*>[^<]*</w:t>', bloque))
        if not wts: return fila

        # Encontrar el primer w:t que contiene digitos de fecha (dd o /mm o /yyyy)
        # Los titulos son texto como "I.13.Fecha de faena/" y "Date of slaughter"
        # La fecha empieza cuando aparece un fragmento con solo digitos o slash
        # Reconstruir texto completo de la celda para encontrar donde empieza la fecha
        txt_completo = ''.join(re.search(r'<w:t[^>]*>([^<]*)</w:t>', wt.group()).group(1) for wt in wts)
        # Buscar posicion del primer dd/ en el texto completo
        m_fecha = re.search(r'\d{2}/', txt_completo)
        if not m_fecha:
            return fila

        # Encontrar que w:t corresponde a esa posicion
        fecha_inicio_idx = None
        pos_acum = 0
        for idx, wt in enumerate(wts):
            txt_wt = re.search(r'<w:t[^>]*>([^<]*)</w:t>', wt.group()).group(1)
            if pos_acum + len(txt_wt) > m_fecha.start():
                fecha_inicio_idx = idx
                break
            pos_acum += len(txt_wt)

        if fecha_inicio_idx is None: return fila

        # Reemplazar desde fecha_inicio_idx en adelante
        nuevo_bloque = bloque
        offset = 0
        for idx, wt in enumerate(wts):
            if idx < fecha_inicio_idx: continue
            tag = re.match(r'<w:t[^>]*>', wt.group()).group()
            old_wt = wt.group()
            pos = nuevo_bloque.find(old_wt, offset)
            if idx == fecha_inicio_idx:
                # Primer fragmento de fecha: poner el valor nuevo con preserve
                new_wt = '<w:t xml:space="preserve">' + nuevo + '</w:t>'
            else:
                # Fragmentos siguientes: vaciar
                new_wt = tag + '</w:t>'
            nuevo_bloque = nuevo_bloque[:pos] + new_wt + nuevo_bloque[pos + len(old_wt):]
            offset = pos + len(new_wt)

        return fila[:cs] + nuevo_bloque + fila[ce:]

    fechas_a_reemplazar = [(f_faena, fmt_func), (f_prod, fmt_func), (f_venc, fmt_func)]
    fecha_idx = 0

    for i, m in enumerate(trs):
        if fecha_idx >= 3: break
        ini = m.start()
        fin = trs[i+1].start() if i+1 < len(trs) else len(xml)
        fila = xml[ini:fin]

        if not ('I.13' in fila or 'I.14' in fila or 'I.15' in fila):
            continue

        celdas = _get_celdas(fila)
        nueva_fila = fila
        offset = 0

        for cs, ce in celdas:
            if fecha_idx >= 3: break
            bloque = nueva_fila[cs+offset:ce+offset]
            txt_celda = ''.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', bloque))
            if re.search(r'\d{2}/\d{2}/\d{4}', txt_celda):
                f_nueva, fmt = fechas_a_reemplazar[fecha_idx]
                if f_nueva:
                    fila_antes = nueva_fila
                    nueva_fila = _reemplazar_fecha_en_celda(nueva_fila, cs+offset, ce+offset, fmt(f_nueva))
                    offset += len(nueva_fila) - len(fila_antes)
                fecha_idx += 1

        xml = xml[:ini] + nueva_fila + xml[fin:]

    return xml

# ── MALASIA AÉREO ────────────────────────────────────────────────────────────

def _gen_malasia_aereo(xml, datos):
    alertas = []
    trs = get_trs(xml)
    xml = _reemplazar_bloque_productos(
        xml, trs, primera_idx=5, total_idx_fallback=19,
        productos=datos.get('productos', []),
        total_cajas=datos.get('total_cajas',''), total_neto=datos.get('total_neto',''), total_bruto=datos.get('total_bruto',''),
        pallets=datos.get('pallets','1'), kg_pallets=datos.get('kg_pallets',''),
        neto_celda=6, bruto_celda=7,
        total_replacer=lambda fila, tc, tn, tb: _reemplazar_total_celdas(fila, tc, tn, tb, cajas_celda=0, neto_celda=2, bruto_celda=3)
    )
    f_faena = datos.get('fecha_faena','')
    f_prod  = datos.get('fecha_produccion','')
    f_venc  = datos.get('fecha_vencimiento','')
    trs2 = get_trs(xml)
    xml = _reemplazar_fechas(xml, trs2, f_faena, f_prod, f_venc, fmt_fecha_al_to)
    transporte = datos.get('transporte','')
    if transporte: xml = xml.replace('>VUELO / FLIGHT: EK248<', '>VUELO / FLIGHT: ' + transporte + '<')
    xml = _set_temperatura_singapur(xml, datos.get('es_congelado', False), tipo_via='aereo')
    fecha_emi = datos.get('fecha_emision') or datetime.datetime.now().strftime('%d/%m/%Y')
    try: dia, mes, anio = fecha_emi.split('/')
    except: dia = mes = anio = ''; alertas.append('Fecha emision no parseada')
    xml = xml.replace('>2026<',  '>' + anio + '<', 1)
    xml = xml.replace('>01 <',   '>' + mes + ' <', 1)
    xml = xml.replace('>23<',    '>' + dia + '<', 1)
    return xml, alertas


# ── MALASIA MARÍTIMO ─────────────────────────────────────────────────────────

def _gen_malasia_maritimo(xml, datos):
    alertas = []
    trs = get_trs(xml)
    xml = _reemplazar_bloque_productos(
        xml, trs, primera_idx=5, total_idx_fallback=15,
        productos=datos.get('productos', []),
        total_cajas=datos.get('total_cajas',''), total_neto=datos.get('total_neto',''), total_bruto=datos.get('total_bruto',''),
        pallets=datos.get('pallets','1'), kg_pallets=datos.get('kg_pallets',''),
        neto_celda=6, bruto_celda=7,
        total_replacer=lambda fila, tc, tn, tb: _reemplazar_total_celdas(fila, tc, tn, tb, cajas_celda=0, neto_celda=2, bruto_celda=3)
    )
    f_faena = datos.get('fecha_faena','')
    f_prod  = datos.get('fecha_produccion','')
    f_venc  = datos.get('fecha_vencimiento','')
    trs2 = get_trs(xml)
    xml = _reemplazar_fechas(xml, trs2, f_faena, f_prod, f_venc, fmt_fecha_al)
    transporte = datos.get('transporte','')
    if transporte: xml = xml.replace('>VAPOR / VESSEL: TIGER PLATA<', '>VAPOR / VESSEL: ' + transporte + '<')
    contenedor    = datos.get('contenedor','')
    precinto_afip = datos.get('precinto_afip','')
    if contenedor:    xml = xml.replace('>TCLU129408-4<', '>' + contenedor + '<')
    if precinto_afip: xml = xml.replace('>BAH79585<',    '>' + precinto_afip + '<')
    if not contenedor:    alertas.append('Contenedor no encontrado - completar manualmente')
    if not precinto_afip: alertas.append('Precinto AFIP no encontrado - completar manualmente')
    xml = _set_temperatura_singapur(xml, datos.get('es_congelado', False), tipo_via='aereo')
    fecha_emi = datos.get('fecha_emision') or datetime.datetime.now().strftime('%d/%m/%Y')
    try: dia, mes, anio = fecha_emi.split('/')
    except: dia = mes = anio = ''; alertas.append('Fecha emision no parseada')
    xml = xml.replace('>:      2026<', '>:      ' + anio + '<')
    xml = re.sub(r'>\)\s+\d{2}\s+<', '>)         ' + mes + ' <', xml, count=1)
    xml = xml.replace('>14<', '>' + dia + '<', 1)
    return xml, alertas


# ── SINGAPUR AÉREO ───────────────────────────────────────────────────────────

def _gen_singapur_aereo(xml, datos):
    alertas = []
    trs = get_trs(xml)
    xml = _reemplazar_bloque_productos(
        xml, trs, primera_idx=5, total_idx_fallback=17,
        productos=datos.get('productos', []),
        total_cajas=datos.get('total_cajas',''), total_neto=datos.get('total_neto',''), total_bruto=datos.get('total_bruto',''),
        pallets=datos.get('pallets','1'), kg_pallets=datos.get('kg_pallets',''),
        neto_celda=6, bruto_celda=7,
        total_replacer=lambda fila, tc, tn, tb: _reemplazar_total_celdas(fila, tc, tn, tb, cajas_celda=0, neto_celda=2, bruto_celda=3)
    )
    # Fechas
    f_faena = datos.get('fecha_faena','')
    f_prod  = datos.get('fecha_produccion','')
    f_venc  = datos.get('fecha_vencimiento','')
    trs2 = get_trs(xml)
    xml = _reemplazar_fechas(xml, trs2, f_faena, f_prod, f_venc, fmt_fecha_al_to)
    # Transporte (vuelo)
    transporte = datos.get('transporte','')
    if transporte: xml = xml.replace('>: LX093<', '>: ' + transporte + '<')
    # Temperatura
    es_congelado = datos.get('es_congelado', False)
    xml = _set_temperatura_singapur(xml, es_congelado, tipo_via='aereo')
    # Consignatario
    xml = xml.replace('>FOODIE MARKET PLACE PTE. LTD<', '>FOODIE MARKET PLACE PTE. LTD<')  # placeholder
    # Fecha emision
    fecha_emi = datos.get('fecha_emision') or datetime.datetime.now().strftime('%d/%m/%Y')
    xml = xml.replace('>04/06/2026<', '>' + fecha_emi + '<')
    return xml, alertas


# ── SINGAPUR MARÍTIMO ────────────────────────────────────────────────────────

def _gen_singapur_maritimo(xml, datos):
    alertas = []
    trs = get_trs(xml)
    xml = _reemplazar_bloque_productos(
        xml, trs, primera_idx=5, total_idx_fallback=10,
        productos=datos.get('productos', []),
        total_cajas=datos.get('total_cajas',''), total_neto=datos.get('total_neto',''), total_bruto=datos.get('total_bruto',''),
        pallets=datos.get('pallets','1'), kg_pallets=datos.get('kg_pallets',''),
        neto_celda=6, bruto_celda=7,
        total_replacer=lambda fila, tc, tn, tb: _reemplazar_total_celdas(fila, tc, tn, tb, cajas_celda=0, neto_celda=2, bruto_celda=3)
    )
    # Fechas
    f_faena = datos.get('fecha_faena','')
    f_prod  = datos.get('fecha_produccion','')
    f_venc  = datos.get('fecha_vencimiento','')
    trs2 = get_trs(xml)
    xml = _reemplazar_fechas(xml, trs2, f_faena, f_prod, f_venc, fmt_fecha_al_to)
    # Transporte (barco)
    transporte = datos.get('transporte','')
    if transporte: xml = xml.replace('>: SAN ANTONIO MAERSK<', '>: ' + transporte + '<')
    # Contenedor y precinto
    contenedor    = datos.get('contenedor','')
    precinto_afip = datos.get('precinto_afip','')
    if contenedor:    xml = xml.replace('>MNBU9179760<', '>' + contenedor + '<')
    if precinto_afip: xml = xml.replace('>BAH66389<',   '>' + precinto_afip + '<')
    if not contenedor:    alertas.append('Contenedor no encontrado - completar manualmente')
    if not precinto_afip: alertas.append('Precinto AFIP no encontrado - completar manualmente')
    # Temperatura
    es_congelado = datos.get('es_congelado', False)
    xml = _set_temperatura_singapur(xml, es_congelado, tipo_via='maritimo')
    # Fecha emision
    fecha_emi = datos.get('fecha_emision') or datetime.datetime.now().strftime('%d/%m/%Y')
    xml = xml.replace('>14/04/2026<', '>' + fecha_emi + '<')
    return xml, alertas


# ── MÉXICO MARÍTIMO ──────────────────────────────────────────────────────────

def _reemplazar_total_celdas(fila_total, total_cajas, total_neto, total_bruto,
                              cajas_celda=0, neto_celda=2, bruto_celda=3):
    """Reemplaza la fila de totales por indice de celda en vez de buscar numeros
    sueltos en el XML. Necesario cuando los totales usan formato con coma
    decimal (ej. '21.692,00'), que no matchea como numero simple."""
    nueva = fila_total
    nueva = _reemplazar_celda(nueva, cajas_celda, str(total_cajas))
    nueva = _reemplazar_celda(nueva, neto_celda, str(total_neto))
    nueva = _reemplazar_celda(nueva, bruto_celda, str(total_bruto))
    return nueva


def _gen_mexico_maritimo(xml, datos):
    alertas = []
    trs = get_trs(xml)

    total_neto_fmt  = formatear_miles(datos.get('total_neto', ''))
    total_bruto_fmt = formatear_miles(datos.get('total_bruto', ''))

    xml = _reemplazar_bloque_productos(
        xml, trs, primera_idx=5, total_idx_fallback=13,
        productos=datos.get('productos', []),
        total_cajas=datos.get('total_cajas', ''), total_neto=total_neto_fmt, total_bruto=total_bruto_fmt,
        pallets=datos.get('pallets', '1'), kg_pallets=datos.get('kg_pallets', ''),
        neto_celda=6, bruto_celda=7,
        lotes_celda=5,
        sumar_pallet_a_bruto=False,  # el total de la plantilla Mexico NO suma el peso de pallets
        total_replacer=lambda fila, tc, tn, tb: _reemplazar_total_celdas(fila, tc, tn, tb),
        armar_nombre_func=armar_nombre_mexico
    )

    # Fechas I.13/I.14/I.15
    f_faena = datos.get('fecha_faena', '')
    f_prod  = datos.get('fecha_produccion', '')
    f_venc  = datos.get('fecha_vencimiento', '')
    trs2 = get_trs(xml)
    xml = _reemplazar_fechas(xml, trs2, f_faena, f_prod, f_venc, fmt_fecha_al)

    # Transporte (buque)
    transporte = datos.get('transporte', '')
    if transporte: xml = xml.replace('>VAPOR/VESSEL:  SUNNY PHOENIX<', '>VAPOR/VESSEL:  ' + transporte + '<')

    # Contenedor
    contenedor = datos.get('contenedor', '')
    if contenedor: xml = xml.replace('>ZMOU8965406<', '>' + contenedor + '<')
    if not contenedor: alertas.append('Contenedor no encontrado - completar manualmente')

    # Precinto combinado AFIP / SENASA
    precinto_afip   = datos.get('precinto_afip', '')
    precinto_senasa = datos.get('precinto_senasa', '')
    if precinto_afip or precinto_senasa:
        combinado = (precinto_afip or '') + ' / ' + (precinto_senasa or '')
        xml = xml.replace('>BAH66487 / 0039365<', '>' + combinado + '<')
    if not precinto_afip:   alertas.append('Precinto AFIP no encontrado - completar manualmente')
    if not precinto_senasa: alertas.append('Precinto SENASA no encontrado - completar manualmente')

    # Temperatura
    es_congelado = datos.get('es_congelado', False)
    xml = _set_temperatura_singapur(xml, es_congelado, tipo_via='maritimo')

    # Fecha de emision
    fecha_emi = datos.get('fecha_emision') or datetime.datetime.now().strftime('%d/%m/%Y')
    xml = xml.replace('>22/04/2026<', '>' + fecha_emi + '<')

    return xml, alertas


# ── TEMPERATURA SINGAPUR ─────────────────────────────────────────────────────

def _set_temperatura_singapur(xml, es_congelado, tipo_via):
    """Todas las plantillas tienen X en refrigeracion por defecto (en la fila en
    español Y en la fila en ingles, que son filas separadas). Si es congelado,
    mover la X a congelacion/frozen en AMBAS filas. Si es enfriado, no hacer nada."""
    if not es_congelado:
        return xml

    cambiado = True
    intentos = 0
    while cambiado and intentos < 10:
        cambiado = False
        intentos += 1
        trs = list(re.finditer(r'<w:tr[ >]', xml))
        for i, m in enumerate(trs):
            ini = m.start()
            fin = trs[i+1].start() if i+1 < len(trs) else len(xml)
            fila = xml[ini:fin]
            txt_low = ''.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', fila)).lower()
            tiene_refrig = 'efrigera' in txt_low          # cubre "Refrigeración" y "Refrigerated"
            tiene_congel = 'ongela' in txt_low or 'rozen' in txt_low  # cubre "congelación" y "Frozen"
            if not (tiene_refrig and tiene_congel):
                continue

            celda_starts = [m2.start() for m2 in re.finditer(r'<w:tc>', fila)]
            celda_ends   = [m2.start() for m2 in re.finditer(r'</w:tc>', fila)]
            idx_refrig = idx_congel = None
            for idx_c, (cs, ce) in enumerate(zip(celda_starts, celda_ends)):
                label_low = ''.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', fila[cs:ce])).lower()
                if 'efrigera' in label_low: idx_refrig = idx_c
                if 'ongela' in label_low or 'rozen' in label_low: idx_congel = idx_c
            if idx_refrig is None or idx_congel is None:
                continue
            if idx_refrig + 1 >= len(celda_starts) or idx_congel + 1 >= len(celda_starts):
                continue

            cs_r, ce_r = celda_starts[idx_refrig + 1], celda_ends[idx_refrig + 1]
            if '<w:t>X</w:t>' not in fila[cs_r:ce_r]:
                continue  # esta fila ya tiene la X movida (evita reprocesarla en la siguiente pasada)

            nueva_fila = fila[:cs_r] + fila[cs_r:ce_r].replace('<w:t>X</w:t>', '<w:t></w:t>', 1) + fila[ce_r:]

            # Recalcular offsets de celdas sobre la fila ya modificada para ubicar el checkbox de congelacion
            celda_starts2 = [m2.start() for m2 in re.finditer(r'<w:tc>', nueva_fila)]
            celda_ends2   = [m2.start() for m2 in re.finditer(r'</w:tc>', nueva_fila)]
            cs2, ce2 = celda_starts2[idx_congel + 1], celda_ends2[idx_congel + 1]
            bloque2 = nueva_fila[cs2:ce2]
            if '<w:t></w:t>' in bloque2:
                nuevo_bloque2 = bloque2.replace('<w:t></w:t>', '<w:t>X</w:t>', 1)
            elif re.search(r'<w:t[^>]*></w:t>', bloque2):
                nuevo_bloque2 = re.sub(r'<w:t[^>]*></w:t>', '<w:t>X</w:t>', bloque2, count=1)
            else:
                insert_pos = bloque2.rfind('</w:p>')
                nuevo_bloque2 = (bloque2[:insert_pos] + '<w:r><w:t>X</w:t></w:r>' + bloque2[insert_pos:]
                                 if insert_pos >= 0 else bloque2)
            nueva_fila = nueva_fila[:cs2] + nuevo_bloque2 + nueva_fila[ce2:]

            xml = xml[:ini] + nueva_fila + xml[fin:]
            cambiado = True
            break  # posiciones de trs cambiaron - reiniciar el escaneo desde cero
    return xml


# ── USA W CLASS ───────────────────────────────────────────────────────────────
# Sin desglose producto por producto: un solo total, y la clave es la
# CONTRAMARCA (I.10 Marca de embarque). Fecha de producción y Lote son el
# mismo rango de fechas, solo que en formatos distintos (dd/mm/yyyy vs YYYYMMDD).

def _gen_usa_wclass(xml, datos):
    alertas = []
    trs = get_trs(xml)

    total_cajas = str(datos.get('total_cajas', '') or '')
    contramarca = datos.get('contramarca', '') or ''
    if not contramarca:
        alertas.append('Contramarca no encontrada en el remito - completar manualmente')

    f_faena = datos.get('fecha_faena', '') or ''
    f_prod  = datos.get('fecha_produccion', '') or ''
    f_venc  = datos.get('fecha_vencimiento', '') or ''
    if not f_prod:
        alertas.append('Fecha de producción no encontrada en el piqueo - completar manualmente')
    if not f_venc:
        alertas.append('Fecha límite de conservación no encontrada - completar manualmente')

    peso_neto_kg  = str(datos.get('total_neto', '') or '')
    peso_neto_lbs = kg_a_lbs(peso_neto_kg)

    f_faena_fmt    = fmt_fecha_al_to_usa(f_faena)
    fecha_prod_fmt = fmt_fecha_al_to_usa(f_prod)
    lote_fmt       = fecha_a_lote_usa(f_prod)
    fecha_venc_fmt = fmt_fecha_al_to_usa(f_venc)
    fecha_emi = datos.get('fecha_emision') or datetime.datetime.now().strftime('%d/%m/%Y')

    # Filas de producto ES/EN - se ubican por el texto fijo de categoria (no
    # por indice de fila fijo, para no depender de cuantas filas de encabezado
    # tenga cada version de la plantilla)
    _, _, _, idx_es = _get_fila_por_contenido(xml, trs, 'PRODUCTO CRUDO INTACTO')
    if idx_es is not None:
        fila_es, ini_es, fin_es = get_fila_xml(xml, trs, idx_es)
        fila_en, ini_en, fin_en = get_fila_xml(xml, trs, idx_es + 1)

        nueva_es = fila_es
        if total_cajas:    nueva_es = _reemplazar_celda(nueva_es, 0, total_cajas)
        if f_faena_fmt:    nueva_es = _reemplazar_celda(nueva_es, 5, f_faena_fmt)
        if fecha_prod_fmt: nueva_es = _reemplazar_celda(nueva_es, 6, fecha_prod_fmt)
        if contramarca:    nueva_es = _reemplazar_celda(nueva_es, 7, contramarca)
        if lote_fmt:       nueva_es = _reemplazar_celda(nueva_es, 8, lote_fmt)
        if peso_neto_kg:   nueva_es = _reemplazar_celda(nueva_es, 9, peso_neto_kg.replace('.', ',') + ' KGS')

        nueva_en = fila_en
        if f_faena_fmt:    nueva_en = _reemplazar_celda(nueva_en, 5, f_faena_fmt)
        if fecha_prod_fmt: nueva_en = _reemplazar_celda(nueva_en, 6, fecha_prod_fmt)
        if contramarca:    nueva_en = _reemplazar_celda(nueva_en, 7, contramarca)
        if lote_fmt:       nueva_en = _reemplazar_celda(nueva_en, 8, lote_fmt)
        if peso_neto_lbs:  nueva_en = _reemplazar_celda(nueva_en, 9, peso_neto_lbs.replace('.', ',') + ' LBS')

        xml = xml[:ini_es] + nueva_es + nueva_en + xml[fin_en:]

    # Fila de Totales - se ubica por el texto "Totales / Total", no por indice fijo
    trs2 = get_trs(xml)
    _, _, _, idx_tot = _get_fila_por_contenido(xml, trs2, 'Totales / Total')
    if idx_tot is not None:
        fila_tot, ini_tot, fin_tot = get_fila_xml(xml, trs2, idx_tot)
        nueva_tot = fila_tot
        if total_cajas: nueva_tot = _reemplazar_celda(nueva_tot, 0, total_cajas)
        celda_starts = [m.start() for m in re.finditer(r'<w:tc>', nueva_tot)]
        celda_ends   = [m.start() for m in re.finditer(r'</w:tc>', nueva_tot)]
        if len(celda_starts) > 2:
            bloque = nueva_tot[celda_starts[2]:celda_ends[2]]
            if peso_neto_kg:
                kg_fmt = peso_neto_kg.replace('.', ',')
                bloque = re.sub(r'[\d\.,]+((?:\s|<[^>]+>)*?KGS)', kg_fmt + r'\1', bloque, count=1)
            if peso_neto_lbs:
                lbs_fmt = peso_neto_lbs.replace('.', ',')
                bloque = re.sub(r'[\d\.,]+((?:\s|<[^>]+>)*?LBS)', lbs_fmt + r'\1', bloque, count=1)
            nueva_tot = nueva_tot[:celda_starts[2]] + bloque + nueva_tot[celda_ends[2]:]
        xml = xml[:ini_tot] + nueva_tot + xml[fin_tot:]

    # Fecha limite de conservacion (I.15) - anclada al label, no al valor de ejemplo
    if fecha_venc_fmt:
        patron_venc = re.compile(
            r'(Limit conservation date:(?:\s|<[^>]+>)*?)\d{2}/\d{2}/\d{4}\s+al/to\s+\d{2}/\d{2}/\d{4}',
            re.IGNORECASE
        )
        xml = patron_venc.sub(lambda m: m.group(1) + fecha_venc_fmt, xml, count=1)

    # Fecha de emision (pie del certificado) - siempre es la ultima fecha
    # dd/mm/yyyy del documento, posicionalmente.
    todas_fechas = list(re.finditer(r'\d{2}/\d{2}/\d{4}', xml))
    if todas_fechas:
        ultima = todas_fechas[-1]
        xml = xml[:ultima.start()] + fecha_emi + xml[ultima.end():]

    return xml, alertas


# ── USA ORLEANS ───────────────────────────────────────────────────────────────
# Detalle producto por producto en un ANEXO (pagina aparte), 2 filas por producto
# (ES/EN) igual que Wclass. A diferencia de Wclass, cada producto tiene su propia
# Contramarca, Fecha de faena y Fecha de produccion/Lote (via piqueo por Cod Prod
# y provisorio por linea, cruzados por cajas+neto+bruto).

def _gen_usa_orleans(xml, datos):
    alertas = []
    trs = get_trs(xml)

    _, _, _, primera_idx = _get_fila_por_contenido(xml, trs, 'PRODUCTO CRUDO INTACTO')
    if primera_idx is None:
        return xml, ['No se encontro la fila modelo de productos en la plantilla Orleans']

    total_idx = None
    for i in range(primera_idx, len(trs)):
        fila, _, _ = get_fila_xml(xml, trs, i)
        if 'Totales' in fila:
            total_idx = i
            break
    if total_idx is None:
        return xml, ['No se encontro la fila de Totales del anexo en la plantilla Orleans']

    fila_es, ini_mod, _ = get_fila_xml(xml, trs, primera_idx)
    fila_en, _, _       = get_fila_xml(xml, trs, primera_idx + 1)
    ini_totales = trs[total_idx].start()

    nuevas_filas = ''
    for prod in datos.get('productos', []):
        contramarca = prod.get('contramarca', '') or ''
        f_faena = prod.get('fecha_faena_prod', '') or ''
        f_prod  = prod.get('fecha_produccion_prod', '') or ''
        if not contramarca: alertas.append('Producto ' + prod.get('codigo', '') + ': contramarca no encontrada - completar manualmente')
        if not f_faena:      alertas.append('Producto ' + prod.get('codigo', '') + ': fecha de faena no encontrada - completar manualmente')
        if not f_prod:       alertas.append('Producto ' + prod.get('codigo', '') + ': fecha de produccion no encontrada - completar manualmente')

        f_faena_fmt = fmt_fecha_al_to_usa(f_faena)
        f_prod_fmt  = fmt_fecha_al_to_usa(f_prod)
        lote_fmt    = fecha_a_lote_usa(f_prod)
        neto_kg  = formatear_miles(prod.get('neto', '')) + ' KGS'
        neto_lbs = formatear_miles(kg_a_lbs(prod.get('neto', ''))) + ' LBS'

        nueva_es = fila_es
        nueva_es = _reemplazar_celda(nueva_es, 0, str(prod.get('cajas', '')))
        nueva_es = _reemplazar_celda(nueva_es, 1, (prod.get('nombre_es', '') or '').strip().upper())
        nueva_es = _reemplazar_celda(nueva_es, 5, f_faena_fmt)
        nueva_es = _reemplazar_celda(nueva_es, 6, f_prod_fmt)
        nueva_es = _reemplazar_celda(nueva_es, 7, contramarca)
        nueva_es = _reemplazar_celda(nueva_es, 8, lote_fmt)
        nueva_es = _reemplazar_celda(nueva_es, 9, neto_kg)

        nueva_en = fila_en
        nueva_en = _reemplazar_celda(nueva_en, 1, (prod.get('nombre_en', '') or '').strip().upper())
        nueva_en = _reemplazar_celda(nueva_en, 5, f_faena_fmt)
        nueva_en = _reemplazar_celda(nueva_en, 6, f_prod_fmt)
        nueva_en = _reemplazar_celda(nueva_en, 7, contramarca)
        nueva_en = _reemplazar_celda(nueva_en, 8, lote_fmt)
        nueva_en = _reemplazar_celda(nueva_en, 9, neto_lbs)

        nuevas_filas += nueva_es + nueva_en

    xml = xml[:ini_mod] + nuevas_filas + xml[ini_totales:]

    # Totales (aparecen 2 veces: resumen en pagina 1 "VER ANEXO" y al pie del
    # anexo). Se ubican por el texto "Totales / Total" en vez de asumir un
    # indice de fila fijo, y el KG/LBS se reemplaza tolerando que el numero y
    # la unidad puedan quedar en runs XML separados.
    total_cajas = str(datos.get('total_cajas', '') or '')
    total_neto_fmt = formatear_miles(datos.get('total_neto', ''))
    total_lbs_fmt  = formatear_miles(kg_a_lbs(datos.get('total_neto', '')))

    trs_tot = get_trs(xml)
    filas_totales = []
    for i, m in enumerate(trs_tot):
        fila_cand, _, _ = get_fila_xml(xml, trs_tot, i)
        if 'Totales' in fila_cand and 'Total' in fila_cand:
            filas_totales.append(i)
    # Procesar de atras para adelante para no invalidar los offsets ya calculados
    for idx_f in reversed(filas_totales):
        fila_t, ini_t, fin_t = get_fila_xml(xml, get_trs(xml), idx_f)
        nueva_t = fila_t
        if total_cajas:
            nueva_t = _reemplazar_celda(nueva_t, 0, total_cajas)
        if total_neto_fmt:
            nueva_t = re.sub(r'[\d\.,]+((?:\s|<[^>]+>)*?KGS)', total_neto_fmt + r'\1', nueva_t, count=1)
        if total_lbs_fmt:
            nueva_t = re.sub(r'[\d\.,]+((?:\s|<[^>]+>)*?LBS)', total_lbs_fmt + r'\1', nueva_t, count=1)
        trs_actual = get_trs(xml)
        _, ini_t2, fin_t2 = get_fila_xml(xml, trs_actual, idx_f)
        xml = xml[:ini_t2] + nueva_t + xml[fin_t2:]

    # Transporte (buque - Orleans es maritimo) - anclado al label
    transporte = datos.get('transporte', '') or ''
    if transporte:
        xml = _reemplazar_tras_label(xml, r'Buque:\s*/\s*Vessel\s*:', transporte)

    # Contenedor - anclado al label
    contenedor = datos.get('contenedor', '') or ''
    if contenedor:
        xml = _reemplazar_tras_label(xml, r'Contenedor N°\)\s*/\s*Container Identification:', contenedor)
    if not contenedor: alertas.append('Contenedor no encontrado - completar manualmente')

    # Precinto (un solo campo en esta plantilla - se usa el de AFIP) - anclado al label
    precinto = datos.get('precinto_afip') or datos.get('precinto_senasa') or ''
    if precinto:
        xml = _reemplazar_tras_label(xml, r'Precinto/s\s*/\s*Seal/s:', precinto)
    if not precinto: alertas.append('Precinto no encontrado - completar manualmente')

    # Fecha limite de conservacion (I.15) - anclada al label
    f_venc_fmt = fmt_fecha_al_to_usa(datos.get('fecha_vencimiento', '') or '')
    if f_venc_fmt:
        patron_venc = re.compile(
            r'(Limit conservation date:(?:\s|<[^>]+>)*?)\d{2}/\d{2}/\d{4}\s+al/to\s+\d{2}/\d{2}/\d{4}',
            re.IGNORECASE
        )
        xml = patron_venc.sub(lambda m: m.group(1) + f_venc_fmt, xml, count=1)

    # Fecha de emision (aparece 2 veces: certificacion pag.2 y firma del anexo
    # pag.3) - son siempre las 2 ultimas fechas dd/mm/yyyy del documento.
    fecha_emi = datos.get('fecha_emision') or datetime.datetime.now().strftime('%d/%m/%Y')
    todas_fechas = list(re.finditer(r'\d{2}/\d{2}/\d{4}', xml))
    for m in reversed(todas_fechas[-2:]):
        xml = xml[:m.start()] + fecha_emi + xml[m.end():]

    return xml, alertas


# ── HONG KONG (4 variantes: congelado/enfriado x aereo/maritimo) ────────────
# Una sola linea por producto (no ES/EN separadas), sin columna de bruto
# (solo peso neto), fecha de produccion como rango unico por envio (no por
# producto), y el checkbox de temperatura ya viene fijo en la plantilla
# correcta (hay un archivo por combinacion, no se mueve dinamicamente).

def _gen_hongkong(xml, datos, es_congelado, tipo_via):
    alertas = []
    trs = get_trs(xml)

    _, _, _, header_idx = _get_fila_por_contenido(xml, trs, 'packages')
    primera_idx = (header_idx + 1) if header_idx is not None else 6

    _, _, _, total_idx = _get_fila_por_contenido(xml, trs, 'Totales / Total')
    if total_idx is None:
        total_idx = primera_idx + 6

    fila_modelo, ini_mod, _ = get_fila_xml(xml, trs, primera_idx)
    fila_total, ini_tot, fin_tot = get_fila_xml(xml, trs, total_idx)

    f_prod_fmt = fmt_fecha_al_to(datos.get('fecha_produccion', '') or '')
    marca = datos.get('marca', '') or ''
    if not marca:
        alertas.append('Marca no encontrada en el remito - completar manualmente')

    nuevas_filas = ''
    for prod in datos.get('productos', []):
        nombre_bi = armar_nombre_hongkong(prod, es_congelado)
        nueva = fila_modelo
        nueva = _reemplazar_celda(nueva, 0, str(prod.get('cajas', '')))
        nueva = _reemplazar_celda(nueva, 1, nombre_bi)
        nueva = _reemplazar_celda(nueva, 2, f_prod_fmt)
        nueva = _reemplazar_celda(nueva, 3, marca)
        nueva = _reemplazar_celda(nueva, 5, str(prod.get('neto', '')))
        nuevas_filas += nueva

    nueva_total = fila_total
    nueva_total = _reemplazar_celda(nueva_total, 0, str(datos.get('total_cajas', '')))
    nueva_total = _reemplazar_celda(nueva_total, 2, str(datos.get('total_neto', '')))

    xml = xml[:ini_mod] + nuevas_filas + nueva_total + xml[fin_tot:]

    # Transporte - Aereo (vuelo + AWB) o Maritimo (buque + referencia BL), segun cual exista en esta plantilla
    if tipo_via == 'aereo':
        vuelo = datos.get('transporte', '') or ''
        if vuelo: xml = xml.replace('LH511', vuelo)
        awb = datos.get('awb', '') or datos.get('pedido_referencia', '') or ''
        if awb: xml = re.sub(r'AWB:\s*020-05990725', 'AWB:  ' + awb, xml)
    else:
        buque = datos.get('transporte', '') or ''
        if buque: xml = xml.replace('SAN LORENZO MAERSK', buque)
        contenedor = datos.get('contenedor', '') or ''
        if contenedor: xml = xml.replace('MNBU386303-5', contenedor)
        precinto = datos.get('precinto_afip') or datos.get('precinto_senasa') or ''
        if precinto: xml = xml.replace('BAH74888', precinto)

    # Fecha limite de conservacion (I.11)
    f_venc_fmt = fmt_fecha_to_hongkong(datos.get('fecha_vencimiento', '') or '')
    if f_venc_fmt:
        xml = re.sub(r'\d{2}/\d{2}/\d{4}\s+to\s+\d{2}/\d{2}/\d{4}', f_venc_fmt, xml, count=1)

    # Fecha de emision (pie del certificado) - siempre es la ultima fecha
    # dd/mm/yyyy del documento, posicionalmente (evita hardcodear el valor de
    # ejemplo de una plantilla puntual, que difiere entre congelado/enfriado).
    fecha_emi = datos.get('fecha_emision') or datetime.datetime.now().strftime('%d/%m/%Y')
    todas_fechas = list(re.finditer(r'\d{2}/\d{2}/\d{4}', xml))
    if todas_fechas:
        ultima = todas_fechas[-1]
        xml = xml[:ultima.start()] + fecha_emi + xml[ultima.end():]

    return xml, alertas


# ── USA ALLE PROCESSING CORP (condimentada / natural) ────────────────────────
# Igual estructura que Wclass (un solo total, sin desglose producto por
# producto), pero es maritimo (no aereo), SI completa fecha de faena, y el
# precinto es un solo campo (no combinado AFIP/SENASA). Las 2 variantes
# (con/sin condimentar) son la MISMA logica - la diferencia es pura descripcion
# fija ya horneada en cada plantilla, elegida por el destino seleccionado.

def _gen_alle_processing(xml, datos):
    alertas = []

    total_cajas = str(datos.get('total_cajas', '') or '')
    contramarca = datos.get('contramarca', '') or ''
    if not contramarca:
        alertas.append('Contramarca no encontrada en el remito - completar manualmente')

    f_faena = datos.get('fecha_faena', '') or ''
    f_prod  = datos.get('fecha_produccion', '') or ''
    f_venc  = datos.get('fecha_vencimiento', '') or ''
    if not f_faena: alertas.append('Fecha de faena no encontrada - completar manualmente')
    if not f_prod:  alertas.append('Fecha de producción no encontrada en el piqueo - completar manualmente')
    if not f_venc:  alertas.append('Fecha límite de conservación no encontrada - completar manualmente')

    peso_neto_kg  = str(datos.get('total_neto', '') or '')
    peso_neto_lbs = kg_a_lbs(peso_neto_kg)

    f_faena_fmt = fmt_fecha_al_to_usa(f_faena)
    f_prod_fmt  = fmt_fecha_al_to_usa(f_prod)
    lote_fmt    = fecha_a_lote_usa(f_prod)
    f_venc_fmt  = fmt_fecha_al_to_usa(f_venc)
    fecha_emi = datos.get('fecha_emision') or datetime.datetime.now().strftime('%d/%m/%Y')

    # Bultos - aparece 2 veces (fila ES, fila EN, mas la fila de Totales)
    if total_cajas:
        # Reemplazo por valor de ejemplo conocido de la plantilla (220 / 1076 segun variante)
        for viejo in ['>220<', '>1076<']:
            if viejo in xml:
                xml = xml.replace(viejo, '>' + total_cajas + '<')

    # Fecha de faena, fecha de produccion y contramarca - aparecen 2 veces (fila ES y fila EN)
    if f_faena_fmt:
        xml = xml.replace('31/08/2026 al/to 07/09/2026', f_faena_fmt)
    if f_prod_fmt:
        xml = xml.replace('01/09/2026 al/to 09/09/2026', f_prod_fmt)
        xml = xml.replace('04/09/2026 al/to 09/09/2026', f_prod_fmt)
    if contramarca:
        xml = xml.replace('C289', contramarca)
        xml = xml.replace('C290', contramarca)
    if lote_fmt:
        xml = xml.replace('20260901 al/to 20260909', lote_fmt)
        xml = xml.replace('20260904 al/to 20260909', lote_fmt)

    # Peso neto en KGS y en LBS (fila ES/EN por separado, y el total en una celda
    # combinada - a veces el numero y la unidad quedan en runs separados por
    # tener distinto color, por eso el regex tolera tags XML de por medio)
    if peso_neto_kg:
        kg_fmt = peso_neto_kg.replace('.', ',')
        for viejo_kg in ['3813,00', '21263,00']:
            xml = re.sub(re.escape(viejo_kg) + r'((?:\s|<[^>]+>)*?)KGS', kg_fmt + r'\1KGS', xml)
    if peso_neto_lbs:
        lbs_fmt = peso_neto_lbs.replace('.', ',')
        for viejo_lbs in ['8406,22', '46876,84']:
            xml = re.sub(re.escape(viejo_lbs) + r'((?:\s|<[^>]+>)*?)LBS', lbs_fmt + r'\1LBS', xml)

    # Transporte (buque - es maritimo)
    transporte = datos.get('transporte', '') or ''
    if transporte: xml = xml.replace('STEPHANIE C', transporte)

    # Contenedor
    contenedor = datos.get('contenedor', '') or ''
    if contenedor: xml = xml.replace('MMAU1250441', contenedor)
    if not contenedor: alertas.append('Contenedor no encontrado - completar manualmente')

    # Precinto (un solo campo en esta plantilla)
    precinto = datos.get('precinto_afip') or datos.get('precinto_senasa') or ''
    if precinto: xml = xml.replace('BAH74945', precinto)
    if not precinto: alertas.append('Precinto no encontrado - completar manualmente')

    # Fecha limite de conservacion (I.15)
    if f_venc_fmt:
        xml = xml.replace('31/08/2028 al/to 07/09/2028', f_venc_fmt)

    # Fecha de emision (pie del certificado) - esta plantilla la deja vacia por
    # defecto (no trae un valor de ejemplo), hay que insertarla.
    xml = xml.replace('Date:   </w:t>', 'Date:   ' + fecha_emi + '</w:t>', 1)

    return xml, alertas


# ── FILIPINAS (aereo/maritimo) ───────────────────────────────────────────────
# Nombre bilingue en UNA sola linea (no ES/EN separadas como Malasia/Singapur).
# Numeros en formato ingles/EEUU (coma miles, punto decimal). Fecha de faena/
# produccion/vencimiento son un rango unico por envio (no por producto). El
# checkbox de temperatura se mueve dinamicamente (un solo archivo cubre
# enfriado y congelado), igual que Malasia/Singapur.

def _gen_filipinas_aereo(xml, datos):
    """Filipinas AEREO - funcion completamente independiente de la de maritimo
    (plantilla propia, con su propia fila de pallets y sus propios valores de
    ejemplo para transporte/contenedor/precinto)."""
    alertas = []
    trs = get_trs(xml)

    _, _, _, header_idx = _get_fila_por_contenido(xml, trs, 'Number of packages')
    primera_idx = (header_idx + 1) if header_idx is not None else 5

    fila_pal, ini_pal, fin_pal, idx_pal = _get_fila_por_contenido(xml, trs, 'ACONDICIONAD')

    _, _, _, total_idx = _get_fila_por_contenido(xml, trs, 'Total / es')
    if total_idx is None:
        total_idx = primera_idx + 6

    fila_modelo, ini_mod, _ = get_fila_xml(xml, trs, primera_idx)
    fila_total, ini_tot, fin_tot = get_fila_xml(xml, trs, total_idx)

    es_congelado = datos.get('es_congelado', False)

    nuevas_filas = ''
    for prod in datos.get('productos', []):
        nombre_bi = armar_nombre_filipinas(prod, es_congelado)
        nueva = fila_modelo
        nueva = _reemplazar_celda(nueva, 0, str(prod.get('cajas', '')))
        nueva = _reemplazar_celda(nueva, 1, nombre_bi)
        nueva = _reemplazar_celda(nueva, 6, formatear_miles_en(prod.get('neto', '')))
        nueva = _reemplazar_celda(nueva, 7, formatear_miles_en(prod.get('bruto', '')))
        nuevas_filas += nueva

    # Filipinas Aereo SIEMPRE tiene la fila de pallets, y su peso SIEMPRE se
    # suma al bruto total (asi es esta plantilla especificamente)
    pallets = datos.get('pallets', '') or ''
    kg_pallets = datos.get('kg_pallets', '') or ''
    nueva_pal = _reemplazar_pallets_en_fila(fila_pal, pallets, kg_pallets) if (fila_pal and pallets) else (fila_pal or '')

    total_bruto = datos.get('total_bruto', '')
    if kg_pallets:
        try:
            total_bruto = '{:.2f}'.format(float(total_bruto) + float(kg_pallets))
        except (TypeError, ValueError):
            pass

    nueva_total = fila_total
    nueva_total = _reemplazar_celda(nueva_total, 0, str(datos.get('total_cajas', '')))
    nueva_total = _reemplazar_celda(nueva_total, 2, formatear_miles_en(datos.get('total_neto', '')))
    nueva_total = _reemplazar_celda(nueva_total, 3, formatear_miles_en(total_bruto))

    xml = xml[:ini_mod] + nuevas_filas + nueva_pal + nueva_total + xml[fin_tot:]

    # Fechas de faena / produccion / vencimiento (rango unico por envio)
    trs2 = get_trs(xml)
    xml = _reemplazar_fechas(xml, trs2, datos.get('fecha_faena', ''), datos.get('fecha_produccion', ''),
                              datos.get('fecha_vencimiento', ''), fmt_fecha_al_to)

    # Temperatura - un solo archivo cubre enfriado y congelado, se mueve la X
    xml = _set_temperatura_singapur(xml, es_congelado, tipo_via='aereo')

    # Transporte (vuelo) - propio de esta plantilla
    transporte = datos.get('transporte', '') or ''
    if transporte: xml = xml.replace('EK248', transporte)

    # Contenedor - esta plantilla lo deja en guiones por defecto (tipico de aereo, sin contenedor)
    contenedor = datos.get('contenedor', '') or ''
    if contenedor:
        xml = _reemplazar_tras_label(xml, r'Container\(s\) number:', contenedor)

    # Precinto - idem, en guiones por defecto
    precinto = datos.get('precinto_afip') or datos.get('precinto_senasa') or ''
    if precinto:
        xml = _reemplazar_tras_label(xml, r'Seal number:', precinto)

    # Fecha de emision (pie del certificado) - la ultima fecha dd/mm/yyyy del documento
    fecha_emi = datos.get('fecha_emision') or datetime.datetime.now().strftime('%d/%m/%Y')
    todas_fechas = list(re.finditer(r'\d{2}/\d{2}/\d{4}', xml))
    if todas_fechas:
        ultima = todas_fechas[-1]
        xml = xml[:ultima.start()] + fecha_emi + xml[ultima.end():]

    return xml, alertas


def _gen_filipinas_maritimo(xml, datos):
    """Filipinas MARITIMO - funcion completamente independiente de la de
    aereo (esta plantilla no tiene fila de pallets separada; el bruto del
    remito ya viene completo tal cual, sin sumarle nada)."""
    alertas = []
    trs = get_trs(xml)

    _, _, _, header_idx = _get_fila_por_contenido(xml, trs, 'Number of packages')
    primera_idx = (header_idx + 1) if header_idx is not None else 5

    _, _, _, total_idx = _get_fila_por_contenido(xml, trs, 'Total / es')
    if total_idx is None:
        total_idx = primera_idx + 6

    fila_modelo, ini_mod, _ = get_fila_xml(xml, trs, primera_idx)
    fila_total, ini_tot, fin_tot = get_fila_xml(xml, trs, total_idx)

    es_congelado = datos.get('es_congelado', False)

    nuevas_filas = ''
    for prod in datos.get('productos', []):
        nombre_bi = armar_nombre_filipinas(prod, es_congelado)
        nueva = fila_modelo
        nueva = _reemplazar_celda(nueva, 0, str(prod.get('cajas', '')))
        nueva = _reemplazar_celda(nueva, 1, nombre_bi)
        nueva = _reemplazar_celda(nueva, 6, formatear_miles_en(prod.get('neto', '')))
        nueva = _reemplazar_celda(nueva, 7, formatear_miles_en(prod.get('bruto', '')))
        nuevas_filas += nueva

    nueva_total = fila_total
    nueva_total = _reemplazar_celda(nueva_total, 0, str(datos.get('total_cajas', '')))
    nueva_total = _reemplazar_celda(nueva_total, 2, formatear_miles_en(datos.get('total_neto', '')))
    nueva_total = _reemplazar_celda(nueva_total, 3, formatear_miles_en(datos.get('total_bruto', '')))

    xml = xml[:ini_mod] + nuevas_filas + nueva_total + xml[fin_tot:]

    # Fechas de faena / produccion / vencimiento (rango unico por envio)
    trs2 = get_trs(xml)
    xml = _reemplazar_fechas(xml, trs2, datos.get('fecha_faena', ''), datos.get('fecha_produccion', ''),
                              datos.get('fecha_vencimiento', ''), fmt_fecha_al_to)

    # Temperatura - un solo archivo cubre enfriado y congelado, se mueve la X
    xml = _set_temperatura_singapur(xml, es_congelado, tipo_via='maritimo')

    # Transporte (buque) - propio de esta plantilla
    transporte = datos.get('transporte', '') or ''
    if transporte: xml = xml.replace('ZIM USA', transporte)

    # Contenedor - propio de esta plantilla
    contenedor = datos.get('contenedor', '') or ''
    if contenedor: xml = xml.replace('MNBU4390003', contenedor)
    if not contenedor: alertas.append('Contenedor no encontrado - completar manualmente')

    # Precinto - propio de esta plantilla
    precinto = datos.get('precinto_afip') or datos.get('precinto_senasa') or ''
    if precinto: xml = xml.replace('BAH74877', precinto)
    if not precinto: alertas.append('Precinto no encontrado - completar manualmente')

    # Fecha de emision (pie del certificado) - la ultima fecha dd/mm/yyyy del documento
    fecha_emi = datos.get('fecha_emision') or datetime.datetime.now().strftime('%d/%m/%Y')
    todas_fechas = list(re.finditer(r'\d{2}/\d{2}/\d{4}', xml))
    if todas_fechas:
        ultima = todas_fechas[-1]
        xml = xml[:ultima.start()] + fecha_emi + xml[ultima.end():]

    return xml, alertas


# ── ECUADOR ───────────────────────────────────────────────────────────────
# Nombre en una sola columna en español (sin bilingue). Fusiona filas que son
# el mismo corte pero difieren solo en rango de peso (ver
# fusionar_productos_ecuador). Un solo archivo cubre enfriado/congelado (se
# mueve la X). El total bruto NO suma el peso de pallets (viene completo tal
# cual del remito, igual que Mexico).

def _gen_ecuador(xml, datos):
    alertas = []
    trs = get_trs(xml)

    _, _, _, header_idx = _get_fila_por_contenido(xml, trs, 'Descripción de la mercadería')
    primera_idx = (header_idx + 1) if header_idx is not None else 5

    fila_pal, ini_pal, fin_pal, idx_pal = _get_fila_por_contenido(xml, trs, 'ACONDICIONADO EN')

    _, _, _, total_idx = _get_fila_por_contenido(xml, trs, 'Total / es')
    if total_idx is None:
        total_idx = primera_idx + 10

    fila_modelo, ini_mod, _ = get_fila_xml(xml, trs, primera_idx)
    fila_total, ini_tot, fin_tot = get_fila_xml(xml, trs, total_idx)

    productos_fusionados = fusionar_productos_ecuador(datos.get('productos', []))

    nuevas_filas = ''
    for prod in productos_fusionados:
        nombre = limpiar_nombre_ecuador(prod.get('desc_original', ''))
        nueva = fila_modelo
        nueva = _reemplazar_celda(nueva, 0, str(prod.get('cajas', '')))
        nueva = _reemplazar_celda(nueva, 1, nombre)
        nueva = _reemplazar_celda(nueva, 6, prod.get('neto', ''))
        nueva = _reemplazar_celda(nueva, 7, prod.get('bruto', ''))
        nuevas_filas += nueva

    pallets = datos.get('pallets', '') or ''
    kg_pallets_raw = datos.get('kg_pallets', '') or ''
    kg_pallets = kg_pallets_raw.replace('.', ',') if kg_pallets_raw else ''
    nueva_pal = _reemplazar_pallets_en_fila(fila_pal, pallets, kg_pallets) if (fila_pal and pallets) else (fila_pal or '')

    nueva_total = fila_total
    nueva_total = _reemplazar_celda(nueva_total, 0, str(datos.get('total_cajas', '')))
    nueva_total = _reemplazar_celda(nueva_total, 2, formatear_miles_en(datos.get('total_neto', '')))
    nueva_total = _reemplazar_celda(nueva_total, 3, formatear_miles_en(datos.get('total_bruto', '')))

    xml = xml[:ini_mod] + nuevas_filas + nueva_pal + nueva_total + xml[fin_tot:]

    # Fechas de faena / produccion / vencimiento (rango unico por envio)
    trs2 = get_trs(xml)
    xml = _reemplazar_fechas(xml, trs2, datos.get('fecha_faena', ''), datos.get('fecha_produccion', ''),
                              datos.get('fecha_vencimiento', ''), fmt_fecha_al_minuscula)

    # Temperatura - un solo archivo cubre enfriado y congelado, se mueve la X
    es_congelado = datos.get('es_congelado', False)
    xml = _set_temperatura_singapur(xml, es_congelado, tipo_via='maritimo')

    # Transporte (buque)
    transporte = datos.get('transporte', '') or ''
    if transporte: xml = xml.replace('SANTA VANESSA', transporte)

    # Contenedor
    contenedor = datos.get('contenedor', '') or ''
    if contenedor: xml = xml.replace('UACU479172-0', contenedor)
    if not contenedor: alertas.append('Contenedor no encontrado - completar manualmente')

    # Precinto (un solo campo)
    precinto = datos.get('precinto_afip') or datos.get('precinto_senasa') or ''
    if precinto: xml = xml.replace('BAH79592', precinto)
    if not precinto: alertas.append('Precinto no encontrado - completar manualmente')

    # Fecha de emision (pie del certificado) - la ultima fecha dd/mm/yyyy del documento
    fecha_emi = datos.get('fecha_emision') or datetime.datetime.now().strftime('%d/%m/%Y')
    todas_fechas = list(re.finditer(r'\d{2}/\d{2}/\d{4}', xml))
    if todas_fechas:
        ultima = todas_fechas[-1]
        xml = xml[:ultima.start()] + fecha_emi + xml[ultima.end():]

    return xml, alertas


# ── EGIPTO ────────────────────────────────────────────────────────────────
# Nombre bilingue de una sola linea (como Malasia/Singapur combinado en un
# solo renglon). Numeros en formato simple (punto decimal, sin separador de
# miles). Sin fila de pallets. El total bruto no suma nada extra (viene
# completo del remito, igual que Mexico/Ecuador).

def _gen_egipto(xml, datos):
    alertas = []
    trs = get_trs(xml)

    _, _, _, header_idx = _get_fila_por_contenido(xml, trs, 'Description of goods')
    primera_idx = (header_idx + 1) if header_idx is not None else 5

    _, _, _, total_idx = _get_fila_por_contenido(xml, trs, 'Total / es')
    if total_idx is None:
        total_idx = primera_idx + 4

    fila_modelo, ini_mod, _ = get_fila_xml(xml, trs, primera_idx)
    fila_total, ini_tot, fin_tot = get_fila_xml(xml, trs, total_idx)

    nuevas_filas = ''
    for prod in datos.get('productos', []):
        nombre_bi = armar_nombre_egipto(prod)
        nueva = fila_modelo
        nueva = _reemplazar_celda(nueva, 0, str(prod.get('cajas', '')))
        nueva = _reemplazar_celda(nueva, 1, nombre_bi)
        nueva = _reemplazar_celda(nueva, 6, prod.get('neto', ''))
        nueva = _reemplazar_celda(nueva, 7, prod.get('bruto', ''))
        nuevas_filas += nueva

    nueva_total = fila_total
    nueva_total = _reemplazar_celda(nueva_total, 0, str(datos.get('total_cajas', '')))
    nueva_total = _reemplazar_celda(nueva_total, 2, str(datos.get('total_neto', '')))
    nueva_total = _reemplazar_celda(nueva_total, 3, str(datos.get('total_bruto', '')))

    xml = xml[:ini_mod] + nuevas_filas + nueva_total + xml[fin_tot:]

    # Fechas de faena / produccion / vencimiento (rango unico por envio)
    trs2 = get_trs(xml)
    xml = _reemplazar_fechas(xml, trs2, datos.get('fecha_faena', ''), datos.get('fecha_produccion', ''),
                              datos.get('fecha_vencimiento', ''), fmt_fecha_al)

    # Temperatura - un solo archivo cubre enfriado y congelado, se mueve la X
    es_congelado = datos.get('es_congelado', False)
    xml = _set_temperatura_singapur(xml, es_congelado, tipo_via='maritimo')

    # Transporte (buque)
    transporte = datos.get('transporte', '') or ''
    if transporte: xml = xml.replace('MAERSK LONDRINA', transporte)

    # Contenedor
    contenedor = datos.get('contenedor', '') or ''
    if contenedor: xml = xml.replace('SUDU624811-7', contenedor)
    if not contenedor: alertas.append('Contenedor no encontrado - completar manualmente')

    # Precinto (un solo campo)
    precinto = datos.get('precinto_afip') or datos.get('precinto_senasa') or ''
    if precinto: xml = xml.replace('BAH74974', precinto)
    if not precinto: alertas.append('Precinto no encontrado - completar manualmente')

    # Fecha de emision (pie del certificado) - la ultima fecha dd/mm/yyyy del documento
    fecha_emi = datos.get('fecha_emision') or datetime.datetime.now().strftime('%d/%m/%Y')
    todas_fechas = list(re.finditer(r'\d{2}/\d{2}/\d{4}', xml))
    if todas_fechas:
        ultima = todas_fechas[-1]
        xml = xml[:ultima.start()] + fecha_emi + xml[ultima.end():]

    return xml, alertas


# ── BRASIL ────────────────────────────────────────────────────────────────
# Nombre bilingue de una sola linea (ES/PT), sin fusionar filas (a diferencia
# de Ecuador, cada linea del remito queda como su propia fila). Cada fila
# lleva el mismo rango de fecha de produccion (no por producto). Es por
# camion (frontera terrestre), no buque/avion - usa la patente del
# camion/acoplado del remito en vez de Buque/Aerolinea. Precinto SENASA y
# AFIP van en 2 campos separados (no combinados en uno solo).

def _gen_brasil(xml, datos):
    alertas = []
    trs = get_trs(xml)

    _, _, _, idx_subheader = _get_fila_por_contenido(xml, trs, 'CARNE BOVINA ENFRIADA SIN HUESO')
    primera_idx = (idx_subheader + 1) if idx_subheader is not None else 6

    fila_pal, ini_pal, fin_pal, idx_pal = _get_fila_por_contenido(xml, trs, 'ACONDICIONADO EM')

    _, _, _, total_idx = _get_fila_por_contenido(xml, trs, 'Total / es')
    if total_idx is None:
        total_idx = primera_idx + 10

    fila_modelo, ini_mod, _ = get_fila_xml(xml, trs, primera_idx)
    fila_total, ini_tot, fin_tot = get_fila_xml(xml, trs, total_idx)

    f_prod_fmt = fmt_fecha_al_minuscula(datos.get('fecha_produccion', '') or '')

    nuevas_filas = ''
    for prod in datos.get('productos', []):
        nombre_bi = armar_nombre_brasil(prod)
        nueva = fila_modelo
        nueva = _reemplazar_celda(nueva, 0, str(prod.get('cajas', '')))
        nueva = _reemplazar_celda(nueva, 1, nombre_bi)
        if f_prod_fmt:
            nueva = _reemplazar_celda(nueva, 5, f_prod_fmt)
        nueva = _reemplazar_celda(nueva, 6, str(prod.get('neto', '')).replace('.', ','))
        nueva = _reemplazar_celda(nueva, 7, str(prod.get('bruto', '')).replace('.', ','))
        nuevas_filas += nueva

    pallets = datos.get('pallets', '') or ''
    kg_pallets_raw = datos.get('kg_pallets', '') or ''
    kg_pallets = kg_pallets_raw.replace('.', ',') if kg_pallets_raw else ''
    nueva_pal = _reemplazar_pallets_en_fila(fila_pal, pallets, kg_pallets) if (fila_pal and pallets) else (fila_pal or '')

    nueva_total = fila_total
    nueva_total = _reemplazar_celda(nueva_total, 0, str(datos.get('total_cajas', '')))
    nueva_total = _reemplazar_celda(nueva_total, 2, str(datos.get('total_neto', '')).replace('.', ','))
    nueva_total = _reemplazar_celda(nueva_total, 3, str(datos.get('total_bruto', '')).replace('.', ','))

    xml = xml[:ini_mod] + nuevas_filas + nueva_pal + nueva_total + xml[fin_tot:]

    # Fechas de faena / produccion / vencimiento (resumen al pie, rango unico por envio)
    trs2 = get_trs(xml)
    xml = _reemplazar_fechas(xml, trs2, datos.get('fecha_faena', ''), datos.get('fecha_produccion', ''),
                              datos.get('fecha_vencimiento', ''), fmt_fecha_al_minuscula)

    # Temperatura - un solo archivo cubre enfriado y congelado, se mueve la X
    es_congelado = datos.get('es_congelado', False)
    xml = _set_temperatura_singapur(xml, es_congelado, tipo_via='maritimo')

    # Transporte - es por camion (frontera terrestre), usa la patente del
    # camion/acoplado del remito en vez de Buque/Aerolinea - anclado al label
    camion_acoplado = datos.get('camion_acoplado', '') or datos.get('camion', '') or ''
    if camion_acoplado:
        xml = _reemplazar_tras_label(xml, r'CAMION CHAPA:', camion_acoplado)
    else:
        alertas.append('Patente de camión/acoplado no encontrada - completar manualmente')

    # Precinto SENASA y AFIP van en 2 campos separados con el MISMO label
    # repetido (ES y PT) - se reemplaza por orden de aparicion, no por texto
    # literal (cada plantilla trae numeros de ejemplo distintos)
    precinto_senasa = datos.get('precinto_senasa', '') or ''
    precinto_afip = datos.get('precinto_afip', '') or ''
    xml = _reemplazar_ocurrencias_por_indice(xml, r'Número de precinto:', [precinto_senasa, precinto_afip])
    if not (precinto_senasa or precinto_afip):
        alertas.append('Precinto no encontrado - completar manualmente')

    # Fecha de emision (pie del certificado) - la ultima fecha dd/mm/yyyy del documento
    fecha_emi = datos.get('fecha_emision') or datetime.datetime.now().strftime('%d/%m/%Y')
    todas_fechas = list(re.finditer(r'\d{2}/\d{2}/\d{4}', xml))
    if todas_fechas:
        ultima = todas_fechas[-1]
        xml = xml[:ultima.start()] + fecha_emi + xml[ultima.end():]

    return xml, alertas


# ── PERU ENFRIADO ─────────────────────────────────────────────────────────
# Nombre en una sola columna en español (sin bilingue), tomado por CODIGO
# desde MAPA_PERU_ENFRIADO (tabla propia de Angie), con regla de limpieza de
# texto como respaldo. Numeros con coma de miles (formato ingles) en todos
# lados (productos y totales). Terrestre con transito por Chile - la segunda
# patente del camion viene del provisorio, no del remito. Precinto SENASA y
# AFIP van combinados en un solo campo con guion.

def _gen_peru_enfriado(xml, datos):
    alertas = []
    trs = get_trs(xml)

    _, _, _, header_idx = _get_fila_por_contenido(xml, trs, 'Descripción de la mercadería')
    primera_idx = (header_idx + 1) if header_idx is not None else 5

    fila_pal, ini_pal, fin_pal, idx_pal = _get_fila_por_contenido(xml, trs, 'ACONDICIONAD')

    _, _, _, total_idx = _get_fila_por_contenido(xml, trs, 'Total / es')
    if total_idx is None:
        total_idx = primera_idx + 9

    fila_modelo, ini_mod, _ = get_fila_xml(xml, trs, primera_idx)
    fila_total, ini_tot, fin_tot = get_fila_xml(xml, trs, total_idx)

    nuevas_filas = ''
    for prod in datos.get('productos', []):
        nombre = armar_nombre_peru(prod, MAPA_PERU_ENFRIADO)
        nueva = fila_modelo
        nueva = _reemplazar_celda(nueva, 0, str(prod.get('cajas', '')))
        nueva = _reemplazar_celda(nueva, 1, nombre)
        nueva = _reemplazar_celda(nueva, 6, formatear_miles_en(prod.get('neto', '')))
        nueva = _reemplazar_celda(nueva, 7, formatear_miles_en(prod.get('bruto', '')))
        nuevas_filas += nueva

    pallets = datos.get('pallets', '') or ''
    kg_pallets = datos.get('kg_pallets', '') or ''
    nueva_pal = _reemplazar_pallets_en_fila(fila_pal, pallets, kg_pallets) if (fila_pal and pallets) else (fila_pal or '')

    nueva_total = fila_total
    nueva_total = _reemplazar_celda(nueva_total, 0, str(datos.get('total_cajas', '')))
    nueva_total = _reemplazar_celda(nueva_total, 2, formatear_miles_en(datos.get('total_neto', '')))
    nueva_total = _reemplazar_celda(nueva_total, 3, formatear_miles_en(datos.get('total_bruto', '')))

    xml = xml[:ini_mod] + nuevas_filas + nueva_pal + nueva_total + xml[fin_tot:]

    # Fechas de faena / produccion / vencimiento (rango unico por envio)
    trs2 = get_trs(xml)
    xml = _reemplazar_fechas(xml, trs2, datos.get('fecha_faena', ''), datos.get('fecha_produccion', ''),
                              datos.get('fecha_vencimiento', ''), fmt_fecha_al)

    # Transporte (terrestre) - patente principal + segunda chapa. La primera
    # se prioriza del REMITO (texto real, no pasa por OCR); la segunda solo
    # esta en el provisorio (OCR), asi que se avisa para que se verifique a
    # mano - el OCR puede confundir digitos con letras en la patente.
    patente1 = datos.get('camion') or datos.get('patente1') or ''
    patente2 = datos.get('patente2') or ''
    if patente1:
        camion_chapa = patente1 + (' / ' + patente2 if patente2 else '')
        xml = _reemplazar_tras_label(xml, r'CAMION CHAPA:', camion_chapa)
        if patente2:
            alertas.append('Segunda patente (' + patente2 + ') leída por OCR del provisorio - verificar que sea correcta')
    else:
        alertas.append('Patente de camión no encontrada - completar manualmente')

    # Precinto SENASA y AFIP combinados en un solo campo con guion
    precinto_senasa = datos.get('precinto_senasa', '') or ''
    precinto_afip = datos.get('precinto_afip', '') or ''
    if precinto_senasa or precinto_afip:
        precinto_combinado = precinto_senasa + (' – ' + precinto_afip if precinto_afip else '')
        xml = _reemplazar_tras_label(xml, r'Número de precinto:', precinto_combinado)
    else:
        alertas.append('Precinto no encontrado - completar manualmente')

    # Fecha de emision (pie del certificado) - la ultima fecha dd/mm/yyyy del documento
    fecha_emi = datos.get('fecha_emision') or datetime.datetime.now().strftime('%d/%m/%Y')
    todas_fechas = list(re.finditer(r'\d{2}/\d{2}/\d{4}', xml))
    if todas_fechas:
        ultima = todas_fechas[-1]
        xml = xml[:ultima.start()] + fecha_emi + xml[ultima.end():]

    return xml, alertas


# ── PERU MENUDENCIAS ─────────────────────────────────────────────────────
# Independiente de Peru Enfriado: es MARITIMO (buque, no camion), precinto en
# UN solo campo (no combinado), numeros en formato argentino (punto miles,
# coma decimal - no ingles como enfriado), y la fila de pallets no lleva
# parentesis alrededor de los KGS.

def _gen_peru_menudencias(xml, datos):
    alertas = []
    trs = get_trs(xml)

    _, _, _, header_idx = _get_fila_por_contenido(xml, trs, 'Descripción de la mercadería')
    primera_idx = (header_idx + 1) if header_idx is not None else 5

    fila_pal, ini_pal, fin_pal, idx_pal = _get_fila_por_contenido(xml, trs, 'ACONDICIONAD')

    _, _, _, total_idx = _get_fila_por_contenido(xml, trs, 'Total / es')
    if total_idx is None:
        total_idx = primera_idx + 5

    fila_modelo, ini_mod, _ = get_fila_xml(xml, trs, primera_idx)
    fila_total, ini_tot, fin_tot = get_fila_xml(xml, trs, total_idx)

    nuevas_filas = ''
    for prod in datos.get('productos', []):
        nombre = armar_nombre_peru(prod, MAPA_PERU_MENUDENCIAS)
        nueva = fila_modelo
        nueva = _reemplazar_celda(nueva, 0, str(prod.get('cajas', '')))
        nueva = _reemplazar_celda(nueva, 1, nombre)
        nueva = _reemplazar_celda(nueva, 6, formatear_miles(prod.get('neto', '')))
        nueva = _reemplazar_celda(nueva, 7, formatear_miles(prod.get('bruto', '')))
        nuevas_filas += nueva

    pallets = datos.get('pallets', '') or ''
    kg_pallets = datos.get('kg_pallets', '') or ''
    nueva_pal = fila_pal or ''
    if fila_pal and pallets:
        nueva_pal = re.sub(r'(ACONDICIONADA EN\s*)\d+', r'\g<1>' + str(pallets), nueva_pal, count=1)
        if kg_pallets:
            kg_fmt = formatear_miles(kg_pallets)
            nueva_pal = re.sub(r'[\d\.,]+((?:\s|<[^>]+>)*?KGS)', kg_fmt + r'\1', nueva_pal, count=1)

    nueva_total = fila_total
    nueva_total = _reemplazar_celda(nueva_total, 0, str(datos.get('total_cajas', '')))
    nueva_total = _reemplazar_celda(nueva_total, 2, formatear_miles(datos.get('total_neto', '')))
    nueva_total = _reemplazar_celda(nueva_total, 3, formatear_miles(datos.get('total_bruto', '')))

    xml = xml[:ini_mod] + nuevas_filas + nueva_pal + nueva_total + xml[fin_tot:]

    # Fechas de faena / produccion / vencimiento (rango unico por envio)
    trs2 = get_trs(xml)
    xml = _reemplazar_fechas(xml, trs2, datos.get('fecha_faena', ''), datos.get('fecha_produccion', ''),
                              datos.get('fecha_vencimiento', ''), fmt_fecha_al)

    # Transporte (buque)
    transporte = datos.get('transporte', '') or ''
    if transporte: xml = xml.replace('SANTA VANESSA', transporte)

    # Contenedor
    contenedor = datos.get('contenedor', '') or ''
    if contenedor: xml = xml.replace('HLBU6152556', contenedor)
    if not contenedor: alertas.append('Contenedor no encontrado - completar manualmente')

    # Precinto (un solo campo - se usa el de AFIP)
    precinto = datos.get('precinto_afip') or datos.get('precinto_senasa') or ''
    if precinto: xml = xml.replace('BAH79560', precinto)
    if not precinto: alertas.append('Precinto no encontrado - completar manualmente')

    # Fecha de emision (pie del certificado) - la ultima fecha dd/mm/yyyy del documento
    fecha_emi = datos.get('fecha_emision') or datetime.datetime.now().strftime('%d/%m/%Y')
    todas_fechas = list(re.finditer(r'\d{2}/\d{2}/\d{4}', xml))
    if todas_fechas:
        ultima = todas_fechas[-1]
        xml = xml[:ultima.start()] + fecha_emi + xml[ultima.end():]

    return xml, alertas


# ── PERU CONGELADO ────────────────────────────────────────────────────────
# Independiente de Enfriado y Menudencias: es MARITIMO (buque, no camion, a
# diferencia de Enfriado), precinto y contenedor en un solo campo cada uno,
# numeros en formato argentino. Cuando el envio mezcla carne congelada con
# menudencia, este es el documento de la parte de carne (ver
# CODIGOS_MENUDENCIA_PERU y la logica de separacion en el route /generar).

def _gen_peru_congelado(xml, datos):
    alertas = []
    trs = get_trs(xml)

    _, _, _, header_idx = _get_fila_por_contenido(xml, trs, 'Descripción de la mercadería')
    primera_idx = (header_idx + 1) if header_idx is not None else 5

    fila_pal, ini_pal, fin_pal, idx_pal = _get_fila_por_contenido(xml, trs, 'ACONDICIONAD')

    _, _, _, total_idx = _get_fila_por_contenido(xml, trs, 'Total / es')
    if total_idx is None:
        total_idx = primera_idx + 7

    fila_modelo, ini_mod, _ = get_fila_xml(xml, trs, primera_idx)
    fila_total, ini_tot, fin_tot = get_fila_xml(xml, trs, total_idx)

    nuevas_filas = ''
    for prod in datos.get('productos', []):
        nombre = armar_nombre_peru(prod, MAPA_PERU_CONGELADO)
        nueva = fila_modelo
        nueva = _reemplazar_celda(nueva, 0, str(prod.get('cajas', '')))
        nueva = _reemplazar_celda(nueva, 1, nombre)
        nueva = _reemplazar_celda(nueva, 6, formatear_miles(prod.get('neto', '')))
        nueva = _reemplazar_celda(nueva, 7, formatear_miles(prod.get('bruto', '')))
        nuevas_filas += nueva

    pallets = datos.get('pallets', '') or ''
    kg_pallets = datos.get('kg_pallets', '') or ''
    nueva_pal = fila_pal or ''
    if fila_pal and pallets:
        nueva_pal = re.sub(r'(ACONDICIONADA EN\s*)\d+', r'\g<1>' + str(pallets), nueva_pal, count=1)
        if kg_pallets:
            kg_fmt = formatear_miles(kg_pallets)
            nueva_pal = re.sub(r'[\d\.,]+((?:\s|<[^>]+>)*?KGS)', kg_fmt + r'\1', nueva_pal, count=1)

    nueva_total = fila_total
    nueva_total = _reemplazar_celda(nueva_total, 0, str(datos.get('total_cajas', '')))
    nueva_total = _reemplazar_celda(nueva_total, 2, formatear_miles(datos.get('total_neto', '')))
    nueva_total = _reemplazar_celda(nueva_total, 3, formatear_miles(datos.get('total_bruto', '')))

    xml = xml[:ini_mod] + nuevas_filas + nueva_pal + nueva_total + xml[fin_tot:]

    # Fechas de faena / produccion / vencimiento (rango unico por envio, o por
    # el subconjunto de codigos si viene de una separacion carne/menudencia)
    trs2 = get_trs(xml)
    xml = _reemplazar_fechas(xml, trs2, datos.get('fecha_faena', ''), datos.get('fecha_produccion', ''),
                              datos.get('fecha_vencimiento', ''), fmt_fecha_al)

    # Transporte (buque)
    transporte = datos.get('transporte', '') or ''
    if transporte: xml = xml.replace('LAKONIA', transporte)

    # Contenedor
    contenedor = datos.get('contenedor', '') or ''
    if contenedor: xml = xml.replace('SEKU9270994', contenedor)
    if not contenedor: alertas.append('Contenedor no encontrado - completar manualmente')

    # Precinto (un solo campo - se usa el de AFIP)
    precinto = datos.get('precinto_afip') or datos.get('precinto_senasa') or ''
    if precinto: xml = xml.replace('BAH74816', precinto)
    if not precinto: alertas.append('Precinto no encontrado - completar manualmente')

    # Fecha de emision (pie del certificado) - la ultima fecha dd/mm/yyyy del documento
    fecha_emi = datos.get('fecha_emision') or datetime.datetime.now().strftime('%d/%m/%Y')
    todas_fechas = list(re.finditer(r'\d{2}/\d{2}/\d{4}', xml))
    if todas_fechas:
        ultima = todas_fechas[-1]
        xml = xml[:ultima.start()] + fecha_emi + xml[ultima.end():]

    return xml, alertas


# ── DUBAI (EMIRATOS ARABES) ──────────────────────────────────────────────
# Nombre bilingue de una sola linea (por codigo, tabla propia). Numeros en
# formato simple (sin separador de miles). El peso de los pallets va en su
# PROPIA celda (columna bruto) en vez de dentro del texto, y SI se suma al
# bruto total (a diferencia de Egipto/Brasil/Ecuador). Aereo, sin camion.
# Fecha de emision en formato especial "YYYY (año-year) MM (mes-month) DD
# (dia-Day)" en vez de dd/mm/yyyy.

def _gen_dubai(xml, datos):
    alertas = []
    trs = get_trs(xml)

    _, _, _, header_idx = _get_fila_por_contenido(xml, trs, 'Description of goods')
    primera_idx = (header_idx + 1) if header_idx is not None else 5

    fila_pal, ini_pal, fin_pal, idx_pal = _get_fila_por_contenido(xml, trs, 'ACONDICIONAD')

    _, _, _, total_idx = _get_fila_por_contenido(xml, trs, 'Total / es')
    if total_idx is None:
        total_idx = primera_idx + 8

    fila_modelo, ini_mod, _ = get_fila_xml(xml, trs, primera_idx)
    fila_total, ini_tot, fin_tot = get_fila_xml(xml, trs, total_idx)

    nuevas_filas = ''
    for prod in datos.get('productos', []):
        nombre_bi = armar_nombre_dubai(prod)
        nueva = fila_modelo
        nueva = _reemplazar_celda(nueva, 0, str(prod.get('cajas', '')))
        nueva = _reemplazar_celda(nueva, 1, nombre_bi)
        nueva = _reemplazar_celda(nueva, 6, str(prod.get('neto', '')))
        nueva = _reemplazar_celda(nueva, 7, str(prod.get('bruto', '')))
        nuevas_filas += nueva

    pallets = datos.get('pallets', '') or ''
    kg_pallets = datos.get('kg_pallets', '') or ''
    nueva_pal = fila_pal or ''
    if fila_pal and pallets:
        nueva_pal = re.sub(r'(ACONDICIONADA EN\s*)\d+', r'\g<1>' + str(pallets), nueva_pal, count=1)
        nueva_pal = re.sub(r'(ACONDITIONED IN\s*)\d+', r'\g<1>' + str(pallets), nueva_pal, count=1, flags=re.IGNORECASE)
        if kg_pallets:
            nueva_pal = _reemplazar_celda(nueva_pal, 7, kg_pallets)

    # El bruto total SI suma el peso de los pallets (a diferencia de otros
    # destinos donde el remito ya trae el bruto completo)
    total_bruto = datos.get('total_bruto', '')
    if kg_pallets:
        try:
            total_bruto = '{:.2f}'.format(float(total_bruto) + float(kg_pallets))
        except (TypeError, ValueError):
            pass

    nueva_total = fila_total
    nueva_total = _reemplazar_celda(nueva_total, 0, str(datos.get('total_cajas', '')))
    nueva_total = _reemplazar_celda(nueva_total, 2, str(datos.get('total_neto', '')))
    nueva_total = _reemplazar_celda(nueva_total, 3, total_bruto)

    xml = xml[:ini_mod] + nuevas_filas + nueva_pal + nueva_total + xml[fin_tot:]

    # Fechas de faena / produccion / vencimiento (rango unico por envio)
    trs2 = get_trs(xml)
    xml = _reemplazar_fechas(xml, trs2, datos.get('fecha_faena', ''), datos.get('fecha_produccion', ''),
                              datos.get('fecha_vencimiento', ''), fmt_fecha_al)

    # Transporte (vuelo)
    transporte = datos.get('transporte', '') or ''
    if transporte: xml = xml.replace('EK9921', transporte)

    # Contenedor y precinto - anclados por label (por defecto vienen en guiones, aereo)
    contenedor = datos.get('contenedor', '') or ''
    if contenedor:
        xml = _reemplazar_tras_label(xml, r'Container\(s\) number:', contenedor)
    precinto = datos.get('precinto_afip') or datos.get('precinto_senasa') or ''
    if precinto:
        xml = _reemplazar_tras_label(xml, r'Seal number:', precinto)

    # Fecha de emision - formato especial "YYYY (año-year) MM (mes-month) DD (dia-Day)"
    fecha_emi = datos.get('fecha_emision') or datetime.datetime.now().strftime('%d/%m/%Y')
    m_femi = re.match(r'(\d{2})/(\d{2})/(\d{4})', fecha_emi)
    if m_femi:
        dd, mm, yyyy = m_femi.groups()
        nuevo_texto = yyyy + ' (año – year) ' + mm + ' (mes – month) ' + dd + ' (día – Day)'
        xml = re.sub(
            r'\d{4}\s*\([^)]*\)\s*\d{2}\s*\([^)]*\)\s*\d{2}\s*\([^)]*\)',
            nuevo_texto, xml, count=1
        )

    return xml, alertas


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
