# -*- coding: utf-8 -*-
"""
Sistema BEX de Control de Credenciales - v2
Fusiona ARCHIVO 1 (datos reales) + ARCHIVO 2 (diseno) en un unico .xlsx funcional.
Genera: tablas estructuradas, motor de dashboard (8 resumenes), graficos conectados,
buscador, filtros interactivos, validaciones, formato condicional y cuadros de ayuda.
Escrito SOLO con libreria estandar de Python.
"""
import zipfile, re, datetime
from xml.etree import ElementTree as ET

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'

# ---------------------------------------------------------------------------
# LECTOR
# ---------------------------------------------------------------------------
def read_all(path):
    z = zipfile.ZipFile(path); shared = []
    if 'xl/sharedStrings.xml' in z.namelist():
        t = ET.fromstring(z.read('xl/sharedStrings.xml'))
        for si in t.findall(NS + 'si'):
            shared.append(''.join(n.text or '' for n in si.iter(NS + 't')))
    wb = ET.fromstring(z.read('xl/workbook.xml'))
    sheets = [s.get('name') for s in wb.iter(NS + 'sheet')]
    def colidx(ref):
        m = re.match(r'([A-Z]+)', ref); n = 0
        for c in m.group(1): n = n * 26 + (ord(c) - 64)
        return n - 1
    out = {}
    sfs = sorted([n for n in z.namelist() if re.match(r'xl/worksheets/sheet\d+\.xml', n)],
                 key=lambda x: int(re.search(r'(\d+)', x).group(1)))
    for idx, sf in enumerate(sfs):
        t = ET.fromstring(z.read(sf)); rows = []
        for r in t.iter(NS + 'row'):
            cells = {}; maxc = -1
            for c in r.findall(NS + 'c'):
                ci = colidx(c.get('r')); v = c.find(NS + 'v'); val = ''
                if v is not None:
                    val = shared[int(v.text)] if c.get('t') == 's' else v.text
                cells[ci] = val; maxc = max(maxc, ci)
            rows.append([cells.get(i, '') for i in range(maxc + 1)])
        out[sheets[idx]] = rows
    return out

A1 = read_all('ARCHIVO 1.xlsx')
A2 = read_all('ARCHIVO 2.xlsx')

def clean(s): return (s or '').strip()
def to_int(s):
    s = clean(s)
    return int(s) if re.fullmatch(r'\d+', s) else None

SUP_MAP = {
    'JOSE GUTIERREZ': 'JOSE GUTIERREZ PEDRAZA',
    'MARIA FERNANDA DAZA': 'MARIA FERNANDA DAZA PALMA',
    'DELIA JORDAN': 'DELIA JORDAN FACUSSE',
}
def norm_sup(s):
    s = clean(s); return SUP_MAP.get(s, s)

# ---------------------------------------------------------------------------
# MODELO
# ---------------------------------------------------------------------------
personal = []
for row in A2['PERSONAL'][3:]:
    if not clean(row[0] if row else ''): continue
    r = (row + [''] * 10)[:10]
    personal.append({'id': clean(r[0]), 'tel': clean(r[1]) or 'POR REGISTRAR',
                     'nombre': clean(r[2]), 'rol': clean(r[3]), 'sup': norm_sup(r[4]),
                     'reg': clean(r[5]), 'estado': clean(r[6]), 'cred': clean(r[7]),
                     'obs': clean(r[9])})
name2tel = {p['nombre']: p['tel'] for p in personal}

movimientos = []
for row in A2['MOVIMIENTOS'][3:]:
    if not clean(row[0] if row else ''): continue
    r = (row + [''] * 9)[:9]
    nombre = clean(r[4])
    movimientos.append({'n': clean(r[0]), 'fecha': to_int(r[1]), 'tipo': clean(r[2]),
                        'cod': clean(r[3]), 'nombre': nombre, 'tel': name2tel.get(nombre, 'POR REGISTRAR'),
                        'sup': norm_sup(r[5]), 'reg': clean(r[6]), 'obs': clean(r[7])})

inventario = []; seen = set()
for row in A2['ACTIVOS'][3:]:
    cod = clean(row[0] if row else '')
    if not cod: continue
    r = (row + [''] * 9)[:9]
    inventario.append({'cod': cod, 'gen': clean(r[1]), 'estado': clean(r[2]), 'resp': clean(r[3]),
                       'sup': norm_sup(r[4]), 'reg': clean(r[5]), 'fecha': to_int(r[6]), 'obs': clean(r[8])})
    seen.add(cod)
for n in range(2051, 2101):
    cod = 'BEX-%04d' % n
    if cod in seen: continue
    inventario.append({'cod': cod, 'gen': 'NUEVA (BEX-2xxx)', 'estado': 'DISPONIBLE',
                       'resp': 'EN OFICINA', 'sup': '', 'reg': 'SANTA CRUZ', 'fecha': 46176,
                       'obs': 'Stock en oficina - recepcionado'})
    seen.add(cod)

regionales = ['COCHABAMBA', 'LA PAZ', 'ORURO', 'SANTA CRUZ', 'SUCRE', 'TARIJA']
estados_cred = ['ASIGNADO', 'DISPONIBLE', 'POR ASIGNAR', 'PERDIDO']
supervisores = ['HASIRA DANIELA OSINAGA CHOQUE', 'PAMELA FANNY CALANI LAURA', 'NATALIA VILLARROEL',
                'CLAUDIA SHASKIA CALLE NINA', 'GERCY EVER ERGUETA KIPPES', 'JOEL DAVID HUANCA HUANCA',
                'MAGALI YESENIA HUANCA HUANCA', 'MILENKO ADRIANA ORDONEZ NUNEZ', 'LIONEL CHABUR',
                'MIRIAM ROSA CHAMBI QUISPE', 'JORGE SAAVEDRA', 'BEATRIZ OVIEDO OVIEDO',
                'DELIA JORDAN FACUSSE', 'JOSE GUTIERREZ PEDRAZA', 'JENNY CRISTINA ECHALAR MONTALVO',
                'MARIA FERNANDA DAZA PALMA']
# orden top por asignadas (descendente) para grafico "top supervisores"
sup_asig = {s: sum(1 for it in inventario if it['sup'] == s and it['estado'] == 'ASIGNADO') for s in supervisores}
supervisores_top = sorted(supervisores, key=lambda s: -sup_asig[s])

tipos_mov = ['ASIGNACION', 'ENTREGA INICIAL', 'DEVOLUCION', 'PERDIDA', 'REPOSICION',
             'REASIGNACION', 'CAMBIO SUPERVISOR', 'CAMBIO REGIONAL', 'BAJA']
estados_persona = ['ACTIVO', 'INACTIVO']
roles = ['AFILIADOR', 'SUPERVISOR']
generaciones = ['NUEVA (BEX-2xxx)', 'LEGACY (BEX-0xxx)']
errores = [
    '1. BEX-0176 asignado a DOS personas: CLAUDIA SHASKIA y REYNA IVONNE',
    '2. BEX-0106 posiblemente duplicado: CAMILA LOZA (GERCY) y JHOSEP SHARUM (MAGALI)',
    '3. Personas con nombres incompletos: TOMAS, LORENA, SELMA',
    '4. BEX-0518 reportada como PERDIDA - pendiente cobro a MARCO ESCOBAR',
    '5. Supervisores sin datos en hoja principal: LIONEL CHABUR, NATALIA VILLARROEL',
    '6. Codigos normalizados (espacios y ceros): BEX-162, BEX-0176, BEX-027, etc.',
    '7. Telefonos marcados como POR REGISTRAR - campo critico faltante',
]

# Buckets mensuales (12) a partir del primer mes con datos
EPOCH = datetime.date(1899, 12, 30)
def serial_to_date(s): return EPOCH + datetime.timedelta(days=s)
def date_to_serial(d): return (d - EPOCH).days
fechas = [m['fecha'] for m in movimientos if m['fecha']]
mind = serial_to_date(min(fechas)) if fechas else datetime.date(2026, 1, 1)
y, mo = mind.year, mind.month
meses = []
for k in range(12):
    yy = y + (mo - 1 + k) // 12
    mm = (mo - 1 + k) % 12 + 1
    start = datetime.date(yy, mm, 1)
    nb = datetime.date(yy + (mm // 12), (mm % 12) + 1, 1)
    end = nb - datetime.timedelta(days=1)
    meses.append((('%04d-%02d' % (yy, mm)), date_to_serial(start), date_to_serial(end)))

print('Personal:', len(personal), '| Movimientos:', len(movimientos), '| Inventario:', len(inventario))
print('Meses:', [m[0] for m in meses])
print('OK modelo v2')



# ===========================================================================
# MOTOR DE ESCRITURA XLSX
# ===========================================================================
def col_letter(c):
    s = ''
    while c > 0:
        c, r = divmod(c - 1, 26); s = chr(65 + r) + s
    return s

def ci(letter):
    c = 0
    for ch in letter: c = c * 26 + (ord(ch) - 64)
    return c

def xesc(s):
    return (str(s).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;'))

def fix_formula(f):
    f = re.sub(r'(?<![A-Za-z0-9_.])XLOOKUP\(', '_xlfn.XLOOKUP(', f)
    f = re.sub(r'(?<![A-Za-z0-9_.])MAXIFS\(', '_xlfn.MAXIFS(', f)
    f = re.sub(r'(?<![A-Za-z0-9_.])MINIFS\(', '_xlfn.MINIFS(', f)
    f = re.sub(r'(?<![A-Za-z0-9_.])LET\(', '_xlfn.LET(', f)
    f = re.sub(r'(?<![A-Za-z0-9_.])SORTBY\(', '_xlfn.SORTBY(', f)
    f = re.sub(r'(?<![A-Za-z0-9_.])SORT\(', '_xlfn._xlws.SORT(', f)
    f = re.sub(r'(?<![A-Za-z0-9_.])UNIQUE\(', '_xlfn.UNIQUE(', f)
    f = re.sub(r'(?<![A-Za-z0-9_.])FILTER\(', '_xlfn._xlws.FILTER(', f)
    f = re.sub(r'(?<![A-Za-z0-9_.])TEXTJOIN\(', '_xlfn.TEXTJOIN(', f)
    f = re.sub(r'(?<![A-Za-z0-9_.])TAKE\(', '_xlfn.TAKE(', f)
    return f

class Sheet:
    def __init__(self, name, hidden=False, tabcolor=None):
        self.name = name; self.hidden = hidden; self.tabcolor = tabcolor
        self.cells = {}; self.cols = []; self.merges = []; self.table = None
        self.freeze = None; self.rowheights = {}; self.drawing_rid = None
        self.cond = []        # (sqref, [(op, formula, dxfId)])
        self.dvs = []         # (sqref, formula1)
        self.hyperlinks = []  # (ref, location, display)

    def set(self, r, c, value, kind='s', style=0):
        self.cells[(r, c)] = (value, kind, style)
    def setc(self, ref, value, kind='s', style=0):
        m = re.match(r'([A-Z]+)(\d+)', ref); self.set(int(m.group(2)), ci(m.group(1)), value, kind, style)

    def xml(self, rid_table=None):
        if self.cells:
            maxr = max(r for r, c in self.cells); maxc = max(c for r, c in self.cells)
        else:
            maxr = maxc = 1
        o = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>']
        o.append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')
        if self.tabcolor:
            o.append('<sheetPr><tabColor rgb="FF%s"/></sheetPr>' % self.tabcolor)
        o.append('<dimension ref="A1:%s%d"/>' % (col_letter(maxc), maxr))
        if self.freeze:
            fr, fc = self.freeze
            top = '%s%d' % (col_letter(fc + 1), fr + 1)
            o.append('<sheetViews><sheetView showGridLines="0" workbookViewId="0">'
                     '<pane xSplit="%d" ySplit="%d" topLeftCell="%s" activePane="bottomRight" state="frozen"/>'
                     '<selection pane="bottomRight"/></sheetView></sheetViews>' % (fc, fr, top))
        else:
            o.append('<sheetViews><sheetView showGridLines="0" workbookViewId="0"/></sheetViews>')
        o.append('<sheetFormatPr defaultRowHeight="15"/>')
        if self.cols:
            o.append('<cols>')
            for mn, mx, w in self.cols:
                o.append('<col min="%d" max="%d" width="%.2f" customWidth="1"/>' % (mn, mx, w))
            o.append('</cols>')
        o.append('<sheetData>')
        for r in range(1, maxr + 1):
            rc = [(c, self.cells[(r, c)]) for c in range(1, maxc + 1) if (r, c) in self.cells]
            if not rc: continue
            hattr = ' ht="%.2f" customHeight="1"' % self.rowheights[r] if r in self.rowheights else ''
            o.append('<row r="%d"%s>' % (r, hattr))
            for c, (val, kind, style) in rc:
                ref = '%s%d' % (col_letter(c), r); sattr = ' s="%d"' % style if style else ''
                if kind == 'f':
                    o.append('<c r="%s"%s><f>%s</f></c>' % (ref, sattr, xesc(fix_formula(val))))
                elif kind == 'fa':
                    o.append('<c r="%s"%s cm="1"><f>%s</f></c>' % (ref, sattr, xesc(fix_formula(val))))
                elif kind == 'n':
                    o.append('<c r="%s"%s><v>%s</v></c>' % (ref, sattr, val))
                else:
                    o.append('<c r="%s"%s t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>' % (ref, sattr, xesc(val)))
            o.append('</row>')
        o.append('</sheetData>')
        if self.merges:
            o.append('<mergeCells count="%d">' % len(self.merges))
            for m in self.merges: o.append('<mergeCell ref="%s"/>' % m)
            o.append('</mergeCells>')
        pr = 1
        for sqref, rules in self.cond:
            o.append('<conditionalFormatting sqref="%s">' % sqref)
            for op, formula, dxf in rules:
                o.append('<cfRule type="cellIs" dxfId="%d" priority="%d" operator="%s"><formula>%s</formula></cfRule>'
                         % (dxf, pr, op, xesc(formula))); pr += 1
            o.append('</conditionalFormatting>')
        if self.dvs:
            o.append('<dataValidations count="%d">' % len(self.dvs))
            for sqref, f1 in self.dvs:
                o.append('<dataValidation type="list" allowBlank="1" showInputMessage="1" '
                         'showErrorMessage="1" sqref="%s"><formula1>%s</formula1></dataValidation>'
                         % (sqref, xesc(f1)))
            o.append('</dataValidations>')
        if self.hyperlinks:
            o.append('<hyperlinks>')
            for ref, loc, disp in self.hyperlinks:
                o.append('<hyperlink ref="%s" location="%s" display="%s"/>' % (ref, xesc(loc), xesc(disp)))
            o.append('</hyperlinks>')
        o.append('<pageMargins left="0.5" right="0.5" top="0.6" bottom="0.6" header="0.3" footer="0.3"/>')
        o.append('<pageSetup orientation="landscape" fitToWidth="1" fitToHeight="0"/>')
        if self.drawing_rid:
            o.append('<drawing r:id="%s"/>' % self.drawing_rid)
        if rid_table:
            o.append('<tableParts count="1"><tablePart r:id="%s"/></tableParts>' % rid_table)
        o.append('</worksheet>')
        return ''.join(o)

def table_xml(tid, name, ref, columns):
    o = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>']
    o.append('<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
             'id="%d" name="%s" displayName="%s" ref="%s" totalsRowShown="0">' % (tid, name, name, ref))
    o.append('<autoFilter ref="%s"/>' % ref)
    o.append('<tableColumns count="%d">' % len(columns))
    for i, cn in enumerate(columns, 1):
        o.append('<tableColumn id="%d" name="%s"/>' % (i, xesc(cn)))
    o.append('</tableColumns>')
    o.append('<tableStyleInfo name="TableStyleMedium9" showFirstColumn="0" showLastColumn="0" '
             'showRowStripes="1" showColumnStripes="0"/></table>')
    return ''.join(o)



STYLES_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="2"><numFmt numFmtId="164" formatCode="dd/mm/yyyy"/><numFmt numFmtId="165" formatCode="#,##0"/></numFmts>
<fonts count="11">
<font><sz val="11"/><name val="Calibri"/></font>
<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
<font><b/><sz val="18"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
<font><b/><sz val="26"/><color rgb="FF1F4E78"/><name val="Calibri"/></font>
<font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
<font><b/><sz val="11"/><color rgb="FF1F4E78"/><name val="Calibri"/></font>
<font><i/><sz val="9"/><color rgb="FF595959"/><name val="Calibri"/></font>
<font><b/><sz val="13"/><color rgb="FF1F4E78"/><name val="Calibri"/></font>
<font><sz val="11"/><color rgb="FF1F4E78"/><name val="Calibri"/></font>
<font><b/><sz val="12"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
</fonts>
<fills count="9">
<fill><patternFill patternType="none"/></fill>
<fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFD9E1F2"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF2E86C1"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFBDD7EE"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFF2F2F2"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF305496"/><bgColor indexed="64"/></patternFill></fill>
</fills>
<borders count="2">
<border><left/><right/><top/><bottom/><diagonal/></border>
<border><left style="thin"><color rgb="FF8EAADB"/></left><right style="thin"><color rgb="FF8EAADB"/></right><top style="thin"><color rgb="FF8EAADB"/></top><bottom style="thin"><color rgb="FF8EAADB"/></bottom><diagonal/></border>
</borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="18">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="2" fillId="8" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
<xf numFmtId="0" fontId="3" fillId="4" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
<xf numFmtId="0" fontId="4" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
<xf numFmtId="0" fontId="5" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
<xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf>
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>
<xf numFmtId="10" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf>
<xf numFmtId="0" fontId="6" fillId="5" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf>
<xf numFmtId="0" fontId="4" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
<xf numFmtId="0" fontId="10" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
<xf numFmtId="0" fontId="7" fillId="7" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="top" wrapText="1"/></xf>
<xf numFmtId="0" fontId="6" fillId="6" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
<xf numFmtId="0" fontId="6" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
<xf numFmtId="0" fontId="8" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
<xf numFmtId="0" fontId="9" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1"/>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
<dxfs count="4">
<dxf><font><color rgb="FF006100"/></font><fill><patternFill><bgColor rgb="FFC6EFCE"/></patternFill></fill></dxf>
<dxf><font><color rgb="FF1F4E78"/></font><fill><patternFill><bgColor rgb="FFBDD7EE"/></patternFill></fill></dxf>
<dxf><font><color rgb="FF9C0006"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf>
<dxf><font><color rgb="FF9C5700"/></font><fill><patternFill><bgColor rgb="FFFFEB9C"/></patternFill></fill></dxf>
</dxfs>
</styleSheet>'''

# indices de estilo
(S_DEF, S_HEAD, S_TITLE, S_SECTION, S_KPINUM, S_KPILBL, S_DATE, S_BORD, S_PCT,
 S_SUB, S_CEN, S_KPIBOX, S_TITLE2, S_HELP, S_INPUT, S_LBLR, S_H2, S_DARKTXT) = range(18)
# dxf ids
DXF = {'ASIGNADO': 0, 'DISPONIBLE': 1, 'PERDIDO': 2, 'POR ASIGNAR': 3}



# ===========================================================================
# CONSTRUCCION DE HOJAS
# ===========================================================================
TIG, TPE, TMO, TPA = 'TablaInventarioGeneral', 'TablaPersonal', 'TablaMovimientos', 'TablaParametros'
ACCENT = 'FF1F4E78'

def banner(sh, title, ncols, help_title, help_text):
    """Titulo + cuadro de ayuda. Devuelve fila de cabecera (4)."""
    last = col_letter(ncols)
    sh.set(1, 1, title, 's', S_TITLE); sh.merges.append('A1:%s1' % last); sh.rowheights[1] = 30
    sh.set(2, 1, '?  ' + help_title, 's', S_TITLE2); sh.merges.append('A2:%s2' % last)
    sh.set(3, 1, help_text, 's', S_HELP); sh.merges.append('A3:%s3' % last); sh.rowheights[3] = 46
    return 4

def build_parametros():
    sh = Sheet('PARAMETROS', tabcolor='808080')
    cols = ['Regionales', 'Estados_Persona', 'Roles', 'Estados_Credencial', 'Tipos_Movimiento', 'Generacion', 'Supervisores']
    widths = [16, 16, 14, 18, 20, 20, 34]
    sh.cols = [(i + 1, i + 1, w) for i, w in enumerate(widths)] + [(9, 9, 70)]
    hr = banner(sh, 'PARAMETROS DEL SISTEMA  -  BEX', 7,
                'COMO USAR ESTA HOJA',
                'Listas maestras que alimentan los menus desplegables y validaciones del sistema. '
                'Edite aqui para agregar regionales, supervisores o estados. Las demas hojas se actualizan solas.')
    data = [regionales, estados_persona, roles, estados_cred, tipos_mov, generaciones, supervisores]
    for i, cn in enumerate(cols, 1):
        sh.set(hr, i, cn, 's', S_HEAD)
    maxlen = max(len(c) for c in data)
    for cidx, vals in enumerate(data, 1):
        for ri in range(maxlen):
            v = vals[ri] if ri < len(vals) else ''
            sh.set(hr + 1 + ri, cidx, v, 's', S_BORD)
    last = hr + maxlen
    sh.table = dict(tid=4, name=TPA, ref='A%d:G%d' % (hr, last), columns=cols)
    sh.set(hr, 9, 'ERRORES / ALERTAS DETECTADAS', 's', S_HEAD)
    for i, e in enumerate(errores, 1):
        sh.set(hr + i, 9, e, 's', S_BORD)
    sh.freeze = (hr, 0)
    return sh, hr, last

def build_inventario(par_rows):
    sh = Sheet('INVENTARIO_GENERAL', tabcolor='1F4E78')
    cols = ['Codigo_Activo', 'Tipo_Credencial', 'Estado_Actual', 'Responsable_Actual', 'Telefono',
            'Supervisor', 'Regional', 'Fecha_Ultima_Asignacion', 'Fecha_Ultima_Devolucion',
            'Cantidad_Perdidas', 'Dias_En_Uso', 'Observaciones']
    widths = [14, 18, 14, 34, 14, 30, 13, 16, 16, 12, 10, 40]
    sh.cols = [(i + 1, i + 1, w) for i, w in enumerate(widths)]
    hr = banner(sh, 'INVENTARIO GENERAL  -  Fuente central del sistema', 12,
                'COMO USAR ESTA HOJA',
                'Cada credencial BEX aparece UNA sola vez. Las columnas en azul (Telefono, fechas, perdidas, dias) '
                'se calculan solas desde MOVIMIENTOS. Cambie Estado/Responsable y todo el sistema se actualiza.')
    for i, cn in enumerate(cols, 1):
        sh.set(hr, i, cn, 's', S_HEAD)
    rr = hr + 1
    for it in inventario:
        sh.set(rr, 1, it['cod'], 's', S_BORD)
        sh.set(rr, 2, it['gen'], 's', S_BORD)
        sh.set(rr, 3, it['estado'], 's', S_CEN)
        sh.set(rr, 4, it['resp'], 's', S_BORD)
        sh.set(rr, 5, 'IFERROR(XLOOKUP([@Responsable_Actual],%s[Nombre_Completo],%s[Telefono]),"POR REGISTRAR")' % (TPE, TPE), 'f', S_CEN)
        sh.set(rr, 6, it['sup'], 's', S_BORD)
        sh.set(rr, 7, it['reg'], 's', S_BORD)
        sh.set(rr, 8, 'IF(COUNTIFS(%s[Codigo_Activo],[@Codigo_Activo])=0,"",MAXIFS(%s[Fecha],%s[Codigo_Activo],[@Codigo_Activo],%s[Tipo_Movimiento],"<>DEVOLUCION"))' % (TMO, TMO, TMO, TMO), 'f', S_DATE)
        sh.set(rr, 9, 'IF(COUNTIFS(%s[Codigo_Activo],[@Codigo_Activo],%s[Tipo_Movimiento],"DEVOLUCION")=0,"",MAXIFS(%s[Fecha],%s[Codigo_Activo],[@Codigo_Activo],%s[Tipo_Movimiento],"DEVOLUCION"))' % (TMO, TMO, TMO, TMO, TMO), 'f', S_DATE)
        sh.set(rr, 10, 'COUNTIFS(%s[Codigo_Activo],[@Codigo_Activo],%s[Tipo_Movimiento],"PERDIDA")' % (TMO, TMO), 'f', S_CEN)
        sh.set(rr, 11, 'IF(NOT(ISNUMBER([@Fecha_Ultima_Asignacion])),"",TODAY()-[@Fecha_Ultima_Asignacion])', 'f', S_CEN)
        sh.set(rr, 12, it['obs'], 's', S_BORD)
        rr += 1
    last = rr - 1
    sh.table = dict(tid=1, name=TIG, ref='A%d:L%d' % (hr, last), columns=cols)
    sh.freeze = (hr, 0)
    # formato condicional por estado (col C)
    sq = 'C%d:C%d' % (hr + 1, last)
    sh.cond.append((sq, [('equal', '"ASIGNADO"', DXF['ASIGNADO']),
                         ('equal', '"DISPONIBLE"', DXF['DISPONIBLE']),
                         ('equal', '"PERDIDO"', DXF['PERDIDO']),
                         ('equal', '"POR ASIGNAR"', DXF['POR ASIGNAR'])]))
    # validaciones
    sh.dvs.append(('C%d:C%d' % (hr + 1, last + 200), 'val_EstadoCred'))
    sh.dvs.append(('B%d:B%d' % (hr + 1, last + 200), 'val_Generacion'))
    sh.dvs.append(('F%d:F%d' % (hr + 1, last + 200), 'val_Supervisor'))
    sh.dvs.append(('G%d:G%d' % (hr + 1, last + 200), 'val_Regional'))
    return sh

def build_personal():
    sh = Sheet('PERSONAL', tabcolor='2E86C1')
    cols = ['Telefono', 'Nombre_Completo', 'Rol', 'Supervisor', 'Regional', 'Estado', 'Credencial_Actual', 'Observaciones']
    widths = [14, 38, 13, 30, 13, 11, 16, 28]
    sh.cols = [(i + 1, i + 1, w) for i, w in enumerate(widths)]
    hr = banner(sh, 'TABLA MAESTRA DE PERSONAL', 8,
                'COMO USAR ESTA HOJA',
                'Registro unico de personas. El identificador principal es el Telefono. '
                'El sistema relaciona a cada persona con su credencial y movimientos por el nombre/telefono.')
    for i, cn in enumerate(cols, 1):
        sh.set(hr, i, cn, 's', S_HEAD)
    rr = hr + 1
    for p in personal:
        sh.set(rr, 1, p['tel'], 's', S_CEN)
        sh.set(rr, 2, p['nombre'], 's', S_BORD)
        sh.set(rr, 3, p['rol'], 's', S_CEN)
        sh.set(rr, 4, p['sup'], 's', S_BORD)
        sh.set(rr, 5, p['reg'], 's', S_BORD)
        sh.set(rr, 6, p['estado'], 's', S_CEN)
        sh.set(rr, 7, p['cred'], 's', S_CEN)
        sh.set(rr, 8, p['obs'], 's', S_BORD)
        rr += 1
    last = rr - 1
    sh.table = dict(tid=2, name=TPE, ref='A%d:H%d' % (hr, last), columns=cols)
    sh.freeze = (hr, 0)
    sh.dvs.append(('C%d:C%d' % (hr + 1, last + 200), 'val_Rol'))
    sh.dvs.append(('D%d:D%d' % (hr + 1, last + 200), 'val_Supervisor'))
    sh.dvs.append(('E%d:E%d' % (hr + 1, last + 200), 'val_Regional'))
    sh.dvs.append(('F%d:F%d' % (hr + 1, last + 200), 'val_EstadoPersona'))
    return sh

def build_movimientos():
    sh = Sheet('MOVIMIENTOS', tabcolor='C0392B')
    cols = ['ID_Movimiento', 'Fecha', 'Tipo_Movimiento', 'Codigo_Activo', 'Telefono', 'Nombre',
            'Supervisor', 'Regional', 'Observaciones']
    widths = [13, 13, 18, 13, 14, 34, 30, 13, 38]
    sh.cols = [(i + 1, i + 1, w) for i, w in enumerate(widths)]
    hr = banner(sh, 'HISTORIAL DE MOVIMIENTOS  -  Registro permanente', 9,
                'COMO USAR ESTA HOJA',
                'Agregue una fila por cada evento (asignacion, devolucion, perdida, etc.). NUNCA borre filas. '
                'Al agregar un movimiento, el inventario, el buscador y el dashboard se actualizan automaticamente.')
    for i, cn in enumerate(cols, 1):
        sh.set(hr, i, cn, 's', S_HEAD)
    rr = hr + 1
    for m in movimientos:
        sh.set(rr, 1, m['n'], 's', S_BORD)
        if m['fecha'] is not None: sh.set(rr, 2, m['fecha'], 'n', S_DATE)
        else: sh.set(rr, 2, '', 's', S_DATE)
        sh.set(rr, 3, m['tipo'], 's', S_CEN)
        sh.set(rr, 4, m['cod'], 's', S_CEN)
        sh.set(rr, 5, m['tel'], 's', S_CEN)
        sh.set(rr, 6, m['nombre'], 's', S_BORD)
        sh.set(rr, 7, m['sup'], 's', S_BORD)
        sh.set(rr, 8, m['reg'], 's', S_BORD)
        sh.set(rr, 9, m['obs'], 's', S_BORD)
        rr += 1
    last = rr - 1
    sh.table = dict(tid=3, name=TMO, ref='A%d:I%d' % (hr, last), columns=cols)
    sh.freeze = (hr, 0)
    sh.dvs.append(('C%d:C%d' % (hr + 1, last + 300), 'val_TipoMov'))
    sh.dvs.append(('G%d:G%d' % (hr + 1, last + 300), 'val_Supervisor'))
    sh.dvs.append(('H%d:H%d' % (hr + 1, last + 300), 'val_Regional'))
    return sh



def build_motor():
    sh = Sheet('MOTOR_DASHBOARD', hidden=True, tabcolor='7F7F7F')
    sh.cols = [(1, 1, 28), (2, 2, 12), (3, 3, 4), (4, 4, 16), (5, 5, 12), (6, 6, 4),
               (7, 7, 32), (8, 8, 12), (9, 9, 4), (10, 10, 10), (11, 11, 12), (12, 12, 12),
               (13, 13, 12), (14, 14, 4), (15, 15, 12), (16, 16, 12)]
    sh.set(1, 1, 'MOTOR DE CALCULO (hoja oculta) - alimenta KPIs, tablas resumen y graficos', 's', S_H2)
    # --- TD1: Resumen general por estado (A3:B7) ---
    sh.set(3, 1, 'Estado', 's', S_HEAD); sh.set(3, 2, 'Cantidad', 's', S_HEAD)
    for i, e in enumerate(estados_cred):
        r = 4 + i
        sh.set(r, 1, e, 's', S_BORD)
        sh.set(r, 2, 'COUNTIF(%s[Estado_Actual],A%d)' % (TIG, r), 'f', S_CEN)
    # --- TD2: Por regional (D3:E9) ---
    sh.set(3, 4, 'Regional', 's', S_HEAD); sh.set(3, 5, 'Cantidad', 's', S_HEAD)
    for i, rg in enumerate(regionales):
        r = 4 + i
        sh.set(r, 4, rg, 's', S_BORD)
        sh.set(r, 5, 'COUNTIF(%s[Regional],D%d)' % (TIG, r), 'f', S_CEN)
    # --- TD3 / TD8: Por supervisor (G3:H18), orden top fijo ---
    sh.set(3, 7, 'Supervisor', 's', S_HEAD); sh.set(3, 8, 'Asignadas', 's', S_HEAD)
    for i, s in enumerate(supervisores_top):
        r = 4 + i
        sh.set(r, 7, s, 's', S_BORD)
        sh.set(r, 8, 'COUNTIFS(%s[Supervisor],G%d,%s[Estado_Actual],"ASIGNADO")' % (TIG, r, TIG), 'f', S_CEN)
    # --- TD4/5/6: por mes (J3:M14) con helper de seriales O/P ---
    sh.set(3, 10, 'Mes', 's', S_HEAD); sh.set(3, 11, 'Movimientos', 's', S_HEAD)
    sh.set(3, 12, 'Entregas', 's', S_HEAD); sh.set(3, 13, 'Perdidas', 's', S_HEAD)
    sh.set(3, 15, 'ini', 's', S_HEAD); sh.set(3, 16, 'fin', 's', S_HEAD)
    for i, (lbl, ini, fin) in enumerate(meses):
        r = 4 + i
        sh.set(r, 10, lbl, 's', S_CEN)
        sh.set(r, 15, ini, 'n', S_CEN); sh.set(r, 16, fin, 'n', S_CEN)
        sh.set(r, 11, 'COUNTIFS(%s[Fecha],">="&O%d,%s[Fecha],"<="&P%d)' % (TMO, r, TMO, r), 'f', S_CEN)
        sh.set(r, 12, 'COUNTIFS(%s[Fecha],">="&O%d,%s[Fecha],"<="&P%d,%s[Tipo_Movimiento],"ENTREGA INICIAL")+COUNTIFS(%s[Fecha],">="&O%d,%s[Fecha],"<="&P%d,%s[Tipo_Movimiento],"ASIGNACION")' % (TMO, r, TMO, r, TMO, TMO, r, TMO, r, TMO), 'f', S_CEN)
        sh.set(r, 13, 'COUNTIFS(%s[Fecha],">="&O%d,%s[Fecha],"<="&P%d,%s[Tipo_Movimiento],"PERDIDA")' % (TMO, r, TMO, r, TMO), 'f', S_CEN)
    # --- TD7: Estado por regional (A20:E26) ---
    sh.set(20, 1, 'Regional / Estado', 's', S_HEAD)
    for j, e in enumerate(estados_cred):
        sh.set(20, 2 + j, e, 's', S_HEAD)
    for i, rg in enumerate(regionales):
        r = 21 + i
        sh.set(r, 1, rg, 's', S_BORD)
        for j, e in enumerate(estados_cred):
            sh.set(r, 2 + j, 'COUNTIFS(%s[Regional],$A%d,%s[Estado_Actual],%s$20)' % (TIG, r, TIG, col_letter(2 + j)), 'f', S_CEN)
    # --- Listas para filtros con (TODOS) en U..X (21..24) ---
    fl = {21: ['(TODOS)'] + regionales, 22: ['(TODOS)'] + supervisores,
          23: ['(TODOS)'] + estados_cred, 24: ['(TODOS)'] + [m[0] for m in meses]}
    for cidx, vals in fl.items():
        sh.set(3, cidx, {21: 'fReg', 22: 'fSup', 23: 'fEst', 24: 'fMes'}[cidx], 's', S_HEAD)
        for i, v in enumerate(vals):
            sh.set(4 + i, cidx, v, 's', S_BORD)
    sh._fl = fl
    return sh



def build_dashboard():
    sh = Sheet('DASHBOARD', tabcolor='305496')
    sh.cols = [(i, i, 17) for i in range(1, 15)]
    hr = banner(sh, 'DASHBOARD EJECUTIVO  -  Control de Credenciales BEX', 14,
                'COMO USAR ESTA HOJA',
                'Todos los indicadores se calculan solos desde INVENTARIO_GENERAL y MOVIMIENTOS. '
                'Al final de la hoja hay un PANEL DE FILTROS (menus desplegables) que actua como segmentadores. '
                'Los graficos se actualizan automaticamente.')
    M = 'MOTOR_DASHBOARD'
    # ---- KPIs ----
    sh.set(5, 1, 'INDICADORES CLAVE', 's', S_SECTION); sh.merges.append('A5:N5')
    kpis = [
        ('TOTAL CREDENCIALES', 'COUNTA(%s[Codigo_Activo])' % TIG),
        ('ASIGNADAS', 'COUNTIF(%s[Estado_Actual],"ASIGNADO")' % TIG),
        ('DISPONIBLES', 'COUNTIF(%s[Estado_Actual],"DISPONIBLE")' % TIG),
        ('POR ASIGNAR', 'COUNTIF(%s[Estado_Actual],"POR ASIGNAR")' % TIG),
        ('PERDIDAS', 'COUNTIF(%s[Estado_Actual],"PERDIDO")' % TIG),
        ('DEVUELTAS', 'COUNTIF(%s[Tipo_Movimiento],"DEVOLUCION")' % TMO),
        ('STOCK EN OFICINA', 'COUNTIF(%s[Estado_Actual],"DISPONIBLE")' % TIG),
        ('STOCK EN USO', 'COUNTIF(%s[Estado_Actual],"ASIGNADO")' % TIG),
        ('PERSONAL ACTIVO', 'COUNTIF(%s[Estado],"ACTIVO")' % TPE),
        ('SUPERVISORES', 'COUNTIF(%s[Rol],"SUPERVISOR")' % TPE),
    ]
    positions = [1, 4, 7, 10, 13]
    for idx, (lbl, f) in enumerate(kpis):
        block = idx // 5
        lblrow = 6 + block * 3
        col = positions[idx % 5]
        c2 = col_letter(col + 1)
        sh.set(lblrow, col, lbl, 's', S_KPILBL); sh.merges.append('%s%d:%s%d' % (col_letter(col), lblrow, c2, lblrow))
        sh.set(lblrow + 1, col, f, 'f', S_KPIBOX); sh.merges.append('%s%d:%s%d' % (col_letter(col), lblrow + 1, c2, lblrow + 1))
        sh.rowheights[lblrow + 1] = 30
    # ---- Tablas resumen (referencian MOTOR) ----
    base = 13
    sh.set(base, 1, 'CREDENCIALES POR REGIONAL', 's', S_SECTION); sh.merges.append('A%d:E%d' % (base, base))
    sh.set(base + 1, 1, 'Regional', 's', S_HEAD); sh.set(base + 1, 2, 'Cantidad', 's', S_HEAD)
    for i, rg in enumerate(regionales):
        r = base + 2 + i
        sh.set(r, 1, "%s!D%d" % (M, 4 + i), 'f', S_BORD)
        sh.set(r, 2, "%s!E%d" % (M, 4 + i), 'f', S_CEN)
    sh.set(base, 8, 'DISTRIBUCION POR ESTADO', 's', S_SECTION); sh.merges.append('H%d:L%d' % (base, base))
    sh.set(base + 1, 8, 'Estado', 's', S_HEAD); sh.set(base + 1, 9, 'Cantidad', 's', S_HEAD)
    for i, e in enumerate(estados_cred):
        r = base + 2 + i
        sh.set(r, 8, "%s!A%d" % (M, 4 + i), 'f', S_BORD)
        sh.set(r, 9, "%s!B%d" % (M, 4 + i), 'f', S_CEN)
    # ---- Ultimos movimientos ----
    mrow = base + 10
    sh.set(mrow, 1, 'ULTIMOS MOVIMIENTOS REGISTRADOS', 's', S_SECTION); sh.merges.append('A%d:I%d' % (mrow, mrow))
    mc = ['ID_Movimiento', 'Fecha', 'Tipo_Movimiento', 'Codigo_Activo', 'Nombre', 'Supervisor', 'Regional']
    for i, c in enumerate(mc, 1):
        sh.set(mrow + 1, i, c, 's', S_HEAD)
    for k in range(10):
        r = mrow + 2 + k
        off = k - 9
        for i, col in enumerate(mc, 1):
            st = S_DATE if col == 'Fecha' else S_BORD
            sh.set(r, i, 'IFERROR(INDEX(%s[%s],COUNTA(%s[ID_Movimiento])%+d),"")' % (TMO, col, TMO, off), 'f', st)
    # ---- Graficos (drawing flotante) ----
    charts_row0 = mrow + 13            # fila 0-indexed para anclas
    sh._charts_anchor_row = charts_row0
    # ---- Panel de filtros (segmentadores) al FINAL, con vista limitada ----
    fr = charts_row0 + 66             # debajo de los graficos (cell row, 1-indexed aprox)
    sh.set(fr, 1, 'PANEL DE FILTROS INTERACTIVOS  (funcionan como segmentadores)', 's', S_SECTION); sh.merges.append('A%d:N%d' % (fr, fr))
    sh.set(fr + 1, 1, 'Regional:', 's', S_LBLR); sh.set(fr + 1, 2, '(TODOS)', 's', S_INPUT)
    sh.set(fr + 1, 4, 'Supervisor:', 's', S_LBLR); sh.set(fr + 1, 5, '(TODOS)', 's', S_INPUT)
    sh.set(fr + 1, 7, 'Estado:', 's', S_LBLR); sh.set(fr + 1, 8, '(TODOS)', 's', S_INPUT)
    sh.dvs.append(('B%d' % (fr + 1), 'val_fReg'))
    sh.dvs.append(('E%d' % (fr + 1), 'val_fSup'))
    sh.dvs.append(('H%d' % (fr + 1), 'val_fEst'))
    cond = ('(IF($B$%d="(TODOS)",1,--(%s[Regional]=$B$%d)))*(IF($E$%d="(TODOS)",1,--(%s[Supervisor]=$E$%d)))*(IF($H$%d="(TODOS)",1,--(%s[Estado_Actual]=$H$%d)))'
            % (fr + 1, TIG, fr + 1, fr + 1, TIG, fr + 1, fr + 1, TIG, fr + 1))
    sh.set(fr + 1, 10, 'En filtro:', 's', S_LBLR)
    sh.set(fr + 1, 11, 'SUMPRODUCT(%s)' % cond, 'f', S_KPIBOX)
    sh.set(fr + 3, 1, 'RESULTADO FILTRADO (primeras 30 filas; cambie los menus de arriba)', 's', S_SECTION); sh.merges.append('A%d:N%d' % (fr + 3, fr + 3))
    invcols = ['Codigo_Activo', 'Tipo_Credencial', 'Estado_Actual', 'Responsable_Actual', 'Telefono',
               'Supervisor', 'Regional', 'Fecha_Ultima_Asignacion', 'Fecha_Ultima_Devolucion',
               'Cantidad_Perdidas', 'Dias_En_Uso', 'Observaciones']
    for i, cn in enumerate(invcols, 1):
        sh.set(fr + 4, i, cn, 's', S_HEAD)
    sh.set(fr + 5, 1, 'IFERROR(TAKE(FILTER(%s,%s),30),"Sin resultados para este filtro")' % (TIG, cond), 'fa', S_BORD)
    sh.freeze = (hr, 0)
    return sh



def build_buscador():
    sh = Sheet('BUSCADOR', tabcolor='2E86C1')
    sh.cols = [(1, 1, 3), (2, 2, 26), (3, 3, 40), (4, 4, 18), (5, 5, 16),
               (6, 6, 16), (7, 7, 30), (8, 8, 30), (9, 9, 16)]
    hr = banner(sh, 'BUSCADOR INTELIGENTE  -  Sistema BEX', 9,
                'COMO USAR ESTA HOJA',
                'Escriba un codigo BEX en C5 (ej: BEX-0152) o un nombre/telefono en C6. '
                'Los resultados y el historial se muestran automaticamente. Si esta en oficina vera DISPONIBLE EN OFICINA; si se perdio, PERDIDA.')
    sh.set(5, 2, 'BUSCAR POR CODIGO BEX:', 's', S_LBLR)
    sh.set(5, 3, '', 's', S_INPUT); sh.merges.append('C5:E5')
    sh.set(6, 2, 'BUSCAR POR NOMBRE / TELEFONO:', 's', S_LBLR)
    sh.set(6, 3, '', 's', S_INPUT); sh.merges.append('C6:E6')
    K = '$C$5'
    def lk(col):
        return 'XLOOKUP(%s,%s[Codigo_Activo],%s[%s])' % (K, TIG, TIG, col)
    sh.set(8, 1, 'RESULTADO  -  BUSQUEDA POR CODIGO BEX', 's', S_SECTION); sh.merges.append('A8:I8')
    fields = [
        ('CODIGO BEX:', 'IF(%s="","-",IFERROR(%s,"NO ENCONTRADO"))' % (K, lk('Codigo_Activo')), S_BORD),
        ('TIPO DE CREDENCIAL:', 'IF(%s="","-",IFERROR(%s,"-"))' % (K, lk('Tipo_Credencial')), S_BORD),
        ('ESTADO ACTUAL:', 'IF(%s="","-",LET(e,IFERROR(%s,"NO ENCONTRADO"),IF(e="DISPONIBLE","DISPONIBLE EN OFICINA",IF(e="PERDIDO","PERDIDA",e))))' % (K, lk('Estado_Actual')), S_BORD),
        ('RESPONSABLE ACTUAL:', 'IF(%s="","-",IFERROR(%s,"-"))' % (K, lk('Responsable_Actual')), S_BORD),
        ('TELEFONO:', 'IF(%s="","-",IFERROR(%s,"-"))' % (K, lk('Telefono')), S_BORD),
        ('SUPERVISOR:', 'IF(%s="","-",IFERROR(%s,"-"))' % (K, lk('Supervisor')), S_BORD),
        ('REGIONAL:', 'IF(%s="","-",IFERROR(%s,"-"))' % (K, lk('Regional')), S_BORD),
        ('FECHA ASIGNACION:', 'IF(%s="","-",IFERROR(%s,"-"))' % (K, lk('Fecha_Ultima_Asignacion')), S_DATE),
        ('FECHA DEVOLUCION:', 'IF(%s="","-",IFERROR(%s,"-"))' % (K, lk('Fecha_Ultima_Devolucion')), S_DATE),
        ('CANTIDAD PERDIDAS:', 'IF(%s="","-",IFERROR(%s,"-"))' % (K, lk('Cantidad_Perdidas')), S_CEN),
        ('OBSERVACIONES:', 'IF(%s="","-",IFERROR(%s,"-"))' % (K, lk('Observaciones')), S_BORD),
    ]
    r = 9
    for lbl, f, st in fields:
        sh.set(r, 2, lbl, 's', S_LBLR); sh.set(r, 3, f, 'f', st); sh.merges.append('C%d:E%d' % (r, r)); r += 1
    base = r + 1
    sh.set(base, 1, 'RESULTADO  -  BUSQUEDA POR NOMBRE / TELEFONO', 's', S_SECTION); sh.merges.append('A%d:I%d' % (base, base))
    nombre_f = ('IF($C$6="","-",IFERROR(IFERROR(XLOOKUP($C$6,%s[Nombre_Completo],%s[Nombre_Completo]),'
                'XLOOKUP($C$6,%s[Telefono],%s[Nombre_Completo])),"NO ENCONTRADO"))' % (TPE, TPE, TPE, TPE))
    rn = base + 1
    sh.set(rn, 2, 'NOMBRE COMPLETO:', 's', S_LBLR); sh.set(rn, 3, nombre_f, 'f', S_BORD); sh.merges.append('C%d:E%d' % (rn, rn))
    key2 = '$C$%d' % rn
    def lkp(col):
        return 'IF($C$6="","-",IFERROR(XLOOKUP(%s,%s[Nombre_Completo],%s[%s]),"-"))' % (key2, TPE, TPE, col)
    for lbl, col in [('TELEFONO:', 'Telefono'), ('ROL:', 'Rol'), ('SUPERVISOR:', 'Supervisor'),
                     ('REGIONAL:', 'Regional'), ('ESTADO:', 'Estado'),
                     ('CREDENCIAL ACTUAL:', 'Credencial_Actual'), ('OBSERVACIONES:', 'Observaciones')]:
        rn += 1
        sh.set(rn, 2, lbl, 's', S_LBLR); sh.set(rn, 3, lkp(col), 'f', S_BORD); sh.merges.append('C%d:E%d' % (rn, rn))
    h = rn + 2
    sh.set(h, 1, 'HISTORIAL DE MOVIMIENTOS DE LA CREDENCIAL (codigo en C5)', 's', S_SECTION); sh.merges.append('A%d:I%d' % (h, h))
    for i, c in enumerate(['ID_Movimiento', 'Fecha', 'Tipo_Movimiento', 'Codigo_Activo', 'Telefono', 'Nombre', 'Supervisor', 'Regional', 'Observaciones'], 1):
        sh.set(h + 1, i, c, 's', S_HEAD)
    sh.set(h + 2, 1, 'IFERROR(FILTER(%s,%s[Codigo_Activo]=$C$5),"Escriba un codigo BEX en C5 para ver su historial")' % (TMO, TMO), 'fa', S_BORD)
    return sh

def build_inicio():
    sh = Sheet('INICIO', tabcolor='1F4E78')
    sh.cols = [(1, 1, 3), (2, 2, 40), (3, 3, 50), (4, 4, 20)]
    sh.set(1, 2, 'SISTEMA BEX  -  CONTROL DE CREDENCIALES', 's', S_TITLE); sh.merges.append('B1:D1'); sh.rowheights[1] = 34
    sh.set(2, 2, '?  COMO USAR ESTE ARCHIVO', 's', S_TITLE2); sh.merges.append('B2:D2')
    sh.set(3, 2, 'Sistema integrado y automatico. Ingrese datos solo en MOVIMIENTOS, PERSONAL e INVENTARIO_GENERAL. '
                 'El DASHBOARD, el BUSCADOR y los indicadores se calculan solos. Use los botones para navegar.',
           's', S_HELP); sh.merges.append('B3:D3'); sh.rowheights[3] = 46
    sh.set(5, 2, 'NAVEGACION', 's', S_SECTION); sh.merges.append('B5:D5')
    nav = [
        ('DASHBOARD', 'Indicadores, graficos y filtros'),
        ('BUSCADOR', 'Buscar credencial por codigo, nombre o telefono'),
        ('INVENTARIO_GENERAL', 'Fuente central: todas las credenciales'),
        ('PERSONAL', 'Tabla maestra de personas'),
        ('MOVIMIENTOS', 'Historial de eventos (entradas/salidas)'),
        ('PARAMETROS', 'Listas maestras del sistema'),
    ]
    r = 6
    for dst, desc in nav:
        sh.set(r, 2, 'Ir a  ->  ' + dst, 's', S_INPUT)
        sh.hyperlinks.append(('B%d' % r, "%s!A1" % dst, 'Ir a ' + dst))
        sh.set(r, 3, desc, 's', S_BORD)
        r += 1
    sh.set(r + 1, 2, 'RESUMEN RAPIDO', 's', S_SECTION); sh.merges.append('B%d:D%d' % (r + 1, r + 1))
    quick = [('Total credenciales', 'COUNTA(%s[Codigo_Activo])' % TIG),
             ('Asignadas', 'COUNTIF(%s[Estado_Actual],"ASIGNADO")' % TIG),
             ('Disponibles', 'COUNTIF(%s[Estado_Actual],"DISPONIBLE")' % TIG),
             ('Perdidas', 'COUNTIF(%s[Estado_Actual],"PERDIDO")' % TIG),
             ('Total personal', 'COUNTA(%s[Telefono])' % TPE)]
    rr = r + 2
    for lbl, f in quick:
        sh.set(rr, 2, lbl, 's', S_SUB); sh.set(rr, 3, f, 'f', S_CEN); rr += 1
    return sh



# ===========================================================================
# GRAFICOS
# ===========================================================================
def _pts(values):
    return '<c:ptCount val="%d"/>%s' % (len(values), ''.join(
        '<c:pt idx="%d"><c:v>%s</c:v></c:pt>' % (i, xesc(v)) for i, v in enumerate(values)))

CHART_HEAD = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:chart>')

def _title(t):
    return ('<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r>'
            '<a:rPr lang="es" b="1" sz="1200"/><a:t>%s</a:t></a:r></a:p></c:rich></c:tx>'
            '<c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>' % xesc(t))

def bar_chart(title, cat_ref, cats, val_ref, vals, color, bardir='col'):
    return (CHART_HEAD + _title(title) + '<c:plotArea><c:layout/>'
        '<c:barChart><c:barDir val="%s"/><c:grouping val="clustered"/><c:varyColors val="0"/>'
        '<c:ser><c:idx val="0"/><c:order val="0"/>'
        '<c:spPr><a:solidFill><a:srgbClr val="%s"/></a:solidFill></c:spPr>'
        '<c:cat><c:strRef><c:f>%s</c:f><c:strCache>%s</c:strCache></c:strRef></c:cat>'
        '<c:val><c:numRef><c:f>%s</c:f><c:numCache><c:formatCode>General</c:formatCode>%s</c:numCache></c:numRef></c:val>'
        '</c:ser><c:axId val="11"/><c:axId val="22"/></c:barChart>'
        '<c:catAx><c:axId val="11"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/>'
        '<c:axPos val="b"/><c:crossAx val="22"/></c:catAx>'
        '<c:valAx><c:axId val="22"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/>'
        '<c:axPos val="l"/><c:crossAx val="11"/></c:valAx></c:plotArea>'
        '<c:plotVisOnly val="1"/></c:chart></c:chartSpace>'
        % (bardir, color, xesc(cat_ref), _pts(cats), xesc(val_ref), _pts(vals)))

def line_chart(title, cat_ref, cats, val_ref, vals, color):
    return (CHART_HEAD + _title(title) + '<c:plotArea><c:layout/>'
        '<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>'
        '<c:ser><c:idx val="0"/><c:order val="0"/>'
        '<c:spPr><a:ln w="28575"><a:solidFill><a:srgbClr val="%s"/></a:solidFill></a:ln></c:spPr>'
        '<c:marker><c:symbol val="circle"/></c:marker>'
        '<c:cat><c:strRef><c:f>%s</c:f><c:strCache>%s</c:strCache></c:strRef></c:cat>'
        '<c:val><c:numRef><c:f>%s</c:f><c:numCache><c:formatCode>General</c:formatCode>%s</c:numCache></c:numRef></c:val>'
        '</c:ser><c:marker val="1"/><c:axId val="11"/><c:axId val="22"/></c:lineChart>'
        '<c:catAx><c:axId val="11"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/>'
        '<c:axPos val="b"/><c:crossAx val="22"/></c:catAx>'
        '<c:valAx><c:axId val="22"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/>'
        '<c:axPos val="l"/><c:crossAx val="11"/></c:valAx></c:plotArea>'
        '<c:plotVisOnly val="1"/></c:chart></c:chartSpace>'
        % (color, xesc(cat_ref), _pts(cats), xesc(val_ref), _pts(vals)))

def pie_chart(title, cat_ref, cats, val_ref, vals):
    return (CHART_HEAD + _title(title) + '<c:plotArea><c:layout/>'
        '<c:pieChart><c:varyColors val="1"/><c:ser><c:idx val="0"/><c:order val="0"/>'
        '<c:dLbls><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/>'
        '<c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>'
        '<c:cat><c:strRef><c:f>%s</c:f><c:strCache>%s</c:strCache></c:strRef></c:cat>'
        '<c:val><c:numRef><c:f>%s</c:f><c:numCache><c:formatCode>General</c:formatCode>%s</c:numCache></c:numRef></c:val>'
        '</c:ser><c:firstSliceAng val="0"/></c:pieChart>'
        '</c:plotArea>'
        '<c:legend><c:legendPos val="r"/><c:overlay val="0"/></c:legend>'
        '<c:plotVisOnly val="1"/></c:chart></c:chartSpace>'
        % (xesc(cat_ref), _pts(cats), xesc(val_ref), _pts(vals)))

def stacked_bar(title, cat_ref, cats, series, colors):
    # series: list of (name, val_ref, vals)
    sers = []
    for i, (nm, vref, vals) in enumerate(series):
        sers.append('<c:ser><c:idx val="%d"/><c:order val="%d"/>'
            '<c:tx><c:v>%s</c:v></c:tx>'
            '<c:spPr><a:solidFill><a:srgbClr val="%s"/></a:solidFill></c:spPr>'
            '<c:cat><c:strRef><c:f>%s</c:f><c:strCache>%s</c:strCache></c:strRef></c:cat>'
            '<c:val><c:numRef><c:f>%s</c:f><c:numCache><c:formatCode>General</c:formatCode>%s</c:numCache></c:numRef></c:val>'
            '</c:ser>' % (i, i, xesc(nm), colors[i % len(colors)], xesc(cat_ref), _pts(cats), xesc(vref), _pts(vals)))
    return (CHART_HEAD + _title(title) + '<c:plotArea><c:layout/>'
        '<c:barChart><c:barDir val="col"/><c:grouping val="stacked"/><c:varyColors val="0"/>'
        + ''.join(sers) + '<c:overlap val="100"/><c:axId val="11"/><c:axId val="22"/></c:barChart>'
        '<c:catAx><c:axId val="11"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/>'
        '<c:axPos val="b"/><c:crossAx val="22"/></c:catAx>'
        '<c:valAx><c:axId val="22"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/>'
        '<c:axPos val="l"/><c:crossAx val="11"/></c:valAx></c:plotArea>'
        '<c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend>'
        '<c:plotVisOnly val="1"/></c:chart></c:chartSpace>')

def anchor(c1, r1, c2, r2, rid, cid, name):
    return ('<xdr:twoCellAnchor><xdr:from><xdr:col>%d</xdr:col><xdr:colOff>0</xdr:colOff>'
        '<xdr:row>%d</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>'
        '<xdr:to><xdr:col>%d</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>%d</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>'
        '<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="%d" name="%s"/>'
        '<xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>'
        '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>'
        '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">'
        '<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="%s"/>'
        '</a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>'
        % (c1, r1, c2, r2, cid, name, rid))



# ===========================================================================
# EMPAQUETADO
# ===========================================================================
def package(path):
    inicio = build_inicio()
    dash = build_dashboard()
    busc = build_buscador()
    par, par_hr, par_last = build_parametros()
    inv = build_inventario(par_hr)
    per = build_personal()
    mov = build_movimientos()
    motor = build_motor()
    dash.drawing_rid = 'rId1'

    sheets = [inicio, dash, busc, inv, per, mov, par, motor]

    CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
    RELNS = 'http://schemas.openxmlformats.org/package/2006/relationships'
    DOC = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

    z = zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED)

    # ---- caches de graficos ----
    est_v = [sum(1 for it in inventario if it['estado'] == e) for e in estados_cred]
    reg_v = [sum(1 for it in inventario if it['reg'] == r) for r in regionales]
    sup_v = [sup_asig[s] for s in supervisores_top]
    def mcount(ini, fin, pred=lambda m: True):
        return sum(1 for m in movimientos if m['fecha'] and ini <= m['fecha'] <= fin and pred(m))
    mov_v = [mcount(i, f) for (_, i, f) in meses]
    ent_v = [mcount(i, f, lambda m: m['tipo'] in ('ENTREGA INICIAL', 'ASIGNACION')) for (_, i, f) in meses]
    per_v = [mcount(i, f, lambda m: m['tipo'] == 'PERDIDA') for (_, i, f) in meses]
    matrix = {e: [sum(1 for it in inventario if it['reg'] == rg and it['estado'] == e) for rg in regionales] for e in estados_cred}

    M = 'MOTOR_DASHBOARD'
    charts = {}
    charts[1] = pie_chart('Distribucion por Estado', '%s!$A$4:$A$7' % M, estados_cred, '%s!$B$4:$B$7' % M, est_v)
    charts[2] = bar_chart('Inventario por Regional', '%s!$D$4:$D$9' % M, regionales, '%s!$E$4:$E$9' % M, reg_v, '2E86C1')
    charts[3] = bar_chart('Top Supervisores (asignadas)', '%s!$G$4:$G$19' % M, supervisores_top, '%s!$H$4:$H$19' % M, sup_v, '27AE60', bardir='bar')
    colors = ['27AE60', '2E86C1', 'E67E22', 'C0392B']
    series = [(e, '%s!$%s$21:$%s$26' % (M, col_letter(2 + j), col_letter(2 + j)), matrix[e]) for j, e in enumerate(estados_cred)]
    charts[4] = stacked_bar('Estado por Regional', '%s!$A$21:$A$26' % M, regionales, series, colors)
    charts[5] = line_chart('Movimientos por Mes', '%s!$J$4:$J$15' % M, [m[0] for m in meses], '%s!$K$4:$K$15' % M, mov_v, '1F4E78')
    charts[6] = bar_chart('Entregas por Mes', '%s!$J$4:$J$15' % M, [m[0] for m in meses], '%s!$L$4:$L$15' % M, ent_v, '27AE60')
    charts[7] = bar_chart('Perdidas por Mes', '%s!$J$4:$J$15' % M, [m[0] for m in meses], '%s!$M$4:$M$15' % M, per_v, 'C0392B')

    # ---- [Content_Types] ----
    ct = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>', '<Types xmlns="%s">' % CT,
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
          '<Default Extension="xml" ContentType="application/xml"/>',
          '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
          '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
          '<Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/>']
    for i in range(1, len(sheets) + 1):
        ct.append('<Override PartName="/xl/worksheets/sheet%d.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' % i)
    for s in sheets:
        if s.table:
            ct.append('<Override PartName="/xl/tables/table%d.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>' % s.table['tid'])
    ct.append('<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>')
    for cid in charts:
        ct.append('<Override PartName="/xl/charts/chart%d.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>' % cid)
    ct.append('</Types>')
    z.writestr('[Content_Types].xml', ''.join(ct))

    z.writestr('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
               '<Relationships xmlns="%s"><Relationship Id="rId1" Type="%s/officeDocument" Target="xl/workbook.xml"/></Relationships>' % (RELNS, DOC))

    # ---- defined names ----
    def rng(col, r1, r2): return "PARAMETROS!$%s$%d:$%s$%d" % (col, r1, col, r2)
    h = par_hr
    dn = {
        'val_Regional': rng('A', h + 1, h + len(regionales)),
        'val_EstadoPersona': rng('B', h + 1, h + len(estados_persona)),
        'val_Rol': rng('C', h + 1, h + len(roles)),
        'val_EstadoCred': rng('D', h + 1, h + len(estados_cred)),
        'val_TipoMov': rng('E', h + 1, h + len(tipos_mov)),
        'val_Generacion': rng('F', h + 1, h + len(generaciones)),
        'val_Supervisor': rng('G', h + 1, h + len(supervisores)),
        'val_fReg': "%s!$U$4:$U$%d" % (M, 3 + len(motor._fl[21])),
        'val_fSup': "%s!$V$4:$V$%d" % (M, 3 + len(motor._fl[22])),
        'val_fEst': "%s!$W$4:$W$%d" % (M, 3 + len(motor._fl[23])),
        'val_fMes': "%s!$X$4:$X$%d" % (M, 3 + len(motor._fl[24])),
    }

    # ---- workbook.xml ----
    wb = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
          '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="%s"><sheets>' % DOC]
    for i, s in enumerate(sheets, 1):
        st = ' state="hidden"' if s.hidden else ''
        wb.append('<sheet name="%s" sheetId="%d"%s r:id="rId%d"/>' % (xesc(s.name), i, st, i))
    wb.append('</sheets><definedNames>')
    for nm, rf in dn.items():
        wb.append('<definedName name="%s">%s</definedName>' % (nm, xesc(rf)))
    wb.append('</definedNames><calcPr calcId="0" fullCalcOnLoad="1"/></workbook>')
    z.writestr('xl/workbook.xml', ''.join(wb))

    # ---- workbook rels ----
    wr = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>', '<Relationships xmlns="%s">' % RELNS]
    for i in range(1, len(sheets) + 1):
        wr.append('<Relationship Id="rId%d" Type="%s/worksheet" Target="worksheets/sheet%d.xml"/>' % (i, DOC, i))
    wr.append('<Relationship Id="rId%d" Type="%s/styles" Target="styles.xml"/>' % (len(sheets) + 1, DOC))
    wr.append('<Relationship Id="rId%d" Type="%s/sheetMetadata" Target="metadata.xml"/>' % (len(sheets) + 2, DOC))
    wr.append('</Relationships>')
    z.writestr('xl/_rels/workbook.xml.rels', ''.join(wr))

    z.writestr('xl/styles.xml', STYLES_XML)
    z.writestr('xl/metadata.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">'
        '<metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" '
        'pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" '
        'clearComments="1" assign="1" coerce="1" cellMeta="1"/></metadataTypes>'
        '<futureMetadata name="XLDAPR" count="1"><bk><extLst>'
        '<ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}">'
        '<xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata>'
        '<cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>')

    # ---- worksheets + tables ----
    for i, s in enumerate(sheets, 1):
        rid_table = 'rId1' if s.table else None
        z.writestr('xl/worksheets/sheet%d.xml' % i, s.xml(rid_table))
        rels = []
        if s.drawing_rid:
            rels.append('<Relationship Id="rId1" Type="%s/drawing" Target="../drawings/drawing1.xml"/>' % DOC)
        if s.table:
            tid_local = 2 if s.drawing_rid else 1
            rid = 'rId%d' % tid_local
            # asegurar el r:id usado en tableParts coincide
            rels.append('<Relationship Id="%s" Type="%s/table" Target="../tables/table%d.xml"/>' % (rid, DOC, s.table['tid']))
        if rels:
            z.writestr('xl/worksheets/_rels/sheet%d.xml.rels' % i,
                       '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="%s">%s</Relationships>'
                       % (RELNS, ''.join(rels)))
        if s.table:
            t = s.table
            z.writestr('xl/tables/table%d.xml' % t['tid'], table_xml(t['tid'], t['name'], t['ref'], t['columns']))

    # ---- charts + drawing (dashboard) ----
    for cid, xml in charts.items():
        z.writestr('xl/charts/chart%d.xml' % cid, xml)
    b0 = dash._charts_anchor_row
    layout = [(0, 0), (7, 0), (0, 16), (7, 16), (0, 32), (7, 32), (0, 48)]
    dr = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
          '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
          'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">']
    names = ['Estado', 'Regional', 'Supervisor', 'EstadoRegional', 'MovMes', 'EntregasMes', 'PerdidasMes']
    for k in range(7):
        cx, ry = layout[k]
        dr.append(anchor(cx, b0 + ry, cx + 7, b0 + ry + 15, 'rId%d' % (k + 1), k + 2, 'Grafico ' + names[k]))
    dr.append('</xdr:wsDr>')
    z.writestr('xl/drawings/drawing1.xml', ''.join(dr))
    z.writestr('xl/drawings/_rels/drawing1.xml.rels',
               '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="%s">%s</Relationships>'
               % (RELNS, ''.join('<Relationship Id="rId%d" Type="%s/chart" Target="../charts/chart%d.xml"/>' % (k + 1, DOC, k + 1) for k in range(7))))

    z.close()
    print('Generado:', path)

package('SISTEMA_BEX_CREDENCIALES.xlsx')
