# -*- coding: utf-8 -*-
"""
Generador del Sistema BEX de Control de Credenciales.
Fusiona ARCHIVO 1 (datos reales) + ARCHIVO 2 (diseño) en un unico .xlsx
funcional con tablas estructuradas, formulas dinamicas, buscador y dashboard.
Escrito con libreria estandar (sin openpyxl / pandas).
"""
import zipfile, re, os
from xml.etree import ElementTree as ET

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'

# ---------------------------------------------------------------------------
# LECTOR de los archivos de origen
# ---------------------------------------------------------------------------
def read_all(path):
    z = zipfile.ZipFile(path)
    shared = []
    if 'xl/sharedStrings.xml' in z.namelist():
        t = ET.fromstring(z.read('xl/sharedStrings.xml'))
        for si in t.findall(NS + 'si'):
            shared.append(''.join(n.text or '' for n in si.iter(NS + 't')))
    wb = ET.fromstring(z.read('xl/workbook.xml'))
    sheets = [s.get('name') for s in wb.iter(NS + 'sheet')]

    def colidx(ref):
        m = re.match(r'([A-Z]+)', ref); n = 0
        for c in m.group(1):
            n = n * 26 + (ord(c) - 64)
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

def clean(s):
    return (s or '').strip()

def to_int(s):
    s = clean(s)
    return int(s) if re.fullmatch(r'\d+', s) else None

# ---------------------------------------------------------------------------
# NORMALIZACION de supervisores (nombres cortos -> canonicos)
# ---------------------------------------------------------------------------
SUP_MAP = {
    'JOSE GUTIERREZ': 'JOSE GUTIERREZ PEDRAZA',
    'MARIA FERNANDA DAZA': 'MARIA FERNANDA DAZA PALMA',
    'DELIA JORDAN': 'DELIA JORDAN FACUSSE',
}
def norm_sup(s):
    s = clean(s)
    return SUP_MAP.get(s, s)

# ---------------------------------------------------------------------------
# MODELO: PERSONAL
# ---------------------------------------------------------------------------
personal = []
for row in A2['PERSONAL'][3:]:
    if not clean(row[0] if len(row) > 0 else ''):
        continue
    r = (row + [''] * 10)[:10]
    personal.append({
        'id': clean(r[0]),
        'tel': clean(r[1]) or 'POR REGISTRAR',
        'nombre': clean(r[2]),
        'rol': clean(r[3]),
        'sup': norm_sup(r[4]),
        'reg': clean(r[5]),
        'estado': clean(r[6]),
        'cred': clean(r[7]),
        'fing': clean(r[8]),
        'obs': clean(r[9]),
    })

# ---------------------------------------------------------------------------
# MODELO: MOVIMIENTOS
# ---------------------------------------------------------------------------
movimientos = []
for row in A2['MOVIMIENTOS'][3:]:
    if not clean(row[0] if len(row) > 0 else ''):
        continue
    r = (row + [''] * 9)[:9]
    movimientos.append({
        'n': clean(r[0]),
        'fecha': to_int(r[1]),
        'tipo': clean(r[2]),
        'cod': clean(r[3]),
        'afil': clean(r[4]),
        'sup': norm_sup(r[5]),
        'reg': clean(r[6]),
        'obs': clean(r[7]),
        'por': clean(r[8]) or 'SISTEMA',
    })

# ---------------------------------------------------------------------------
# MODELO: INVENTARIO_GENERAL  (base ACTIVOS + BEX-2051..2100)
# ---------------------------------------------------------------------------
inventario = []
seen = set()
for row in A2['ACTIVOS'][3:]:
    cod = clean(row[0] if len(row) > 0 else '')
    if not cod:
        continue
    r = (row + [''] * 9)[:9]
    fecha = to_int(r[6])
    inventario.append({
        'cod': cod,
        'gen': clean(r[1]),
        'estado': clean(r[2]),
        'resp': clean(r[3]),
        'sup': norm_sup(r[4]),
        'reg': clean(r[5]),
        'fecha': fecha,
        'obs': clean(r[8]),
    })
    seen.add(cod)

# Agregar credenciales nuevas en oficina BEX-2051..BEX-2100 (del INVENTARIO real)
for n in range(2051, 2101):
    cod = 'BEX-%04d' % n
    if cod in seen:
        continue
    inventario.append({
        'cod': cod,
        'gen': 'NUEVA (BEX-2xxx)',
        'estado': 'DISPONIBLE',
        'resp': 'EN OFICINA',
        'sup': '',
        'reg': 'SANTA CRUZ',
        'fecha': 46176,
        'obs': 'Stock en oficina - recepcionado',
    })
    seen.add(cod)

# ---------------------------------------------------------------------------
# MODELO: PARAMETROS
# ---------------------------------------------------------------------------
param_cols = {
    'Regionales': ['COCHABAMBA', 'LA PAZ', 'ORURO', 'SANTA CRUZ', 'SUCRE', 'TARIJA'],
    'Estados_Persona': ['ACTIVO', 'INACTIVO'],
    'Roles': ['AFILIADOR', 'SUPERVISOR'],
    'Estados_Credencial': ['ASIGNADO', 'DISPONIBLE', 'POR ASIGNAR', 'PERDIDO', 'EN TRAMITE'],
    'Tipos_Movimiento': ['ENTREGA INICIAL', 'ASIGNACION', 'DEVOLUCION', 'PERDIDA',
                         'REPOSICION', 'REASIGNACION', 'CAMBIO SUPERVISOR', 'CAMBIO REGIONAL', 'BAJA'],
    'Generacion': ['NUEVA (BEX-2xxx)', 'LEGACY (BEX-0xxx)'],
}
errores = [
    '1. BEX-0176 asignado a DOS personas: CLAUDIA SHASKIA y REYNA IVONNE',
    '2. BEX-0106 posiblemente duplicado: CAMILA LOZA (GERCY) y JHOSEP SHARUM (MAGALI)',
    '3. 3 personas con nombres incompletos: TOMAS, LORENA, SELMA',
    '4. BEX-0518 reportada como PERDIDA - pendiente cobro a MARCO ESCOBAR',
    '5. SUPERVISORES sin datos en hoja principal: LIONEL CHABUR, NATALIA VILLARROEL',
    '6. Codigos normalizados (espacios y ceros): BEX-162, BEX- 0176, BEX-027, etc.',
    '7. Todos los telefonos marcados como POR REGISTRAR - campo critico faltante',
]

# Supervisores canonicos (orden y regional) para el dashboard
supervisores = [
    ('HASIRA DANIELA OSINAGA CHOQUE', 'COCHABAMBA'),
    ('PAMELA FANNY CALANI LAURA', 'COCHABAMBA'),
    ('NATALIA VILLARROEL', 'COCHABAMBA'),
    ('CLAUDIA SHASKIA CALLE NINA', 'LA PAZ'),
    ('GERCY EVER ERGUETA KIPPES', 'LA PAZ'),
    ('JOEL DAVID HUANCA HUANCA', 'LA PAZ'),
    ('MAGALI YESENIA HUANCA HUANCA', 'LA PAZ'),
    ('MILENKO ADRIANA ORDONEZ NUNEZ', 'LA PAZ'),
    ('LIONEL CHABUR', 'LA PAZ'),
    ('MIRIAM ROSA CHAMBI QUISPE', 'ORURO'),
    ('JORGE SAAVEDRA', 'SANTA CRUZ'),
    ('BEATRIZ OVIEDO OVIEDO', 'SANTA CRUZ'),
    ('DELIA JORDAN FACUSSE', 'SANTA CRUZ'),
    ('JOSE GUTIERREZ PEDRAZA', 'SANTA CRUZ'),
    ('JENNY CRISTINA ECHALAR MONTALVO', 'SUCRE'),
    ('MARIA FERNANDA DAZA PALMA', 'TARIJA'),
]
regionales = ['COCHABAMBA', 'LA PAZ', 'ORURO', 'SANTA CRUZ', 'SUCRE', 'TARIJA']

print('Personal:', len(personal))
print('Movimientos:', len(movimientos))
print('Inventario:', len(inventario))
print('Estados inventario:', sorted(set(i['estado'] for i in inventario)))
print('OK modelo')



# ===========================================================================
# MOTOR DE ESCRITURA XLSX (libreria estandar)
# ===========================================================================
def col_letter(c):
    s = ''
    while c > 0:
        c, r = divmod(c - 1, 26)
        s = chr(65 + r) + s
    return s

def xesc(s):
    return (str(s).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            .replace('"', '&quot;'))

def fix_formula(f):
    """Excel almacena funciones modernas con prefijo _xlfn. (y _xlws. para FILTER)."""
    f = re.sub(r'(?<![A-Za-z0-9_.])XLOOKUP\(', '_xlfn.XLOOKUP(', f)
    f = re.sub(r'(?<![A-Za-z0-9_.])MAXIFS\(', '_xlfn.MAXIFS(', f)
    f = re.sub(r'(?<![A-Za-z0-9_.])MINIFS\(', '_xlfn.MINIFS(', f)
    f = re.sub(r'(?<![A-Za-z0-9_.])FILTER\(', '_xlfn._xlws.FILTER(', f)
    return f

class Sheet:
    def __init__(self, name):
        self.name = name
        self.cells = {}          # (row,col) -> (value, kind, style)  kind: s,n,f
        self.cols = []           # (minc, maxc, width)
        self.merges = []         # "A1:B2"
        self.table = None        # dict
        self.freeze = None       # (row, col) panes freeze top-left
        self.rowheights = {}     # row -> height
        self.drawing_rid = None  # r:id del drawing (graficos)

    def set(self, r, c, value, kind='s', style=0):
        self.cells[(r, c)] = (value, kind, style)

    def setc(self, ref, value, kind='s', style=0):
        m = re.match(r'([A-Z]+)(\d+)', ref)
        col = 0
        for ch in m.group(1):
            col = col * 26 + (ord(ch) - 64)
        self.set(int(m.group(2)), col, value, kind, style)

    def xml(self, rid_table=None):
        if self.cells:
            maxr = max(r for r, c in self.cells)
            maxc = max(c for r, c in self.cells)
        else:
            maxr = maxc = 1
        dim = 'A1:%s%d' % (col_letter(maxc), maxr)
        out = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>']
        out.append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                   'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')
        out.append('<dimension ref="%s"/>' % dim)
        # sheetViews
        if self.freeze:
            fr, fc = self.freeze
            top = '%s%d' % (col_letter(fc + 1), fr + 1)
            pane = ('<pane xSplit="%d" ySplit="%d" topLeftCell="%s" activePane="bottomRight" state="frozen"/>'
                    % (fc, fr, top))
            out.append('<sheetViews><sheetView workbookViewId="0">%s'
                       '<selection pane="bottomRight"/></sheetView></sheetViews>' % pane)
        else:
            out.append('<sheetViews><sheetView workbookViewId="0"/></sheetViews>')
        out.append('<sheetFormatPr defaultRowHeight="15"/>')
        if self.cols:
            out.append('<cols>')
            for mn, mx, w in self.cols:
                out.append('<col min="%d" max="%d" width="%.2f" customWidth="1"/>' % (mn, mx, w))
            out.append('</cols>')
        out.append('<sheetData>')
        for r in range(1, maxr + 1):
            rowcells = [(c, self.cells[(r, c)]) for c in range(1, maxc + 1) if (r, c) in self.cells]
            if not rowcells:
                continue
            hattr = ''
            if r in self.rowheights:
                hattr = ' ht="%.2f" customHeight="1"' % self.rowheights[r]
            out.append('<row r="%d"%s>' % (r, hattr))
            for c, (val, kind, style) in rowcells:
                ref = '%s%d' % (col_letter(c), r)
                sattr = ' s="%d"' % style if style else ''
                if kind == 'f':
                    out.append('<c r="%s"%s><f>%s</f></c>' % (ref, sattr, xesc(fix_formula(val))))
                elif kind == 'fa':
                    out.append('<c r="%s"%s cm="1"><f>%s</f></c>'
                               % (ref, sattr, xesc(fix_formula(val))))
                elif kind == 'n':
                    out.append('<c r="%s"%s><v>%s</v></c>' % (ref, sattr, val))
                else:
                    out.append('<c r="%s"%s t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>'
                               % (ref, sattr, xesc(val)))
            out.append('</row>')
        out.append('</sheetData>')
        if self.merges:
            out.append('<mergeCells count="%d">' % len(self.merges))
            for m in self.merges:
                out.append('<mergeCell ref="%s"/>' % m)
            out.append('</mergeCells>')
        if self.drawing_rid:
            out.append('<drawing r:id="%s"/>' % self.drawing_rid)
        if rid_table:
            out.append('<tableParts count="1"><tablePart r:id="%s"/></tableParts>' % rid_table)
        out.append('</worksheet>')
        return ''.join(out)


def table_xml(tid, name, ref, columns):
    out = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>']
    out.append('<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
               'id="%d" name="%s" displayName="%s" ref="%s" totalsRowShown="0">' % (tid, name, name, ref))
    out.append('<autoFilter ref="%s"/>' % ref)
    out.append('<tableColumns count="%d">' % len(columns))
    for i, cn in enumerate(columns, 1):
        out.append('<tableColumn id="%d" name="%s"/>' % (i, xesc(cn)))
    out.append('</tableColumns>')
    out.append('<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" '
               'showRowStripes="1" showColumnStripes="0"/>')
    out.append('</table>')
    return ''.join(out)


STYLES_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="1"><numFmt numFmtId="164" formatCode="dd/mm/yyyy"/></numFmts>
<fonts count="7">
<font><sz val="11"/><name val="Calibri"/></font>
<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
<font><b/><sz val="16"/><color rgb="FF1F4E78"/><name val="Calibri"/></font>
<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
<font><b/><sz val="22"/><color rgb="FF1F4E78"/><name val="Calibri"/></font>
<font><b/><sz val="11"/><name val="Calibri"/></font>
<font><b/><sz val="12"/><color rgb="FF1F4E78"/><name val="Calibri"/></font>
</fonts>
<fills count="6">
<fill><patternFill patternType="none"/></fill>
<fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFD9E1F2"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF2E86C1"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFBDD7EE"/><bgColor indexed="64"/></patternFill></fill>
</fills>
<borders count="2">
<border><left/><right/><top/><bottom/><diagonal/></border>
<border><left style="thin"><color rgb="FF8EAADB"/></left><right style="thin"><color rgb="FF8EAADB"/></right><top style="thin"><color rgb="FF8EAADB"/></top><bottom style="thin"><color rgb="FF8EAADB"/></bottom><diagonal/></border>
</borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="12">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>
<xf numFmtId="0" fontId="6" fillId="5" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment vertical="center"/></xf>
<xf numFmtId="0" fontId="4" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="center"/></xf>
<xf numFmtId="0" fontId="3" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
<xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf>
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>
<xf numFmtId="10" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf>
<xf numFmtId="0" fontId="5" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf>
<xf numFmtId="0" fontId="4" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
<dxfs count="0"/>
</styleSheet>'''

# style indices
S_DEF, S_HEAD, S_TITLE, S_SECTION, S_KPINUM, S_KPILBL, S_DATE, S_BORD, S_PCT, S_SUB, S_CEN, S_KPIBOX = range(12)



# ===========================================================================
# CONSTRUCCION DE HOJAS
# ===========================================================================
TIG = 'TablaInventarioGeneral'
TPE = 'TablaPersonal'
TMO = 'TablaMovimientos'
TPA = 'TablaParametros'

def build_inventario():
    sh = Sheet('INVENTARIO_GENERAL')
    cols = ['Codigo_Activo', 'Generacion', 'Estado_Actual', 'Responsable_Actual', 'Telefono',
            'Supervisor', 'Regional', 'Fecha_Ultima_Asignacion', 'Fecha_Ultima_Devolucion',
            'Cantidad_Perdidas', 'Dias_En_Uso', 'Observaciones']
    widths = [14, 18, 14, 34, 14, 30, 13, 16, 16, 12, 11, 40]
    sh.cols = [(i + 1, i + 1, w) for i, w in enumerate(widths)]
    for i, cn in enumerate(cols, 1):
        sh.set(1, i, cn, 's', S_HEAD)
    rr = 2
    for it in inventario:
        sh.set(rr, 1, it['cod'], 's', S_BORD)
        sh.set(rr, 2, it['gen'], 's', S_BORD)
        sh.set(rr, 3, it['estado'], 's', S_CEN)
        sh.set(rr, 4, it['resp'], 's', S_BORD)
        sh.set(rr, 5, 'IFERROR(XLOOKUP([@Responsable_Actual],%s[Nombres_Apellidos],%s[Telefono]),"POR REGISTRAR")' % (TPE, TPE), 'f', S_CEN)
        sh.set(rr, 6, it['sup'], 's', S_BORD)
        sh.set(rr, 7, it['reg'], 's', S_BORD)
        sh.set(rr, 8, 'IF(COUNTIFS(%s[Codigo_BEX],[@Codigo_Activo])=0,"",MAXIFS(%s[Fecha],%s[Codigo_BEX],[@Codigo_Activo]))' % (TMO, TMO, TMO), 'f', S_DATE)
        sh.set(rr, 9, 'IF(COUNTIFS(%s[Codigo_BEX],[@Codigo_Activo],%s[Tipo_Movimiento],"DEVOLUCION")=0,"",MAXIFS(%s[Fecha],%s[Codigo_BEX],[@Codigo_Activo],%s[Tipo_Movimiento],"DEVOLUCION"))' % (TMO, TMO, TMO, TMO, TMO), 'f', S_DATE)
        sh.set(rr, 10, 'COUNTIFS(%s[Codigo_BEX],[@Codigo_Activo],%s[Tipo_Movimiento],"PERDIDA")' % (TMO, TMO), 'f', S_CEN)
        sh.set(rr, 11, 'IF(NOT(ISNUMBER([@Fecha_Ultima_Asignacion])),"",TODAY()-[@Fecha_Ultima_Asignacion])', 'f', S_CEN)
        sh.set(rr, 12, it['obs'], 's', S_BORD)
        rr += 1
    last = rr - 1
    sh.table = dict(tid=1, name=TIG, ref='A1:%s%d' % (col_letter(12), last), columns=cols)
    sh.freeze = (1, 0)
    return sh

def build_personal():
    sh = Sheet('PERSONAL')
    cols = ['ID_Persona', 'Telefono', 'Nombres_Apellidos', 'Rol', 'Supervisor', 'Regional',
            'Estado', 'Credencial_Actual', 'Fecha_Ingreso', 'Observaciones']
    widths = [12, 14, 38, 13, 30, 13, 11, 16, 14, 28]
    sh.cols = [(i + 1, i + 1, w) for i, w in enumerate(widths)]
    for i, cn in enumerate(cols, 1):
        sh.set(1, i, cn, 's', S_HEAD)
    rr = 2
    for p in personal:
        sh.set(rr, 1, p['id'], 's', S_BORD)
        sh.set(rr, 2, p['tel'], 's', S_CEN)
        sh.set(rr, 3, p['nombre'], 's', S_BORD)
        sh.set(rr, 4, p['rol'], 's', S_CEN)
        sh.set(rr, 5, p['sup'], 's', S_BORD)
        sh.set(rr, 6, p['reg'], 's', S_BORD)
        sh.set(rr, 7, p['estado'], 's', S_CEN)
        sh.set(rr, 8, p['cred'], 's', S_CEN)
        sh.set(rr, 9, p['fing'], 's', S_CEN)
        sh.set(rr, 10, p['obs'], 's', S_BORD)
        rr += 1
    last = rr - 1
    sh.table = dict(tid=2, name=TPE, ref='A1:%s%d' % (col_letter(10), last), columns=cols)
    sh.freeze = (1, 0)
    return sh

def build_movimientos():
    sh = Sheet('MOVIMIENTOS')
    cols = ['N_Mov', 'Fecha', 'Tipo_Movimiento', 'Codigo_BEX', 'Afiliador', 'Supervisor',
            'Regional', 'Observaciones', 'Registrado_Por']
    widths = [11, 13, 18, 13, 34, 30, 13, 40, 14]
    sh.cols = [(i + 1, i + 1, w) for i, w in enumerate(widths)]
    for i, cn in enumerate(cols, 1):
        sh.set(1, i, cn, 's', S_HEAD)
    rr = 2
    for m in movimientos:
        sh.set(rr, 1, m['n'], 's', S_BORD)
        if m['fecha'] is not None:
            sh.set(rr, 2, m['fecha'], 'n', S_DATE)
        else:
            sh.set(rr, 2, '', 's', S_DATE)
        sh.set(rr, 3, m['tipo'], 's', S_CEN)
        sh.set(rr, 4, m['cod'], 's', S_CEN)
        sh.set(rr, 5, m['afil'], 's', S_BORD)
        sh.set(rr, 6, m['sup'], 's', S_BORD)
        sh.set(rr, 7, m['reg'], 's', S_BORD)
        sh.set(rr, 8, m['obs'], 's', S_BORD)
        sh.set(rr, 9, m['por'], 's', S_CEN)
        rr += 1
    last = rr - 1
    sh.table = dict(tid=3, name=TMO, ref='A1:%s%d' % (col_letter(9), last), columns=cols)
    sh.freeze = (1, 0)
    return sh

def build_parametros():
    sh = Sheet('PARAMETROS')
    cols = ['Regionales', 'Estados_Persona', 'Roles', 'Estados_Credencial',
            'Tipos_Movimiento', 'Generacion']
    widths = [16, 16, 14, 18, 20, 20]
    sh.cols = [(i + 1, i + 1, w) for i, w in enumerate(widths)] + [(8, 8, 70)]
    for i, cn in enumerate(cols, 1):
        sh.set(1, i, cn, 's', S_HEAD)
    order = ['Regionales', 'Estados_Persona', 'Roles', 'Estados_Credencial', 'Tipos_Movimiento', 'Generacion']
    maxlen = max(len(param_cols[k]) for k in order)
    for ci, k in enumerate(order, 1):
        vals = param_cols[k]
        for ri in range(maxlen):
            v = vals[ri] if ri < len(vals) else ''
            sh.set(2 + ri, ci, v, 's', S_BORD)
    last = 1 + maxlen
    sh.table = dict(tid=4, name=TPA, ref='A1:%s%d' % (col_letter(6), last), columns=cols)
    # Errores criticos (fuera de la tabla)
    sh.set(1, 8, 'ERRORES_CRITICOS_DETECTADOS', 's', S_HEAD)
    for i, e in enumerate(errores, 2):
        sh.set(i, 8, e, 's', S_BORD)
    return sh



def build_buscador():
    sh = Sheet('BUSCADOR')
    sh.cols = [(1, 1, 3), (2, 2, 26), (3, 3, 42), (4, 4, 18), (5, 5, 16),
               (6, 6, 16), (7, 7, 30), (8, 8, 30), (9, 9, 16)]
    sh.setc('A1', 'BUSCADOR INTELIGENTE DE ACTIVOS  -  Sistema BEX', 's', S_TITLE)
    sh.merges.append('A1:I1')
    # entradas
    sh.setc('B3', 'BUSCAR CODIGO BEX:', 's', S_SUB)
    sh.setc('C3', '', 's', S_KPIBOX)
    sh.setc('D3', '<- escriba el codigo BEX (ej: BEX-0152)', 's', S_DEF)
    sh.setc('B4', 'BUSCAR NOMBRE / TELEFONO:', 's', S_SUB)
    sh.setc('C4', '', 's', S_KPIBOX)
    sh.setc('D4', '<- escriba nombre o telefono', 's', S_DEF)

    # ---- resultado por codigo ----
    sh.setc('A6', 'RESULTADO  -  BUSQUEDA POR CODIGO BEX', 's', S_SECTION)
    sh.merges.append('A6:I6')
    def lk(col):
        return 'XLOOKUP($C$3,%s[Codigo_Activo],%s[%s])' % (TIG, TIG, col)
    rows_cod = [
        ('CODIGO BEX:', 'IF($C$3="","-",IFERROR(%s,"NO ENCONTRADO"))' % lk('Codigo_Activo'), S_BORD),
        ('GENERACION:', 'IF($C$3="","-",IFERROR(%s,"-"))' % lk('Generacion'), S_BORD),
        ('ESTADO:', 'IF($C$3="","-",IFERROR(IF(%s="DISPONIBLE","DISPONIBLE EN OFICINA",IF(%s="PERDIDO","PERDIDA",%s)),"NO ENCONTRADO"))' % (lk('Estado_Actual'), lk('Estado_Actual'), lk('Estado_Actual')), S_BORD),
        ('RESPONSABLE ACTUAL:', 'IF($C$3="","-",IFERROR(%s,"-"))' % lk('Responsable_Actual'), S_BORD),
        ('TELEFONO:', 'IF($C$3="","-",IFERROR(%s,"-"))' % lk('Telefono'), S_BORD),
        ('SUPERVISOR:', 'IF($C$3="","-",IFERROR(%s,"-"))' % lk('Supervisor'), S_BORD),
        ('REGIONAL:', 'IF($C$3="","-",IFERROR(%s,"-"))' % lk('Regional'), S_BORD),
        ('FECHA ASIGNACION:', 'IF($C$3="","-",IFERROR(%s,"-"))' % lk('Fecha_Ultima_Asignacion'), S_DATE),
        ('CANTIDAD PERDIDAS:', 'IF($C$3="","-",IFERROR(%s,"-"))' % lk('Cantidad_Perdidas'), S_CEN),
        ('OBSERVACIONES:', 'IF($C$3="","-",IFERROR(%s,"-"))' % lk('Observaciones'), S_BORD),
    ]
    r = 7
    for lbl, f, st in rows_cod:
        sh.setc('B%d' % r, lbl, 's', S_SUB)
        sh.set(r, 3, f, 'f', st)
        sh.merges.append('C%d:E%d' % (r, r))
        r += 1

    # ---- resultado por nombre/telefono ----
    base = r + 1
    sh.setc('A%d' % base, 'RESULTADO  -  BUSQUEDA POR NOMBRE / TELEFONO', 's', S_SECTION)
    sh.merges.append('A%d:I%d' % (base, base))
    nombre_f = ('IF($C$4="","-",IFERROR(IFERROR(XLOOKUP($C$4,%s[Nombres_Apellidos],%s[Nombres_Apellidos]),'
                'XLOOKUP($C$4,%s[Telefono],%s[Nombres_Apellidos])),"NO ENCONTRADO"))' % (TPE, TPE, TPE, TPE))
    rn = base + 1
    sh.setc('B%d' % rn, 'NOMBRE COMPLETO:', 's', S_SUB)
    sh.set(rn, 3, nombre_f, 'f', S_BORD); sh.merges.append('C%d:E%d' % (rn, rn))
    key = '$C$%d' % rn  # nombre resuelto
    def lkp(col):
        return 'IF($C$4="","-",IFERROR(XLOOKUP(%s,%s[Nombres_Apellidos],%s[%s]),"-"))' % (key, TPE, TPE, col)
    person_rows = [
        ('TELEFONO:', 'Telefono'), ('ROL:', 'Rol'), ('SUPERVISOR:', 'Supervisor'),
        ('REGIONAL:', 'Regional'), ('ESTADO:', 'Estado'),
        ('CREDENCIAL ACTUAL:', 'Credencial_Actual'), ('OBSERVACIONES:', 'Observaciones'),
    ]
    rr = rn + 1
    for lbl, col in person_rows:
        sh.setc('B%d' % rr, lbl, 's', S_SUB)
        sh.set(rr, 3, lkp(col), 'f', S_BORD); sh.merges.append('C%d:E%d' % (rr, rr))
        rr += 1

    # ---- historial ----
    h = rr + 1
    sh.setc('A%d' % h, 'HISTORIAL DE MOVIMIENTOS DE LA CREDENCIAL (codigo en C3)', 's', S_SECTION)
    sh.merges.append('A%d:I%d' % (h, h))
    hd = ['N_Mov', 'Fecha', 'Tipo', 'Codigo_BEX', 'Afiliador', 'Supervisor', 'Regional', 'Observaciones', 'Registrado_Por']
    for i, c in enumerate(hd, 1):
        sh.set(h + 1, i, c, 's', S_HEAD)
    sh.set(h + 2, 1, 'IFERROR(FILTER(%s,%s[Codigo_BEX]=$C$3),"Sin movimientos para este codigo / escriba un BEX en C3")' % (TMO, TMO), 'fa', S_BORD)
    return sh


def build_dashboard():
    sh = Sheet('DASHBOARD')
    sh.cols = [(1, 1, 30), (2, 2, 12), (3, 3, 12), (4, 4, 12), (5, 5, 12), (6, 6, 10),
               (7, 7, 10), (8, 8, 3), (9, 9, 32), (10, 10, 13), (11, 11, 11),
               (12, 12, 9), (13, 13, 10), (14, 14, 10)]
    sh.setc('A1', 'DASHBOARD EJECUTIVO  -  Control de Credenciales BEX  (actualizacion automatica)', 's', S_TITLE)
    sh.merges.append('A1:N1')
    sh.setc('A3', 'INDICADORES CLAVE DE CREDENCIALES', 's', S_SECTION); sh.merges.append('A3:N3')
    # KPIs credenciales
    kpis = [
        ('A', 'TOTAL CREDENCIALES', 'COUNTA(%s[Codigo_Activo])' % TIG),
        ('D', 'ASIGNADAS', 'COUNTIF(%s[Estado_Actual],"ASIGNADO")' % TIG),
        ('G', 'DISPONIBLES', 'COUNTIF(%s[Estado_Actual],"DISPONIBLE")' % TIG),
        ('J', 'POR ASIGNAR', 'COUNTIF(%s[Estado_Actual],"POR ASIGNAR")' % TIG),
        ('M', 'PERDIDAS', 'COUNTIF(%s[Estado_Actual],"PERDIDO")' % TIG),
    ]
    for colL, lbl, f in kpis:
        sh.setc('%s4' % colL, lbl, 's', S_KPILBL)
        sh.set(4, _ci(colL) + 1 - 1, lbl, 's', S_KPILBL)  # ensure
    for colL, lbl, f in kpis:
        c = _ci(colL)
        sh.merges.append('%s4:%s4' % (colL, col_letter(c + 1)))
        sh.set(5, c, f, 'f', S_KPIBOX)
        sh.merges.append('%s5:%s5' % (colL, col_letter(c + 1)))

    sh.setc('A7', 'INDICADORES DE PERSONAL', 's', S_SECTION); sh.merges.append('A7:N7')
    kpis2 = [
        ('A', 'TOTAL PERSONAL', 'COUNTA(%s[ID_Persona])' % TPE),
        ('D', 'ACTIVOS', 'COUNTIF(%s[Estado],"ACTIVO")' % TPE),
        ('G', 'INACTIVOS', 'COUNTIF(%s[Estado],"INACTIVO")' % TPE),
        ('J', 'SUPERVISORES', 'COUNTIF(%s[Rol],"SUPERVISOR")' % TPE),
        ('M', 'ALERTAS', 'COUNTIF(%s[Estado_Actual],"PERDIDO")+COUNTIF(%s[Observaciones],"*DUPLICADO*")' % (TIG, TIG)),
    ]
    for colL, lbl, f in kpis2:
        c = _ci(colL)
        sh.set(8, c, lbl, 's', S_KPILBL)
        sh.merges.append('%s8:%s8' % (colL, col_letter(c + 1)))
        sh.set(9, c, f, 'f', S_KPIBOX)
        sh.merges.append('%s9:%s9' % (colL, col_letter(c + 1)))

    # Tabla por regional
    sh.setc('A11', 'CREDENCIALES POR REGIONAL', 's', S_SECTION); sh.merges.append('A11:G11')
    rh = ['REGIONAL', 'ASIGNADO', 'DISPONIBLE', 'POR ASIGNAR', 'PERDIDO', 'TOTAL', '% ASIG.']
    for i, c in enumerate(rh, 1):
        sh.set(12, i, c, 's', S_HEAD)
    rstart = 13
    for k, reg in enumerate(regionales):
        rr = rstart + k
        sh.set(rr, 1, reg, 's', S_BORD)
        sh.set(rr, 2, 'COUNTIFS(%s[Regional],$A%d,%s[Estado_Actual],"ASIGNADO")' % (TIG, rr, TIG), 'f', S_CEN)
        sh.set(rr, 3, 'COUNTIFS(%s[Regional],$A%d,%s[Estado_Actual],"DISPONIBLE")' % (TIG, rr, TIG), 'f', S_CEN)
        sh.set(rr, 4, 'COUNTIFS(%s[Regional],$A%d,%s[Estado_Actual],"POR ASIGNAR")' % (TIG, rr, TIG), 'f', S_CEN)
        sh.set(rr, 5, 'COUNTIFS(%s[Regional],$A%d,%s[Estado_Actual],"PERDIDO")' % (TIG, rr, TIG), 'f', S_CEN)
        sh.set(rr, 6, 'COUNTIF(%s[Regional],$A%d)' % (TIG, rr), 'f', S_CEN)
        sh.set(rr, 7, 'IFERROR(B%d/F%d,0)' % (rr, rr), 'f', S_PCT)
    trow = rstart + len(regionales)
    sh.set(trow, 1, 'TOTAL GENERAL', 's', S_SUB)
    for col in range(2, 7):
        L = col_letter(col)
        sh.set(trow, col, 'SUM(%s%d:%s%d)' % (L, rstart, L, trow - 1), 'f', S_CEN)
    sh.set(trow, 7, 'IFERROR(B%d/F%d,0)' % (trow, trow), 'f', S_PCT)

    # Tabla por supervisor
    sh.setc('I11', 'CREDENCIALES POR SUPERVISOR', 's', S_SECTION); sh.merges.append('I11:N11')
    sh2 = ['SUPERVISOR', 'REGIONAL', 'ASIGNADO', 'TOTAL', 'INACTIVOS', '% ACTIV.']
    for i, c in enumerate(sh2, 9):
        sh.set(12, i, c, 's', S_HEAD)
    for k, (sup, reg) in enumerate(supervisores):
        rr = 13 + k
        sh.set(rr, 9, sup, 's', S_BORD)
        sh.set(rr, 10, reg, 's', S_CEN)
        sh.set(rr, 11, 'COUNTIFS(%s[Supervisor],$I%d,%s[Estado_Actual],"ASIGNADO")' % (TIG, rr, TIG), 'f', S_CEN)
        sh.set(rr, 12, 'COUNTIF(%s[Supervisor],$I%d)' % (TIG, rr), 'f', S_CEN)
        sh.set(rr, 13, 'COUNTIFS(%s[Supervisor],$I%d,%s[Estado],"INACTIVO")' % (TPE, rr, TPE), 'f', S_CEN)
        sh.set(rr, 14, 'IFERROR(K%d/L%d,0)' % (rr, rr), 'f', S_PCT)

    # Ultimos movimientos (dinamico: ultimos 10)
    mrow = 13 + len(supervisores) + 2
    sh.set(mrow, 1, 'ULTIMOS MOVIMIENTOS REGISTRADOS', 's', S_SECTION); sh.merges.append('A%d:G%d' % (mrow, mrow))
    mh = ['N_Mov', 'Fecha', 'Tipo', 'Codigo_BEX', 'Afiliador', 'Supervisor', 'Regional']
    for i, c in enumerate(mh, 1):
        sh.set(mrow + 1, i, c, 's', S_HEAD)
    mcols = ['N_Mov', 'Fecha', 'Tipo_Movimiento', 'Codigo_BEX', 'Afiliador', 'Supervisor', 'Regional']
    for k in range(10):
        rr = mrow + 2 + k
        off = k - 9  # -9..0  => ultimos 10
        for i, mc in enumerate(mcols, 1):
            st = S_DATE if mc == 'Fecha' else S_BORD
            f = ('IFERROR(INDEX(%s[%s],COUNTA(%s[N_Mov])%+d),"")' % (TMO, mc, TMO, off))
            sh.set(rr, i, f, 'f', st)
    # Bloque auxiliar para grafico de torta por estado (col P/Q = 16/17)
    sh.set(2, 16, 'DISTRIBUCION POR ESTADO', 's', S_SUB)
    estados_kpi = ['ASIGNADO', 'DISPONIBLE', 'POR ASIGNAR', 'PERDIDO']
    for k, e in enumerate(estados_kpi):
        rr = 3 + k
        sh.set(rr, 16, e, 's', S_BORD)
        sh.set(rr, 17, 'COUNTIF(%s[Estado_Actual],"%s")' % (TIG, e), 'f', S_CEN)
    sh.freeze = (1, 0)
    return sh

def _ci(letter):
    c = 0
    for ch in letter:
        c = c * 26 + (ord(ch) - 64)
    return c


# ===========================================================================
# GRAFICOS (charts + drawing)
# ===========================================================================
def _pts(values, numeric=True):
    out = []
    for i, v in enumerate(values):
        out.append('<c:pt idx="%d"><c:v>%s</c:v></c:pt>' % (i, xesc(v)))
    return '<c:ptCount val="%d"/>%s' % (len(values), ''.join(out))

def bar_chart_xml(title, cat_ref, cats, val_ref, vals, color='4472C4'):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<c:chart>'
        '<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r>'
        '<a:rPr lang="es" b="1"/><a:t>%s</a:t></a:r></a:p></c:rich></c:tx>'
        '<c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>'
        '<c:plotArea><c:layout/>'
        '<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>'
        '<c:ser><c:idx val="0"/><c:order val="0"/>'
        '<c:spPr><a:solidFill><a:srgbClr val="%s"/></a:solidFill></c:spPr>'
        '<c:cat><c:strRef><c:f>%s</c:f><c:strCache>%s</c:strCache></c:strRef></c:cat>'
        '<c:val><c:numRef><c:f>%s</c:f><c:numCache><c:formatCode>General</c:formatCode>%s</c:numCache></c:numRef></c:val>'
        '</c:ser>'
        '<c:axId val="111"/><c:axId val="222"/></c:barChart>'
        '<c:catAx><c:axId val="111"/><c:scaling><c:orientation val="minMax"/></c:scaling>'
        '<c:delete val="0"/><c:axPos val="b"/><c:crossAx val="222"/></c:catAx>'
        '<c:valAx><c:axId val="222"/><c:scaling><c:orientation val="minMax"/></c:scaling>'
        '<c:delete val="0"/><c:axPos val="l"/><c:crossAx val="111"/></c:valAx>'
        '</c:plotArea><c:plotVisOnly val="1"/></c:chart></c:chartSpace>'
        % (xesc(title), color, xesc(cat_ref), _pts(cats), xesc(val_ref), _pts(vals)))

def pie_chart_xml(title, cat_ref, cats, val_ref, vals):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<c:chart>'
        '<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r>'
        '<a:rPr lang="es" b="1"/><a:t>%s</a:t></a:r></a:p></c:rich></c:tx>'
        '<c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>'
        '<c:plotArea><c:layout/>'
        '<c:pieChart><c:varyColors val="1"/>'
        '<c:ser><c:idx val="0"/><c:order val="0"/>'
        '<c:dLbls><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/>'
        '<c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>'
        '<c:cat><c:strRef><c:f>%s</c:f><c:strCache>%s</c:strCache></c:strRef></c:cat>'
        '<c:val><c:numRef><c:f>%s</c:f><c:numCache><c:formatCode>General</c:formatCode>%s</c:numCache></c:numRef></c:val>'
        '</c:ser><c:firstSliceAng val="0"/></c:pieChart>'
        '</c:plotArea><c:legend><c:legendPos val="r"/><c:overlay val="0"/></c:legend>'
        '<c:plotVisOnly val="1"/></c:chart></c:chartSpace>'
        % (xesc(title), xesc(cat_ref), _pts(cats), xesc(val_ref), _pts(vals)))

def anchor(c1, r1, c2, r2, rid, cid, name):
    return ('<xdr:twoCellAnchor>'
        '<xdr:from><xdr:col>%d</xdr:col><xdr:colOff>0</xdr:colOff>'
        '<xdr:row>%d</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>'
        '<xdr:to><xdr:col>%d</xdr:col><xdr:colOff>0</xdr:colOff>'
        '<xdr:row>%d</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>'
        '<xdr:graphicFrame macro="">'
        '<xdr:nvGraphicFramePr><xdr:cNvPr id="%d" name="%s"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>'
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
    dash = build_dashboard()
    busc = build_buscador()
    inv = build_inventario()
    per = build_personal()
    mov = build_movimientos()
    par = build_parametros()
    sheets = [dash, busc, inv, per, mov, par]  # orden de pestañas
    dash.drawing_rid = 'rId1'  # graficos en el dashboard

    CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
    RELNS = 'http://schemas.openxmlformats.org/package/2006/relationships'
    DOC = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

    z = zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED)

    # [Content_Types].xml
    ct = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>']
    ct.append('<Types xmlns="%s">' % CT)
    ct.append('<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>')
    ct.append('<Default Extension="xml" ContentType="application/xml"/>')
    ct.append('<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>')
    ct.append('<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>')
    ct.append('<Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/>')
    for i in range(1, len(sheets) + 1):
        ct.append('<Override PartName="/xl/worksheets/sheet%d.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' % i)
    for s in sheets:
        if s.table:
            ct.append('<Override PartName="/xl/tables/table%d.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>' % s.table['tid'])
    ct.append('<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>')
    for ci in (1, 2, 3):
        ct.append('<Override PartName="/xl/charts/chart%d.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>' % ci)
    ct.append('</Types>')
    z.writestr('[Content_Types].xml', ''.join(ct))

    # _rels/.rels
    z.writestr('_rels/.rels',
               '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
               '<Relationships xmlns="%s">'
               '<Relationship Id="rId1" Type="%s/officeDocument" Target="xl/workbook.xml"/>'
               '</Relationships>' % (RELNS, DOC))

    # xl/workbook.xml
    wb = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>']
    wb.append('<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
              'xmlns:r="%s">' % DOC)
    wb.append('<sheets>')
    for i, s in enumerate(sheets, 1):
        wb.append('<sheet name="%s" sheetId="%d" r:id="rId%d"/>' % (xesc(s.name), i, i))
    wb.append('</sheets>')
    wb.append('<calcPr calcId="0" fullCalcOnLoad="1"/>')
    wb.append('</workbook>')
    z.writestr('xl/workbook.xml', ''.join(wb))

    # xl/_rels/workbook.xml.rels
    wr = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>']
    wr.append('<Relationships xmlns="%s">' % RELNS)
    for i in range(1, len(sheets) + 1):
        wr.append('<Relationship Id="rId%d" Type="%s/worksheet" Target="worksheets/sheet%d.xml"/>' % (i, DOC, i))
    wr.append('<Relationship Id="rId%d" Type="%s/styles" Target="styles.xml"/>' % (len(sheets) + 1, DOC))
    wr.append('<Relationship Id="rId%d" Type="%s/sheetMetadata" Target="metadata.xml"/>' % (len(sheets) + 2, DOC))
    wr.append('</Relationships>')
    z.writestr('xl/_rels/workbook.xml.rels', ''.join(wr))

    # styles
    z.writestr('xl/styles.xml', STYLES_XML)

    # metadata.xml (propiedades de arrays dinamicos para FILTER/XLOOKUP con desbordamiento)
    z.writestr('xl/metadata.xml',
               '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
               '<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
               'xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" '
               'xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">'
               '<metadataTypes count="1">'
               '<metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" '
               'pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" '
               'clearComments="1" assign="1" coerce="1" cellMeta="1"/>'
               '</metadataTypes>'
               '<futureMetadata name="XLDAPR" count="1"><bk><extLst>'
               '<ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}">'
               '<xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext>'
               '</extLst></bk></futureMetadata>'
               '<cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata>'
               '</metadata>')

    # worksheets + tables + sheet rels
    for i, s in enumerate(sheets, 1):
        rid_table = 'rId1' if s.table else None
        z.writestr('xl/worksheets/sheet%d.xml' % i, s.xml(rid_table))
        if s.table:
            t = s.table
            z.writestr('xl/worksheets/_rels/sheet%d.xml.rels' % i,
                       '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                       '<Relationships xmlns="%s">'
                       '<Relationship Id="rId1" Type="%s/table" Target="../tables/table%d.xml"/>'
                       '</Relationships>' % (RELNS, DOC, t['tid']))
            z.writestr('xl/tables/table%d.xml' % t['tid'],
                       table_xml(t['tid'], t['name'], t['ref'], t['columns']))

    # ---- Graficos del dashboard (sheet1) ----
    est = ['ASIGNADO', 'DISPONIBLE', 'POR ASIGNAR', 'PERDIDO']
    est_vals = [sum(1 for it in inventario if it['estado'] == e) for e in est]
    reg_vals = [sum(1 for it in inventario if it['reg'] == r) for r in regionales]
    sup_names = [s for s, _ in supervisores]
    sup_vals = [sum(1 for it in inventario if it['sup'] == s and it['estado'] == 'ASIGNADO') for s in sup_names]

    z.writestr('xl/charts/chart1.xml',
               pie_chart_xml('Distribucion por Estado', 'DASHBOARD!$P$3:$P$6', est,
                             'DASHBOARD!$Q$3:$Q$6', est_vals))
    z.writestr('xl/charts/chart2.xml',
               bar_chart_xml('Credenciales por Regional', 'DASHBOARD!$A$13:$A$18', regionales,
                             'DASHBOARD!$F$13:$F$18', reg_vals, color='2E86C1'))
    sup_last = 12 + len(supervisores)
    z.writestr('xl/charts/chart3.xml',
               bar_chart_xml('Asignadas por Supervisor', 'DASHBOARD!$I$13:$I$%d' % sup_last, sup_names,
                             'DASHBOARD!$K$13:$K$%d' % sup_last, sup_vals, color='27AE60'))

    # drawing con 3 anclas (debajo de las tablas)
    dr = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>']
    dr.append('<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
              'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">')
    base = 12 + len(supervisores) + 14   # debajo de ultimos movimientos
    dr.append(anchor(0, base, 6, base + 16, 'rId1', 2, 'Grafico Estado'))
    dr.append(anchor(7, base, 13, base + 16, 'rId2', 3, 'Grafico Regional'))
    dr.append(anchor(0, base + 17, 13, base + 33, 'rId3', 4, 'Grafico Supervisor'))
    dr.append('</xdr:wsDr>')
    z.writestr('xl/drawings/drawing1.xml', ''.join(dr))
    z.writestr('xl/drawings/_rels/drawing1.xml.rels',
               '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
               '<Relationships xmlns="%s">'
               '<Relationship Id="rId1" Type="%s/chart" Target="../charts/chart1.xml"/>'
               '<Relationship Id="rId2" Type="%s/chart" Target="../charts/chart2.xml"/>'
               '<Relationship Id="rId3" Type="%s/chart" Target="../charts/chart3.xml"/>'
               '</Relationships>' % (RELNS, DOC, DOC, DOC))
    # rels de sheet1 (dashboard) -> drawing
    z.writestr('xl/worksheets/_rels/sheet1.xml.rels',
               '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
               '<Relationships xmlns="%s">'
               '<Relationship Id="rId1" Type="%s/drawing" Target="../drawings/drawing1.xml"/>'
               '</Relationships>' % (RELNS, DOC))
    z.close()
    print('Generado:', path)

package('SISTEMA_BEX_CREDENCIALES.xlsx')
