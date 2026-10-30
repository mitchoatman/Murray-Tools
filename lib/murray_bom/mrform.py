# -*- coding: utf-8 -*-
"""Fills the Murray Material Requisition (MR) form (templates/MR-BOM.xlsx) with export rows.
Port of mrForm.js: same cells, same styles, same header fields. Standard library only (zipfile, re)."""
import re
import zipfile
import io
import datetime

COLUMNS = ['Qty', 'Length', 'Size', 'Description', 'Material', 'Unit Cost', 'Total']
LETTERS = ['A', 'B', 'C', 'D', 'E', 'F', 'G']
FIRST_ROW = 17          # row 16 holds the column headers
TEMPLATE_LAST = 228     # the form carries empty bordered rows down to here
SHRINK_STYLE = '46'     # left aligned, shrink to fit (used for a long Job Name); present in the template
LEVELS_STYLE = '47'     # A15 style with shrink to fit


def _esc(s):
    return str(s).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;')


def _cell_xml(ref, style, value):
    s = ' s="%s"' % style if style else ''
    if value == '' or value is None:
        return '<c r="%s"%s/>' % (ref, s)
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return '<c r="%s"%s><v>%s</v></c>' % (ref, s, repr(value) if isinstance(value, float) else value)
    return '<c r="%s"%s t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>' % (ref, s, _esc(value))


def _date_serial(d):
    return (datetime.date(d.year, d.month, d.day) - datetime.date(1899, 12, 30)).days


def _set_cell(sheet, ref, value, style=None):
    rx = re.compile(r'<c r="%s"([^>]*?)(?:/>|>.*?</c>)' % ref)
    m = rx.search(sheet)
    if not m:
        return sheet
    sm = re.search(r' s="(\d+)"', m.group(1))
    new = _cell_xml(ref, style or (sm.group(1) if sm else ''), value)
    return sheet[:m.start()] + new + sheet[m.end():]


def job_from_project(project_name):
    """'23456 - LAWA TBIT', 'LAWA TBIT (23456)' or '23456_LAWA TBIT' -> job number and job name.
    The first standalone 5-digit code is the job number; no code means no job number."""
    name = ('' if project_name is None else str(project_name)).strip()
    m = re.search(r'(^|\D)(\d{5})(?!\d)', name)
    if not m:
        return {'jobName': name, 'jobNumber': ''}
    num = m.group(2)
    sep = r'[\s\-\u2013\u2014_:.,|/]*'
    rx = re.compile(sep + r'[(\[]?' + num + r'[)\]]?' + sep)
    rest = rx.sub(' ', name, count=1)
    rest = re.sub(r'\s+', ' ', rest)
    rest = re.sub(r'^[\s\-\u2013\u2014_:.,|/]+|[\s\-\u2013\u2014_:.,|/]+$', '', rest)
    return {'jobName': rest, 'jobNumber': num}


def build(template_path, rows, info=None):
    """Returns the finished .xlsx as bytes.
    info: {'date': date, 'jobName': text, 'jobNumber': text, 'levels': [names], 'exportName': text,
           'selectionCount': n, 'rodRows': [rows]}"""
    info = info or {}
    with zipfile.ZipFile(template_path) as zin:
        items = [(i, zin.read(i.filename)) for i in zin.infolist()]
    files = {i.filename: data for i, data in items}
    if 'xl/worksheets/sheet1.xml' not in files or 'xl/workbook.xml' not in files:
        raise ValueError('The MR form template is missing its sheet.')
    sheet = files['xl/worksheets/sheet1.xml'].decode('utf-8')
    book = files['xl/workbook.xml'].decode('utf-8')

    row_start = sheet.find('<row r="%d"' % FIRST_ROW)
    data_end = sheet.find('</sheetData>')
    if row_start < 0 or data_end < 0:
        raise ValueError('The MR form template has no data rows.')
    first_row = sheet[row_start:sheet.index('</row>', row_start)]
    style = {}
    for L in LETTERS:
        m = re.search(r'<c r="%s%d"[^>]*? s="(\d+)"' % (L, FIRST_ROW), first_row)
        style[L] = m.group(1) if m else ''

    all_rows = list(rows)
    if info.get('rodRows'):
        all_rows.append({})
        all_rows.extend(info['rodRows'])
    last = max(TEMPLATE_LAST, FIRST_ROW - 1 + len(all_rows))
    xml = []
    for r in range(FIRST_ROW, last + 1):
        idx = r - FIRST_ROW
        item = all_rows[idx] if idx < len(all_rows) else None
        cells = ''
        for c in range(5):      # A to E, as in the form; F and G keep the column default
            v = item.get(COLUMNS[c], '') if item else ''
            cells += _cell_xml('%s%d' % (LETTERS[c], r), style[LETTERS[c]], v)
        xml.append('<row r="%d" spans="1:7" x14ac:dyDescent="0.2">%s</row>' % (r, cells))
    sheet = sheet[:row_start] + ''.join(xml) + sheet[data_end:]

    if info.get('date'):
        sheet = _set_cell(sheet, 'F7', _date_serial(info['date']))
    if info.get('jobName'):
        sheet = _set_cell(sheet, 'F8', info['jobName'], SHRINK_STYLE)
    if info.get('jobNumber'):
        sheet = _set_cell(sheet, 'F9', info['jobNumber'])
    line = []
    if info.get('exportName'):
        line.append(info['exportName'])
    lv = info.get('levels') or []
    if lv:
        line.append(('Levels: ' if len(lv) > 1 else 'Level: ') + ', '.join(lv))
    n = info.get('selectionCount') or 0
    if n:
        line.append('Selection (%s %s)' % (format(n, ','), 'part' if n == 1 else 'parts'))
    ex = info.get('servicesExcluded') or []
    if ex:
        names = ', '.join(sorted(ex))
        line.append(('Services excluded: ' + names) if len(ex) <= 3 and len(names) <= 60 else '%d services excluded' % len(ex))
    if line:
        sheet = _set_cell(sheet, 'A15', ' | '.join(line), LEVELS_STYLE)
    sheet = re.sub(r'<dimension ref="[^"]*"/>', '<dimension ref="A1:G%d"/>' % last, sheet, count=1)

    last_data = max(FIRST_ROW, FIRST_ROW - 1 + len(all_rows))
    book = re.sub(r'(Copper_Pipe_And_Fitting_List_Total_BOM"[^>]*>BOM!\$A\$16:\$G\$)\d+',
                  lambda m: m.group(1) + str(last_data), book, count=1)

    files['xl/worksheets/sheet1.xml'] = sheet.encode('utf-8')
    files['xl/workbook.xml'] = book.encode('utf-8')
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zout:
        for i, _ in items:
            zout.writestr(i.filename, files[i.filename])
    return buf.getvalue()


def rod_rows(rods):
    """The hanger rod totals as extra form rows (written under the hanger rows after one blank row)."""
    if not rods:
        return []
    out = [{'Description': '*Values not accurate if user has not synced FP data*'}]
    for l in rods:
        out.append({'Length': l['length'], 'Size': l['size'], 'Description': l['description']})
    return out
