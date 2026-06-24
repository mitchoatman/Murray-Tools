# -*- coding: UTF-8 -*-
#! python3

import zipfile
import xml.etree.ElementTree as ET
import traceback

import clr
clr.AddReference('System')
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    ViewSchedule,
    ScheduleFilter,
    ScheduleFilterType,
    ScheduleHorizontalAlignment,
    TableCellStyle,
    TableCellStyleOverrideOptions,
    TableMergedCell,
    SectionType,
    ElementId,
    BuiltInCategory,
    BuiltInParameter,
    GraphicsStyleType,
    VerticalAlignmentStyle,
    Color,
    Transaction,
)

from pyrevit import forms


# ==============================================================
# XLSX namespaces
# ==============================================================
_NS_SS      = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
_NS_DOC_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

def _ss_tag(local):
    return '{%s}%s' % (_NS_SS, local)

def _col_letter_to_index(letters):
    idx = 0
    for ch in letters.upper():
        idx = idx * 26 + (ord(ch) - ord('A') + 1)
    return idx - 1

def _cell_ref_to_rowcol(ref):
    letters = ''.join(ch for ch in ref if ch.isalpha())
    digits = ''.join(ch for ch in ref if ch.isdigit())
    col_idx = _col_letter_to_index(letters) if letters else 0
    row_idx = (int(digits) - 1) if digits else 0
    return row_idx, col_idx


# ==============================================================
# Constants
# ==============================================================
MAX_COLS = 50
MAX_ROWS = 200

DEFAULT_COL_WIDTH_FT = 0.7
DEFAULT_ROW_HEIGHT_FT = 0.18
EXCEL_DEFAULT_CHAR_WIDTH = 8.43
EXCEL_DEFAULT_ROW_HEIGHT_PT = 15.0

FILTER_SENTINEL = '__TABLEGEN_NO_MATCH__'

DEFAULT_INDEXED_COLORS = [
    '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
    '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
    '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080',
    '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF',
    '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF',
    '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99',
    '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696',
    '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333'
]

_EMPTY_BORDER = {'left': None, 'right': None, 'top': None, 'bottom': None}
_EMPTY_ALIGNMENT = {'horizontal': None, 'vertical': None, 'wrapText': False}
_EMPTY_FONT = {'size_pt': 11.0, 'bold': False, 'italic': False}
_EMPTY_FILL = {'patternType': None, 'rgb': None}


# ==============================================================
# Safe helpers
# ==============================================================
def safe_str(v):
    return '' if v is None else str(v)

def get_or_default(seq, idx, default=None):
    try:
        return seq[idx]
    except:
        return default

def sanitize_schedule_name(name):
    bad = ['\\', '/', ':', '{', '}', '[', ']', '|', ';', '<', '>', '?', '`', '~']
    out = safe_str(name).strip()
    if not out:
        out = 'Excel Schedule'
    for ch in bad:
        out = out.replace(ch, '_')
    return out

def index_to_col_letter(idx):
    s = ''
    n = idx + 1
    while n > 0:
        n, rem = divmod(n - 1, 26)
        s = chr(ord('A') + rem) + s
    return s

def rect_to_a1(rect):
    if rect is None:
        return '<none>'
    min_r, min_c, max_r, max_c = rect
    a = '{}{}'.format(index_to_col_letter(min_c), min_r + 1)
    b = '{}{}'.format(index_to_col_letter(max_c), max_r + 1)
    return '{}:{}'.format(a, b)


# ==============================================================
# Color helpers
# ==============================================================
def _hex_to_rgb_tuple(hex6):
    if not hex6 or len(hex6) != 6:
        return None
    try:
        return (int(hex6[0:2], 16), int(hex6[2:4], 16), int(hex6[4:6], 16))
    except:
        return None

def _clamp_u8(v):
    return max(0, min(255, int(round(v))))

def _apply_tint_to_rgb(rgb, tint):
    if rgb is None:
        return None
    if tint is None:
        return '{:02X}{:02X}{:02X}'.format(rgb[0], rgb[1], rgb[2])

    try:
        tint = float(tint)
    except:
        return '{:02X}{:02X}{:02X}'.format(rgb[0], rgb[1], rgb[2])

    out = []
    for c in rgb:
        if tint < 0:
            c2 = c * (1.0 + tint)
        else:
            c2 = c + (255.0 - c) * tint
        out.append(_clamp_u8(c2))
    return '{:02X}{:02X}{:02X}'.format(out[0], out[1], out[2])

def rgb_string_to_color(rgb):
    if not rgb or len(rgb) != 6:
        return None
    try:
        return Color(int(rgb[0:2], 16), int(rgb[2:4], 16), int(rgb[4:6], 16))
    except:
        return None

def read_theme_colors(zf, zip_names):
    if 'xl/theme/theme1.xml' not in zip_names:
        return []

    try:
        tree = ET.fromstring(zf.read('xl/theme/theme1.xml'))
    except:
        return []

    ns_a = 'http://schemas.openxmlformats.org/drawingml/2006/main'

    def a_tag(local):
        return '{%s}%s' % (ns_a, local)

    scheme = tree.find('.//' + a_tag('clrScheme'))
    if scheme is None:
        return []

    theme_colors = []
    for child in list(scheme):
        rgb = None
        srgb = child.find(a_tag('srgbClr'))
        if srgb is not None:
            rgb = srgb.get('val')
        sysc = child.find(a_tag('sysClr'))
        if rgb is None and sysc is not None:
            rgb = sysc.get('lastClr')
        if rgb:
            rgb = rgb.strip().upper()
            if len(rgb) == 8:
                rgb = rgb[2:]
        theme_colors.append(rgb if rgb and len(rgb) == 6 else None)

    return theme_colors

def read_indexed_colors(styles_tree):
    indexed = list(DEFAULT_INDEXED_COLORS)

    colors_el = styles_tree.find(_ss_tag('colors'))
    if colors_el is None:
        return indexed

    indexed_el = colors_el.find(_ss_tag('indexedColors'))
    if indexed_el is None:
        return indexed

    vals = []
    for rgb_el in indexed_el.findall(_ss_tag('rgbColor')):
        rgb = rgb_el.get('rgb')
        if rgb:
            rgb = rgb.strip().upper()
            if len(rgb) == 8:
                rgb = rgb[2:]
            if len(rgb) == 6:
                vals.append(rgb)

    if vals:
        indexed = vals

    return indexed

def resolve_color_element(color_el, theme_colors, indexed_colors):
    if color_el is None:
        return None

    tint = color_el.get('tint')

    rgb = color_el.get('rgb')
    if rgb:
        rgb = rgb.strip().upper()
        if len(rgb) == 8:
            rgb = rgb[2:]
        if len(rgb) == 6:
            return _apply_tint_to_rgb(_hex_to_rgb_tuple(rgb), tint)

    theme = color_el.get('theme')
    if theme is not None:
        try:
            idx = int(theme)
            if 0 <= idx < len(theme_colors):
                base = theme_colors[idx]
                if base:
                    return _apply_tint_to_rgb(_hex_to_rgb_tuple(base), tint)
        except:
            pass

    indexed = color_el.get('indexed')
    if indexed is not None:
        try:
            idx = int(indexed)
            if 0 <= idx < len(indexed_colors):
                base = indexed_colors[idx]
                if base:
                    return _apply_tint_to_rgb(_hex_to_rgb_tuple(base), tint)
        except:
            pass

    return None


# ==============================================================
# XLSX style parsing
# ==============================================================
def read_xlsx_styles(zf, zip_names):
    if 'xl/styles.xml' not in zip_names:
        return (
            [_EMPTY_BORDER.copy()],
            [0],
            [_EMPTY_FONT.copy()],
            [0],
            [_EMPTY_ALIGNMENT.copy()],
            [_EMPTY_FILL.copy()],
            [0]
        )

    styles_tree = ET.fromstring(zf.read('xl/styles.xml'))
    theme_colors = read_theme_colors(zf, zip_names)
    indexed_colors = read_indexed_colors(styles_tree)

    border_defs = []
    borders_el = styles_tree.find(_ss_tag('borders'))
    if borders_el is not None:
        for border_el in borders_el.findall(_ss_tag('border')):
            bd = {}
            for edge in ('left', 'right', 'top', 'bottom'):
                edge_el = border_el.find(_ss_tag(edge))
                style = edge_el.get('style') if edge_el is not None else None
                bd[edge] = style
            border_defs.append(bd)
    if not border_defs:
        border_defs = [_EMPTY_BORDER.copy()]

    font_defs = []
    fonts_el = styles_tree.find(_ss_tag('fonts'))
    if fonts_el is not None:
        for font_el in fonts_el.findall(_ss_tag('font')):
            fd = _EMPTY_FONT.copy()
            sz_el = font_el.find(_ss_tag('sz'))
            if sz_el is not None:
                try:
                    fd['size_pt'] = float(sz_el.get('val', 11.0))
                except:
                    pass
            fd['bold'] = font_el.find(_ss_tag('b')) is not None
            fd['italic'] = font_el.find(_ss_tag('i')) is not None
            font_defs.append(fd)
    if not font_defs:
        font_defs = [_EMPTY_FONT.copy()]

    fill_defs = []
    fills_el = styles_tree.find(_ss_tag('fills'))
    if fills_el is not None:
        for fill_el in fills_el.findall(_ss_tag('fill')):
            fd = _EMPTY_FILL.copy()
            pattern_el = fill_el.find(_ss_tag('patternFill'))
            if pattern_el is not None:
                fd['patternType'] = pattern_el.get('patternType')
                fg_el = pattern_el.find(_ss_tag('fgColor'))
                bg_el = pattern_el.find(_ss_tag('bgColor'))
                fd['rgb'] = resolve_color_element(fg_el, theme_colors, indexed_colors)
                if not fd['rgb']:
                    fd['rgb'] = resolve_color_element(bg_el, theme_colors, indexed_colors)
            fill_defs.append(fd)
    if not fill_defs:
        fill_defs = [_EMPTY_FILL.copy()]

    xf_border_ids = []
    xf_font_ids = []
    xf_fill_ids = []
    xf_alignments = []

    cellxfs_el = styles_tree.find(_ss_tag('cellXfs'))
    if cellxfs_el is not None:
        for xf_el in cellxfs_el.findall(_ss_tag('xf')):
            try:
                xf_border_ids.append(int(xf_el.get('borderId', 0)))
            except:
                xf_border_ids.append(0)

            try:
                xf_font_ids.append(int(xf_el.get('fontId', 0)))
            except:
                xf_font_ids.append(0)

            try:
                xf_fill_ids.append(int(xf_el.get('fillId', 0)))
            except:
                xf_fill_ids.append(0)

            align_data = _EMPTY_ALIGNMENT.copy()
            align_el = xf_el.find(_ss_tag('alignment'))
            if align_el is not None:
                align_data['horizontal'] = align_el.get('horizontal')
                align_data['vertical'] = align_el.get('vertical')
                wrap_val = align_el.get('wrapText')
                align_data['wrapText'] = wrap_val in ('1', 'true', 'True')
            xf_alignments.append(align_data)

    if not xf_border_ids:
        xf_border_ids = [0]
    if not xf_font_ids:
        xf_font_ids = [0]
    if not xf_fill_ids:
        xf_fill_ids = [0]
    if not xf_alignments:
        xf_alignments = [_EMPTY_ALIGNMENT.copy()]

    return border_defs, xf_border_ids, font_defs, xf_font_ids, xf_alignments, fill_defs, xf_fill_ids

def get_border_for_style(style_idx, xf_border_ids, border_defs):
    try:
        return border_defs[xf_border_ids[style_idx]]
    except:
        return _EMPTY_BORDER

def get_font_def_for_style(style_idx, xf_font_ids, font_defs):
    try:
        return font_defs[xf_font_ids[style_idx]]
    except:
        return _EMPTY_FONT

def get_alignment_for_style(style_idx, xf_alignments):
    try:
        return xf_alignments[style_idx]
    except:
        return _EMPTY_ALIGNMENT

def get_fill_for_style(style_idx, xf_fill_ids, fill_defs):
    try:
        return fill_defs[xf_fill_ids[style_idx]]
    except:
        return _EMPTY_FILL


# ==============================================================
# XLSX reader
# ==============================================================
def read_xlsx(path):
    with zipfile.ZipFile(path, 'r') as zf:
        zip_names = set(zf.namelist())

        border_defs, xf_border_ids, font_defs, xf_font_ids, xf_alignments, fill_defs, xf_fill_ids = read_xlsx_styles(zf, zip_names)

        sst = []
        for sst_path in ('xl/sharedStrings.xml', 'xl/SharedStrings.xml'):
            if sst_path in zip_names:
                tree = ET.fromstring(zf.read(sst_path))
                for si in tree.findall(_ss_tag('si')):
                    text = ''.join(t.text or '' for t in si.iter(_ss_tag('t')))
                    sst.append(text)
                break

        wb_tree = ET.fromstring(zf.read('xl/workbook.xml'))
        sheet_nodes = wb_tree.findall('.//' + _ss_tag('sheet'))

        rels_tree = ET.fromstring(zf.read('xl/_rels/workbook.xml.rels'))
        rid_map = {}
        for rel in rels_tree:
            target = rel.get('Target', '')
            if target.startswith('/'):
                target = target.lstrip('/')
            else:
                target = 'xl/' + target
            rid_map[rel.get('Id')] = target

        sheet_names = []
        sheets = {}
        col_widths = {}
        sheet_styles = {}
        row_heights = {}
        default_row_heights = {}
        merge_ranges = {}

        for sn in sheet_nodes:
            name = sn.get('name')
            rid = sn.get('{%s}id' % _NS_DOC_REL)
            sheet_path = rid_map.get(rid, '')
            sheet_names.append(name)

            if not sheet_path or sheet_path not in zip_names:
                sheets[name] = []
                col_widths[name] = []
                sheet_styles[name] = []
                row_heights[name] = []
                default_row_heights[name] = EXCEL_DEFAULT_ROW_HEIGHT_PT
                merge_ranges[name] = []
                continue

            sheet_tree = ET.fromstring(zf.read(sheet_path))

            default_row_pt = EXCEL_DEFAULT_ROW_HEIGHT_PT
            sf = sheet_tree.find('.//' + _ss_tag('sheetFormatPr'))
            if sf is not None:
                try:
                    default_row_pt = float(sf.get('defaultRowHeight', EXCEL_DEFAULT_ROW_HEIGHT_PT))
                except:
                    pass
            default_row_heights[name] = default_row_pt

            merges = []
            for merge_el in sheet_tree.findall('.//' + _ss_tag('mergeCell')):
                ref = merge_el.get('ref', '')
                if not ref:
                    continue
                try:
                    if ':' in ref:
                        a, b = ref.split(':', 1)
                    else:
                        a, b = ref, ref
                    r1, c1 = _cell_ref_to_rowcol(a)
                    r2, c2 = _cell_ref_to_rowcol(b)
                    if r2 < r1:
                        r1, r2 = r2, r1
                    if c2 < c1:
                        c1, c2 = c2, c1
                    merges.append((r1, c1, r2, c2))
                except:
                    pass
            merge_ranges[name] = merges

            cw = []
            for col_el in sheet_tree.findall('.//' + _ss_tag('col')):
                mn = int(col_el.get('min', 1))
                mx = int(col_el.get('max', 1))
                width = float(col_el.get('width', 8.43))
                hidden = col_el.get('hidden')
                if hidden in ('1', 'true', 'True'):
                    width = 0.0
                for _ in range(mx - mn + 1):
                    cw.append(width)
            col_widths[name] = cw

            rows_data = []
            styles_data = []
            row_heights_data = []

            for row_el in sheet_tree.findall('.//' + _ss_tag('row')):
                row_idx = int(row_el.get('r', 1)) - 1

                while len(rows_data) <= row_idx:
                    rows_data.append({})
                while len(styles_data) <= row_idx:
                    styles_data.append({})
                while len(row_heights_data) <= row_idx:
                    row_heights_data.append(None)

                try:
                    if row_el.get('hidden') in ('1', 'true', 'True'):
                        row_heights_data[row_idx] = 0.0
                    else:
                        row_heights_data[row_idx] = float(row_el.get('ht')) if row_el.get('ht') is not None else None
                except:
                    row_heights_data[row_idx] = None

                for c_el in row_el.findall(_ss_tag('c')):
                    ref = c_el.get('r', '')
                    ctype = c_el.get('t', 'n')
                    style_idx = int(c_el.get('s', 0)) if c_el.get('s') is not None else 0

                    letters = ''.join(ch for ch in ref if ch.isalpha())
                    col_idx = _col_letter_to_index(letters) if letters else 0
                    styles_data[row_idx][col_idx] = style_idx

                    val = None
                    if ctype == 's':
                        v = c_el.find(_ss_tag('v'))
                        if v is not None and v.text is not None:
                            try:
                                val = sst[int(v.text)]
                            except:
                                val = v.text
                    elif ctype == 'inlineStr':
                        is_el = c_el.find(_ss_tag('is'))
                        if is_el is not None:
                            val = ''.join(t.text or '' for t in is_el.iter(_ss_tag('t')))
                    elif ctype == 'b':
                        v = c_el.find(_ss_tag('v'))
                        val = bool(int(v.text)) if (v is not None and v.text) else None
                    else:
                        v = c_el.find(_ss_tag('v'))
                        if v is not None and v.text is not None:
                            try:
                                val = int(v.text) if '.' not in v.text else float(v.text)
                            except:
                                val = v.text

                    if val is not None:
                        rows_data[row_idx][col_idx] = val

            dense_rows = []
            dense_styles = []
            row_total = max(len(rows_data), len(styles_data))
            for i in range(row_total):
                rd = rows_data[i] if i < len(rows_data) else {}
                sd = styles_data[i] if i < len(styles_data) else {}

                keys = []
                if rd:
                    keys.extend(rd.keys())
                if sd:
                    keys.extend(sd.keys())

                if not keys:
                    dense_rows.append([])
                    dense_styles.append([])
                    continue

                max_col = max(keys) + 1
                dense_rows.append([rd.get(c) for c in range(max_col)])
                dense_styles.append([sd.get(c, 0) for c in range(max_col)])

            while len(row_heights_data) < len(dense_rows):
                row_heights_data.append(None)

            sheets[name] = dense_rows
            sheet_styles[name] = dense_styles
            row_heights[name] = row_heights_data

    return (
        sheet_names,
        sheets,
        col_widths,
        sheet_styles,
        row_heights,
        default_row_heights,
        merge_ranges,
        border_defs,
        xf_border_ids,
        font_defs,
        xf_font_ids,
        xf_alignments,
        fill_defs,
        xf_fill_ids
    )


# ==============================================================
# Crop
# ==============================================================
def cell_has_meaningful_value(val):
    if val is None:
        return False
    try:
        s = str(val)
    except:
        return True
    return s.strip() != ''

def get_meaningful_rect(rows, merges, max_rows=MAX_ROWS, max_cols=MAX_COLS):
    used = []

    row_count = min(len(rows), max_rows)
    for r in range(row_count):
        row = rows[r] if r < len(rows) else []
        col_count = min(len(row), max_cols)
        for c in range(col_count):
            if cell_has_meaningful_value(row[c]):
                used.append((r, c))

    for mr in merges or []:
        try:
            r1, c1, r2, c2 = mr
            if r1 >= max_rows or c1 >= max_cols:
                continue
            r2 = min(r2, max_rows - 1)
            c2 = min(c2, max_cols - 1)
            used.append((r1, c1))
            used.append((r2, c2))
        except:
            pass

    if not used:
        return None

    min_r = min(x[0] for x in used)
    min_c = min(x[1] for x in used)
    max_r = max(x[0] for x in used)
    max_c = max(x[1] for x in used)
    return min_r, min_c, max_r, max_c

def crop_sheet_to_rect(rows, styles, row_heights, merges, rect):
    if rect is None:
        return [], [], [], []

    min_r, min_c, max_r, max_c = rect

    new_rows = []
    new_styles = []
    new_row_heights = []

    for r in range(min_r, max_r + 1):
        src_row = rows[r] if r < len(rows) else []
        src_sty = styles[r] if r < len(styles) else []

        new_rows.append([src_row[c] if c < len(src_row) else None for c in range(min_c, max_c + 1)])
        new_styles.append([src_sty[c] if c < len(src_sty) else 0 for c in range(min_c, max_c + 1)])
        new_row_heights.append(row_heights[r] if r < len(row_heights) else None)

    new_merges = []
    for mr in merges or []:
        try:
            r1, c1, r2, c2 = mr
            if r2 < min_r or r1 > max_r or c2 < min_c or c1 > max_c:
                continue

            nr1 = max(r1, min_r) - min_r
            nc1 = max(c1, min_c) - min_c
            nr2 = min(r2, max_r) - min_r
            nc2 = min(c2, max_c) - min_c

            if nr2 >= nr1 and nc2 >= nc1:
                new_merges.append((nr1, nc1, nr2, nc2))
        except:
            pass

    return new_rows, new_styles, new_row_heights, new_merges

def crop_col_widths_to_rect(col_widths_chars, rect):
    if rect is None:
        return []
    _, min_c, _, max_c = rect
    return [col_widths_chars[c] if c < len(col_widths_chars) else None for c in range(min_c, max_c + 1)]

def crop_to_meaningful_content(rows, styles, row_heights, merges, col_widths_chars):
    rect = get_meaningful_rect(rows, merges, MAX_ROWS, MAX_COLS)
    if rect is None:
        return [], [], [], [], [], None

    new_rows, new_styles, new_row_heights, new_merges = crop_sheet_to_rect(rows, styles, row_heights, merges, rect)
    new_col_widths = crop_col_widths_to_rect(col_widths_chars, rect)

    return new_rows, new_styles, new_row_heights, new_merges, new_col_widths, rect


# ==============================================================
# Geometry helpers
# ==============================================================
def get_sheet_bounds(rows, styles, merges):
    row_count = min(max(len(rows), len(styles)), MAX_ROWS)
    col_count = 0
    for i in range(row_count):
        col_count = max(col_count,
                        len(rows[i]) if i < len(rows) else 0,
                        len(styles[i]) if i < len(styles) else 0)

    for r1, c1, r2, c2 in merges:
        row_count = max(row_count, min(r2 + 1, MAX_ROWS))
        col_count = max(col_count, min(c2 + 1, MAX_COLS))

    return min(row_count, MAX_ROWS), min(col_count, MAX_COLS)

def col_width_to_ft(char_width):
    if char_width is None:
        return DEFAULT_COL_WIDTH_FT
    if float(char_width) <= 0.0:
        return 0.0
    return (float(char_width) / EXCEL_DEFAULT_CHAR_WIDTH) * DEFAULT_COL_WIDTH_FT

def row_height_pt_to_ft(row_pt, default_row_pt):
    base_pt = default_row_pt if default_row_pt else EXCEL_DEFAULT_ROW_HEIGHT_PT
    use_pt = row_pt if row_pt is not None else base_pt
    if float(use_pt) <= 0.0:
        return 0.0
    return (float(use_pt) / float(base_pt)) * DEFAULT_ROW_HEIGHT_FT


# ==============================================================
# Revit style helpers
# ==============================================================
def map_schedule_alignment(xl_h):
    s = (xl_h or '').strip().lower()
    if s in ('center', 'centercontinuous', 'distributed', 'justify', 'fill'):
        return ScheduleHorizontalAlignment.Center
    if s in ('right',):
        return ScheduleHorizontalAlignment.Right
    return ScheduleHorizontalAlignment.Left

def map_vertical_alignment(xl_v):
    s = (xl_v or '').strip().lower()
    if s in ('center', 'distributed', 'justify'):
        return VerticalAlignmentStyle.Middle
    if s in ('bottom',):
        return VerticalAlignmentStyle.Bottom
    return VerticalAlignmentStyle.Top

def excel_border_weight(style_name):
    if not style_name:
        return 0
    s = str(style_name).strip().lower()
    if s in ('thick', 'double'):
        return 3
    if 'medium' in s:
        return 2
    return 1

def _find_line_style(doc, preferred_names):
    try:
        lines_cat = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Lines)
        subcats = [sc for sc in lines_cat.SubCategories]
    except:
        return None

    by_name = {}
    for sc in subcats:
        try:
            by_name[sc.Name.strip().lower()] = sc
        except:
            pass

    for nm in preferred_names:
        sc = by_name.get(nm.strip().lower())
        if sc:
            try:
                return sc.GetGraphicsStyle(GraphicsStyleType.Projection)
            except:
                pass

    for sc in subcats:
        try:
            sc_name = sc.Name.strip().lower()
        except:
            continue
        for nm in preferred_names:
            token = nm.strip().lower().replace(' lines', '')
            if token and token in sc_name:
                try:
                    return sc.GetGraphicsStyle(GraphicsStyleType.Projection)
                except:
                    pass
    return None

def get_line_style_map(doc):
    thin = _find_line_style(doc, ['Thin Lines', 'Thin'])
    medium = _find_line_style(doc, ['Medium Lines', 'Medium'])
    wide = _find_line_style(doc, ['Wide Lines', 'Wide', 'Thick Lines', 'Thick'])

    if medium is None:
        medium = thin
    if wide is None:
        wide = medium or thin

    return {1: thin, 2: medium, 3: wide}

def build_cell_style(doc, font_def, fill_def, border_def, align_data, line_style_map):
    tcs = TableCellStyle()
    opts = TableCellStyleOverrideOptions()

    try:
        opts.FontSize = True
        tcs.TextSize = float(font_def.get('size_pt', 11.0))
    except:
        pass

    try:
        if bool(font_def.get('bold', False)):
            opts.Bold = True
            tcs.IsFontBold = True
    except:
        pass

    try:
        if bool(font_def.get('italic', False)):
            opts.Italics = True
            tcs.IsFontItalic = True
    except:
        pass

    try:
        rgb = fill_def.get('rgb') if fill_def else None
        color_obj = rgb_string_to_color(rgb)
        if color_obj is not None:
            opts.BackgroundColor = True
            tcs.BackgroundColor = color_obj
    except:
        pass

    try:
        if border_def:
            for side, prop_name, opt_name in [
                ('left', 'BorderLeftLineStyle', 'BorderLeftLineStyle'),
                ('right', 'BorderRightLineStyle', 'BorderRightLineStyle'),
                ('top', 'BorderTopLineStyle', 'BorderTopLineStyle'),
                ('bottom', 'BorderBottomLineStyle', 'BorderBottomLineStyle'),
            ]:
                weight = excel_border_weight(border_def.get(side))
                gs = line_style_map.get(weight)
                if gs is not None:
                    try:
                        setattr(opts, opt_name, True)
                        setattr(tcs, prop_name, gs.Id)
                    except:
                        pass
    except:
        pass

    try:
        setattr(opts, 'FontVerticalAlignment', True)
        tcs.FontVerticalAlignment = map_vertical_alignment(
            align_data.get('vertical') if align_data else None
        )
    except:
        pass

    try:
        tcs.SetCellStyleOverrideOptions(opts)
    except:
        pass

    return tcs

def try_set_header_cell_text(header, row, col, text):
    try:
        header.SetCellText(row, col, safe_str(text))
        return True
    except:
        return False

def try_set_header_cell_style(header, row, col, style):
    try:
        header.SetCellStyle(row, col, style)
        return True
    except:
        return False

def try_merge_header_cells(header, r1, c1, r2, c2):
    try:
        tm = TableMergedCell()
        tm.Top = r1
        tm.Left = c1
        tm.Bottom = r2
        tm.Right = c2
        header.MergeCells(tm)
        return True
    except:
        return False

def try_set_header_col_width(header, col_idx, width_ft):
    try:
        if width_ft > 0:
            header.SetColumnWidth(col_idx, width_ft)
            return True
    except:
        pass
    return False

def try_set_header_row_height(header, row_idx, height_ft):
    try:
        if height_ft > 0:
            header.SetRowHeight(row_idx, height_ft)
            return True
    except:
        pass
    return False


# ==============================================================
# Schedule creation
# ==============================================================
def delete_schedule_by_name(doc, name):
    for vs in FilteredElementCollector(doc).OfClass(ViewSchedule):
        try:
            if vs.Name == name:
                doc.Delete(vs.Id)
                return True
        except:
            pass
    return False

def find_schedulable_field_by_bip(definition, bip):
    target_pid = ElementId(bip)
    for sf in definition.GetSchedulableFields():
        try:
            if sf.ParameterId == target_pid:
                return sf
        except:
            pass
    return None

def create_schedule_with_fields(doc, schedule_name, col_count):
    bic = BuiltInCategory.OST_GenericModel
    vs = ViewSchedule.CreateSchedule(doc, ElementId(bic))
    vs.Name = schedule_name

    d = vs.Definition
    d.ShowTitle = False
    d.ShowHeaders = True

    try:
        d.IsItemized = True
    except:
        pass

    preferred_bips = [
        BuiltInParameter.ALL_MODEL_MARK,
        BuiltInParameter.ALL_MODEL_TYPE_MARK,
        BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS,
        BuiltInParameter.ALL_MODEL_TYPE_COMMENTS,
        BuiltInParameter.ALL_MODEL_DESCRIPTION,
        BuiltInParameter.ALL_MODEL_MANUFACTURER,
        BuiltInParameter.ALL_MODEL_MODEL,
        BuiltInParameter.UNIFORMAT_CODE,
        BuiltInParameter.KEYNOTE_PARAM,
        BuiltInParameter.ALL_MODEL_COST,
        BuiltInParameter.ALL_MODEL_URL,
        BuiltInParameter.ALL_MODEL_IMAGE,
    ]

    added_fields = []

    for bip in preferred_bips:
        if len(added_fields) >= col_count:
            break

        sf = find_schedulable_field_by_bip(d, bip)
        if sf is None:
            continue

        try:
            fld = d.AddField(sf)
            added_fields.append((sf, fld, bip))
        except:
            pass

    if len(added_fields) < col_count:
        doc.Delete(vs.Id)
        raise Exception('Not enough reusable schedulable fields for {} columns. Needed {}, got {}.'
                        .format(bic, col_count, len(added_fields)))

    # blank/minimal native headings
    for triple in added_fields[:col_count]:
        fld = triple[1]
        try:
            fld.ColumnHeading = ''
        except:
            pass
        try:
            fld.HorizontalAlignment = ScheduleHorizontalAlignment.Left
        except:
            pass

    filter_field = None
    for triple in added_fields:
        if triple[2] == BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS:
            filter_field = triple[1]
            break

    if filter_field is None:
        sf_comments = find_schedulable_field_by_bip(d, BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
        if sf_comments is not None:
            try:
                filter_field = d.AddField(sf_comments)
                try:
                    filter_field.IsHidden = True
                except:
                    pass
            except:
                filter_field = None

    if filter_field is None:
        doc.Delete(vs.Id)
        raise Exception('Could not create hidden Comments filter field.')

    try:
        filt = ScheduleFilter(filter_field.FieldId, ScheduleFilterType.Equal, FILTER_SENTINEL)
        d.InsertFilter(filt, 0)
    except Exception as ex:
        doc.Delete(vs.Id)
        raise Exception('Could not apply empty-body filter: {}'.format(ex))

    return vs, bic, added_fields[:col_count]


# ==============================================================
# Header build
# ==============================================================
def insert_rows_above_native_heading(header, count):
    inserted = 0
    base_row = header.FirstRowNumber
    for _ in range(count):
        try:
            if header.CanInsertRow(base_row):
                header.InsertRow(base_row)
                inserted += 1
                continue
        except:
            pass

        try:
            header.InsertRow(base_row)
            inserted += 1
        except:
            pass
    return inserted

def build_schedule_header_from_entire_sheet(doc, vs, sheet_rows, style_rows, row_heights_pt, default_row_pt,
                                            col_widths_ch, merge_ranges,
                                            xf_border_ids, border_defs, xf_font_ids, font_defs,
                                            xf_alignments, xf_fill_ids, fill_defs):
    row_count, col_count = get_sheet_bounds(sheet_rows, style_rows, merge_ranges)
    if row_count <= 0 or col_count <= 0:
        return {'inserted_rows': 0, 'merged': 0, 'styled': 0, 'texted': 0, 'colwidths': 0, 'rowheights': 0}

    header = vs.GetTableData().GetSectionData(SectionType.Header)
    line_style_map = get_line_style_map(doc)

    inserted_rows = insert_rows_above_native_heading(header, row_count)

    first_row = header.FirstRowNumber
    first_col = header.FirstColumnNumber

    styled = 0
    texted = 0
    merged = 0
    colwidths = 0
    rowheights = 0

    # set widths on actual header section
    for c in range(col_count):
        width_ft = col_width_to_ft(get_or_default(col_widths_ch, c, None))
        if try_set_header_col_width(header, first_col + c, width_ft):
            colwidths += 1

    # write/style/height custom header rows
    for r in range(row_count):
        rr = first_row + r
        h_ft = row_height_pt_to_ft(get_or_default(row_heights_pt, r, None), default_row_pt)
        if try_set_header_row_height(header, rr, h_ft):
            rowheights += 1

        row_vals = get_or_default(sheet_rows, r, [])
        sty_vals = get_or_default(style_rows, r, [])

        for c in range(col_count):
            cc = first_col + c
            val = get_or_default(row_vals, c, None)
            style_idx = get_or_default(sty_vals, c, 0)

            border_def = get_border_for_style(style_idx, xf_border_ids, border_defs)
            font_def = get_font_def_for_style(style_idx, xf_font_ids, font_defs)
            align_def = get_alignment_for_style(style_idx, xf_alignments)
            fill_def = get_fill_for_style(style_idx, xf_fill_ids, fill_defs)

            style = build_cell_style(doc, font_def, fill_def, border_def, align_def, line_style_map)

            if try_set_header_cell_style(header, rr, cc, style):
                styled += 1

            if val is not None:
                if try_set_header_cell_text(header, rr, cc, val):
                    texted += 1

    # merge custom rows
    for r1, c1, r2, c2 in merge_ranges:
        sr1 = first_row + r1
        sr2 = first_row + r2
        sc1 = first_col + c1
        sc2 = first_col + c2

        if (sr1 != sr2) or (sc1 != sc2):
            if try_merge_header_cells(header, sr1, sc1, sr2, sc2):
                merged += 1

    # shrink native bottom heading row if possible
    try:
        native_row = first_row + row_count
        try_set_header_row_height(header, native_row, 0.01)
        for c in range(col_count):
            try_set_header_cell_text(header, native_row, first_col + c, '')
    except:
        pass

    return {
        'inserted_rows': inserted_rows,
        'merged': merged,
        'styled': styled,
        'texted': texted,
        'colwidths': colwidths,
        'rowheights': rowheights
    }


# ==============================================================
# Main
# ==============================================================
def main():
    uidoc = __revit__.ActiveUIDocument
    doc = uidoc.Document

    xlsx_path = forms.pick_file(file_ext='xlsx', title='Select Excel File')
    if not xlsx_path:
        return

    try:
        (
            sheet_names,
            sheets,
            col_widths,
            sheet_styles,
            row_heights,
            default_row_heights,
            merge_ranges,
            border_defs,
            xf_border_ids,
            font_defs,
            xf_font_ids,
            xf_alignments,
            fill_defs,
            xf_fill_ids
        ) = read_xlsx(xlsx_path)
    except Exception as e:
        forms.alert('Could not read Excel file:\n{}'.format(str(e)), title='Error', exitscript=True)
        return

    if not sheet_names:
        forms.alert('No sheets found in workbook.', title='Error', exitscript=True)
        return

    if len(sheet_names) == 1:
        ws_name = sheet_names[0]
    else:
        ws_name = forms.SelectFromList.show(sheet_names, title='Select Worksheet', button_name='Import')
        if not ws_name:
            return

    sheet_rows = sheets[ws_name]
    style_rows = sheet_styles[ws_name]
    row_heights_pt = row_heights[ws_name]
    default_row_pt = default_row_heights[ws_name]
    sheet_merges = merge_ranges[ws_name]
    col_widths_ch = col_widths[ws_name]

    sheet_rows, style_rows, row_heights_pt, sheet_merges, col_widths_ch, meaningful_rect = crop_to_meaningful_content(
        sheet_rows, style_rows, row_heights_pt, sheet_merges, col_widths_ch
    )

    if not sheet_rows or not style_rows:
        forms.alert('Selected sheet has no meaningful text/merged content to import.',
                    title='Warning', exitscript=True)
        return

    row_count, col_count = get_sheet_bounds(sheet_rows, style_rows, sheet_merges)
    if row_count == 0 or col_count == 0:
        forms.alert('Selected sheet appears to be empty after cropping.',
                    title='Warning', exitscript=True)
        return

    schedule_name = forms.ask_for_string(
        default=sanitize_schedule_name('Excel - {}'.format(ws_name)),
        prompt='Name of new schedule:',
        title='Schedule Name'
    )
    if not schedule_name:
        return
    schedule_name = sanitize_schedule_name(schedule_name)

    t = Transaction(doc, 'TableGen - Import Excel to Schedule Header')

    try:
        t.Start()

        delete_schedule_by_name(doc, schedule_name)

        vs, bic, fields = create_schedule_with_fields(doc, schedule_name, col_count)

        stats = build_schedule_header_from_entire_sheet(
            doc,
            vs,
            sheet_rows,
            style_rows,
            row_heights_pt,
            default_row_pt,
            col_widths_ch,
            sheet_merges,
            xf_border_ids,
            border_defs,
            xf_font_ids,
            font_defs,
            xf_alignments,
            xf_fill_ids,
            fill_defs
        )

        t.Commit()

    except Exception:
        try:
            if t.HasStarted() and not t.HasEnded():
                t.RollBack()
        except:
            pass

        forms.alert('Import failed:\n\n{}'.format(traceback.format_exc()),
                    title='TableGen Error', exitscript=True)
        return

    forms.alert(
        'Done!\n\n'
        'Schedule: {}\n'
        'Worksheet: {}\n'
        'Imported source range: {}\n'
        'Category used: {}\n'
        'Columns created: {}\n'
        'Excel rows imported: {}\n'
        'Custom header rows inserted: {}\n'
        'Header text cells written: {}\n'
        'Header cells styled: {}\n'
        'Header merges created: {}\n'
        'Header column widths set: {}\n'
        'Header row heights set: {}\n\n'
        'This is the closest schedule-header approach for Revit 2021+; exact Excel parity may still vary by Revit version/view behavior.'.format(
            vs.Name,
            ws_name,
            rect_to_a1(meaningful_rect),
            bic,
            col_count,
            row_count,
            stats['inserted_rows'],
            stats['texted'],
            stats['styled'],
            stats['merged'],
            stats['colwidths'],
            stats['rowheights']
        ),
        title='TableGen Schedule Header Import'
    )

main()