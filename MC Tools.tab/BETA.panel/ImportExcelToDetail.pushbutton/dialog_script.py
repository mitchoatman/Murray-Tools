# -*- coding: UTF-8 -*-
#! python3
"""
TableGen Phase 3.5.1 - Import Excel to Current Revit Drafting View
(NO GROUPS: tagged replacement + ranked text-size mapping + row-heights
 + alignment + merged cells + hidden lines + bold/italic + fills + wrap
 + improved Excel fill color resolution)

Features:
- Imports into CURRENT ACTIVE DRAFTING VIEW
- Replaces previous TableGen import in that same view
- No detail groups created
- Tags imported elements with Extensible Storage
- Reads Excel row heights and variable row spacing
- Reads Excel horizontal / vertical alignment
- Reads merged cells and suppresses internal merged-cell lines
- Draws ONLY explicit Excel borders by default
- Maps Excel font sizes to available Revit TextNoteTypes by rank
- Auto-creates derived text types for bold / italic as needed
- Reads Excel fill colors and places Filled Regions behind text/lines
- Uses wrapped TextNotes when Excel wrap is enabled

Notes:
- If you want legacy "full thin grid everywhere", set DRAW_BASE_GRID = True
"""

import zipfile
import xml.etree.ElementTree as ET

import clr
clr.AddReference('System')
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

from System import Guid, String, Int32
from System.Collections.Generic import List

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    TextNote,
    TextNoteOptions,
    TextNoteType,
    XYZ,
    Transaction,
    Line,
    BuiltInCategory,
    BuiltInParameter,
    ElementId,
    GraphicsStyleType,
    ViewType,
    Group,
    HorizontalTextAlignment,
    VerticalTextAlignment,
    Color,
    FilledRegion,
    FilledRegionType,
    FillPatternElement,
    FillPatternTarget,
    CurveLoop,
)

from Autodesk.Revit.DB.ExtensibleStorage import (
    Schema,
    SchemaBuilder,
    Entity,
    AccessLevel,
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
DEFAULT_COL_WIDTH_FT = 0.7
DEFAULT_ROW_HEIGHT_FT = 0.18
EXCEL_DEFAULT_CHAR_WIDTH = 8.43
EXCEL_DEFAULT_ROW_HEIGHT_PT = 15.0

CELL_PADDING_X_FT = 0.05
CELL_PADDING_Y_FT = 0.03

MAX_COLS = 50
MAX_ROWS = 200

TABLEGEN_GROUP_PREFIX = '__TABLEGEN_VIEW_'

TABLEGEN_SCHEMA_GUID = Guid('6F0A71F2-9E17-4D1F-8D5D-9E4E9D0A1123')
TABLEGEN_SCHEMA_NAME = 'TableGenImportMarker'

EXCEL_TO_REVIT_TEXT_SCALE = 0.60
TEXT_SIZE_TIE_EPS_IN = 0.005

DRAW_BASE_GRID = False
CREATE_FILL_REGIONS = True
CREATE_DERIVED_TEXT_TYPES = True

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

_EMPTY_BORDER = {
    'left': None,
    'right': None,
    'top': None,
    'bottom': None,
}

_EMPTY_ALIGNMENT = {
    'horizontal': None,
    'vertical': None,
    'wrapText': False,
}

_EMPTY_FONT = {
    'size_pt': 11.0,
    'bold': False,
    'italic': False,
}

_EMPTY_FILL = {
    'patternType': None,
    'rgb': None,
}

# ==============================================================
# Extensible Storage helpers
# ==============================================================
def get_or_create_tablegen_schema():
    schema = Schema.Lookup(TABLEGEN_SCHEMA_GUID)
    if schema is not None:
        return schema

    sb = SchemaBuilder(TABLEGEN_SCHEMA_GUID)
    sb.SetSchemaName(TABLEGEN_SCHEMA_NAME)
    sb.SetReadAccessLevel(AccessLevel.Public)
    sb.SetWriteAccessLevel(AccessLevel.Public)
    sb.AddSimpleField('App', String)
    sb.AddSimpleField('ViewId', Int32)
    return sb.Finish()

def mark_tablegen_element(element, view_id_int):
    schema = get_or_create_tablegen_schema()
    entity = Entity(schema)

    f_app = schema.GetField('App')
    f_view = schema.GetField('ViewId')

    try:
        entity.Set[String](f_app, 'TableGen')
        entity.Set[Int32](f_view, view_id_int)
        element.SetEntity(entity)
        return True
    except:
        return False

def is_tablegen_element_in_view(element, view_id_int):
    try:
        schema = Schema.Lookup(TABLEGEN_SCHEMA_GUID)
        if schema is None:
            return False

        entity = element.GetEntity(schema)
        if entity is None or not entity.IsValid():
            return False

        f_app = schema.GetField('App')
        f_view = schema.GetField('ViewId')

        app = entity.Get[String](f_app)
        vid = entity.Get[Int32](f_view)

        return (app == 'TableGen' and vid == view_id_int)
    except:
        return False

def delete_previous_tagged_tablegen(doc, view):
    ids = List[ElementId]()
    view_id_int = view.Id.IntegerValue

    for el in FilteredElementCollector(doc, view.Id).WhereElementIsNotElementType():
        if is_tablegen_element_in_view(el, view_id_int):
            ids.Add(el.Id)

    deleted = ids.Count
    if deleted:
        doc.Delete(ids)
    return deleted

# ==============================================================
# Old group cleanup helpers
# ==============================================================
def get_tablegen_group_name(view):
    return '{}{}'.format(TABLEGEN_GROUP_PREFIX, view.Id.IntegerValue)

def is_old_tablegen_group_name(name, view):
    if not name:
        return False
    target = get_tablegen_group_name(view)
    return name == target or name.startswith(target)

def delete_previous_group_based_tablegen(doc, view):
    ids = List[ElementId]()
    found = {}

    for grp in FilteredElementCollector(doc).OfClass(Group):
        try:
            gtype = doc.GetElement(grp.GetTypeId())
            gname = gtype.Name if gtype else ''
            if not is_old_tablegen_group_name(gname, view):
                continue

            keep = False

            try:
                if grp.OwnerViewId == view.Id:
                    keep = True
            except:
                pass

            if not keep:
                try:
                    for mid in grp.GetMemberIds():
                        mel = doc.GetElement(mid)
                        if mel is None:
                            continue
                        try:
                            if mel.OwnerViewId == view.Id:
                                keep = True
                                break
                        except:
                            keep = True
                            break
                except:
                    keep = True

            if keep:
                found[grp.Id.IntegerValue] = grp.Id
        except:
            pass

    for gid in found.values():
        ids.Add(gid)

    deleted = ids.Count
    if deleted:
        doc.Delete(ids)

    return deleted

# ==============================================================
# XLSX color helpers
# ==============================================================
def _hex_to_rgb_tuple(hex6):
    if not hex6 or len(hex6) != 6:
        return None
    try:
        return (
            int(hex6[0:2], 16),
            int(hex6[2:4], 16),
            int(hex6[4:6], 16),
        )
    except:
        return None

def _rgb_tuple_to_hex(rgb):
    if rgb is None:
        return None
    return '{:02X}{:02X}{:02X}'.format(rgb[0], rgb[1], rgb[2])

def _clamp_u8(v):
    return max(0, min(255, int(round(v))))

def _apply_tint_to_rgb(rgb, tint):
    if rgb is None:
        return None
    if tint is None:
        return _rgb_tuple_to_hex(rgb)

    try:
        tint = float(tint)
    except:
        return _rgb_tuple_to_hex(rgb)

    out = []
    for c in rgb:
        if tint < 0:
            c2 = c * (1.0 + tint)
        else:
            c2 = c + (255.0 - c) * tint
        out.append(_clamp_u8(c2))
    return _rgb_tuple_to_hex(tuple(out))

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

    # borders
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

    # fonts
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

    # fills
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
        if style_idx is None:
            return _EMPTY_BORDER
        if style_idx < 0 or style_idx >= len(xf_border_ids):
            return _EMPTY_BORDER
        border_id = xf_border_ids[style_idx]
        if border_id < 0 or border_id >= len(border_defs):
            return _EMPTY_BORDER
        return border_defs[border_id]
    except:
        return _EMPTY_BORDER

def get_font_def_for_style(style_idx, xf_font_ids, font_defs):
    try:
        if style_idx is None:
            return _EMPTY_FONT
        if style_idx < 0 or style_idx >= len(xf_font_ids):
            return _EMPTY_FONT
        font_id = xf_font_ids[style_idx]
        if font_id < 0 or font_id >= len(font_defs):
            return _EMPTY_FONT
        return font_defs[font_id]
    except:
        return _EMPTY_FONT

def get_font_size_for_style(style_idx, xf_font_ids, font_defs):
    return float(get_font_def_for_style(style_idx, xf_font_ids, font_defs).get('size_pt', 11.0))

def get_alignment_for_style(style_idx, xf_alignments):
    try:
        if style_idx is None:
            return _EMPTY_ALIGNMENT
        if style_idx < 0 or style_idx >= len(xf_alignments):
            return _EMPTY_ALIGNMENT
        return xf_alignments[style_idx]
    except:
        return _EMPTY_ALIGNMENT

def get_fill_for_style(style_idx, xf_fill_ids, fill_defs):
    try:
        if style_idx is None:
            return _EMPTY_FILL
        if style_idx < 0 or style_idx >= len(xf_fill_ids):
            return _EMPTY_FILL
        fill_id = xf_fill_ids[style_idx]
        if fill_id < 0 or fill_id >= len(fill_defs):
            return _EMPTY_FILL
        return fill_defs[fill_id]
    except:
        return _EMPTY_FILL

def excel_border_weight(style_name):
    if not style_name:
        return 0

    s = str(style_name).strip().lower()

    if s in ('thick', 'double'):
        return 3
    if 'medium' in s:
        return 2
    return 1

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
            sheet_format_el = sheet_tree.find('.//' + _ss_tag('sheetFormatPr'))
            if sheet_format_el is not None:
                try:
                    default_row_pt = float(sheet_format_el.get('defaultRowHeight', EXCEL_DEFAULT_ROW_HEIGHT_PT))
                except:
                    default_row_pt = EXCEL_DEFAULT_ROW_HEIGHT_PT
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
                            except (ValueError, TypeError):
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
# Revit helpers
# ==============================================================
def parse_named_text_size_inches(name):
    if not name:
        return None

    txt = str(name).replace('"', ' ').replace("'", ' ')
    txt = txt.replace('-', ' ').replace('_', ' ')

    for token in txt.split():
        try:
            if '/' in token:
                a, b = token.split('/', 1)
                val = float(a) / float(b)
            else:
                val = float(token)

            if 0.03 <= val <= 1.0:
                return val
        except:
            pass

    return None

def excel_points_to_revit_inches(excel_pt):
    try:
        return (float(excel_pt) * EXCEL_TO_REVIT_TEXT_SCALE) / 72.0
    except:
        return (11.0 * EXCEL_TO_REVIT_TEXT_SCALE) / 72.0

def _get_param_bool(element, names):
    for nm in names:
        try:
            p = element.LookupParameter(nm)
            if p is None:
                continue
            return p.AsInteger() == 1
        except:
            pass
    return False

def _set_param_bool(element, names, value):
    for nm in names:
        try:
            p = element.LookupParameter(nm)
            if p is None or p.IsReadOnly:
                continue
            p.Set(1 if value else 0)
            return True
        except:
            pass
    return False

def _set_param_elementid(element, names, eid):
    for nm in names:
        try:
            p = element.LookupParameter(nm)
            if p is None or p.IsReadOnly:
                continue
            p.Set(eid)
            return True
        except:
            pass
    return False

def get_text_type_catalog(doc):
    catalog = []
    fallback_id = None

    for tt in FilteredElementCollector(doc).OfClass(TextNoteType):
        if fallback_id is None:
            fallback_id = tt.Id

        try:
            name = tt.Name
        except:
            name = '<Unknown>'

        size_in = None

        try:
            p = tt.get_Parameter(BuiltInParameter.TEXT_SIZE)
            if p is not None:
                size_ft = p.AsDouble()
                size_in = size_ft * 12.0
        except:
            pass

        if not size_in or size_in <= 0:
            size_in = parse_named_text_size_inches(name)

        if size_in and size_in > 0:
            catalog.append({
                'id': tt.Id,
                'name': name,
                'size_in': float(size_in),
                'bold': _get_param_bool(tt, ['Bold']),
                'italic': _get_param_bool(tt, ['Italic']),
            })

    catalog = sorted(catalog, key=lambda x: x['size_in'])
    return catalog, fallback_id

def choose_text_type_id(excel_pt, text_type_catalog, fallback_id):
    if not text_type_catalog:
        return fallback_id

    target_in = excel_points_to_revit_inches(excel_pt)

    best = None
    best_delta = None

    for item in text_type_catalog:
        delta = abs(item['size_in'] - target_in)

        if best is None:
            best = item
            best_delta = delta
            continue

        if delta < (best_delta - TEXT_SIZE_TIE_EPS_IN):
            best = item
            best_delta = delta
            continue

        if abs(delta - best_delta) <= TEXT_SIZE_TIE_EPS_IN:
            if item['size_in'] < best['size_in']:
                best = item
                best_delta = delta

    return best['id'] if best else fallback_id

def collect_used_excel_font_sizes(sheet_styles, font_defs, xf_font_ids, row_count, col_count):
    used = set()

    for row_idx in range(row_count):
        style_row = sheet_styles[row_idx] if row_idx < len(sheet_styles) else []
        for col_idx in range(col_count):
            style_idx = style_row[col_idx] if col_idx < len(style_row) else 0
            pt = get_font_size_for_style(style_idx, xf_font_ids, font_defs)
            try:
                used.add(float(pt))
            except:
                pass

    if not used:
        used.add(11.0)

    return sorted(used)

def build_excel_to_revit_text_map(used_excel_pts, text_type_catalog, fallback_id):
    mapping = {}

    if not used_excel_pts:
        return mapping

    if not text_type_catalog:
        for pt in used_excel_pts:
            mapping[pt] = fallback_id
        return mapping

    if len(text_type_catalog) == 1:
        only_id = text_type_catalog[0]['id']
        for pt in used_excel_pts:
            mapping[pt] = only_id
        return mapping

    if len(used_excel_pts) == 1:
        pt = used_excel_pts[0]
        mapping[pt] = choose_text_type_id(pt, text_type_catalog, fallback_id)
        return mapping

    excel_count = len(used_excel_pts)
    revit_count = len(text_type_catalog)

    for i, pt in enumerate(used_excel_pts):
        t = float(i) / float(excel_count - 1)
        revit_index = int(round(t * (revit_count - 1)))
        revit_index = max(0, min(revit_index, revit_count - 1))
        mapping[pt] = text_type_catalog[revit_index]['id']

    return mapping

def get_type_id_for_excel_size_ranked(excel_pt, excel_to_revit_map, text_type_catalog, fallback_id):
    try:
        key = float(excel_pt)
        if key in excel_to_revit_map:
            return excel_to_revit_map[key]
    except:
        pass

    return choose_text_type_id(excel_pt, text_type_catalog, fallback_id)

def get_text_type_name(doc, type_id):
    try:
        tt = doc.GetElement(type_id)
        return tt.Name if tt else '<Unknown>'
    except:
        return '<Unknown>'

def col_width_to_ft(char_width):
    if char_width is None:
        return DEFAULT_COL_WIDTH_FT
    if float(char_width) <= 0.0:
        return 0.0
    return (char_width / EXCEL_DEFAULT_CHAR_WIDTH) * DEFAULT_COL_WIDTH_FT

def row_height_pt_to_ft(row_pt, default_row_pt):
    base_pt = default_row_pt if default_row_pt else EXCEL_DEFAULT_ROW_HEIGHT_PT
    use_pt = row_pt if row_pt is not None else base_pt
    if float(use_pt) <= 0.0:
        return 0.0
    return (float(use_pt) / float(base_pt)) * DEFAULT_ROW_HEIGHT_FT

def get_table_geometry(sheet_rows, sheet_styles, col_widths_chars, row_heights_pt, default_row_pt, merge_ranges):
    row_count = min(max(len(sheet_rows), len(sheet_styles)), MAX_ROWS)

    col_count = 0
    for i in range(row_count):
        row_len = len(sheet_rows[i]) if i < len(sheet_rows) else 0
        sty_len = len(sheet_styles[i]) if i < len(sheet_styles) else 0
        col_count = max(col_count, row_len, sty_len)

    for r1, c1, r2, c2 in merge_ranges:
        row_count = max(row_count, min(r2 + 1, MAX_ROWS))
        col_count = max(col_count, min(c2 + 1, MAX_COLS))

    col_count = min(col_count, MAX_COLS)
    row_count = min(row_count, MAX_ROWS)

    cum_x = [0.0]
    for i in range(col_count):
        w = col_width_to_ft(col_widths_chars[i]) if i < len(col_widths_chars) else DEFAULT_COL_WIDTH_FT
        cum_x.append(cum_x[-1] + w)

    row_heights_ft = []
    cum_y = [0.0]
    for i in range(row_count):
        row_pt = row_heights_pt[i] if i < len(row_heights_pt) else None
        h = row_height_pt_to_ft(row_pt, default_row_pt)
        row_heights_ft.append(h)
        cum_y.append(cum_y[-1] - h)

    total_width = cum_x[-1] if cum_x else 0.0
    total_height = -cum_y[-1] if cum_y else 0.0

    return row_count, col_count, cum_x, cum_y, row_heights_ft, total_width, total_height

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

# ==============================================================
# Alignment mapping
# ==============================================================
def map_excel_horizontal_alignment(xl_h):
    s = (xl_h or '').strip().lower()

    if s in ('center', 'centercontinuous', 'distributed', 'justify', 'fill'):
        return HorizontalTextAlignment.Center

    if s in ('right',):
        return HorizontalTextAlignment.Right

    return HorizontalTextAlignment.Left

def map_excel_vertical_alignment(xl_v):
    s = (xl_v or '').strip().lower()

    if s in ('center', 'distributed', 'justify'):
        return VerticalTextAlignment.Middle

    if s in ('bottom',):
        return VerticalTextAlignment.Bottom

    return VerticalTextAlignment.Top

def build_text_options(type_id, xl_h, xl_v):
    opts = TextNoteOptions(type_id)
    opts.HorizontalAlignment = map_excel_horizontal_alignment(xl_h)
    try:
        opts.VerticalAlignment = map_excel_vertical_alignment(xl_v)
    except:
        pass
    return opts

# ==============================================================
# Merge helpers
# ==============================================================
def clip_merge_range(mr, row_count, col_count):
    r1, c1, r2, c2 = mr
    if r1 >= row_count or c1 >= col_count:
        return None
    r2 = min(r2, row_count - 1)
    c2 = min(c2, col_count - 1)
    if r2 < r1 or c2 < c1:
        return None
    return (r1, c1, r2, c2)

def build_merge_lookup(merge_ranges, row_count, col_count):
    merge_by_cell = {}
    master_cells = {}

    for mr in merge_ranges:
        clipped = clip_merge_range(mr, row_count, col_count)
        if clipped is None:
            continue

        r1, c1, r2, c2 = clipped
        master_cells[(r1, c1)] = clipped

        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                merge_by_cell[(r, c)] = clipped

    return merge_by_cell, master_cells

def get_text_anchor_from_bounds(x_left, x_right, y_top, y_bottom, h_align, v_align):
    width = x_right - x_left
    height = y_top - y_bottom

    if h_align == HorizontalTextAlignment.Right:
        x = x_right - CELL_PADDING_X_FT
    elif h_align == HorizontalTextAlignment.Center:
        x = x_left + (width * 0.5)
    else:
        x = x_left + CELL_PADDING_X_FT

    pad_y = min(CELL_PADDING_Y_FT, height * 0.20) if height > 0 else CELL_PADDING_Y_FT

    if v_align == VerticalTextAlignment.Bottom:
        y = y_bottom + pad_y
    elif v_align == VerticalTextAlignment.Middle:
        y = y_top - (height * 0.5)
    else:
        y = y_top - pad_y

    return XYZ(x, y, 0)

# ==============================================================
# Text type derivation
# ==============================================================
def _sanitize_name(txt):
    bad = ['\\', '/', ':', '{', '}', '[', ']', '|', ';', '<', '>', '?', '`', '~']
    out = str(txt)
    for ch in bad:
        out = out.replace(ch, '_')
    return out.strip()

def ensure_derived_text_type(doc, base_type_id, bold, italic, cache):
    key = '{}|{}|{}'.format(base_type_id.IntegerValue, int(bool(bold)), int(bool(italic)))
    if key in cache:
        return cache[key]

    base_type = doc.GetElement(base_type_id)
    if base_type is None:
        cache[key] = base_type_id
        return base_type_id

    base_bold = _get_param_bool(base_type, ['Bold'])
    base_italic = _get_param_bool(base_type, ['Italic'])

    if base_bold == bool(bold) and base_italic == bool(italic):
        cache[key] = base_type_id
        return base_type_id

    for tt in FilteredElementCollector(doc).OfClass(TextNoteType):
        try:
            if _get_param_bool(tt, ['Bold']) == bool(bold) and _get_param_bool(tt, ['Italic']) == bool(italic):
                p1 = tt.get_Parameter(BuiltInParameter.TEXT_SIZE)
                p2 = base_type.get_Parameter(BuiltInParameter.TEXT_SIZE)
                if p1 and p2 and abs(p1.AsDouble() - p2.AsDouble()) < 1e-9:
                    cache[key] = tt.Id
                    return tt.Id
        except:
            pass

    if not CREATE_DERIVED_TEXT_TYPES:
        cache[key] = base_type_id
        return base_type_id

    try:
        suffix = []
        if bold:
            suffix.append('Bold')
        if italic:
            suffix.append('Italic')
        suffix_txt = ' '.join(suffix) if suffix else 'Regular'

        base_name = _sanitize_name(base_type.Name)
        new_name = 'TableGen {} {}'.format(base_name, suffix_txt)

        existing_names = set()
        for tt in FilteredElementCollector(doc).OfClass(TextNoteType):
            try:
                existing_names.add(tt.Name)
            except:
                pass
        if new_name in existing_names:
            i = 1
            test_name = '{} {}'.format(new_name, i)
            while test_name in existing_names:
                i += 1
                test_name = '{} {}'.format(new_name, i)
            new_name = test_name

        new_id = base_type.Duplicate(new_name)
        new_type = doc.GetElement(new_id)

        _set_param_bool(new_type, ['Bold'], bool(bold))
        _set_param_bool(new_type, ['Italic'], bool(italic))

        cache[key] = new_id
        return new_id
    except:
        cache[key] = base_type_id
        return base_type_id

# ==============================================================
# Fill helpers
# ==============================================================
def rgb_string_to_color(rgb):
    if not rgb or len(rgb) != 6:
        return None
    try:
        return Color(int(rgb[0:2], 16), int(rgb[2:4], 16), int(rgb[4:6], 16))
    except:
        return None

def is_effective_fill(fill_def):
    if not fill_def:
        return False

    pattern = (fill_def.get('patternType') or '').strip().lower()
    rgb = fill_def.get('rgb')

    if not rgb:
        return False

    if pattern in ('none',):
        return False

    if rgb.upper() == 'FFFFFF':
        return False

    return True

def get_solid_fill_pattern_id(doc):
    for fpe in FilteredElementCollector(doc).OfClass(FillPatternElement):
        try:
            fp = fpe.GetFillPattern()
            if fp and fp.Target == FillPatternTarget.Drafting and fp.IsSolidFill:
                return fpe.Id
        except:
            pass
    return ElementId.InvalidElementId

def get_or_create_filled_region_type_for_color(doc, rgb, cache):
    if rgb in cache:
        return cache[rgb]

    color_obj = rgb_string_to_color(rgb)
    if color_obj is None:
        cache[rgb] = None
        return None

    base_type = None
    for frt in FilteredElementCollector(doc).OfClass(FilledRegionType):
        base_type = frt
        break
    if base_type is None:
        cache[rgb] = None
        return None

    solid_fill_id = get_solid_fill_pattern_id(doc)
    if solid_fill_id == ElementId.InvalidElementId:
        cache[rgb] = None
        return None

    for frt in FilteredElementCollector(doc).OfClass(FilledRegionType):
        try:
            if frt.ForegroundPatternId == solid_fill_id:
                c = frt.ForegroundPatternColor
                if c and c.Red == color_obj.Red and c.Green == color_obj.Green and c.Blue == color_obj.Blue:
                    cache[rgb] = frt.Id
                    return frt.Id
        except:
            pass

    try:
        existing_names = set()
        for frt in FilteredElementCollector(doc).OfClass(FilledRegionType):
            try:
                existing_names.add(frt.Name)
            except:
                pass

        new_name = 'TableGen Fill {}'.format(rgb)
        if new_name in existing_names:
            i = 1
            test_name = '{} {}'.format(new_name, i)
            while test_name in existing_names:
                i += 1
                test_name = '{} {}'.format(new_name, i)
            new_name = test_name

        new_id = base_type.Duplicate(new_name)
        frt = doc.GetElement(new_id)

        frt.ForegroundPatternId = solid_fill_id
        frt.ForegroundPatternColor = color_obj

        try:
            frt.IsMasking = True
        except:
            pass

        cache[rgb] = new_id
        return new_id
    except:
        cache[rgb] = None
        return None

def create_cell_fill_region(doc, view, x1, x2, y_top, y_bottom, fr_type_id):
    if fr_type_id is None:
        return None

    try:
        loop = CurveLoop()
        loop.Append(Line.CreateBound(XYZ(x1, y_top, 0), XYZ(x2, y_top, 0)))
        loop.Append(Line.CreateBound(XYZ(x2, y_top, 0), XYZ(x2, y_bottom, 0)))
        loop.Append(Line.CreateBound(XYZ(x2, y_bottom, 0), XYZ(x1, y_bottom, 0)))
        loop.Append(Line.CreateBound(XYZ(x1, y_bottom, 0), XYZ(x1, y_top, 0)))

        loops = List[CurveLoop]()
        loops.Add(loop)

        fr = FilledRegion.Create(doc, fr_type_id, view.Id, loops)
        return fr
    except:
        return None

# ==============================================================
# Placement - fills
# ==============================================================
def place_table_fills(doc, view, sheet_rows, sheet_styles, col_widths_chars, row_heights_pt, default_row_pt, merge_ranges, fill_defs, xf_fill_ids):
    if not CREATE_FILL_REGIONS:
        return 0, [], 0

    row_count, col_count, cum_x, cum_y, row_heights_ft, _, _ = get_table_geometry(
        sheet_rows, sheet_styles, col_widths_chars, row_heights_pt, default_row_pt, merge_ranges
    )

    merge_by_cell, master_cells = build_merge_lookup(merge_ranges, row_count, col_count)
    fill_type_cache = {}

    created_ids = []
    fill_count = 0
    tag_count = 0

    handled = set()

    for row_idx in range(row_count):
        style_row = sheet_styles[row_idx] if row_idx < len(sheet_styles) else []

        for col_idx in range(col_count):
            mr = merge_by_cell.get((row_idx, col_idx))
            if mr is not None:
                if (row_idx, col_idx) != (mr[0], mr[1]):
                    continue
                if mr in handled:
                    continue
                handled.add(mr)
                r1, c1, r2, c2 = mr
                style_idx = style_row[col_idx] if col_idx < len(style_row) else 0
            else:
                r1 = row_idx
                c1 = col_idx
                r2 = row_idx
                c2 = col_idx
                style_idx = style_row[col_idx] if col_idx < len(style_row) else 0

            fill_def = get_fill_for_style(style_idx, xf_fill_ids, fill_defs)
            if not is_effective_fill(fill_def):
                continue

            rgb = fill_def.get('rgb')
            fr_type_id = get_or_create_filled_region_type_for_color(doc, rgb, fill_type_cache)
            if fr_type_id is None:
                continue

            x1 = cum_x[c1]
            x2 = cum_x[c2 + 1]
            y_top = cum_y[r1]
            y_bottom = cum_y[r2 + 1]

            if abs(x2 - x1) <= 1e-9 or abs(y_top - y_bottom) <= 1e-9:
                continue

            fr = create_cell_fill_region(doc, view, x1, x2, y_top, y_bottom, fr_type_id)
            if fr is None:
                continue

            created_ids.append(fr.Id)
            if mark_tablegen_element(fr, view.Id.IntegerValue):
                tag_count += 1
            fill_count += 1

    return fill_count, created_ids, tag_count

# ==============================================================
# Placement - text
# ==============================================================
def place_table_text(doc, view, sheet_rows, sheet_styles, col_widths_chars, row_heights_pt, default_row_pt, merge_ranges, font_defs, xf_font_ids, xf_alignments):
    row_count, col_count, cum_x, cum_y, row_heights_ft, _, _ = get_table_geometry(
        sheet_rows, sheet_styles, col_widths_chars, row_heights_pt, default_row_pt, merge_ranges
    )

    merge_by_cell, master_cells = build_merge_lookup(merge_ranges, row_count, col_count)

    text_type_catalog, fallback_type_id = get_text_type_catalog(doc)
    if fallback_type_id is None:
        raise Exception('No TextNoteType found in project.')

    used_excel_pts = collect_used_excel_font_sizes(
        sheet_styles, font_defs, xf_font_ids, row_count, col_count
    )
    excel_to_revit_map = build_excel_to_revit_text_map(
        used_excel_pts, text_type_catalog, fallback_type_id
    )

    style_to_type_id = {}
    options_cache = {}
    derived_type_cache = {}
    used_type_names = set()

    created_ids = []
    count = 0
    tag_count = 0

    for row_idx in range(row_count):
        row = sheet_rows[row_idx] if row_idx < len(sheet_rows) else []
        style_row = sheet_styles[row_idx] if row_idx < len(sheet_styles) else []

        for col_idx, val in enumerate(row[:col_count] if row else []):
            if val is None:
                continue

            mr = merge_by_cell.get((row_idx, col_idx))
            if mr is not None:
                if (row_idx, col_idx) != (mr[0], mr[1]):
                    continue
                r1, c1, r2, c2 = mr
            else:
                r1 = row_idx
                c1 = col_idx
                r2 = row_idx
                c2 = col_idx

            style_idx = style_row[col_idx] if col_idx < len(style_row) else 0
            font_def = get_font_def_for_style(style_idx, xf_font_ids, font_defs)

            if style_idx not in style_to_type_id:
                excel_pt = float(font_def.get('size_pt', 11.0))
                base_type_id = get_type_id_for_excel_size_ranked(
                    excel_pt,
                    excel_to_revit_map,
                    text_type_catalog,
                    fallback_type_id
                )
                final_type_id = ensure_derived_text_type(
                    doc,
                    base_type_id,
                    font_def.get('bold', False),
                    font_def.get('italic', False),
                    derived_type_cache
                )
                style_to_type_id[style_idx] = final_type_id

            align_data = get_alignment_for_style(style_idx, xf_alignments)
            xl_h = align_data.get('horizontal')
            xl_v = align_data.get('vertical')
            wrap_text = bool(align_data.get('wrapText', False))

            type_id = style_to_type_id[style_idx]
            cache_key = '{}|{}|{}'.format(type_id.IntegerValue, xl_h or '', xl_v or '')

            if cache_key not in options_cache:
                options_cache[cache_key] = build_text_options(type_id, xl_h, xl_v)

            opts = options_cache[cache_key]

            x_left = cum_x[c1]
            x_right = cum_x[c2 + 1]
            y_top = cum_y[r1]
            y_bottom = cum_y[r2 + 1]

            pos = get_text_anchor_from_bounds(
                x_left,
                x_right,
                y_top,
                y_bottom,
                opts.HorizontalAlignment,
                opts.VerticalAlignment
            )

            note = None
            text_value = str(val)

            if wrap_text:
                usable_width = max((x_right - x_left) - (CELL_PADDING_X_FT * 2.0), 0.01)
                try:
                    min_w = TextNote.GetMinimumAllowedWidth(doc, type_id)
                    max_w = TextNote.GetMaximumAllowedWidth(doc, type_id)
                    width = max(min_w, min(usable_width, max_w))
                except:
                    width = usable_width

                try:
                    note = TextNote.Create(doc, view.Id, pos, width, text_value, opts)
                except:
                    note = TextNote.Create(doc, view.Id, pos, text_value, opts)
            else:
                note = TextNote.Create(doc, view.Id, pos, text_value, opts)

            created_ids.append(note.Id)

            if mark_tablegen_element(note, view.Id.IntegerValue):
                tag_count += 1

            used_type_names.add(get_text_type_name(doc, type_id))
            count += 1

    return count, created_ids, sorted(list(used_type_names)), tag_count

# ==============================================================
# Border helpers
# ==============================================================
def apply_merge_outline_from_master(h_segments, v_segments, merge_ranges, sheet_styles, border_defs, xf_border_ids, row_count, col_count):
    for mr in merge_ranges:
        clipped = clip_merge_range(mr, row_count, col_count)
        if clipped is None:
            continue

        r1, c1, r2, c2 = clipped

        style_row = sheet_styles[r1] if r1 < len(sheet_styles) else []
        style_idx = style_row[c1] if c1 < len(style_row) else 0
        border = get_border_for_style(style_idx, xf_border_ids, border_defs)

        top_w = excel_border_weight(border.get('top'))
        bot_w = excel_border_weight(border.get('bottom'))
        left_w = excel_border_weight(border.get('left'))
        right_w = excel_border_weight(border.get('right'))

        if top_w:
            for c in range(c1, c2 + 1):
                h_segments[(r1, c)] = max(h_segments.get((r1, c), 0), top_w)

        if bot_w:
            for c in range(c1, c2 + 1):
                h_segments[(r2 + 1, c)] = max(h_segments.get((r2 + 1, c), 0), bot_w)

        if left_w:
            for r in range(r1, r2 + 1):
                v_segments[(c1, r)] = max(v_segments.get((c1, r), 0), left_w)

        if right_w:
            for r in range(r1, r2 + 1):
                v_segments[(c2 + 1, r)] = max(v_segments.get((c2 + 1, r), 0), right_w)

def apply_merge_suppression(h_segments, v_segments, merge_ranges, row_count, col_count):
    for mr in merge_ranges:
        clipped = clip_merge_range(mr, row_count, col_count)
        if clipped is None:
            continue

        r1, c1, r2, c2 = clipped

        for r in range(r1 + 1, r2 + 1):
            for c in range(c1, c2 + 1):
                h_segments[(r, c)] = 0

        for c in range(c1 + 1, c2 + 1):
            for r in range(r1, r2 + 1):
                v_segments[(c, r)] = 0

def build_border_segments(sheet_rows, sheet_styles, col_widths_chars, row_heights_pt, default_row_pt, merge_ranges, border_defs, xf_border_ids):
    row_count, col_count, _, _, _, _, _ = get_table_geometry(
        sheet_rows, sheet_styles, col_widths_chars, row_heights_pt, default_row_pt, merge_ranges
    )

    h_segments = {}
    v_segments = {}

    if DRAW_BASE_GRID:
        for r in range(row_count + 1):
            for c in range(col_count):
                h_segments[(r, c)] = 1

        for c in range(col_count + 1):
            for r in range(row_count):
                v_segments[(c, r)] = 1

    for row_idx in range(row_count):
        style_row = sheet_styles[row_idx] if row_idx < len(sheet_styles) else []

        for col_idx in range(col_count):
            style_idx = style_row[col_idx] if col_idx < len(style_row) else 0
            border = get_border_for_style(style_idx, xf_border_ids, border_defs)

            top_w = excel_border_weight(border.get('top'))
            bot_w = excel_border_weight(border.get('bottom'))
            left_w = excel_border_weight(border.get('left'))
            right_w = excel_border_weight(border.get('right'))

            if top_w:
                h_segments[(row_idx, col_idx)] = max(h_segments.get((row_idx, col_idx), 0), top_w)
            if bot_w:
                h_segments[(row_idx + 1, col_idx)] = max(h_segments.get((row_idx + 1, col_idx), 0), bot_w)
            if left_w:
                v_segments[(col_idx, row_idx)] = max(v_segments.get((col_idx, row_idx), 0), left_w)
            if right_w:
                v_segments[(col_idx + 1, row_idx)] = max(v_segments.get((col_idx + 1, row_idx), 0), right_w)

    apply_merge_outline_from_master(
        h_segments, v_segments, merge_ranges, sheet_styles, border_defs, xf_border_ids, row_count, col_count
    )

    apply_merge_suppression(h_segments, v_segments, merge_ranges, row_count, col_count)

    return h_segments, v_segments

def draw_table_lines(doc, view, sheet_rows, sheet_styles, col_widths_chars, row_heights_pt, default_row_pt, merge_ranges, border_defs, xf_border_ids):
    row_count, col_count, cum_x, cum_y, row_heights_ft, _, _ = get_table_geometry(
        sheet_rows, sheet_styles, col_widths_chars, row_heights_pt, default_row_pt, merge_ranges
    )

    style_map = get_line_style_map(doc)
    h_segments, v_segments = build_border_segments(
        sheet_rows, sheet_styles, col_widths_chars, row_heights_pt, default_row_pt, merge_ranges, border_defs, xf_border_ids
    )

    created_ids = []
    line_count = 0
    tag_count = 0

    for (row_edge, col_idx), weight in h_segments.items():
        if weight <= 0:
            continue

        y = cum_y[row_edge]
        x1 = cum_x[col_idx]
        x2 = cum_x[col_idx + 1]

        if abs(x2 - x1) <= 1e-9:
            continue

        crv = Line.CreateBound(XYZ(x1, y, 0), XYZ(x2, y, 0))
        dc = doc.Create.NewDetailCurve(view, crv)

        gs = style_map.get(weight)
        if gs is not None:
            try:
                dc.LineStyle = gs
            except:
                pass

        created_ids.append(dc.Id)
        if mark_tablegen_element(dc, view.Id.IntegerValue):
            tag_count += 1
        line_count += 1

    for (col_edge, row_idx), weight in v_segments.items():
        if weight <= 0:
            continue

        x = cum_x[col_edge]
        y1 = cum_y[row_idx]
        y2 = cum_y[row_idx + 1]

        if abs(y2 - y1) <= 1e-9:
            continue

        crv = Line.CreateBound(XYZ(x, y1, 0), XYZ(x, y2, 0))
        dc = doc.Create.NewDetailCurve(view, crv)

        gs = style_map.get(weight)
        if gs is not None:
            try:
                dc.LineStyle = gs
            except:
                pass

        created_ids.append(dc.Id)
        if mark_tablegen_element(dc, view.Id.IntegerValue):
            tag_count += 1
        line_count += 1

    return line_count, created_ids, tag_count

# ==============================================================
# Main
# ==============================================================
def main():
    uidoc = __revit__.ActiveUIDocument
    doc = uidoc.Document
    view = uidoc.ActiveView

    if view.ViewType != ViewType.DraftingView:
        forms.alert(
            'Please run this command from a Drafting View.\n\n'
            'This version updates the CURRENT active drafting view.',
            title='TableGen',
            exitscript=True
        )
        return

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
        forms.alert(
            'Could not read Excel file:\n{}'.format(str(e)),
            title='Error',
            exitscript=True
        )
        return

    if not sheet_names:
        forms.alert('No sheets found in workbook.', title='Error', exitscript=True)
        return

    if len(sheet_names) == 1:
        ws_name = sheet_names[0]
    else:
        ws_name = forms.SelectFromList.show(
            sheet_names,
            title='Select Worksheet',
            button_name='Import'
        )
        if not ws_name:
            return

    sheet_rows = sheets[ws_name]
    style_rows = sheet_styles[ws_name]
    col_widths_ch = col_widths[ws_name]
    row_heights_pt = row_heights[ws_name]
    default_row_pt = default_row_heights[ws_name]
    sheet_merges = merge_ranges[ws_name]

    row_count, col_count, _, _, _, _, _ = get_table_geometry(
        sheet_rows, style_rows, col_widths_ch, row_heights_pt, default_row_pt, sheet_merges
    )
    if row_count == 0 or col_count == 0:
        forms.alert(
            'Selected sheet appears to be empty.',
            title='Warning',
            exitscript=True
        )
        return

    t1 = Transaction(doc, 'TableGen - Clear Previous Import')
    t1.Start()

    deleted_tagged_count = delete_previous_tagged_tablegen(doc, view)
    deleted_old_group_count = delete_previous_group_based_tablegen(doc, view)

    t1.Commit()

    t2 = Transaction(doc, 'TableGen - Import Excel Table')
    t2.Start()

    fill_count, fill_ids, fill_tag_count = place_table_fills(
        doc,
        view,
        sheet_rows,
        style_rows,
        col_widths_ch,
        row_heights_pt,
        default_row_pt,
        sheet_merges,
        fill_defs,
        xf_fill_ids
    )

    line_count, line_ids, line_tag_count = draw_table_lines(
        doc,
        view,
        sheet_rows,
        style_rows,
        col_widths_ch,
        row_heights_pt,
        default_row_pt,
        sheet_merges,
        border_defs,
        xf_border_ids
    )

    text_count, text_ids, used_type_names, text_tag_count = place_table_text(
        doc,
        view,
        sheet_rows,
        style_rows,
        col_widths_ch,
        row_heights_pt,
        default_row_pt,
        sheet_merges,
        font_defs,
        xf_font_ids,
        xf_alignments
    )

    t2.Commit()

    forms.alert(
        'Done!\n\n'
        'View updated: {}\n'
        'Worksheet: {}\n'
        'Prior tagged elements removed: {}\n'
        'Prior old-group imports removed: {}\n'
        'Merged ranges found: {}\n'
        'Fill regions created: {}\n'
        'Cells placed: {}\n'
        'Line segments drawn: {}\n'
        'Tagged new fills: {}\n'
        'Tagged new text: {}\n'
        'Tagged new lines: {}\n'
        'Text types used: {}\n'
        'Draw base grid: {}'.format(
            view.Name,
            ws_name,
            deleted_tagged_count,
            deleted_old_group_count,
            len(sheet_merges),
            fill_count,
            text_count,
            line_count,
            fill_tag_count,
            text_tag_count,
            line_tag_count,
            ', '.join(used_type_names) if used_type_names else '<none>',
            DRAW_BASE_GRID
        ),
        title='TableGen Phase 3.5.1'
    )

main()