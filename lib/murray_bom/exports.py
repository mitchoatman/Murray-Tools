# -*- coding: utf-8 -*-
"""The four Bulk Fab exports (Copper, Non Copper, Valve, Hanger): which parts are kept, how rows are rolled up,
and how the MR form rows come out. Port of exportTypes.js (BULK_FAB) and ExportXLS.js (buildRow, aggregate,
combinePipeRows, summarizeRods, prepareTables) from the web app.

A Part is what the Revit reader hands over: id, name, props (parameter name -> value), straight (bool or None),
level, workset. Lengths arrive as text with a unit ("12.500000 in"), so the same parsing and rounding rules apply
as in the web app. Nothing in this module needs Revit."""
import re
from functools import cmp_to_key

from .units import (strip_id, parse_inches, to_feet_inches, to_fractional_inches, size_numbers, natural_key,
                    FRACTION_DENOMINATOR)

RAW = '\x01raw:Length'
LEVEL_PROPS = ['Reference Level', 'Level', 'Schedule Level']
SERVICE_PROPS = ['FP_Service Name', 'Fabrication Service Name', 'Fabrication Service', 'Service Name']


class Part(object):
    def __init__(self, id, name, props, straight=None, workset=''):
        self.id = id
        self.name = name or ''
        self.props = props or {}
        self.straight = straight
        self.workset = workset or ''
        self.element = None          # kept only by tests; the reader keeps eid and doc instead
        self.eid = None              # the Revit ElementId (geometry is read from it)
        self.doc = None


def first_value(props, names):
    for n in names:
        v = props.get(n)
        if v is None or isinstance(v, (list, dict)):
            continue
        if str(v).strip() != '':
            return v
    return ''


def _s(v):
    return '' if v is None else str(v)


# ---- predicates and computed cells ----

NO_SERVICE = '(no service)'


def service_text(part):
    """The part's fabrication service, text after the colon ("PLUMBING: STORM DRAIN" -> "STORM DRAIN"). Same text as the 3D
    legend and the system color list use. Blank when the part has no service."""
    v = _s(first_value(part.props, SERVICE_PROPS)).strip()
    i = v.find(':')
    return (v if i == -1 else v[i + 1:]).strip()


def service_label(part):
    return service_text(part) or NO_SERVICE


def description(part):
    """The part's description: its family name (Revit's Element.Name on a fabrication part is the type, 'Default').
    Falls back to the part name when no family name was read."""
    fam = _s(first_value(part.props, ['Family', 'Family Name'])).strip()
    return fam if fam else strip_id(part.name)



BUILD_ORDER = ['lines', 'straights', 'valves', 'hangers', 'fittings']
STAGE_LABEL = {'lines': 'Centerlines', 'straights': 'Straights', 'fittings': 'Fittings', 'valves': 'Valves', 'hangers': 'Hangers'}
_SERVICE_TYPE_PROPS = ['Service Type', 'FP_Service Type', 'Fabrication Service Type', 'ServiceType']


def part_kind(part, is_hanger=False):
    """Where a part sits in the 3D build order: 'straight', 'fitting', 'valve' or 'hanger' (see BUILD_ORDER)."""
    if is_hanger:
        return 'hanger'
    if is_straight(part):
        return 'straight'
    if re.match(r'^valves?$', _s(first_value(part.props, _SERVICE_TYPE_PROPS)).strip(), re.I):
        return 'valve'
    return 'fitting'


def is_straight(part):
    # Revit's FabricationPart.IsAStraight() when the reader could call it; otherwise a property that only straights have.
    if part.straight is not None:
        return bool(part.straight)
    return first_value(part.props, ['Length Option']) != ''


def is_bulk_straight(part):
    # Some fittings (olets) also report straight, so the part name must also say pipe or tube.
    return is_straight(part) and re.search(r'\b(pipe|tube|tubing)\b', description(part), re.I) is not None


def _material_after_colon(part):
    v = _s(first_value(part.props, ['Part Material', 'FP_Part Material']))
    i = v.find(':')
    return (v if i == -1 else v[i + 1:]).strip()


def _sort_key(r):
    length = _s(r.get('Length')).strip()
    is_pipe = length != ''
    return [0 if is_pipe else 1, size_numbers(r.get('Size')), r.get('Description'), r.get('Material'),
            r.get('Part Material')]


def bulk_fab(key, label, opts):
    cols = [
        {'header': 'Size', 'props': ['Product Entry', 'Size']},      # [CONFIRM] 'Size' fallback on a real model
        {'header': 'Length', 'props': ['Length'], 'format': 'feet-inches', 'when': is_bulk_straight, 'sum': True,
         'blank': bool(opts.get('noLength'))},
        {'header': 'Description', 'compute': description},
        {'header': 'Service Type', 'props': ['Service Type', 'FP_Service Type', 'Fabrication Service Type', 'ServiceType']},
        {'header': 'Part Material', 'props': ['Part Material', 'FP_Part Material']},
        {'header': 'Material', 'compute': _material_after_colon},
        {'header': 'Reference Level', 'props': ['Reference Level'], 'blank': True},
        {'header': 'Fabrication Service Name',
         'props': ['Fabrication Service Name', 'Fabrication Service', 'FP_Service Name'], 'blank': True},
        {'header': 'Product Code', 'props': ['Product Code'], 'blank': True},
        {'header': 'Unit Cost', 'compute': lambda p: '', 'blank': True},
        {'header': 'Total', 'compute': lambda p: '', 'blank': True},
    ]
    d = {
        'key': key, 'label': label,
        'source': opts.get('source', 'pipe'),
        'skipWhenBlank': opts.get('skipWhenBlank'),
        'includeWhen': opts.get('includeWhen'),
        'excludeWhen': opts.get('excludeWhen'),
        'skipNote': opts.get('skipNote'),
        'quietCleared': bool(opts.get('noLength')),
        'columns': cols,
        'groupBy': ['Description', 'Size', 'Part Material'],
        'sortKey': _sort_key,
        'countColumn': 'Qty',
        'output': ['Qty', 'Length', 'Size', 'Description', 'Material', 'Unit Cost', 'Total'],
        'sumColumns': ['Length'],
        'combinePipe': bool(opts.get('combinePipe')),
        'noPipeQty': bool(opts.get('noPipeQty')),
        'rodSummary': bool(opts.get('rodSummary')),
    }
    if d['rodSummary']:
        cols.append({'header': 'Rod Size', 'props': ['FP_Rod Size'], 'blank': True})
        cols.append({'header': 'Length A', 'props': ['FP_Rod Length A'], 'blank': True})
        cols.append({'header': 'Length B', 'props': ['FP_Rod Length B'], 'blank': True})
    return d


_EXCL = re.compile(r'^(joints?|welds?|valves?|gaskets?|hangers?)$', re.I)

EXPORT_TYPES = [
    bulk_fab('bulkcopper', 'Bulk Fab Copper', {
        'combinePipe': True, 'noPipeQty': True,
        'includeWhen': {'Part Material': re.compile(r'copper|bronze|brass', re.I)},
        'excludeWhen': {'Service Type': _EXCL},
        'skipNote': 'parts left out (not copper, bronze or brass; or Joint, Weld, Valve, Gasket or Hanger service type)'}),
    bulk_fab('bulknoncopper', 'Bulk Fab Non Copper', {
        'combinePipe': True, 'noPipeQty': True,
        'includeWhen': {'Part Material': re.compile(r'^(?!.*(?:copper|bronze|brass)).+$', re.I)},
        'excludeWhen': {'Service Type': _EXCL},
        'skipNote': 'parts left out (copper, bronze or brass; no material; or Joint, Weld, Valve, Gasket or Hanger service type)'}),
    bulk_fab('bulkvalves', 'Bulk Fab Valve', {
        'includeWhen': {'Service Type': re.compile(r'^valves?$', re.I)},
        'noLength': True,
        'skipNote': 'parts left out (service type is not Valve)'}),
    bulk_fab('bulkhangers', 'Bulk Fab Hanger', {
        'source': 'hangers', 'noLength': True, 'rodSummary': True,
        'skipWhenBlank': ['Reference Level', 'Fabrication Service Name', 'Product Code'],
        'skipNote': 'empty parts (no level, service or product code) left out'}),
]


def type_by_key(key):
    for t in EXPORT_TYPES:
        if t['key'] == key:
            return t
    return None


# ---- one part -> one row ----

def build_row(defn, part):
    row = {}
    notes = {'rounded': 0, 'assumed': 0, 'unparsed': 0, 'cleared': 0}
    for c in defn['columns']:
        when = c.get('when')
        if when is not None and not when(part):
            row[c['header']] = ''
            notes['cleared'] += 1
            continue
        if c.get('compute'):
            v = c['compute'](part)
            v = v if v else ''
        else:
            v = first_value(part.props, c['props'])
        if c.get('sum'):
            pi = parse_inches(v)
            if pi:
                row['\x01raw:' + c['header']] = pi[0]
        fmt = c.get('format')
        if fmt in ('fractional-inches', 'feet-inches'):
            r = to_feet_inches(v, c.get('denom')) if fmt == 'feet-inches' else to_fractional_inches(v, c.get('denom'))
            v = r['text']
            if r['rounded'] and not c.get('sum'):
                notes['rounded'] += 1
            if r['ok'] and r['assumed'] and r['text'] != '':
                notes['assumed'] += 1
            if not r['ok']:
                notes['unparsed'] += 1
        row[c['header']] = v

    level = _s(first_value(part.props, LEVEL_PROPS)).strip()
    skip = False
    swb = defn.get('skipWhenBlank')
    if swb:
        skip = all(_s(row.get(h)).strip() == '' for h in swb)
    if not skip and defn.get('excludeWhen'):
        skip = any(rx.search(_s(row.get(h)).strip()) for h, rx in defn['excludeWhen'].items())
    if not skip and defn.get('includeWhen'):
        skip = not all(rx.search(_s(row.get(h))) for h, rx in defn['includeWhen'].items())
    return {'row': row, 'level': level, 'notes': notes, 'skip': skip, 'part': part}


# ---- rolling up ----

def _cmp_part(x, y):
    if isinstance(x, list) and isinstance(y, list):
        for a, b in zip(x, y):
            if a != b:
                return -1 if a < b else 1
        return len(x) - len(y)
    if isinstance(x, (int, float)) and isinstance(y, (int, float)):
        return (x > y) - (x < y)
    a, b = natural_key(_s(x)), natural_key(_s(y))
    return (a > b) - (a < b)


def _cmp_keys(a, b):
    for x, y in zip(a, b):
        d = _cmp_part(x, y)
        if d:
            return d
    return 0


def aggregate(defn, rows):
    groups, order = {}, []
    sums = defn.get('sumColumns') or []
    rounded_sums = 0
    for r in rows:
        key = tuple(_s(r.get(h)).strip() for h in defn['groupBy'])
        if key not in groups:
            groups[key] = {'count': 0, 'first': r, 'sum': {}, 'sk': defn['sortKey'](r) if defn.get('sortKey') else None}
            order.append(key)
        groups[key]['count'] += 1
        for h in sums:
            raw = r.get('\x01raw:' + h)
            if isinstance(raw, (int, float)):
                groups[key]['sum'][h] = groups[key]['sum'].get(h, 0) + raw

    def order_cmp(a, b):
        ga, gb = groups[a], groups[b]
        if ga['sk'] is not None and gb['sk'] is not None:
            return _cmp_keys(ga['sk'], gb['sk'])
        for x, y in zip(a, b):
            d = _cmp_part(x, y)
            if d:
                return d
        return 0
    order.sort(key=cmp_to_key(order_cmp))

    out = []
    for key in order:
        g = groups[key]
        o = {}
        for h in defn['output']:
            if h == defn.get('countColumn'):
                o[h] = g['count']
                continue
            if h in sums:
                if h not in g['sum']:
                    o[h] = ''
                    continue
                col = next((c for c in defn['columns'] if c['header'] == h), None)
                res = to_feet_inches(g['sum'][h], col.get('denom') if col else None)
                o[h] = res['text']
                if res['rounded']:
                    rounded_sums += 1
                continue
            o[h] = g['first'].get(h)
        for h in sums:
            o['\x01raw:' + h] = g['sum'].get(h)
        out.append(o)
    return out, rounded_sums


def combine_pipe_rows(rows, defn):
    """Pipe rows whose names differ only in end conditions (BEV X THRD, PE X PE...) become one line per size,
    material and name root, with the total length. Stub-ins count under their branch size."""
    def base_of(name):
        n = _s(name).strip()
        m = re.match(r'^STUB[\s-]*IN\s*\((.*)\)$', n, re.I)
        if m:
            n = m.group(1).strip()
        else:
            n = re.sub(r'^STUB[\s-]*IN\s+', '', n, flags=re.I)
        e = re.match(r'^(.*\S)\s+\S{1,8}\s+X\s+\S{1,8}$', n, re.I)
        return e.group(1) if e else n

    def size_of(size):
        parts = [p for p in re.split(r'\s*[xX\u00d7]\s*', _s(size)) if p.strip() != '']
        return parts[-1].strip() if parts else ''

    def norm(t):
        return re.sub(r'["\'\s]', '', _s(t)).lower()

    pipe = [r for r in rows if isinstance(r.get(RAW), (int, float)) and r[RAW] > 0]
    others = [r for r in rows if not (isinstance(r.get(RAW), (int, float)) and r[RAW] > 0)]
    if not pipe:
        return rows, None

    groups, order = {}, []
    for r in pipe:
        k = norm(size_of(r.get('Size'))) + '\x01' + _s(r.get('Material'))
        if k not in groups:
            groups[k] = []
            order.append(k)
        groups[k].append(r)

    lines, rounded = [], 0
    for k in order:
        lst = groups[k]
        base_for = [base_of(r.get('Description')) for r in lst]
        distinct = sorted(set(base_for), key=len)
        root = {}
        for b in distinct:
            root[b] = b
            for a in distinct:
                if a != b and b.startswith(a + ' ') and len(b[len(a):].strip().split()) <= 2:
                    root[b] = root[a]
                    break
        by_set, set_order = {}, []
        for i, r in enumerate(lst):
            rb = root[base_for[i]]
            if rb not in by_set:
                by_set[rb] = []
                set_order.append(rb)
            by_set[rb].append(r)
        for rb in set_order:
            members = by_set[rb]
            total = sum(r[RAW] for r in members)
            res = to_feet_inches(total)
            if res['rounded']:
                rounded += 1
            plain = next((r for r in members if not re.search(r'[xX\u00d7]', _s(r.get('Size')))), None)
            line = {'Qty': '', 'Length': res['text'], 'Size': plain['Size'] if plain else size_of(members[0].get('Size')),
                    'Description': rb, 'Material': members[0].get('Material'), 'Unit Cost': '', 'Total': '',
                    '_members': len(members), RAW: total}
            lines.append(line)

    if defn.get('sortKey'):
        for l in lines:
            l['_sk'] = defn['sortKey'](l)
        lines.sort(key=cmp_to_key(lambda a, b: _cmp_keys(a['_sk'], b['_sk'])))
    merged = sum(l['_members'] for l in lines if l['_members'] > 1)
    return lines + others, {'sets': len(lines), 'rows': len(pipe), 'merged': merged, 'rounded': rounded}


# ---- hanger rods ----

def rod_inches(v):
    t = _s(v).strip()
    if t == '':
        return None
    if re.match(r'^-?\d+(\.\d+)?$', t):
        return float(t) * 12          # a bare number is feet, Revit's own unit
    p = parse_inches(t)
    return p[0] if p else None


def rod_size_inches(v):
    t = _s(v).strip()
    if t == '':
        return None
    if re.match(r'^-?\d+(\.\d+)?$', t):
        n = float(t)
        if 0.25 <= n * 12 <= 2.5:
            return n * 12
        if 0.25 <= n <= 2.5:
            return n
        return n * 12
    p = parse_inches(t)
    return p[0] if p else None


def summarize_rods(rows):
    by, order = {}, []
    with_a = with_b = rounded = unreadable = 0

    def present(v):
        return v is not None and _s(v).strip() != ''
    for r in rows:
        a = rod_inches(r.get('Length A'))
        b = rod_inches(r.get('Length B'))
        if present(r.get('Length A')) and a is None:
            unreadable += 1
        if present(r.get('Length B')) and b is None:
            unreadable += 1
        if a and a > 0:
            with_a += 1
        if b and b > 0:
            with_b += 1
        ln = (a if a and a > 0 else 0) + (b if b and b > 0 else 0)
        if not ln:
            continue
        s = rod_size_inches(r.get('Rod Size'))
        key = '?' if s is None else str(int(s * 64 + 0.5))
        if key not in by:
            by[key] = {'size': s, 'inches': 0}
            order.append(key)
        by[key]['inches'] += ln
    order.sort(key=lambda k: (1, 0) if k == '?' else (0, int(k)))
    no_size = 0
    lines = []
    for k in order:
        e = by[k]
        tot = to_feet_inches(e['inches'])
        if tot['rounded']:
            rounded += 1
        if e['size'] is None:
            no_size += 1
        lines.append({
            'size': '' if e['size'] is None else to_fractional_inches(e['size'], 16)['text'],
            'description': 'ALL THREAD ROD (rod size not found)' if e['size'] is None else 'ALL THREAD ROD',
            'length': tot['text']})
    return {'lines': lines, 'withA': with_a, 'withB': with_b, 'noSize': no_size, 'rounded': rounded,
            'unreadable': unreadable}


def _level_key(v):
    return _s(v).strip().lower()


def _top(counter, n=8):
    items = sorted(counter.items(), key=lambda kv: -kv[1])[:n]
    return ', '.join('%s (%d)' % (k, c) for k, c in items)


def select_parts(d, parts, wanted=None, hidden=None, sel=None, excluded_services=None):
    """Which parts an export includes. The same rule feeds the MR forms and the 3D preview.
    wanted: {level key: name}; hidden: set of workset names; sel: set of element ids or None.
    Returns (items, kept, skipped, no_level, other_level, notes); kept are the build_row results that stay."""
    if hidden:
        parts = [p for p in parts if not p.workset or p.workset not in hidden]
    if sel is not None:
        parts = [p for p in parts if p.id in sel]
    if excluded_services:
        parts = [p for p in parts if service_label(p) not in excluded_services]
    items = [build_row(d, p) for p in parts]
    notes = {'rounded': 0, 'assumed': 0, 'unparsed': 0, 'cleared': 0}
    kept, skipped, no_level, other_level = [], 0, 0, 0
    for it in items:
        if it['skip']:
            skipped += 1
            continue
        if wanted:
            k = _level_key(it['level'])
            if not k:
                no_level += 1
                continue
            if k not in wanted:
                other_level += 1
                continue
        kept.append(it)
        for n in notes:
            notes[n] += it['notes'][n]
    return items, kept, skipped, no_level, other_level, notes


def run_exports(sources, keys, levels=None, hidden_worksets=None, selection_ids=None, excluded_services=None):
    """sources: {'pipe': [Part], 'hangers': [Part]}; keys: export keys.
    Returns {label: {'rows': [...], 'rods': [...], 'warnings': [...], 'parts': n}} and a list of summary warnings."""
    wanted = {_level_key(l): l for l in (levels or [])}
    hidden = set(hidden_worksets or [])
    sel = set(selection_ids) if selection_ids else None
    excl = set(excluded_services or [])
    den = FRACTION_DENOMINATOR
    results = {}
    for key in keys:
        d = type_by_key(key)
        if not d:
            continue
        warn = []
        parts = list(sources.get(d['source'], []))
        # Rows are built once. The service filter is applied to the result, so excluding a service costs nothing extra and the
        # note can say how many parts it removed.
        items, would_keep, skipped, no_level, other_level, _ = select_parts(d, parts, wanted, hidden, sel, None)
        kept = [it for it in would_keep if service_label(it['part']) not in excl] if excl else would_keep
        notes = {'rounded': 0, 'assumed': 0, 'unparsed': 0, 'cleared': 0}
        for it in kept:
            for n in notes:
                notes[n] += it['notes'][n]
        if excl:
            names = ', '.join(sorted(excl))
            left_out = len(would_keep) - len(kept)
            if left_out:
                warn.append('Service filter: %d parts left out (%s)' % (left_out, names))
            else:
                seen = sorted(set(service_label(it['part']) for it in would_keep))
                warn.append('Service filter on (%s) but it removed no parts from this export. Services in this export: %s'
                            % (names, ', '.join(seen) or 'none'))
        f_rows = [it['row'] for it in kept]

        rows, rounded_sums = aggregate(d, f_rows)
        if not rows and items:
            seen = {}
            for it in items:
                m = _s(it['row'].get('Part Material')).strip() or '(none)'
                seen[m] = seen.get(m, 0) + 1
            warn.append('no parts matched. Part Material on the %d parts read: %s' % (len(items), _top(seen)))
        if d.get('excludeWhen'):
            for h in d['excludeWhen']:
                seen = {}
                for it in items:
                    v = _s(it['row'].get(h)).strip() or '(none)'
                    seen[v] = seen.get(v, 0) + 1
                warn.append('%s values found: %s' % (h, _top(seen)))
        if d.get('combinePipe'):
            rows, comb = combine_pipe_rows(rows, d)
            if comb and comb['rows'] != comb['sets']:
                warn.append('%d pipe rows (end conditions dropped, stub-ins counted under their branch size) combined into %d line%s%s'
                            % (comb['rows'], comb['sets'], '' if comb['sets'] == 1 else 's',
                               '; %d totals rounded to the nearest 1/%d in' % (comb['rounded'], den) if comb['rounded'] else ''))
        if d.get('noPipeQty'):
            for r in rows:
                has_len = _s(r.get('Length')).strip() != '' or (isinstance(r.get(RAW), (int, float)) and r[RAW] > 0)
                if has_len:
                    r['Qty'] = ''
        rods = []
        if d.get('rodSummary'):
            rs = summarize_rods(f_rows)
            rods = rs['lines']
            if rs['unreadable']:
                warn.append('%d FP_Rod Length A / B values could not be read as lengths and were not added to the rod totals' % rs['unreadable'])
            if not rods:
                warn.append('no rod lengths found (FP_Rod Length A and FP_Rod Length B were not found on the hangers, or are all zero)')
            else:
                warn.append('rod lengths added up from FP_Rod Length A on %d and FP_Rod Length B on %d hangers%s%s'
                            % (rs['withA'], rs['withB'],
                               '; some have no rod size (shown as "rod size not found")' if rs['noSize'] else '',
                               '; %d totals rounded to the nearest 1/%d in' % (rs['rounded'], den) if rs['rounded'] else ''))
        if not parts:
            warn.append('no elements found in this model for the current filters')
        if notes['rounded']:
            warn.append('%d sizes/lengths rounded to the nearest 1/%d in' % (notes['rounded'], den))
        if notes['assumed']:
            warn.append('%d sizes/lengths had no unit and were read as inches' % notes['assumed'])
        if skipped:
            warn.append('%d %s' % (skipped, d.get('skipNote') or 'empty parts left out'))
        if wanted and (no_level or other_level):
            warn.append('level filter: %d parts on other levels and %d with no level left out' % (other_level, no_level))
        if rounded_sums:
            warn.append('%d summed lengths rounded to the nearest 1/%d in' % (rounded_sums, den))
        if notes['cleared'] and not d.get('quietCleared'):
            warn.append('%d fittings and olets shown without a length' % notes['cleared'])
        if notes['unparsed']:
            warn.append('%d sizes/lengths could not be read as inches and were left as given' % notes['unparsed'])
        blank = [c['header'] for c in d['columns'] if not c.get('blank') and c['header'] not in ('Comments', 'Workset')
                 and f_rows and all(_s(r.get(c['header'])).strip() == '' for r in f_rows)]
        if blank:
            warn.append('%s blank on every element (property name may differ)' % ', '.join(blank))

        results[d['label']] = {'key': key, 'rows': rows, 'rods': rods, 'warnings': warn, 'parts': len(kept)}
    return results
