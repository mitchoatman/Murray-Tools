#!/usr/bin/env python3
r"""
fabmap_editor.py - Fabrication service.map palette editor (2026 format). ONE FILE, no other files needed.

pyRevit: run with no arguments. A browser page opens: pick a profile, drag buttons to reorder, drop a button on
another tab's name to copy it, right-click a tile for Delete / Exclude From Fill, drag parts in from the Libraries pane.
Every change is written straight into the picked profile's service.map (no backup files, no temp files).
Undo (button in the page) restores the previous state for the current session.

Settings come from C:\temp\Ribbon_FABCADMEP-Path.json (database_path) and MAP.INI in that folder.
Command line (no arguments needed in pyRevit):
    python fabmap_editor.py                            open the editor
    python fabmap_editor.py service.map --html out.html   read-only viewer page
    python fabmap_editor.py service.map --selftest        consistency tests of the writer
    python fabmap_editor.py service.map --move|--copy|--delete ... --out new.map   command-line edits
"""
import json, os, re, struct, sys, time, zlib

r"""
fabmap.py - READ-ONLY reader for the Fabrication 'service.map' file (2026 format).

pyRevit use: run with no arguments. The database folder comes from C:\temp\Ribbon_FABCADMEP-Path.json
(key database_path; created on first run). The Items and Images folders and the Database folder holding
service.map are read from MAP.INI in that folder, and the viewer opens in your browser. Shift-click picks a
different database folder, ctrl-click picks a different service.map. Nothing is written to the map file.

Command line:
    python fabmap.py service.map --summary
    python fabmap.py service.map --json services.json
    python fabmap.py service.map --html viewer.html
    python fabmap.py service.map --find "storm"

Palettes (service templates) are read from the second half of the file: template > tab > button >
item path (.ITM) and image path (.png). Images are loaded from your Fabrication ITEMS folder.

Format as decoded so far (see FINDINGS in the chat; items marked [CONFIRM] are inferred):
  * 30-byte header "MAP Compressed File (c) 2005..." followed by zlib blocks.
  * Decompressed data is a tagged object stream. Tags used here (little endian dwords):
      0x006C090B  service collection : [tag][size][size-4][count]...
      0x006C090A  service            : [tag][size][payload of size bytes]
      0x006C0909  service entry      : [tag][size][payload]
      0x001118F6  button list        : [tag][size][size-4][count][buttons...]
      0x001118F5  button             : [tag][size][payload]
  * Strings are [int32 n incl. NUL][n UTF-16LE chars].
This module never writes. Writing is deliberately not implemented yet.
"""
import json, re, struct, sys, zlib

TAG_COLLECTION = b'\x0b\x09\x6c\x00'
TAG_SERVICE    = b'\x0a\x09\x6c\x00'
TAG_ENTRY      = b'\x09\x09\x6c\x00'
TAG_BTNLIST    = b'\xf6\x18\x11\x00'
TAG_BUTTON     = b'\xf5\x18\x11\x00'
TPL_TAG = b'\xdd\x03\x5c\x00'   # service template  [tag][size][payload]
TPL_BTN = b'\xdf\x03\x5c\x00'   # palette button    [tag][size][payload]
TPL_TAB = b'\xc1\x33\x20\x00'   # palette tab       [tag][size][payload]  (button count sits just before the first button)


def decompress(path):
    with open(path, 'rb') as f:
        return decompress_bytes(f.read())


def decompress_bytes(d):
    if not d.startswith(b'MAP Compressed File'):
        raise ValueError('Not a MAP Compressed File')
    pos = d.find(b'\x78\x9c')
    out = []
    while pos >= 0 and pos < len(d):
        o = zlib.decompressobj()
        out.append(o.decompress(d[pos:]))
        pos = len(d) - len(o.unused_data)
        nxt = d.find(b'\x78\x9c', pos, pos + 64)
        if nxt < 0:
            break
        pos = nxt
    return b''.join(out)


def _str_at(raw, o, limit):
    """Return (string, end) if a valid length-prefixed UTF-16 string starts at o, else (None, o)."""
    if o + 4 > limit:
        return None, o
    n = struct.unpack_from('<I', raw, o)[0]
    if not (0 < n < 400) or o + 4 + 2 * n > limit:
        return None, o
    b = raw[o + 4:o + 4 + 2 * n]
    if b[-2:] != b'\x00\x00' or any(b[i + 1] != 0 for i in range(0, len(b), 2)):
        return None, o
    s = b.decode('utf-16le').rstrip('\0')
    if any(ord(c) < 32 for c in s):
        return None, o
    return s, o + 4 + 2 * n


def _strings(raw, start, end):
    o, out = start, []
    while o < end:
        s, e = _str_at(raw, o, end)
        if s is not None:
            out.append(s)
            o = e
        else:
            o += 1
    return out


def parse(raw):
    """Parse all services. Returns (services, checks)."""
    c = raw.find(TAG_COLLECTION)
    if c < 0:
        raise ValueError('service collection tag not found')
    coll_size, _, coll_count = struct.unpack_from('<III', raw, c + 4)
    coll_end = c + 8 + coll_size
    starts = []
    i = c + 4
    while True:
        i = raw.find(TAG_SERVICE, i, coll_end)
        if i < 0:
            break
        starts.append(i)
        i += 4
    services, checks = [], {'collection_count': coll_count, 'service_tags': len(starts), 'problems': []}
    for k, s in enumerate(starts):
        size = struct.unpack_from('<I', raw, s + 4)[0]
        end = s + 8 + size
        if k + 1 < len(starts) and end != starts[k + 1]:
            checks['problems'].append('service %d size does not reach next record' % k)
        # header strings: name, group, code (first three strings before first entry/button list)
        first_entry = raw.find(TAG_ENTRY, s + 8, end)
        first_btn = raw.find(TAG_BTNLIST, s + 8, end)
        hdr_end = min([x for x in (first_entry, first_btn, end) if x > 0])
        # empty strings are ambiguous with small integers in the header, so only real text counts here
        hs = [x for x in _strings(raw, s + 8, hdr_end) if x]
        svc = {'index': k, 'offset': s, 'name': hs[0] if hs else '', 'group': hs[1] if len(hs) > 1 else '',
               'code': hs[2] if len(hs) > 2 else '', 'entries': [], 'buttons': []}
        # entries
        j = s + 8
        while True:
            j = raw.find(TAG_ENTRY, j, end)
            if j < 0:
                break
            esz = struct.unpack_from('<I', raw, j + 4)[0]
            eend = min(j + 8 + esz, end)
            svc['entries'].append({'strings': _strings(raw, j + 8, eend)})
            j = eend
        # buttons
        bl = raw.find(TAG_BTNLIST, s + 8, end)
        declared = None
        if bl >= 0:
            lsz = struct.unpack_from('<I', raw, bl + 4)[0]
            declared = struct.unpack_from('<I', raw, bl + 12)[0]
            lend = min(bl + 8 + lsz, end)
            j = bl + 16
            while True:
                j = raw.find(TAG_BUTTON, j, lend)
                if j < 0:
                    break
                bsz = struct.unpack_from('<I', raw, j + 4)[0]
                bend = min(j + 8 + bsz, lend)
                st = _strings(raw, j + 8, bend)
                svc['buttons'].append({'name': st[0] if st else '', 'code': st[1] if len(st) > 1 else '',
                                       'separator': bool(st) and st[0].startswith('---')})
                j = bend
            if declared != len(svc['buttons']):
                checks['problems'].append('service %r: declared %s buttons, parsed %d' % (svc['name'], declared, len(svc['buttons'])))
        svc['declared_buttons'] = declared
        services.append(svc)
    checks['services'] = len(services)
    checks['entries'] = sum(len(x['entries']) for x in services)
    checks['buttons'] = sum(len(x['buttons']) for x in services)
    checks['last_record_reaches_collection_end'] = bool(starts) and (
        struct.unpack_from('<I', raw, starts[-1] + 4)[0] + starts[-1] + 8 == coll_end)
    return services, checks


_RUN = re.compile(rb'(?:[\x20-\x7e]\x00){3,}')


def _first_string(raw, a, b, limit=90):
    o, stop = a, min(b, a + limit)
    while o < stop:
        s, e = _str_at(raw, o, b)
        if s:
            return s
        o += 1
    return ''


def _all(raw, tag, a, b):
    out, i = [], a
    while True:
        i = raw.find(tag, i, b)
        if i < 0:
            return out
        out.append(i)
        i += 1


def _button_info(raw, b, be):
    """Name, .ITM path(s) and stored image of one palette button record raw[b:be]."""
    bname = _first_string(raw, b + 8, be, 70)
    itms, pngs = [], []
    for m in _RUN.finditer(raw[b:be]):
        s = m.group(0).decode('utf-16le')
        l = s.lower()
        if l.endswith('.itm') and s not in itms and len(itms) < 8:
            itms.append(s)
        elif l.endswith(('.png', '.bmp', '.jpg', '.jpeg')) and s not in pngs and len(pngs) < 2:
            pngs.append(s)
    png = pngs[0] if pngs else ''          # key kept as 'png' for compatibility; may be .bmp/.jpg
    if itms and png.lower().endswith('.png') and png.rsplit('.', 1)[0].lower() == itms[0].rsplit('.', 1)[0].lower():
        png = ''          # same stem and .png as the item: the viewer derives it
    return {'name': bname, 'items': itms, 'png': png, 'png_missing': not pngs}


def parse_templates(raw, start_search=0):
    """Service templates -> tabs -> buttons. Returns (templates, checks)."""
    first = raw.find(TPL_TAG, start_search)
    if first < 0:
        return [], {'templates': 0, 'problems': ['no service template section found']}
    templates, problems = [], []
    o = first
    end_data = len(raw.rstrip(b'\0'))
    while raw[o:o + 4] == TPL_TAG:
        size = struct.unpack_from('<I', raw, o + 4)[0]
        te = o + 8 + size
        if te > len(raw):
            problems.append('template at %d runs past end of data' % o)
            break
        tabs_at = _all(raw, TPL_TAB, o + 8, te)
        hdr_end = tabs_at[0] if tabs_at else te
        hs, p = [], o + 8
        while p < hdr_end and len(hs) < 4:
            s, e = _str_at(raw, p, hdr_end)
            if s:
                hs.append(s); p = e
            else:
                p += 1
        tpl = {'name': hs[0] if hs else '', 'meta': hs[1:], 'offset': o, 'tabs': []}
        for x in tabs_at:
            xsz = struct.unpack_from('<I', raw, x + 4)[0]
            xe = min(x + 8 + xsz, te)
            nm = _first_string(raw, x + 8, xe, 60)
            if not nm:
                continue
            tab = {'name': nm, 'buttons': []}
            btns = _all(raw, TPL_BTN, x, xe)
            if btns:
                declared = struct.unpack_from('<I', raw, btns[0] - 4)[0]
                if declared != len(btns):
                    problems.append('template %r tab %r: declared %d buttons, found %d' % (tpl['name'], nm, declared, len(btns)))
            for b in btns:
                bsz = struct.unpack_from('<I', raw, b + 4)[0]
                be = min(b + 8 + bsz, xe)
                tab['buttons'].append(_button_info(raw, b, be))
            tpl['tabs'].append(tab)
        templates.append(tpl)
        o = te
    trailing = end_data - o
    total_btn_tags = len(_all(raw, TPL_BTN, first, o))
    got = sum(len(tb['buttons']) for tp in templates for tb in tp['tabs'])
    if got != total_btn_tags:
        problems.append('buttons parsed %d but %d button tags in section' % (got, total_btn_tags))
    checks = {'templates': len(templates), 'tabs': sum(len(tp['tabs']) for tp in templates), 'buttons': got,
              'button_tags_in_section': total_btn_tags, 'trailing_bytes_after_templates': trailing,
              'problems': problems}
    return templates, checks


_IMG_EXT = ('.png', '.bmp', '.jpg', '.jpeg')


_EXT16 = re.compile(rb'\.\x00(?:[pP]\x00[nN]\x00[gG]\x00|[bB]\x00[mM]\x00[pP]\x00|[jJ]\x00[pP]\x00[gG]\x00|[jJ]\x00[pP]\x00[eE]\x00[gG]\x00)')
_EXT8 = re.compile(rb'\.(?:png|bmp|jpe?g)', re.I)


def _itm_image_refs(itm_file):
    """Best effort: image names stored inside an .ITM. Format not confirmed [CONFIRM]; read-only."""
    import os
    try:
        if os.path.getsize(itm_file) > 30000000:
            return []
        with open(itm_file, 'rb') as f:
            data = f.read()
        if data.startswith(b'MAP Compressed File'):
            data = decompress_bytes(data)
    except Exception:
        return []
    refs = []
    # look for the picture file extensions first (a fast literal-style search), then read the name backwards from there
    for m in _EXT16.finditer(data):
        j, st = m.start(), m.start()
        while j >= 2 and (st - j) // 2 < 260 and data[j - 1] == 0 and 0x20 <= data[j - 2] < 0x7f:
            j -= 2
        s = data[j:m.end()].decode('utf-16le', 'replace')
        if len(s) >= 5 and s not in refs:
            refs.append(s)
    for m in _EXT8.finditer(data):
        j, st = m.start(), m.start()
        while j >= 1 and (st - j) < 260 and 0x20 <= data[j - 1] < 0x7f:
            j -= 1
        s = data[j:m.end()].decode('ascii', 'replace')
        if len(s) >= 5 and s not in refs:
            refs.append(s)
    return refs[:6]


def resolve_images(templates, items_root, images_root='', itm_first=False, cache_path='', budget_s=20, memo=None, shared=None, itm_store=None):
    """Find the real image file for each palette button. Adds b['img'] (absolute path or '').
    Default (fast): what service.map stores and the item's own name beside it; an .ITM is opened only when no
    picture was found that way. itm_first=True reads the .ITM of every button first (slow on a large database).
    .ITM reading stops after budget_s seconds. Results are cached in cache_path (keyed by file size and date)."""
    import os, time
    t_start = time.time()
    stats = {'found': 0, 'missing': 0, 'via': {}, 'missing_examples': [], 'missing_rows': [],
             'disagree_rows': [], 'itm_files_read': 0, 'itm_cache_hits': 0, 'itm_first': bool(itm_first), 'itm_skipped': 0, 'dirs_listed': 0}
    seen = shared.setdefault('seen', {}) if shared is not None else {}

    dir_names = shared.setdefault('dirs', {}) if shared is not None else {}

    def isf(p):
        # one folder listing answers every lookup in that folder (much faster than a call per file on a network share)
        if p not in seen:
            d, f = os.path.split(p)
            names = dir_names.get(d)
            if names is None:
                _yield_to_user()
                try:
                    names = frozenset(n.lower() for n in os.listdir(d))
                except Exception:
                    names = frozenset()
                dir_names[d] = names
                stats['dirs_listed'] += 1
            seen[p] = f.lower() in names
        return seen[p]

    def j(base, rel):
        rel = rel.replace('\\', '/')
        if rel.startswith('./'):
            rel = rel[2:]
        return os.path.normpath(os.path.join(base, rel))

    if itm_store is not None:                # caller keeps the .ITM scan results (in memory, saved elsewhere)
        itm_cache = itm_store.setdefault('d', {})
    else:
        itm_cache = _read_json(cache_path) if cache_path else {}
    dirty = [False]

    def refs_for(itm_abs):
        """Image names stored in the .ITM, or None if the file is missing."""
        _yield_to_user()
        try:
            st = os.stat(itm_abs)
        except Exception:
            return None
        ent = itm_cache.get(itm_abs)
        if ent and ent.get('m') == st.st_mtime and ent.get('s') == st.st_size:
            stats['itm_cache_hits'] += 1
            return ent.get('refs', [])
        refs = _itm_image_refs(itm_abs)
        stats['itm_files_read'] += 1
        itm_cache[itm_abs] = {'m': st.st_mtime, 's': st.st_size, 'refs': refs}
        dirty[0] = True
        return refs

    def resolve_ref(ref, itm_abs):
        r = ref.replace('\\', '/')
        opts = []
        if re.match(r'^[A-Za-z]:[\\/]|^[\\/]{2}', ref):
            opts.append(ref)
        opts.append(os.path.normpath(os.path.join(os.path.dirname(itm_abs), r)))
        if images_root:
            opts.append(os.path.join(images_root, os.path.basename(r)))
        opts.append(j(items_root, ref))
        if images_root:
            opts.append(j(images_root, ref))
            opts.append(j(os.path.dirname(images_root.rstrip('\\/')), ref))
        for p in opts:
            if isf(p):
                return p
        return ''

    itm_result = {}      # itm_abs -> (chosen path, refs list or None)

    def from_itm(itm_abs):
        if itm_abs not in itm_result:
            if time.time() - t_start > budget_s:
                stats['itm_skipped'] += 1
                itm_result[itm_abs] = ('', 'SKIP')
                return itm_result[itm_abs]
            refs = refs_for(itm_abs)
            hit = ''
            for ref in (refs or []):
                hit = resolve_ref(ref, itm_abs)
                if hit:
                    break
            itm_result[itm_abs] = (hit, refs)
        return itm_result[itm_abs]

    for tp in templates:
        for tb in tp['tabs']:
            for b in tb['buttons']:
                png = b['png']
                itm = b['items'][0] if b['items'] else ''
                if memo is not None and (png, itm) in memo:       # already looked up in this session
                    b['img'] = memo[(png, itm)]
                    stats['found' if b['img'] else 'missing'] += 1
                    continue
                itm_abs = j(items_root, itm) if itm else ''
                cands = []          # (path, how) from service.map and the item's own name
                if png:
                    if re.match(r'^[A-Za-z]:[\\/]|^[\\/]{2}', png):
                        cands.append((png, 'stored absolute'))
                        if images_root:
                            cands.append((os.path.join(images_root, os.path.basename(png.replace('\\', '/'))), 'stored absolute, name in Images folder'))
                    elif '/' in png.replace('\\', '/'):
                        cands.append((j(items_root, png), 'stored path'))
                    else:
                        if images_root:
                            cands.append((os.path.join(images_root, png), 'name in Images folder'))
                        if itm:
                            cands.append((os.path.join(os.path.dirname(itm_abs), png), 'beside item'))
                        cands.append((os.path.join(items_root, png), 'ITEMS root'))
                if itm:
                    for ext in ('.png', '.bmp', '.jpg', '.jpeg'):
                        cands.append((os.path.splitext(itm_abs)[0] + ext, 'beside item'))
                map_hit, map_how = '', ''
                for p, h in cands:
                    if isf(p):
                        map_hit, map_how = p, h
                        break
                itm_hit, refs = ('', None)
                if itm and (itm_first or not map_hit):
                    itm_hit, refs = from_itm(itm_abs)
                if itm_first:
                    chosen, how = (itm_hit, 'named inside item') if itm_hit else (map_hit, map_how)
                else:
                    chosen, how = (map_hit, map_how) if map_hit else (itm_hit, 'named inside item')
                if itm_hit and map_hit and os.path.normcase(os.path.normpath(itm_hit)) != os.path.normcase(os.path.normpath(map_hit)):
                    stats['disagree_rows'].append([tp['name'], tb['name'], b['name'], itm, itm_hit, map_hit, map_how])
                b['img'] = chosen.replace('\\', '/') if chosen else ''
                if memo is not None and refs != 'SKIP':
                    memo[(png, itm)] = b['img']
                if chosen:
                    stats['found'] += 1
                    stats['via'][how] = stats['via'].get(how, 0) + 1
                else:
                    stats['missing'] += 1
                    stats['missing_rows'].append([tp['name'], tb['name'], b['name'], itm, png,
                                                  ('not checked (time limit)' if refs == 'SKIP' else ('yes' if refs is not None else 'NO')),
                                                  '' if refs == 'SKIP' else '; '.join((refs or [])[:3]),
                                                  cands[0][0] if cands else ''])
                    if len(stats['missing_examples']) < 12:
                        stats['missing_examples'].append('%s | %s' % (tp['name'], b['name']))
    if dirty[0] and itm_store is not None:
        itm_store['dirty'] = True
    if dirty[0] and cache_path and itm_store is None:
        _write_json(cache_path, itm_cache)
    stats['distinct_items_missing'] = len({r[3] for r in stats['missing_rows']})
    stats['disagree'] = len(stats['disagree_rows'])
    stats['seconds'] = round(time.time() - t_start, 1)
    return stats


HTML = r"""<!DOCTYPE html><html lang="en"><head><meta charset="utf-8"><title>Fabrication service.map viewer</title>
<style>
:root{--y:#FFE716;--bg:#fff;--fg:#000}
*{box-sizing:border-box}body{margin:0;font-family:Arial,sans-serif;color:var(--fg);background:var(--bg)}
header{background:var(--y);padding:8px 14px;border-bottom:2px solid #000;display:flex;gap:14px;align-items:center;flex-wrap:wrap}
header b{font-size:15px}header .m{font-size:12px}
.tabsrow button,#bar button{font:inherit;border:2px solid #000;background:#fff;padding:4px 10px;cursor:pointer}
.tabsrow button.on{background:var(--y);font-weight:bold}
#bar{padding:6px 14px;border-bottom:1px solid #000;display:flex;gap:8px;align-items:center;font-size:12px}
#root{flex:1;font:inherit;padding:4px;border:1px solid #000}
#wrap{display:flex;height:calc(100vh - 100px)}
#left{width:330px;border-right:2px solid #000;display:flex;flex-direction:column}
#q{margin:8px;padding:6px;font:inherit;border:1px solid #000}
#list{overflow:auto;flex:1}.row{padding:6px 10px;border-bottom:1px solid #000;cursor:pointer;font-size:13px}
.row small{font-size:11px}.row.sel{background:var(--y)}
#main{flex:1;overflow:auto;padding:12px 16px}
h2{margin:0 0 2px}.note{font-size:12px}
.ptabs{display:flex;gap:0;margin:10px 0 0}.ptabs button{font:inherit;font-size:12px;border:1px solid #000;border-bottom:0;background:#fff;padding:4px 14px;cursor:pointer}
.ptabs button.on{background:var(--y);font-weight:bold}
.grid{display:flex;flex-wrap:wrap;border-top:2px solid #000;padding-top:2px;gap:2px}
.tile{width:112px;height:128px;border:2px solid #000;background:#fff;display:flex;flex-direction:column;align-items:center;cursor:default;position:relative}
.tile:hover{background:var(--y)}
.tile .im{height:74px;width:100%;display:flex;align-items:center;justify-content:center;padding:4px}
.tile img{max-width:100%;max-height:100%}
.tile .no{font-size:10px;border:1px dashed #000;padding:14px 6px}
.tile .lb{font-size:10px;text-transform:uppercase;text-align:center;line-height:1.15;overflow:hidden;padding:0 3px;flex:1;width:100%;word-break:break-word}
.tile .x{position:absolute;right:2px;top:2px;background:var(--y);border:1px solid #000;font-size:10px;padding:0 3px}
.tile.sep{background:#000;color:#fff;width:auto;min-width:112px;justify-content:center}
.tile.sep:hover{background:#000}.tile.sep .lb{color:#fff;flex:0;font-weight:bold}
table{border-collapse:collapse;margin:8px 0 16px}th{background:var(--y);text-align:left;border:1px solid #000;padding:3px 8px}td{border:1px solid #000;padding:3px 8px}
</style></head><body>
<header><b>Fabrication service.map viewer (read-only)</b><span class="m" id="meta"></span>
<span class="tabsrow"><button id="vP" class="on">Palettes</button><button id="vS">Services</button></span></header>
<div id="bar"><span>Items folder:</span><input id="root" placeholder="Folder that contains 'Piping Components' (your Fabrication ITEMS folder)"><button id="apply">Apply</button><span id="rootnote"></span></div>
<div id="wrap"><div id="left"><input id="q" placeholder="Search"><div id="list"></div></div><div id="main"></div></div>
<script>
const D=__DATA__;
const $=id=>document.getElementById(id);
let view='P',selT=-1,selTab=0,selS=-1,root=D.root||'';
try{if(!root)root=localStorage.getItem('fabmap_root')||''}catch(e){}
$('root').value=root;
$('meta').textContent=D.templates.length+' palettes, '+D.checks.tabs+' tabs, '+D.checks.buttons+' buttons | '+D.services.length+' services';
function esc(s){return String(s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]))}
function enc(p){return p.split('/').map(encodeURIComponent).join('/')}
function fileUrl(rel){
  if(!root)return '';let r=root.replace(/\\/g,'/').replace(/\/+$/,'');let u;
  if(r.startsWith('//'))u='file:'+'//'+r.slice(2).split('/').map(encodeURIComponent).join('/');
  else{const parts=r.split('/');const h=/^[A-Za-z]:$/.test(parts[0])?parts.shift()+'/':'';u='file:///'+h+parts.map(encodeURIComponent).join('/')}
  return u+'/'+enc(rel.replace(/^\.\//,''));}
function absUrl(p){p=p.replace(/\\/g,'/');if(p.startsWith('//'))return 'file:'+'//'+p.slice(2).split('/').map(encodeURIComponent).join('/');
  const parts=p.split('/');const h=/^[A-Za-z]:$/.test(parts[0])?parts.shift()+'/':'';return 'file:///'+h+parts.map(encodeURIComponent).join('/')}
function pngOf(b){if(b.png)return b.png;if(b.items.length)return b.items[0].replace(/\.itm$/i,'.png');return ''}
function tile(b){
  const d=document.createElement('div');d.className='tile'+(b.name.startsWith('---')?' sep':'');
  d.title=b.items.length?b.items.join('\n'):'(no item path)';
  if(b.items.length>1){const x=document.createElement('span');x.className='x';x.textContent='x'+b.items.length;d.appendChild(x)}
  const im=document.createElement('div');im.className='im';
  if(!d.classList.contains('sep')){const p=pngOf(b),u=p?fileUrl(p):'';
    if(u){const i=document.createElement('img');i.loading='lazy';i.src=u;i.onerror=()=>{im.innerHTML='';const n=document.createElement('div');n.className='no';n.textContent='no image';im.appendChild(n)};im.appendChild(i)}
    else{const n=document.createElement('div');n.className='no';n.textContent=(root||D.resolved)?'no image':'set items folder';im.appendChild(n)}
    d.appendChild(im)}
  const l=document.createElement('div');l.className='lb';l.textContent=b.name;d.appendChild(l);return d}
function showP(i){selT=i;selTab=0;drawP()}
function drawP(){
  document.querySelectorAll('#list .row').forEach(r=>r.classList.toggle('sel',+r.dataset.i===selT));
  const m=$('main');m.innerHTML='';if(selT<0){m.textContent='Select a palette.';return}
  const T=D.templates[selT];
  const h=document.createElement('div');h.innerHTML='<h2>'+esc(T.name)+'</h2><div class="note">'+esc(T.meta.join(' | '))+' | '+T.tabs.length+' tabs</div>';m.appendChild(h);
  const tb=document.createElement('div');tb.className='ptabs';
  T.tabs.forEach((x,k)=>{const b=document.createElement('button');b.textContent=x.name+' ('+x.buttons.length+')';if(k===selTab)b.className='on';b.onclick=()=>{selTab=k;drawP()};tb.appendChild(b)});m.appendChild(tb);
  const g=document.createElement('div');g.className='grid';
  if(T.tabs[selTab])T.tabs[selTab].buttons.forEach(b=>g.appendChild(tile(b)));m.appendChild(g)}
function showS(i){selS=i;drawS()}
function drawS(){
  document.querySelectorAll('#list .row').forEach(r=>r.classList.toggle('sel',+r.dataset.i===selS));
  const m=$('main');if(selS<0){m.textContent='Select a service.';return}
  const s=D.services[selS];const pal=D.templates.findIndex(t=>t.name===s.name);
  let h='<h2>'+esc(s.name)+'</h2><div class="note">Group: '+esc(s.group)+' | Service code: '+esc(s.code)+' | Entries: '+s.entries.length+' | Button mappings: '+s.buttons.length+(pal>=0?' | A palette with the same name exists':'')+'</div>';
  h+='<h3>Button mappings</h3><div class="note">Meaning of the second column is not confirmed.</div><table><tr><th>#</th><th>Name</th><th>Code</th></tr>';
  s.buttons.forEach((b,n)=>{h+='<tr><td>'+(n+1)+'</td><td>'+esc(b.name)+'</td><td>'+esc(b.code)+'</td></tr>'});
  h+='</table><h3>Service entries</h3><div class="note">Columns follow the API property order. Mapping is inferred, not confirmed.</div><table><tr><th>#</th><th>LayerTag1</th><th>LayerTag2</th><th>LevelBlock</th><th>SizeBlock</th></tr>';
  s.entries.forEach((e,n)=>{h+='<tr><td>'+(n+1)+'</td>'+[0,1,2,3].map(k=>'<td>'+esc(e.strings[k]||'')+'</td>').join('')+'</tr>'});
  m.innerHTML=h+'</table>'}
function renderList(){const t=$('q').value.toLowerCase(),L=$('list');L.innerHTML='';
  if(view==='P')D.templates.forEach((p,i)=>{const hay=(p.name+' '+p.meta.join(' ')+' '+p.tabs.map(x=>x.name+' '+x.buttons.map(b=>b.name).join(' ')).join(' ')).toLowerCase();
    if(t&&!hay.includes(t))return;const r=document.createElement('div');r.className='row';r.dataset.i=i;
    r.innerHTML=esc(p.name)+'<br><small>'+p.tabs.length+' tabs, '+p.tabs.reduce((a,x)=>a+x.buttons.length,0)+' buttons</small>';r.onclick=()=>showP(i);L.appendChild(r)});
  else D.services.forEach((s,i)=>{const hay=(s.name+' '+s.group+' '+s.code).toLowerCase();if(t&&!hay.includes(t))return;
    const r=document.createElement('div');r.className='row';r.dataset.i=i;r.innerHTML=esc(s.name)+'<br><small>'+esc(s.group)+' | '+esc(s.code)+'</small>';r.onclick=()=>showS(i);L.appendChild(r)})}
function setView(v){view=v;$('vP').className=v==='P'?'on':'';$('vS').className=v==='S'?'on':'';renderList();v==='P'?drawP():drawS()}
$('vP').onclick=()=>setView('P');$('vS').onclick=()=>setView('S');$('q').oninput=renderList;
$('apply').onclick=()=>{root=$('root').value.trim();try{localStorage.setItem('fabmap_root',root)}catch(e){}$('rootnote').textContent=root?'Images load from this folder plus the png path stored in service.map.':'';if(view==='P')drawP()};
$('rootnote').textContent=root?'Images load from this folder plus the png path stored in service.map.':'Set the items folder to show button images.';
setView('P');
</script></body></html>"""


def _bundle(raw, root='', images_root='', itm_first=False, cache_path=''):
    import os, time
    t0 = time.time()
    services, checks = parse(raw)
    templates, tchecks = parse_templates(raw)
    checks['parse_seconds'] = round(time.time() - t0, 1)
    resolved = False
    if root and os.path.isdir(root):
        checks['images'] = resolve_images(templates, root, images_root if images_root and os.path.isdir(images_root) else '', itm_first, cache_path)
        resolved = True
    checks['templates'] = tchecks
    allp = checks['problems'] + tchecks['problems']
    checks['tabs'] = tchecks['tabs']
    checks['buttons'] = tchecks['buttons']          # palette buttons (what the stock editor shows)
    checks['mapping_buttons'] = sum(len(s['buttons']) for s in services)
    checks['all_problems'] = allp
    return {'services': services, 'templates': templates, 'checks': checks, 'root': root, 'resolved': resolved}


def _write_html(bundle, path):
    im = bundle['checks'].get('images')
    if im and 'missing_rows' in im:
        bundle = dict(bundle)
        bundle['checks'] = dict(bundle['checks'])
        bundle['checks']['images'] = {k: v for k, v in im.items() if k not in ('missing_rows', 'disagree_rows')}
    data = json.dumps(bundle).replace('</', '<\\/')
    with open(path, 'w', encoding='utf-8') as f:
        f.write(HTML.replace('__DATA__', data))


def main(argv):
    if len(argv) < 2:
        print(__doc__); return 1
    args = argv[2:]
    root = args[args.index('--items') + 1] if '--items' in args else ''
    imgs = args[args.index('--images') + 1] if '--images' in args else ''
    cache = args[args.index('--cache') + 1] if '--cache' in args else ''
    raw = decompress(argv[1])
    bundle = _bundle(raw, root, imgs, '--itm-first' in args, cache)
    services, checks = bundle['services'], bundle['checks']
    if '--json' in args:
        p = args[args.index('--json') + 1]
        b2 = dict(bundle)
        b2['checks'] = {k: (({kk: vv for kk, vv in v.items() if kk not in ('missing_rows', 'disagree_rows')}) if k == 'images' else v) for k, v in bundle['checks'].items()}
        json.dump(b2, open(p, 'w', encoding='utf-8'), indent=1)
        print('wrote', p)
    if '--html' in args:
        p = args[args.index('--html') + 1]
        _write_html(bundle, p)
        print('wrote', p)
    if '--find' in args:
        t = args[args.index('--find') + 1].lower()
        for tp in bundle['templates']:
            nb = sum(len(x['buttons']) for x in tp['tabs'])
            if t in tp['name'].lower() or any(t in b['name'].lower() for x in tp['tabs'] for b in x['buttons']):
                print('palette  %-50s tabs=%d buttons=%d' % (tp['name'], len(tp['tabs']), nb))
        for s in services:
            if t in s['name'].lower() or t in s['group'].lower():
                print('service  %-50s entries=%d' % (s['name'], len(s['entries'])))
    if '--summary' in args or len(args) == 0:
        brief = {k: v for k, v in checks.items() if k not in ('templates',)}
        if 'images' in brief:
            brief['images'] = {k: v for k, v in brief['images'].items() if k not in ('missing_rows', 'disagree_rows')}
        print(json.dumps(brief, indent=1))
    return 0


RIBBON_JSON = os.environ.get('FABMAP_JSON') or r'C:\temp\Ribbon_FABCADMEP-Path.json'          # (the environment variables are for testing)
DEFAULT_DB = os.environ.get('FABMAP_DB') or r'C:\Egnyte\Shared\Engineering\031407MEP\MURRAY FAB DATABASE'

# Only these MAP.INI entries are ever read. Everything else in the file (logins, recent files) is ignored.
_INI_PATH_KEYS = {
    'PROGRAM PATHS': None,    # all keys: Items, Images, Database, Parts, Reports, ...
    'GENERAL': {'PART DIR', 'IMPORT DIR', 'PROFILE DIR', 'PROJECT DIR'},
}


def parse_map_ini(ini_path):
    """Read only the folder settings from MAP.INI. Returns {'SECTION': {'KEY': 'value'}} (upper-case names)."""
    out = {}
    with open(ini_path, 'rb') as f:
        data = f.read()
    # Fabrication writes MAP.INI as UTF-16 with a BOM; accept UTF-8 and ANSI too
    if data[:2] in (b'\xff\xfe', b'\xfe\xff'):
        text = data.decode('utf-16', errors='replace')
    elif data[:3] == b'\xef\xbb\xbf':
        text = data[3:].decode('utf-8', errors='replace')
    else:
        try:
            text = data.decode('utf-8')
        except UnicodeDecodeError:
            text = data.decode('cp1252', errors='replace')
    if True:
        section = None
        for line in text.splitlines():
            line = line.strip()
            if line.startswith('[') and line.endswith(']'):
                section = line[1:-1].strip().upper()
                continue
            if section not in _INI_PATH_KEYS or '=' not in line:
                continue
            k, v = line.split('=', 1)
            k = k.strip().upper()
            allowed = _INI_PATH_KEYS[section]
            if allowed is None or k in allowed:
                out.setdefault(section, {})[k] = v.strip()
    return out


def _find_ci(folder, name):
    """Case-insensitive file lookup inside a folder. Returns full path or ''."""
    import os
    try:
        for n in os.listdir(folder):
            if n.lower() == name.lower():
                return os.path.join(folder, n)
    except Exception:
        pass
    return ''


def _resolve(db, value):
    import os
    if not value:
        return ''
    v = value.replace('/', os.sep).replace('\\', os.sep)
    if os.path.isabs(v) or (len(v) > 1 and v[1] == ':'):
        return os.path.normpath(v)
    return os.path.normpath(os.path.join(db, v))


def _read_json(path):
    try:
        with open(path, 'r', encoding='utf-8') as f:
            d = json.load(f)
        return d if isinstance(d, dict) else {}
    except Exception:
        return {}


def _write_json(path, cfg):
    import os
    try:
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(cfg, f, indent=2)
        return True
    except Exception as ex:
        _log('Could not write', path, '-', ex)
        return False


def resolve_environment(json_path=None, default_db=None, reselect=False, pick_folder=None, progress=None):
    """Find the database folder (JSON, else default), read MAP.INI, return the working folders.
    The JSON keeps any other keys other tools put in it. Optional override keys the user may add:
    items_path, images_path, service_map_path."""
    import os
    _p = progress or (lambda text: None)
    json_path = json_path or RIBBON_JSON
    default_db = default_db or DEFAULT_DB
    _p('Reading the settings file (%s)' % json_path)
    cfg = _read_json(json_path)
    db = cfg.get('database_path') or default_db
    _p('Looking for MAP.INI in %s' % db)

    def valid(p):
        return bool(p) and os.path.isdir(p) and bool(_find_ci(p, 'MAP.INI'))

    if reselect or not valid(db):
        if pick_folder:
            msg = 'Select the Fabrication database folder (the one that contains MAP.INI)'
            picked = pick_folder(msg)
            if picked:
                db = picked
    env = {'db': db, 'ini': '', 'items': '', 'images': '', 'database': '', 'service_map': '',
           'json': json_path, 'problems': [], 'ini_paths': {},
           'image_source': 'itm_first' if str(cfg.get('image_source', '')).lower() == 'itm_first' else 'map_first'}
    if not valid(db):
        env['problems'].append('MAP.INI not found in %r. Edit database_path in %s or shift-click the button.' % (db, json_path))
    else:
        ini = _find_ci(db, 'MAP.INI')
        env['ini'] = ini
        _p('Reading %s' % ini)
        if cfg.get('database_path') != db:
            cfg['database_path'] = db
            _write_json(json_path, cfg)
        elif not os.path.isfile(json_path):
            _write_json(json_path, cfg)
        try:
            paths = parse_map_ini(ini)
        except Exception as ex:
            paths = {}
            env['problems'].append('Could not read MAP.INI: %s' % ex)
        pp = paths.get('PROGRAM PATHS', {})
        env['ini_paths'] = {k: _resolve(db, v) for k, v in pp.items()}
        env['items'] = env['ini_paths'].get('ITEMS', '')
        env['images'] = env['ini_paths'].get('IMAGES', '')
        env['database'] = env['ini_paths'].get('DATABASE', '')
        if env['database'] and os.path.isdir(env['database']):
            env['service_map'] = _find_ci(env['database'], 'service.map')
        for key, ov in (('items', 'items_path'), ('images', 'images_path'), ('service_map', 'service_map_path')):
            if cfg.get(ov):
                env[key] = cfg[ov]
        _p('Checking the Items and Images folders')
        if env['items'] and not os.path.isdir(env['items']):
            env['problems'].append('Items folder from MAP.INI does not exist: %s' % env['items'])
        if env['images'] and not os.path.isdir(env['images']):
            env['problems'].append('Images folder from MAP.INI does not exist: %s' % env['images'])
    return env


def _pick_file(init_dir=''):
    """File picker for pyRevit (CPython) or plain Windows Python. Returns a path or None."""
    try:
        from pyrevit import forms
        return forms.pick_file(file_ext='map', init_dir=init_dir or '', title='Select service.map')
    except Exception:
        pass
    try:
        import clr
        clr.AddReference('System.Windows.Forms')
        from System.Windows.Forms import OpenFileDialog, DialogResult
        dlg = OpenFileDialog()
        dlg.Title = 'Select service.map'
        dlg.Filter = 'Fabrication map (*.map)|*.map|All files (*.*)|*.*'
        if init_dir:
            dlg.InitialDirectory = init_dir
        return dlg.FileName if dlg.ShowDialog() == DialogResult.OK else None
    except Exception:
        pass
    try:
        import tkinter
        from tkinter import filedialog
        r = tkinter.Tk(); r.withdraw()
        return filedialog.askopenfilename(title='Select service.map', initialdir=init_dir or None, filetypes=[('map', '*.map')]) or None
    except Exception:
        return None


def _pick_folder(title):
    try:
        from pyrevit import forms
        return forms.pick_folder(title=title)
    except Exception:
        pass
    try:
        import clr
        clr.AddReference('System.Windows.Forms')
        from System.Windows.Forms import FolderBrowserDialog, DialogResult
        dlg = FolderBrowserDialog()
        dlg.Description = title
        return dlg.SelectedPath if dlg.ShowDialog() == DialogResult.OK else None
    except Exception:
        pass
    try:
        import tkinter
        from tkinter import filedialog
        r = tkinter.Tk(); r.withdraw()
        return filedialog.askdirectory(title=title) or None
    except Exception:
        return None


def interactive():
    import os, tempfile, webbrowser
    g = globals()
    shift = bool(g.get('__shiftclick__', False))   # pyRevit: shift-click = choose the database folder again
    ctrl = bool(g.get('__ctrlclick__', False))     # pyRevit: ctrl-click = choose a different service.map
    env = resolve_environment(reselect=shift, pick_folder=_pick_folder)
    path = '' if ctrl else env['service_map']
    if not path or not os.path.isfile(path):
        path = _pick_file(env['database'] or env['db']) or ''
    if not path:
        print('No service.map selected.')
        for p in env['problems']:
            print('  problem:', p)
        return 0
    items = env['items'] if os.path.isdir(env['items'] or '') else ''
    images = env['images'] if os.path.isdir(env['images'] or '') else ''
    if not items:
        items = _pick_folder("MAP.INI did not give a usable Items folder. Select your ITEMS folder (contains 'Piping Components'), or cancel") or ''
    raw = decompress(path)
    cache_path = os.path.join(os.path.dirname(env['json']), 'fabmap_itm_cache.json')
    bundle = _bundle(raw, items, images, env['image_source'] == 'itm_first', cache_path)
    out = os.path.join(tempfile.gettempdir(), 'fabmap_viewer.html')
    _write_html(bundle, out)
    c = bundle['checks']
    print('Read %s: %d palettes, %d tabs, %d palette buttons, %d services, %d problems' % (
        os.path.basename(path), c['templates']['templates'], c['tabs'], c['buttons'], c['services'], len(c['all_problems'])))
    for p in c['all_problems'][:10]:
        print('  problem:', p)
    print('Database folder:', env['db'])
    print('MAP.INI        :', env['ini'] or '(not found)')
    print('Items folder   :', items or '(not set)')
    print('Images folder  :', images or '(not set)')
    print('service.map    :', path)
    print('Path settings  :', env['json'], ' (edit database_path to redirect, or shift-click the button)')
    for p in env['problems']:
        print('  problem:', p)
    im = c.get('images')
    if im:
        print('Image source   : %s (add "image_source": "itm_first" to the JSON for the slow, thorough mode)' % ('.ITM first (slow)' if im['itm_first'] else 'service.map first, .ITM only when no picture is found'))
        print('Time           : read and parse %s s, pictures %s s' % (c.get('parse_seconds', '?'), im.get('seconds', '?')))
        print('Images found %d, not found %d (%d distinct items)  (%s)' % (im['found'], im['missing'], im.get('distinct_items_missing', 0), ', '.join('%s: %d' % kv for kv in sorted(im['via'].items()))))
        print('.ITM files read %d, from cache %d%s' % (im['itm_files_read'], im['itm_cache_hits'],
              (', skipped %d (20 s limit; run again to continue, results are cached)' % im['itm_skipped']) if im.get('itm_skipped') else ''))
        import csv
        for rows, name, head in ((im.get('missing_rows'), 'fabmap_missing_images.csv',
                                  ['palette', 'tab', 'button', 'item_path', 'image_stored_in_service_map', 'itm_file_found', 'image_names_inside_itm', 'first_path_tried']),
                                 (im.get('disagree_rows'), 'fabmap_image_disagreements.csv',
                                  ['palette', 'tab', 'button', 'item_path', 'image_from_itm', 'image_from_service_map', 'service_map_how'])):
            if rows:
                p = os.path.join(tempfile.gettempdir(), name)
                try:
                    with open(p, 'w', newline='', encoding='utf-8-sig') as f:
                        w = csv.writer(f)
                        w.writerow(head)
                        w.writerows(rows)
                    print('%s (%d rows): %s' % ('Missing images' if 'missing' in name else 'Item and service.map point to different images', len(rows), p))
                except Exception as ex:
                    print('Could not write', p, ex)
        for m in im['missing_examples'][:8]:
            print('  no image:', m)
    print('Viewer:', out)
    try:
        os.startfile(out)
    except Exception:
        webbrowser.open('file:///' + out.replace(os.sep, '/'))
    return 0




# ======================================================================================
#  WRITER
# ======================================================================================
CHUNK = 1 << 25
HEADER_LEN = 26
U = lambda raw, o: struct.unpack_from('<I', raw, o)[0]


# ---------------------------------------------------------------- container
def read_map(path):
    with open(path, 'rb') as f:
        data = f.read()
    return data, decompress_bytes(data)


def build_container(raw, header):
    out = [header]
    for i in range(0, len(raw), CHUNK):
        chunk = bytes(raw[i:i + CHUNK])
        c = zlib.compressobj(6, zlib.DEFLATED, 15, 8)
        z = c.compress(chunk) + c.flush()
        out.append(struct.pack('<I', len(z) + 5))
        out.append(z)
        out.append(b'\x01' + struct.pack('<I', len(chunk)))
    return b''.join(out)


# ---------------------------------------------------------------- layout
_LAY = {'raw': None, 'lay': None}


def layout(raw):
    """Locate every template, tab and button (cached for the most recent stream: same object, same answer)."""
    if _LAY['raw'] is raw:
        return _LAY['lay']
    lay = _layout(raw)
    _LAY['raw'], _LAY['lay'] = raw, lay
    return lay


def _layout(raw):
    """Locate every template, tab and button. Offsets are into the decompressed stream."""
    first = raw.find(TPL_TAG)
    if first < 0:
        raise ValueError('no service template section')
    lay = {'first': first, 'templates': []}
    o = first
    while raw[o:o + 4] == TPL_TAG:
        te = o + 8 + U(raw, o + 4)
        tpl = {'t': o, 'te': te, 'tabs': []}
        tpl['name'] = _first_string(raw, o + 8, te, 160)
        for x in _all(raw, TPL_TAB, o + 8, te):
            xe = min(x + 8 + U(raw, x + 4), te)
            n = U(raw, x + 12) if x + 16 <= xe else 0
            nm = ''
            if 1 <= n < 300 and U(raw, x + 8) == 2 * n + 28 and x + 16 + 2 * n <= xe and raw[x + 14 + 2 * n:x + 16 + 2 * n] == b'\x00\x00':
                nm = raw[x + 16:x + 14 + 2 * n].decode('utf-16le', 'replace') or '(unnamed)'     # the name sits at a fixed place in the tab header
            else:
                nm = _first_string(raw, x + 8, xe, 60)
            if not nm:
                continue
            bs = _all(raw, TPL_BTN, x, xe)
            blocks = [(b, b + 8 + U(raw, b + 4)) for b in bs]
            tpl['tabs'].append({'x': x, 'xe': xe, 'name': nm, 'blocks': blocks})
        lay['templates'].append(tpl)
        o = te
    lay['chain_end'] = o
    return lay


def validate(raw, deep=True):
    """Structural checks. Returns a list of problems (empty = consistent). deep=False skips the full re-read."""
    p = []
    lay = layout(raw)
    first, chain_end = lay['first'], lay['chain_end']
    span = chain_end - first
    if U(raw, first - 12) != span + 12:
        p.append('collection size A is %d, expected %d' % (U(raw, first - 12), span + 12))
    if U(raw, first - 8) != span + 8:
        p.append('collection size B is %d, expected %d' % (U(raw, first - 8), span + 8))
    if U(raw, first - 4) != len(lay['templates']):
        p.append('template count field %d, found %d' % (U(raw, first - 4), len(lay['templates'])))
    if len(raw) - chain_end != 4:
        p.append('%d bytes after the template section, expected 4' % (len(raw) - chain_end))
    for tpl in lay['templates']:
        if tpl['tabs'] and tpl['te'] - (tpl['tabs'][-1]['xe']) != 20:
            p.append('template %r: footer is %d bytes, expected 20' % (tpl['name'], tpl['te'] - tpl['tabs'][-1]['xe']))
        if tpl['tabs'] and U(raw, tpl['tabs'][0]['x'] - 4) != len(tpl['tabs']):
            p.append('template %r: tab count field %d, found %d tabs' % (tpl['name'], U(raw, tpl['tabs'][0]['x'] - 4), len(tpl['tabs'])))
        for tab in tpl['tabs']:
            bl = tab['blocks']
            szp, cnp, start = _list_pos(raw, tab)
            if not bl:
                if U(raw, szp) != 4 or U(raw, cnp) != 0 or tab['xe'] != start:
                    p.append('%r/%r: empty tab layout is wrong' % (tpl['name'], tab['name']))
                continue
            b0 = bl[0][0]
            if b0 != start:
                p.append('%r/%r: first button is not where the list header says' % (tpl['name'], tab['name']))
            for k in range(len(bl) - 1):
                if bl[k][1] != bl[k + 1][0]:
                    p.append('%r/%r: buttons %d and %d are not back to back' % (tpl['name'], tab['name'], k + 1, k + 2))
            if U(raw, b0 - 4) != len(bl):
                p.append('%r/%r: count field %d, found %d buttons' % (tpl['name'], tab['name'], U(raw, b0 - 4), len(bl)))
            if U(raw, b0 - 8) != 4 + sum(e - s for s, e in bl):
                p.append('%r/%r: list size field wrong' % (tpl['name'], tab['name']))
            if tab['xe'] != bl[-1][1]:
                p.append('%r/%r: tab end differs from last button end' % (tpl['name'], tab['name']))
    if deep:
        p.extend(enclosing_problems(raw))
        # the reader used by the viewer must agree
        tp, chk = parse_templates(raw)
        p.extend(chk['problems'])
    return p


# ---------------------------------------------------------------- enclosing sizes, button ids
# Verified against a before/after pair written by the stock editor (one button added): besides the tab, list, template and
# collection sizes it also changes three more size fields - the stream size at offset 20, the record that encloses all the
# templates (24 bytes in front of the first one) and the template's own inner record (right in front of its condition list) -
# and it keeps a button-id counter in the template header (the 42 bytes in front of that inner size).
_DE_ = re.compile(re.escape(b'\xde\x03\x5c\x00'))
_E0_ = re.compile(re.escape(b'\xe0\x03\x5c\x00'))


def _cond_size_pos(raw, tpl):
    m = _DE_.search(raw, tpl['t'], tpl['tabs'][0]['x'] if tpl['tabs'] else tpl['te'])
    return None if not m else m.start() - 8


def _bump_enclosing(new, raw, lay, tpl, delta):
    """Add delta to the stream size, the record around all templates and this template's inner record."""
    def bump(off):
        struct.pack_into('<I', new, off, (U(raw, off) + delta) & 0xFFFFFFFF)
    bump(20)
    bump(lay['first'] - 24)
    sp = _cond_size_pos(raw, tpl)
    if sp is not None:
        bump(sp)


def enclosing_problems(raw):
    p = []
    lay = layout(raw)
    L = len(raw)
    if U(raw, 20) != L - 24:
        p.append('stream size field is %d, expected %d' % (U(raw, 20), L - 24))
    f = lay['first'] - 24
    if U(raw, f) != L - f - 4:
        p.append('size of the record around the templates is %d, expected %d' % (U(raw, f), L - f - 4))
    for tpl in lay['templates']:
        sp = _cond_size_pos(raw, tpl)
        if sp is not None and U(raw, sp) != tpl['te'] - (sp + 4):
            p.append('template %r: inner record size %d, expected %d' % (tpl['name'], U(raw, sp), tpl['te'] - (sp + 4)))
    return p


def fix_enclosing_sizes(raw):
    """Recompute those three fields from the actual structure (returns raw itself when they are already right)."""
    lay = layout(raw)
    L = len(raw)
    want = [(20, L - 24), (lay['first'] - 24, L - (lay['first'] - 24) - 4)]
    for tpl in lay['templates']:
        sp = _cond_size_pos(raw, tpl)
        if sp is not None:
            want.append((sp, tpl['te'] - (sp + 4)))
    new = None
    for off, v in want:
        if U(raw, off) != v:
            if new is None:
                new = bytearray(raw)
            struct.pack_into('<I', new, off, v)
    return raw if new is None else new


def counter_pos(raw, tpl):
    sp = _cond_size_pos(raw, tpl)
    return None if sp is None else sp - 42


def template_counter(raw, ti):
    cp = counter_pos(raw, layout(raw)['templates'][ti])
    return None if cp is None else U(raw, cp)


def set_template_counter(raw, ti, value):
    cp = counter_pos(raw, layout(raw)['templates'][ti])
    if cp is None:
        raise ValueError('this template has no button counter')
    new = bytearray(raw)
    struct.pack_into('<I', new, cp, value)
    return new


def button_format(block):
    """'old' or 'new': which record version a button uses, from its first size record (None when it has none)."""
    items = []
    m = _E0_.search(block, 8)
    if not m:
        return None
    sz = U(block, m.start() + 4)
    inner = _E0_.search(block, m.start() + 8, m.start() + 8 + sz)
    if not inner:
        return None
    v = U(block, inner.start() + 8)
    return {34: 'old', 74: 'new'}.get(v)


_ID_BACK = {'old': 28, 'new': 44}          # the id sits this many bytes in front of the item record (it differs between the versions)
_STD_BACK = {'old': 44, 'new': None}       # in the older version the 4 bytes 44 in front hold a constant, never an id


def template_format(raw, tpl):
    """The record version most of a template's buttons use (None if none can be recognised). Buttons with no size records cannot
    tell their own version, and the buttons of one template share it."""
    seen = {}
    for tab in tpl['tabs']:
        for s, e in tab['blocks']:
            f = button_format(bytes(raw[s:e]))
            if f:
                seen[f] = seen.get(f, 0) + 1
    return max(seen.items(), key=lambda kv: kv[1])[0] if seen else None


def button_id(block, hint=None):
    f = button_format(block) or hint
    m = _E0_.search(block, 8)
    if not f or not m or m.start() < _ID_BACK[f]:
        return None
    return U(block, m.start() - _ID_BACK[f])


def with_button_id(block, bid, hint=None):
    f = button_format(block) or hint
    m = _E0_.search(block, 8)
    if not f or not m or m.start() < _ID_BACK[f]:
        raise ValueError('this button has a layout I do not recognise')
    b = bytearray(block)
    struct.pack_into('<I', b, m.start() - _ID_BACK[f], bid)
    return bytes(b)


def template_ids(raw, tpl):
    """The ids of all buttons of a template, read the way their version stores them (None where a button is not recognised)."""
    out = []
    hint = template_format(raw, tpl)
    for tab in tpl['tabs']:
        for s, e in tab['blocks']:
            out.append(button_id(bytes(raw[s:e]), hint))
    return out


def ids_look_right(raw, tpl, cp):
    """Sanity check before the counter is touched: every button has an id, they are unique, and the counter is not below the highest."""
    ids = [i for i in template_ids(raw, tpl) if i is not None]        # buttons with no size records have no recognisable version: ignored
    return bool(ids) and len(set(ids)) == len(ids) and U(raw, cp) >= max(ids)


def find_damaged_ids(raw):
    """Older-format buttons whose constant (44 bytes in front of the item) was overwritten with an id by the id feature of earlier
    editors: [(ti, tabi, idx, found, expected)]. Expected = the most common value among the same kind of button in the file."""
    lay = layout(raw)
    seen, by = [], {}
    for ti, tpl in enumerate(lay['templates']):
        for tabi, tab in enumerate(tpl['tabs']):
            for idx, (s, e) in enumerate(tab['blocks']):
                blk = bytes(raw[s:e])
                if button_format(blk) != 'old':
                    continue
                m = _E0_.search(blk, 8)
                if m.start() < 44:
                    continue
                v = U(blk, m.start() - 44)
                key = m.start() - (16 + 2 * U(blk, 12))          # length of the header block (a button 'variant')
                seen.append((ti, tabi, idx, v, key))
                by.setdefault(key, {}).setdefault(v, 0)
                by[key][v] += 1
    out = []
    for ti, tabi, idx, v, key in seen:
        std = max(by[key].items(), key=lambda kv: kv[1])[0]
        if v != std and v > 255:
            out.append((ti, tabi, idx, v, std))
    return out


def repair_damaged_ids(raw, damaged):
    """Put the constant back where an earlier editor wrote an id. Same size. Returns (new bytes, [(position, old value)])."""
    lay = layout(raw)
    new = bytearray(raw)
    olds = []
    for ti, tabi, idx, found, std in damaged:
        s, e = lay['templates'][ti]['tabs'][tabi]['blocks'][idx]
        m = _E0_.search(raw, s + 8, e)
        pos = m.start() - 44
        olds.append((pos, U(raw, pos)))
        struct.pack_into('<I', new, pos, std)
    return new, olds


def add_block(raw, ti, tabi, pos, block):
    """Insert a NEW button the way the stock editor does: it gets the template's next id (counter + 1) and the counter moves up
    to it. Returns (new bytes, old counter) - the old counter is needed to undo this exactly."""
    lay = layout(raw)
    tpl = lay['templates'][ti]
    cp = counter_pos(raw, tpl)
    if cp is None:
        raise ValueError('this template has no button counter')
    old = U(raw, cp)
    hint = template_format(raw, tpl)
    if not ids_look_right(raw, tpl, cp) or (button_format(block) or hint) is None:
        return insert_block(raw, ti, tabi, pos, block), None          # unknown layout: copy exactly, change no ids
    new = insert_block(raw, ti, tabi, pos, with_button_id(block, old + 1, hint))
    struct.pack_into('<I', new, cp, old + 1)
    return new, old


# ---------------------------------------------------------------- edit helpers
def _list_pos(raw, tab):
    """(offset of the button-list size, of the button count, of the first button) of a tab. The list header sits 24 bytes
    after the tab name, so it is in the same place whether the tab has buttons or none (an empty tab is [4][0] and ends there)."""
    x = tab['x']
    a = x + 16 + 2 * U(raw, x + 12)
    return a + 24, a + 28, a + 32


def _apply(raw, lay, tpl, tab, delta, count_delta, splice):
    """Return new bytearray: sizes fixed for tab/template/collection, then one splice (off, dellen, data)."""
    new = bytearray(raw)
    szp, cnp, _ = _list_pos(raw, tab)

    def bump(off, d):
        struct.pack_into('<I', new, off, (U(raw, off) + d) & 0xFFFFFFFF)

    bump(tab['x'] + 4, delta)                 # tab size
    bump(szp, delta)                          # button list size
    bump(cnp, count_delta)                    # button count
    bump(tpl['t'] + 4, delta)                 # template size
    bump(lay['first'] - 12, delta)            # collection size A
    bump(lay['first'] - 8, delta)             # collection size B
    _bump_enclosing(new, raw, lay, tpl, delta)
    off, dellen, data = splice
    new[off:off + dellen] = data
    return new


def _at(lay, ti, tabi):
    try:
        tpl = lay['templates'][ti]
        return tpl, tpl['tabs'][tabi]
    except (IndexError, TypeError):
        raise ValueError('no such palette or tab')


def move_idx(raw, ti, tabi, frm, to):
    """0-based positions. Moves one button inside a tab; no sizes change."""
    lay = layout(raw)
    tpl, tab = _at(lay, ti, tabi)
    n = len(tab['blocks'])
    if not (0 <= frm < n and 0 <= to < n):
        raise ValueError('positions must be 0..%d' % (n - 1))
    blocks = [bytes(raw[s:e]) for s, e in tab['blocks']]
    blk = blocks.pop(frm)
    blocks.insert(to, blk)
    new = bytearray(raw)
    new[tab['blocks'][0][0]:tab['xe']] = b''.join(blocks)
    return new


def delete_idx(raw, ti, tabi, idx):
    lay = layout(raw)
    tpl, tab = _at(lay, ti, tabi)
    n = len(tab['blocks'])
    if not 0 <= idx < n:
        raise ValueError('button position must be 0..%d' % (n - 1))
    s, e = tab['blocks'][idx]
    return _apply(raw, lay, tpl, tab, -(e - s), -1, (s, e - s, b''))


def block_bytes(raw, ti, tabi, idx):
    lay = layout(raw)
    tpl, tab = _at(lay, ti, tabi)
    if not 0 <= idx < len(tab['blocks']):
        raise ValueError('button position must be 0..%d' % (len(tab['blocks']) - 1))
    s, e = tab['blocks'][idx]
    return bytes(raw[s:e])


def insert_block(raw, ti, tabi, pos, block):
    """Insert one whole button record at 0-based position pos (None = end) of a tab; fixes every size and count."""
    lay = layout(raw)
    tpl, tab = _at(lay, ti, tabi)
    n = len(tab['blocks'])
    pos = n if pos is None else pos
    if not 0 <= pos <= n:
        raise ValueError('destination position must be 0..%d' % n)
    at = tab['blocks'][pos][0] if pos < n else tab['xe']
    return _apply(raw, lay, tpl, tab, len(block), 1, (at, 0, block))


def copy_idx(raw, ti, tabi, idx, dti, dtabi, pos=None):
    """Copy button idx of (ti, tabi) into tab dtabi of palette dti at 0-based position pos (None = end)."""
    return add_block(raw, dti, dtabi, pos, block_bytes(raw, ti, tabi, idx))[0]



# ---------------------------------------------------------------- Exclude From Fill (buttons and tabs)
def btn_tail(raw, s, e):
    """(start of the button-code string, the code) found by counting back from the end of a button record, or None."""
    for n in range(1, 60):
        st = e - 28 - (4 + 2 * n)
        if st <= s:
            break
        if U(raw, st) == n and raw[st + 4 + 2 * (n - 1):st + 4 + 2 * n] == b'\x00\x00':
            ch = raw[st + 4:st + 4 + 2 * (n - 1)]
            if all(ch[i + 1] == 0 and 32 <= ch[i] < 127 for i in range(0, len(ch), 2)):
                return st, ch[::2].decode('ascii')
    return None


def btn_flag_pos(raw, s, e):
    """Offset of the button's 'excluded from fill' value (16 bit, 0 or 1): 8 bytes before its code string."""
    t = btn_tail(raw, s, e)
    if t is None:
        return None
    p = t[0] - 8
    return p if struct.unpack_from('<H', raw, p)[0] in (0, 1) else None


def btn_excluded(raw, s, e):
    p = btn_flag_pos(raw, s, e)
    return None if p is None else struct.unpack_from('<H', raw, p)[0] == 1


def tab_flag_pos(raw, x):
    """Offset of the tab's mode value (4 bytes, 0 or 1 = excluded from fill): right after the tab's name."""
    return x + 16 + 2 * U(raw, x + 12)


def tab_excluded(raw, tab):
    return U(raw, tab_flag_pos(raw, tab['x'])) == 1


def set_button_flags(raw, ti, tabi, idxs, value):
    """Set 'excluded from fill' on the given buttons. Same size; returns (new bytes, old values) so it can be undone."""
    lay = layout(raw)
    tpl, tab = _at(lay, ti, tabi)
    new, olds = bytearray(raw), []
    for i in _picked(len(tab['blocks']), idxs):
        s, e = tab['blocks'][i]
        p = btn_flag_pos(raw, s, e)
        if p is None:
            raise ValueError('button %d has a layout I do not recognise; nothing changed' % (i + 1))
        olds.append((i, struct.unpack_from('<H', raw, p)[0]))
        struct.pack_into('<H', new, p, 1 if value else 0)
    return new, olds


def restore_button_flags(raw, ti, tabi, olds):
    lay = layout(raw)
    tpl, tab = _at(lay, ti, tabi)
    new = bytearray(raw)
    for i, v in olds:
        s, e = tab['blocks'][i]
        p = btn_flag_pos(raw, s, e)
        if p is None:
            raise ValueError('button layout not recognised')
        struct.pack_into('<H', new, p, v)
    return new


def set_tab_flag(raw, ti, tabi, value):
    """Set the tab's mode (excluded from fill). Same size; returns (new bytes, old value)."""
    lay = layout(raw)
    tpl, tab = _at(lay, ti, tabi)
    p = tab_flag_pos(raw, tab['x'])
    old = U(raw, p)
    if old not in (0, 1):
        raise ValueError('this tab has an unexpected mode value (%d); nothing changed' % old)
    new = bytearray(raw)
    struct.pack_into('<I', new, p, 1 if value else 0)
    return new, old


# ---------------------------------------------------------------- size restrictions (conditions)
_E0 = re.compile(re.escape(b'\xe0\x03\x5c\x00'))
_DE = re.compile(re.escape(b'\xde\x03\x5c\x00'))
_ASCII16 = re.compile(rb'(?:[ -~]\x00){2,}')


def _png_end(raw, o, end):
    """Where the item's picture path ends (None if the item does not look as expected). Everything after it is the item's own
    settings: the part id, the restriction id (8 bytes after it), the date, ..."""
    try:
        pn = U(raw, o + 16)
        if not 1 <= pn <= 600:
            return None
        pe = o + 20 + 2 * pn
        qn = U(raw, pe)
        if not 1 <= qn <= 600 or pe + 4 + 2 * qn + 24 > end:
            return None
        return pe + 4 + 2 * qn
    except struct.error:
        return None


def btn_items(raw, s, e):
    """The item records of a button: [{'start','end','path','cid_pos','cid'}].
    An item record holds the part's path, its picture path, its size restriction (a condition id, 8 bytes after the picture path in
    both record versions) and its size records. Earlier editors counted back 140 bytes from the size list, which is only the same
    place for one common layout."""
    out = []
    m = _E0.search(raw, s + 8, e)
    while m:
        o = m.start()
        sz = U(raw, o + 4)
        end = o + 8 + sz
        if end > e or sz < 100:
            break
        inner = _E0.search(raw, o + 8, end)
        rows_start = inner.start() if inner else end - 16
        pm = _ASCII16.match(raw, o + 20)
        path = pm.group().decode('utf-16le') if pm else ''
        pe = _png_end(raw, o, end)
        cp = pe + 8 if pe else rows_start - 140
        out.append({'start': o, 'end': end, 'path': path, 'cid_pos': cp, 'cid': U(raw, cp) if cp >= o + 16 else None, 'png_end': pe,
                    'rows_start': rows_start})
        m = _E0.search(raw, end, e)
    return out


def template_conditions(raw, tpl):
    """The template's size conditions: [{'id','name','gt','lte','gt_pos','lte_pos'}] (-1.0 = no restriction)."""
    out = []
    a, b = tpl['t'], tpl['tabs'][0]['x'] if tpl['tabs'] else tpl['te']
    for m in _DE.finditer(raw, a, b):
        o = m.start()
        sz = U(raw, o + 4)
        body = raw[o:o + 8 + sz]
        txt = [x.decode('utf-16le') for x in _ASCII16.findall(body, 32)]
        out.append({'id': U(raw, o + 12), 'name': txt[0] if txt else '', 'gt': struct.unpack_from('<d', raw, o + 16)[0],
                    'lte': struct.unpack_from('<d', raw, o + 24)[0], 'gt_pos': o + 16, 'lte_pos': o + 24})
    return out


def set_item_condition(raw, ti, tabi, idx, k, cid):
    """Make item k of button idx use condition id cid (one 4-byte value; same size). Returns (new bytes, old id)."""
    lay = layout(raw)
    tpl, tab = _at(lay, ti, tabi)
    if not 0 <= idx < len(tab['blocks']):
        raise ValueError('button position must be 0..%d' % (len(tab['blocks']) - 1))
    s, e = tab['blocks'][idx]
    items = btn_items(raw, s, e)
    if not 0 <= k < len(items) or items[k]['cid'] is None:
        raise ValueError('that button has no item %d' % (k + 1))
    ids = [c['id'] for c in template_conditions(raw, tpl)]
    if cid not in ids:
        raise ValueError('this template has no condition number %s' % cid)
    new = bytearray(raw)
    struct.pack_into('<I', new, items[k]['cid_pos'], int(cid))
    return new, items[k]['cid']


def set_condition_limits(raw, ti, cid, gt, lte, check=True):
    """Change a template condition's two numbers (greater than, less than or equal; None = no restriction). Same size.
    Returns (new bytes, (old gt, old lte))."""
    lay = layout(raw)
    tpl = lay['templates'][ti]
    cs = [c for c in template_conditions(raw, tpl) if c['id'] == cid]
    if not cs:
        raise ValueError('this template has no condition number %s' % cid)
    c = cs[0]
    g = -1.0 if gt is None else float(gt)
    l = -1.0 if lte is None else float(lte)
    for v in (g, l):
        if check and v != -1.0 and not 0 <= v <= 10000:
            raise ValueError('a size limit must be between 0 and 10000 (or empty for no limit)')
    if check and g != -1.0 and l != -1.0 and g >= l:
        raise ValueError('"greater than" must be smaller than "less than or equal"')
    new = bytearray(raw)
    struct.pack_into('<d', new, c['gt_pos'], g)
    struct.pack_into('<d', new, c['lte_pos'], l)
    return new, (c['gt'], c['lte'])


# ---------------------------------------------------------------- the template's list of size restrictions (add, rename, delete)
def s16(s):
    """A length-prefixed UTF-16 string with its terminating NUL, the way the file stores text."""
    return struct.pack('<I', len(s) + 1) + s.encode('utf-16le') + b'\x00\x00'


def _cond_recs(raw, tpl):
    """[(start, end, id)] of the template's condition records (they sit back to back; the count is the dword in front of the first)."""
    a, b = tpl['t'], (tpl['tabs'][0]['x'] if tpl['tabs'] else tpl['te'])
    return [(m.start(), m.start() + 8 + U(raw, m.start() + 4), U(raw, m.start() + 12)) for m in _DE_.finditer(raw, a, b)]


def _check_cond_name(raw, tpl, name, skip_id=None):
    name = (name or '').strip()
    if not 1 <= len(name) <= 40 or not name.isprintable():
        raise ValueError('A restriction needs a name of 1 to 40 characters.')
    for c in template_conditions(raw, tpl):
        if c['id'] != skip_id and c['name'].lower() == name.lower():
            raise ValueError('This template already has a restriction called "%s".' % c['name'])
    return name


def _bump_all(new, raw, lay, tpl, delta):
    def bump(off):
        struct.pack_into('<I', new, off, (U(raw, off) + delta) & 0xFFFFFFFF)
    bump(tpl['t'] + 4)
    bump(lay['first'] - 12)
    bump(lay['first'] - 8)
    _bump_enclosing(new, raw, lay, tpl, delta)


def add_condition(raw, ti, name, gt, lte):
    """Add a size restriction to the template (id = highest id + 1). Returns (new bytes, new id)."""
    lay = layout(raw)
    tpl = lay['templates'][ti]
    recs = _cond_recs(raw, tpl)
    if not recs:
        raise ValueError('this template has no restriction list I can extend')
    name = _check_cond_name(raw, tpl, name)
    g = -1.0 if gt is None else float(gt)
    l = -1.0 if lte is None else float(lte)
    for v in (g, l):
        if v != -1.0 and not 0 <= v <= 10000:
            raise ValueError('a size limit must be between 0 and 10000 (or empty for no limit)')
    if g != -1.0 and l != -1.0 and g >= l:
        raise ValueError('"greater than" must be smaller than "less than or equal"')
    nid = max(r[2] for r in recs) + 1
    body = struct.pack('<I', nid) + struct.pack('<dd', g, l) + s16(name) + b'\x00' * 20       # nothing assigned under 'Supports'
    rec = struct.pack('<III', 0x5c03de, len(body) + 4, len(body)) + body
    new = bytearray(raw)
    struct.pack_into('<I', new, recs[0][0] - 4, len(recs) + 1)
    _bump_all(new, raw, lay, tpl, len(rec))
    new[recs[-1][1]:recs[-1][1]] = rec
    return new, nid


def delete_condition(raw, ti, cid):
    """Remove a restriction. Buttons whose items used it go back to 'Unrestricted' (id 0). Returns (new, record bytes, record offset,
    [(position, old id)]) - everything needed to put it back exactly."""
    lay = layout(raw)
    tpl = lay['templates'][ti]
    recs = _cond_recs(raw, tpl)
    hit = [r for r in recs if r[2] == cid]
    if cid == 0 or not hit:
        raise ValueError('The "Unrestricted" entry cannot be deleted.' if cid == 0 else 'this template has no restriction number %s' % cid)
    start, end, _ = hit[0]
    new = bytearray(raw)
    patches = []
    for tab in tpl['tabs']:
        for s, e in tab['blocks']:
            for it in btn_items(raw, s, e):
                if it['cid'] == cid:
                    patches.append((it['cid_pos'], cid))
                    struct.pack_into('<I', new, it['cid_pos'], 0)
    rec = bytes(raw[start:end])
    struct.pack_into('<I', new, recs[0][0] - 4, len(recs) - 1)
    _bump_all(new, raw, lay, tpl, -(end - start))
    new[start:end] = b''
    return new, rec, start, patches


def restore_condition(raw, ti, start, rec, patches):
    """The exact reverse of delete_condition."""
    lay = layout(raw)
    tpl = lay['templates'][ti]
    recs = _cond_recs(raw, tpl)
    new = bytearray(raw)
    struct.pack_into('<I', new, recs[0][0] - 4, len(recs) + 1)
    _bump_all(new, raw, lay, tpl, len(rec))
    new[start:start] = rec
    for pos, val in patches:
        struct.pack_into('<I', new, pos, val)
    return new


def rename_condition(raw, ti, cid, newname):
    """Rename a restriction. Returns (new bytes, old name)."""
    lay = layout(raw)
    tpl = lay['templates'][ti]
    hit = [r for r in _cond_recs(raw, tpl) if r[2] == cid]
    if not hit:
        raise ValueError('this template has no restriction number %s' % cid)
    start, end, _ = hit[0]
    newname = _check_cond_name(raw, tpl, newname, skip_id=cid)
    n = U(raw, start + 32)
    old = raw[start + 36:start + 36 + 2 * (n - 1)].decode('utf-16le')
    if old == newname:
        return bytearray(raw), old
    delta = 2 * (len(newname) - len(old))
    new = bytearray(raw)
    struct.pack_into('<I', new, start + 4, U(raw, start + 4) + delta)
    struct.pack_into('<I', new, start + 8, U(raw, start + 8) + delta)
    _bump_all(new, raw, lay, tpl, delta)
    new[start + 32:start + 36 + 2 * n] = s16(newname)
    return new, old


def set_item_conditions_many(raw, ti, tabi, idxs, cid, k=None):
    """Give the item of every selected button (that has exactly one item) the restriction cid. Buttons with several items are left
    alone (each of their items has its own). With k, the k-th item (0 = first) of every selected button that has one instead, whatever
    the number of items. Returns (new bytes, [(position, old id)], skipped count)."""
    lay = layout(raw)
    tpl, tab = _at(lay, ti, tabi)
    if cid not in [c['id'] for c in template_conditions(raw, tpl)]:
        raise ValueError('this template has no restriction number %s' % cid)
    new = bytearray(raw)
    olds, skipped = [], 0
    for i in _picked(len(tab['blocks']), idxs):
        s, e = tab['blocks'][i]
        its = btn_items(raw, s, e)
        pick = its[0] if (k is None and len(its) == 1) else (its[k] if k is not None and 0 <= k < len(its) else None)
        if pick is None or pick['cid'] is None:
            skipped += 1
            continue
        if pick['cid'] != cid:
            olds.append((pick['cid_pos'], pick['cid']))
            struct.pack_into('<I', new, pick['cid_pos'], int(cid))
    return new, olds, skipped


# ---------------------------------------------------------------- damage done by the old id position
def _png_string(raw, o):
    pn = U(raw, o + 16)
    pe = o + 20 + 2 * pn
    qn = U(raw, pe)
    return pe + 4, qn      # where the picture path's characters start, and its length including the terminator


def find_damaged_items(raw):
    """Items an earlier editor damaged by writing a restriction id at the wrong place. What can be recognised for sure:
    a picture path with a control character in it. Returns (repairable, lost): repairable = [(position, new bytes)] (the whole picture
    path, taken from a picture path elsewhere in the file that fits every other character), lost = [(template index, button name)]"""
    lay = layout(raw)
    known = {}
    items = []
    for ti, tpl in enumerate(lay['templates']):
        for tab in tpl['tabs']:
            for s, e in tab['blocks']:
                for it in btn_items(raw, s, e):
                    if not it['png_end']:
                        continue
                    cs, qn = _png_string(raw, it['start'])
                    chars = raw[cs:cs + 2 * qn]
                    ok = qn >= 1 and chars[-2:] == b'\x00\x00' and all(chars[i] >= 0x20 or chars[i + 1] != 0 for i in range(0, 2 * qn - 2, 2))
                    if ok:
                        known.setdefault(qn, set()).add(bytes(chars))
                    else:
                        items.append((ti, it, cs, qn, bytes(chars)))
    repairable, lost = [], []
    for ti, it, cs, qn, chars in items:
        stem = it['path'].rsplit('.', 1)[0] if '.' in it['path'] else it['path']
        # the damaged characters are control characters: compare on the characters that are not
        bad = {i // 2 for i in range(0, 2 * qn - 2, 2) if chars[i + 1] == 0 and chars[i] < 0x20}
        cands = set(c for c in known.get(qn, ()) if all(c[2 * k:2 * k + 2] == chars[2 * k:2 * k + 2] for k in range(qn - 1) if k not in bad))
        derived = (stem + '.png').encode('utf-16le') + b'\x00\x00'
        if len(derived) == 2 * qn and all(derived[2 * k:2 * k + 2] == chars[2 * k:2 * k + 2] for k in range(qn - 1) if k not in bad):
            cands.add(derived)
        if len(cands) == 1:
            repairable.append((cs, next(iter(cands)), chars))
        else:
            lost.append((ti, it['path'].rsplit('/', 1)[-1]))
    return repairable, lost


def repair_damaged_items(raw, repairable):
    new = bytearray(raw)
    olds = []
    for pos, good, old in repairable:
        olds.append((pos, old))
        new[pos:pos + len(good)] = good
    return new, olds

# ---------------------------------------------------------------- copying buttons from one template into another
def copy_across(raw, sti, stabi, idxs, dti, dtabi, pos=None):
    """Copy the picked buttons of one template's tab into a tab of another template (or the same one).
    - every copy is a new button (new id from the destination's counter),
    - a button that is already in the destination tab (same items) is skipped,
    - size restrictions are per template: an item's restriction is matched in the destination by its limits, and added there (with
      the same name) when the destination has no restriction with those limits.
    Returns (new bytes, {'copied', 'skipped', 'added': [restriction names added]})."""
    lay = layout(raw)
    stpl = lay['templates'][sti]
    src = tab_blocks_bytes(raw, sti, stabi)
    sel = [src[i] for i in _picked(len(src), idxs)]
    if not sel:
        raise ValueError('nothing picked')
    dbase = tab_blocks_bytes(raw, dti, dtabi)

    def paths(b):
        return tuple(it['path'] for it in btn_items(b, 0, len(b)))

    have = {paths(b) for b in dbase}
    fresh, skipped = [], 0
    for b in sel:
        if paths(b) in have:
            skipped += 1
        else:
            fresh.append(b)
            have.add(paths(b))
    if not fresh:
        raise ValueError('Those buttons are already in this tab.' if len(sel) > 1 else 'That button is already in this tab.')
    sconds = {c['id']: c for c in template_conditions(raw, stpl)}
    need = sorted({it['cid'] for b in fresh for it in btn_items(b, 0, len(b)) if it['cid'] not in (None, 0)})
    new = raw
    mapping, added = {0: 0}, []
    key = lambda c: (round(c['gt'], 6), round(c['lte'], 6))
    for cid in need:
        sc = sconds.get(cid)
        if sc is None:
            mapping[cid] = 0
            continue
        dconds = template_conditions(new, layout(new)['templates'][dti])
        hit = [c for c in dconds if key(c) == key(sc)]
        if hit:
            mapping[cid] = hit[0]['id']
            continue
        name, n = sc['name'], 1
        while any(c['name'].lower() == name.lower() for c in dconds):
            n += 1
            name = '%s (%d)' % (sc['name'][:34], n)
        new, nid = add_condition(new, dti, name, None if sc['gt'] == -1.0 else sc['gt'], None if sc['lte'] == -1.0 else sc['lte'])
        mapping[cid] = nid
        added.append(name)
    out = []
    for b in fresh:
        nb = bytearray(b)
        for it in btn_items(b, 0, len(b)):
            if it['cid'] is not None and it['cid_pos'] + 4 <= len(nb):
                struct.pack_into('<I', nb, it['cid_pos'], mapping.get(it['cid'], 0))
        out.append(bytes(nb))
    lay2 = layout(new)
    dtpl = lay2['templates'][dti]
    cp = counter_pos(new, dtpl)
    if cp is None:
        raise ValueError('the destination template has no button counter')
    base = tab_blocks_bytes(new, dti, dtabi)
    at = len(base) if pos is None else pos
    if not 0 <= at <= len(base):
        raise ValueError('destination position must be 0..%d' % len(base))
    old = U(new, cp)
    hint = template_format(new, dtpl)
    if ids_look_right(new, dtpl, cp) and all((button_format(b) or hint) is not None for b in out):
        out = [with_button_id(b, old + 1 + k, hint) for k, b in enumerate(out)]
        res = bytearray(set_tab_blocks(new, dti, dtabi, base[:at] + out + base[at:]))
        struct.pack_into('<I', res, cp, old + len(out))
    else:
        res = bytearray(set_tab_blocks(new, dti, dtabi, base[:at] + out + base[at:]))
    return res, {'copied': len(out), 'skipped': skipped, 'added': added}


# ---------------------------------------------------------------- adding a part that no button in the profile uses yet
# Fabrication keeps a cache inside every button item (the part's pattern number, item type flags, sizes, a date ...). Before it uses a
# button it compares the stored date with the part file's own date and, when they differ, rebuilds the whole cache from the part file
# (CacheItem::Validate / ReloadCache, found in FabricationCoreDB.dll). So a new button is written with what can be read reliably from the
# part file (name, picture, pattern number, size list), the rest taken from a similar button as a placeholder, and a deliberately OLD date,
# which makes Fabrication rebuild all of it from the part file the first time the button is used.
_STALE_TIME = (2000, 1, 6, 1, 0, 0, 0, 0)          # 2000-01-01, as a Windows SYSTEMTIME: never equal to a part file's date
_LAB = re.compile(r'^\d+(?:[- ]\d+/\d+|/\d+)?(?:[xX]\d+(?:[- ]\d+/\d+|/\d+)?)*$')
_LAB_U16 = re.compile(rb'(?:[0-9xX/ \-]\x00){1,38}\x00\x00')
_LAB_A8 = re.compile(rb'[0-9xX/ \-]{1,38}\x00')
_PSIG = re.compile(re.escape(struct.pack('<5I', 0, 10, 1, 0, 1)) + b'(....)' + re.escape(struct.pack('<3I', 0, 0, 0xffffffff)), re.S)


def _label_value(p):
    from fractions import Fraction
    t = 0.0
    for part in p.replace('-', ' ').split():
        t += float(Fraction(part)) if '/' in part else float(part)
    return t


def itm_stream(data):
    """The decompressed contents of a part file (they are stored compressed, or sometimes plainly)."""
    if data.startswith(b'MAP Compressed File'):
        return decompress_bytes(data)
    return data


def itm_pattern_id(itm):
    """The part's pattern number (what Fabrication calls PatNo), or None."""
    m = _PSIG.search(itm)
    return struct.unpack('<I', m.group(1))[0] if m else None


def itm_size_rows(itm):
    """[(label, size 1, size 2)] listed in a part file: text labels (UTF-16 or 8-bit) that are followed by their own number.
    Checked against the stored lists of 86 real parts: identical for 74, and the others are parts that changed since."""
    best = None
    for enc, rx, step in (('u16', _LAB_U16, 2), ('a8', _LAB_A8, 1)):
        rows = []
        for m in rx.finditer(itm):
            st = m.start()
            if st < 4 or struct.unpack_from('<I', itm, st - 4)[0] != (m.end() - st) // step:
                continue
            try:
                s = itm[st:m.end() - step].decode('utf-16le' if enc == 'u16' else 'ascii')
            except UnicodeDecodeError:
                continue
            if not _LAB.match(s):
                continue
            want = _label_value(re.split('[xX]', s)[0])
            for off in range(0, 120, 4):
                if m.end() + off + 8 <= len(itm) and abs(struct.unpack_from('<d', itm, m.end() + off)[0] - want) < 1e-9:
                    rows.append((s, off, m.end()))
                    break
        if not rows:
            continue
        counts = {}
        for s, off, end in rows:
            counts[off] = counts.get(off, 0) + 1
        common = max(counts.items(), key=lambda kv: kv[1])[0]
        out, seen = [], set()
        for s, off, end in rows:
            if off != common or s in seen:
                continue
            seen.add(s)
            parts = re.split('[xX]', s)
            d1 = struct.unpack_from('<d', itm, end + off)[0]
            out.append((s, d1, _label_value(parts[1]) if len(parts) > 1 else d1))
        if best is None or len(out) > len(best):
            best = out
    return best or []


def _row_layout(row):
    v = U(row, 8)
    return (70, 46) if v == 34 else ((110, 86) if v == 74 else (None, None))


def _row_fields(row):
    lp, _ = _row_layout(row)
    if lp is None or len(row) < lp + 8:
        return None
    n = U(row, lp)
    end = lp + 4 + 2 * n
    if not 1 <= n <= 60 or end + 86 > len(row):
        return None
    return row[lp + 4:lp + 4 + 2 * (n - 1)].decode('utf-16le', 'replace'), struct.unpack_from('<d', row, end + 34)[0], struct.unpack_from('<d', row, end + 78)[0], row[end + 74]


def _make_size_row(donor, label, d1, d2, rb):
    """One size record from a donor's size record: only the label, the two sizes and the part's size byte are changed."""
    lp, lenpos = _row_layout(donor)
    old = _row_fields(donor)[0]
    n = U(donor, lp)
    new = bytearray(donor[:lp] + s16(label) + donor[lp + 4 + 2 * n:])
    struct.pack_into('<I', new, 4, U(donor, 4) + 2 * (len(label) - len(old)))
    struct.pack_into('<H', new, lenpos, 2 * (len(label) + 1) + 24)
    end = lp + 4 + 2 * (len(label) + 1)
    struct.pack_into('<d', new, end + 34, d1)
    struct.pack_into('<d', new, end + 78, d2)
    new[end + 74] = rb
    return bytes(new)


def _has_pstring(b):
    for o in range(0, len(b) - 8, 4):
        n = U(b, o)
        if 2 <= n <= 60 and o + 4 + 2 * n <= len(b) and b[o + 4 + 2 * (n - 1):o + 4 + 2 * n] == b'\x00\x00':
            ch = b[o + 4:o + 4 + 2 * (n - 1)]
            if all(ch[i + 1] == 0 and 32 <= ch[i] < 127 for i in range(0, len(ch), 2)):
                return True
    return False


def split_button(blk):
    """The parts of a single-item button record (the same for both record versions)."""
    n = U(blk, 12)
    if not 1 <= n <= 200:
        raise ValueError('not a button')
    m = _E0_.search(blk, 8)
    if not m:
        raise ValueError('no item')
    io = m.start()
    iend = io + 8 + U(blk, io + 4)
    if iend > len(blk):
        raise ValueError('item longer than the button')
    pn = U(blk, io + 16)
    p_end = io + 20 + 2 * pn
    qn = U(blk, p_end)
    q_end = p_end + 4 + 2 * qn
    rows = []
    for mm in _E0_.finditer(blk, q_end, iend):
        rows.append((mm.start(), mm.start() + 8 + U(blk, mm.start() + 4)))
    if rows:
        if any(rows[i][1] != rows[i + 1][0] for i in range(len(rows) - 1)) or rows[-1][1] != iend - 16:
            raise ValueError('size records do not line up')
        rs, re_ = rows[0][0], rows[-1][1]
    else:
        rs = re_ = iend - 16
    return {'blk': blk, 'name': blk[16:16 + 2 * (n - 1)].decode('utf-16le', 'replace'), 'name_end': 16 + 2 * n, 'io': io, 'iend': iend,
            'path': blk[io + 20:io + 20 + 2 * (pn - 1)].decode('utf-16le', 'replace'), 'png': blk[p_end + 4:p_end + 4 + 2 * (qn - 1)].decode('utf-16le', 'replace'),
            'q_end': q_end, 'rows': [blk[a:b] for a, b in rows], 'tail': blk[q_end:rs], 'h1': blk[16 + 2 * n:io], 'itail': blk[re_:iend], 'btail': blk[iend:]}


def _donor_ok(D):
    """A button can serve as the template for a new one only if it has the common shape and rebuilds one of its own size records exactly."""
    try:
        fmt = 'old' if D['rows'] and _row_layout(D['rows'][0])[0] == 70 else 'new'
        if not D['rows'] or len(D['tail']) != (108 if fmt == 'old' else 148) or len(D['h1']) < 44 or _has_pstring(D['h1']):
            return False
        if btn_tail(D['blk'], 0, len(D['blk'])) is None or not D['png'].startswith(('./', '.')) and D['png'] != '':
            return False
        r = D['rows'][0]
        f = _row_fields(r)
        if f is None or _row_layout(r)[0] is None:
            return False
        again = _make_size_row(r, f[0], f[1], f[2], f[3])
        return again[:-4] == r[:-4] and len(again) == len(r)
    except (struct.error, ValueError, IndexError):
        return False


def _cpl(a, b):
    n = 0
    for x, y in zip(a, b):
        if x != y:
            break
        n += 1
    return n


def _donor_format(D):
    return 'old' if D['rows'] and _row_layout(D['rows'][0])[0] == 70 else 'new'


def find_donor(raw, lay, b, rel, want_fmt=None, dest_ti=None):
    """The existing single-item button most like the new part: same pattern number first, then the closest name and folder.
    Returns (split button, template index) or None."""
    low = rel.lower()
    stem, folder = low.rsplit('/', 1)[-1].rsplit('.', 1)[0], low.rsplit('/', 1)[0]
    cands = []
    for ti, tpl in enumerate(lay['templates']):
        for tab in tpl['tabs']:
            for s, e in tab['blocks']:
                its = btn_items(raw, s, e)
                if len(its) != 1 or not its[0]['png_end']:
                    continue
                p = its[0]['path'].lower()
                score = (20000 if b is not None and U(raw, its[0]['png_end'] + 4) == b else 0) + 3 * _cpl(p.rsplit('/', 1)[-1].rsplit('.', 1)[0], stem) + _cpl(p.rsplit('/', 1)[0], folder)
                cands.append((score, ti, s, e))
    cands.sort(key=lambda c: -c[0])
    if want_fmt:                         # a template's buttons share one record version: that comes first, then the same pattern number, then the template's own buttons
        top = []
        for score, ti, s, e in cands:
            bonus = (50000 if button_format(bytes(raw[s:e])) == want_fmt else 0) + (10000 if ti == dest_ti else 0)
            top.append((score + bonus, ti, s, e))
        top.sort(key=lambda c: -c[0])
        cands = top
    for score, ti, s, e in cands[:700]:
        try:
            D = split_button(bytes(raw[s:e]))
        except (ValueError, struct.error):
            continue
        if _donor_ok(D):
            return D, ti
    return None


def _clean_button_tail(D):
    """The donor's button tail with an empty button code and 'not excluded from fill'."""
    blk, iend = D['blk'], D['iend']
    t = btn_tail(blk, 0, len(blk))
    cs = t[0]
    n = U(blk, cs)
    bt = bytearray(blk[iend:cs] + s16('') + blk[cs + 4 + 2 * n:])
    fp = cs - 8 - iend
    if fp >= 0 and struct.unpack_from('<H', bt, fp)[0] in (0, 1):
        struct.pack_into('<H', bt, fp, 0)
    return bytes(bt)


def build_new_button(D, name, path, png, pattern, rows):
    """A button record for a part, from a donor button of the same shape. rows = [(label, size 1, size 2)] (may be empty)."""
    blk = D['blk']
    tail = bytearray(D['tail'])
    if pattern is not None:
        struct.pack_into('<I', tail, 4, pattern)
    struct.pack_into('<I', tail, 8, 0)                          # no size restriction
    struct.pack_into('<8H', tail, 16, *_STALE_TIME)             # out of date on purpose: Fabrication rebuilds the cache from the part file
    if len(tail) == 148:
        tail[94:102] = b'\x00' * 8                              # the newer version's file CRC: none, so Fabrication does not trust the cache
    f0 = _row_fields(D['rows'][0])
    rb = f0[3]
    rws = [_make_size_row(D['rows'][0], l, d1, d2, rb) for l, d1, d2 in rows]
    rs = b''.join(rws)
    struct.pack_into('<II', tail, len(tail) - 8, len(rs) + 20, len(rws))
    C = U(blk, D['io'] + 8) - (2 * (len(D['path']) + 1) + 2 * (len(D['png']) + 1))
    item = bytearray(struct.pack('<IIII', 0x5c03e0, 0, 2 * (len(path) + 1) + 2 * (len(png) + 1) + C, U(blk, D['io'] + 12)) + s16(path) + s16(png) + bytes(tail) + rs + D['itail'])
    struct.pack_into('<I', item, 4, len(item) - 8)
    bt = _clean_button_tail(D)
    h1 = bytearray(D['h1'])
    struct.pack_into('<I', h1, len(h1) - 8, len(item) - 8 + len(bt) + 12)
    out = bytearray(struct.pack('<III', 0x5c03df, 0, U(blk, 8) + 2 * (len(name) - len(D['name']))) + s16(name) + bytes(h1) + bytes(item) + bt)
    struct.pack_into('<I', out, 4, len(out) - 8)
    return bytes(out)


def library_png(itm_file, rel):
    """The picture path for a part: its own name with .png (or .jpg, .bmp) beside it, else a picture named inside the part file; '' if none."""
    d = os.path.dirname(itm_file)
    stem = os.path.splitext(os.path.basename(itm_file))[0]
    folder = rel.rsplit('/', 1)[0]
    try:
        low = {n.lower(): n for n in os.listdir(d)}
    except OSError:
        return ''
    for ext in ('.png', '.jpg', '.bmp'):
        n = low.get((stem + ext).lower())
        if n:
            return folder + '/' + n
    return ''


def add_new_part(raw, ti, tabi, pos, rel, itm_data, png=''):
    """Add a button for a part file that no button of the profile uses. Returns (new bytes, info). The new bytes are checked before they
    are returned: the whole profile must still pass the structure check and the new button must read back as intended."""
    itm = itm_stream(itm_data)
    lay = layout(raw)
    pattern = itm_pattern_id(itm)
    rows = itm_size_rows(itm)
    found = find_donor(raw, lay, pattern, rel, template_format(raw, lay['templates'][ti]), ti)
    if not found:
        raise ValueError('There is no button in this profile to use as a template for a new one.')
    D, dti = found
    name = rel.rsplit('/', 1)[-1].rsplit('.', 1)[0]
    if any(it['path'].lower() == rel.lower() for s, e in lay['templates'][ti]['tabs'][tabi]['blocks'] for it in btn_items(raw, s, e)):
        raise ValueError('That part is already in this tab.')
    blk = build_new_button(D, name, rel, png, pattern, rows)
    chk = split_button(blk)
    if chk['path'] != rel or chk['name'] != name or chk['png'] != png or len(chk['rows']) != len(rows) or (rows and button_format(blk) is None):
        raise ValueError('The new button did not read back as intended, so nothing was added.')
    dtpl = lay['templates'][ti]
    cp = counter_pos(raw, dtpl)
    old = None
    if cp is not None and ids_look_right(raw, dtpl, cp):          # a new button gets the next id; its version is the donor's, which is known
        old = U(raw, cp)
        new = bytearray(insert_block(raw, ti, tabi, pos, with_button_id(blk, old + 1, _donor_format(D))))
        struct.pack_into('<I', new, cp, old + 1)
    else:
        new = insert_block(raw, ti, tabi, pos, blk)
    probs = validate(bytes(new), deep=False)
    if probs:
        raise ValueError('The profile would not pass its structure check (%s), so nothing was added.' % probs[0])
    tname = lay['templates'][dti]['name']
    return new, {'name': name, 'sizes': len(rows), 'pattern': pattern, 'donor': '%s (template %s)' % (D['name'], tname), 'png': bool(png), 'counter_before': old}


# ---------------------------------------------------------------- stacked buttons: several items (parts) in one button
# A stacked button is an ordinary button record with more than one item record one after the other. Nothing counts them: only the
# button's own size and the length field at the end of its header (= bytes from the first item to the end of the button, plus 4;
# checked on all 1,046 stacked buttons and every single-item one in three real files) depend on how many there are.
def _fix_button_lengths(blk):
    out = bytearray(blk)
    io = _E0_.search(blk, 8).start()
    struct.pack_into('<I', out, 4, len(out) - 8)
    struct.pack_into('<I', out, io - 8, len(out) - io + 4)
    return bytes(out)


def stack_item(blk, item):
    """The button with one more item record after its last one."""
    last = btn_items(blk, 0, len(blk))[-1]['end']
    return _fix_button_lengths(blk[:last] + item + blk[last:])


def unstack_item(blk, k):
    """The button without its k-th item (a button keeps at least one)."""
    items = btn_items(blk, 0, len(blk))
    if len(items) < 2:
        raise ValueError('A button needs at least one part.')
    if not 0 <= k < len(items):
        raise ValueError('this button has no item number %d' % (k + 1))
    return _fix_button_lengths(blk[:items[k]['start']] + blk[items[k]['end']:])


def add_item_to_button(raw, ti, tabi, idx, item):
    blocks = tab_blocks_bytes(raw, ti, tabi)
    if not 0 <= idx < len(blocks):
        raise ValueError('button position must be 0..%d' % (len(blocks) - 1))
    blocks[idx] = stack_item(blocks[idx], item)
    return set_tab_blocks(raw, ti, tabi, blocks)


def remove_item_from_button(raw, ti, tabi, idx, k):
    blocks = tab_blocks_bytes(raw, ti, tabi)
    if not 0 <= idx < len(blocks):
        raise ValueError('button position must be 0..%d' % (len(blocks) - 1))
    blocks[idx] = unstack_item(blocks[idx], k)
    return set_tab_blocks(raw, ti, tabi, blocks)


def existing_item_record(raw, loc, rel, want_fmt):
    """The item record of the button at loc (template, tab, position) that holds the part rel, with no size restriction; None if that
    button uses a different record version than the one it will be stacked on."""
    ti, tabi, n = loc
    s, e = layout(raw)['templates'][ti]['tabs'][tabi]['blocks'][n]
    blk = bytes(raw[s:e])
    f = button_format(blk)
    if f and want_fmt and f != want_fmt:
        return None
    for it in btn_items(blk, 0, len(blk)):
        if it['path'].lower() == rel.lower():
            rec = bytearray(blk[it['start']:it['end']])
            if it['cid_pos'] - it['start'] + 4 <= len(rec):
                struct.pack_into('<I', rec, it['cid_pos'] - it['start'], 0)
            return bytes(rec)
    return None


def new_part_item(raw, ti, want_fmt, rel, itm_data, png=''):
    """An item record for a part no button uses yet, built the same way as a new button's (see add_new_part)."""
    itm = itm_stream(itm_data)
    lay = layout(raw)
    pattern = itm_pattern_id(itm)
    rows = itm_size_rows(itm)
    found = find_donor(raw, lay, pattern, rel, want_fmt, ti)
    if not found:
        raise ValueError('There is no button in this profile to use as a template for a new one.')
    D, _ = found
    blk = build_new_button(D, rel.rsplit('/', 1)[-1].rsplit('.', 1)[0], rel, png, pattern, rows)
    S = split_button(blk)
    return blk[S['io']:S['iend']], {'name': rel.rsplit('/', 1)[-1].rsplit('.', 1)[0], 'sizes': len(rows)}


def tab_bytes(raw, ti, tabi):
    lay = layout(raw)
    tpl, tab = _at(lay, ti, tabi)
    return bytes(raw[tab['x']:tab['xe']])


def delete_tab(raw, ti, tabi):
    """Remove a whole tab with its buttons. The template keeps a tab count in front of its first tab; the template and the
    collection sizes shrink. Tab numbers are ids (native leaves gaps), so the other tabs are not renumbered."""
    lay = layout(raw)
    tpl, tab = _at(lay, ti, tabi)
    if len(tpl['tabs']) <= 1:
        raise ValueError('A template needs at least one tab. Add another tab first, or delete the buttons instead.')
    size = tab['xe'] - tab['x']
    new = bytearray(raw)

    def bump(off, d):
        struct.pack_into('<I', new, off, (U(raw, off) + d) & 0xFFFFFFFF)
    bump(tpl['tabs'][0]['x'] - 4, -1)
    bump(tpl['t'] + 4, -size)
    bump(lay['first'] - 12, -size)
    bump(lay['first'] - 8, -size)
    _bump_enclosing(new, raw, lay, tpl, -size)
    new[tab['x']:tab['xe']] = b''
    return new


def insert_tab(raw, ti, pos, data):
    """Put a whole tab back (as taken out by tab_bytes) at tab position pos (None = after the last tab)."""
    lay = layout(raw)
    tpl = lay['templates'][ti]
    n = len(tpl['tabs'])
    pos = n if pos is None else pos
    if not 0 <= pos <= n or not n:
        raise ValueError('tab position must be 0..%d' % n)
    at = tpl['tabs'][pos]['x'] if pos < n else tpl['tabs'][-1]['xe']
    new = bytearray(raw)

    def bump(off, d):
        struct.pack_into('<I', new, off, (U(raw, off) + d) & 0xFFFFFFFF)
    bump(tpl['tabs'][0]['x'] - 4, 1)
    bump(tpl['t'] + 4, len(data))
    bump(lay['first'] - 12, len(data))
    bump(lay['first'] - 8, len(data))
    _bump_enclosing(new, raw, lay, tpl, len(data))
    new[at:at] = data
    return new


def tab_blocks_bytes(raw, ti, tabi):
    """Every button record of one tab, as bytes (used for bulk edits and to remember what to undo to)."""
    lay = layout(raw)
    tpl, tab = _at(lay, ti, tabi)
    return [bytes(raw[s:e]) for s, e in tab['blocks']]


def set_tab_blocks(raw, ti, tabi, blocks):
    """Replace all button records of a tab with 'blocks' (any order, any number, including none); fixes every size and count."""
    lay = layout(raw)
    tpl, tab = _at(lay, ti, tabi)
    start, end = _list_pos(raw, tab)[2], tab['xe']
    data = b''.join(blocks)
    return _apply(raw, lay, tpl, tab, len(data) - (end - start), len(blocks) - len(tab['blocks']), (start, end - start, data))


def _picked(n, idxs):
    s = sorted(set(int(i) for i in idxs))
    if not s or s[0] < 0 or s[-1] >= n:
        raise ValueError('button positions must be 0..%d' % (n - 1))
    return s


def delete_many(raw, ti, tabi, idxs):
    blocks = tab_blocks_bytes(raw, ti, tabi)
    s = set(_picked(len(blocks), idxs))
    keep = [b for i, b in enumerate(blocks) if i not in s]
    return set_tab_blocks(raw, ti, tabi, keep)


def move_many(raw, ti, tabi, idxs, to):
    """Take the picked buttons out (keeping their order) and put them back as a group at position 'to' of the rest."""
    blocks = tab_blocks_bytes(raw, ti, tabi)
    s = _picked(len(blocks), idxs)
    sel = [blocks[i] for i in s]
    rest = [b for i, b in enumerate(blocks) if i not in set(s)]
    if not 0 <= to <= len(rest):
        raise ValueError('destination position must be 0..%d' % len(rest))
    return set_tab_blocks(raw, ti, tabi, rest[:to] + sel + rest[to:])


def copy_many(raw, ti, tabi, idxs, dtabi, pos=None):
    """Copy the picked buttons into tab dtabi of the same template at position pos (None = end)."""
    src = tab_blocks_bytes(raw, ti, tabi)
    sel = [src[i] for i in _picked(len(src), idxs)]
    base = src if dtabi == tabi else tab_blocks_bytes(raw, ti, dtabi)
    pos = len(base) if pos is None else pos
    if not 0 <= pos <= len(base):
        raise ValueError('destination position must be 0..%d' % len(base))
    cp = counter_pos(raw, layout(raw)['templates'][ti])
    if cp is None:
        raise ValueError('this template has no button counter')
    old = U(raw, cp)
    tpl_ = layout(raw)['templates'][ti]
    hint = template_format(raw, tpl_)
    if ids_look_right(raw, tpl_, cp) and all((button_format(b) or hint) is not None for b in sel):
        sel = [with_button_id(b, old + 1 + k, hint) for k, b in enumerate(sel)]  # every copy is a new button: new ids
        new = set_tab_blocks(raw, ti, dtabi, base[:pos] + sel + base[pos:])
        struct.pack_into('<I', new, cp, old + len(sel))
        return new
    return set_tab_blocks(raw, ti, dtabi, base[:pos] + sel + base[pos:])      # unknown layout: copy exactly, change no ids


_E_TAIL = re.compile(r'\s*\[E\]\s*$')


def tab_is_excluded(name):
    return bool(_E_TAIL.search(name or ''))


def rename_tab(raw, ti, tabi, newname):
    """Change a tab's name. The name sits in the tab header, so five sizes move with it: the tab's size and inner size,
    its name length, the template size and the two sizes of the collection in front of the first template."""
    lay = layout(raw)
    tpl, tab = _at(lay, ti, tabi)
    x = tab['x']
    n, inner = U(raw, x + 12), U(raw, x + 8)
    if not (1 <= n < 300 and inner == 2 * n + 28 and raw[x + 14 + 2 * n:x + 16 + 2 * n] == b'\x00\x00'):
        raise ValueError('this tab header has an unexpected layout; not changed')
    if len(newname) > 200 or any(ord(c) < 32 for c in newname):
        raise ValueError('that tab name is not allowed')
    nn = len(newname) + 1
    delta = 2 * (nn - n)
    new = bytearray(raw)

    def bump(off, d):
        struct.pack_into('<I', new, off, (U(raw, off) + d) & 0xFFFFFFFF)
    bump(x + 4, delta)
    bump(x + 8, delta)
    struct.pack_into('<I', new, x + 12, nn)
    bump(tpl['t'] + 4, delta)
    bump(lay['first'] - 12, delta)
    bump(lay['first'] - 8, delta)
    _bump_enclosing(new, raw, lay, tpl, delta)
    new[x + 16:x + 16 + 2 * n] = newname.encode('utf-16le') + b'\x00\x00'
    return new


def toggle_tab_exclude(raw, ti, tabi):
    """'Exclude From Fill' for a tab: the stock editor marks it with ' [E]' at the end of the tab name."""
    lay = layout(raw)
    tpl, tab = _at(lay, ti, tabi)
    name = '' if tab['name'] == '(unnamed)' else tab['name']
    if tab_is_excluded(name):
        new = _E_TAIL.sub('', name)
    else:
        new = (name + ' [E]') if name else '[E]'
    return rename_tab(raw, ti, tabi, new), name, new


def _find(lay, tname, tabname):
    ts = [i for i, t in enumerate(lay['templates']) if t['name'] == tname]
    if len(ts) != 1:
        raise ValueError('template %r: %d matches' % (tname, len(ts)))
    tabs = [i for i, x in enumerate(lay['templates'][ts[0]]['tabs']) if x['name'] == tabname]
    if len(tabs) != 1:
        raise ValueError('tab %r in %r: %d matches' % (tabname, tname, len(tabs)))
    return ts[0], tabs[0]


# name-based, 1-based wrappers used by the command line
def move_button(raw, tname, tabname, frm, to):
    ti, tb = _find(layout(raw), tname, tabname)
    return move_idx(raw, ti, tb, frm - 1, to - 1)


def delete_button(raw, tname, tabname, idx):
    ti, tb = _find(layout(raw), tname, tabname)
    return delete_idx(raw, ti, tb, idx - 1)


def copy_button(raw, tname, src_tab, idx, dst_tab, pos=None, dst_template=None):
    lay = layout(raw)
    ti, tb = _find(lay, tname, src_tab)
    dti, dtb = _find(lay, dst_template or tname, dst_tab)
    return copy_idx(raw, ti, tb, idx - 1, dti, dtb, None if pos is None else pos - 1)


def write_inplace(path, raw, header):
    """Rewrite the map file where it is. The new content is built first and swapped in as one step."""
    data = build_container(fix_enclosing_sizes(raw), header)      # repairs those three size fields if an older save left them stale
    tmp = path + '.~fabmap'
    try:
        with open(tmp, 'wb') as f:
            f.write(data)
        os.replace(tmp, path)
    finally:
        if os.path.exists(tmp):
            try:
                os.remove(tmp)
            except Exception:
                pass


# ---------------------------------------------------------------- compare
def describe_offset(lay, off):
    for tpl in lay['templates']:
        if tpl['t'] <= off < tpl['te']:
            for tab in tpl['tabs']:
                if tab['x'] <= off < tab['xe']:
                    for k, (s, e) in enumerate(tab['blocks']):
                        if s <= off < e:
                            return 'template %r, tab %r, button %d, +%d bytes in' % (tpl['name'], tab['name'], k + 1, off - s)
                    return 'template %r, tab %r header (+%d)' % (tpl['name'], tab['name'], off - tab['x'])
            return 'template %r (+%d)' % (tpl['name'], off - tpl['t'])
    return 'outside templates (offset %d)' % off


def compare(a, b, lay=None):
    """Describe differences between two decompressed streams."""
    if a == b:
        return ['IDENTICAL: %d bytes' % len(a)]
    out = ['DIFFERENT: %d vs %d bytes (%+d)' % (len(a), len(b), len(b) - len(a))]
    lay = lay or layout(a)
    i = 0
    n = min(len(a), len(b))
    regions = 0
    while i < n and regions < 12:
        if a[i] != b[i]:
            j = i
            while j < n and a[j] != b[j]:
                j += 1
            out.append('  differs at %d..%d (%d bytes): %s | mine=%s theirs=%s' % (
                i, j - 1, j - i, describe_offset(lay, i), bytes(a[i:min(j, i + 8)]).hex(' '), bytes(b[i:min(j, i + 8)]).hex(' ')))
            regions += 1
            i = j
        else:
            i += 1
    if len(a) != len(b) and regions == 0:
        out.append('  one stream is a prefix of the other')
    return out


# ---------------------------------------------------------------- self test
def selftest(path):
    data, raw = read_map(path)
    ok = True

    def check(name, cond, extra=''):
        nonlocal ok
        ok &= bool(cond)
        print(('PASS  ' if cond else 'FAIL  ') + name + (('  ' + extra) if extra else ''))

    check('container: rebuilding the unchanged file gives identical bytes', build_container(raw, data[:HEADER_LEN]) == data)
    probs = validate(raw)
    check('structure of the original file is consistent', not probs, str(probs[:3]))
    lay = layout(raw)
    # pick a template with a tab of >=4 buttons and another tab
    pick = None
    for t in lay['templates']:
        big = [x for x in t['tabs'] if len(x['blocks']) >= 4]
        other = [x for x in t['tabs'] if x is not big[0]] if big else []
        if big and other and other[0]['blocks']:
            pick = (t['name'], big[0]['name'], other[0]['name'])
            break
    tn, ta, tb = pick
    print('      using palette %r, tabs %r and %r' % pick)
    # reorder and reverse
    m = move_button(raw, tn, ta, 3, 1)
    check('move 3->1 changes the file, keeps size, stays consistent', bytes(m) != raw and len(m) == len(raw) and not validate(bytes(m)))
    check('move 3->1 then 1->3 restores the original exactly', bytes(move_button(bytes(m), tn, ta, 1, 3)) == raw)
    # copy to other tab, then delete the copy
    c = copy_button(raw, tn, ta, 2, tb)
    cv = validate(bytes(c))
    check('copy button 2 to the other tab: structure consistent', not cv, str(cv[:3]))
    nb = len(layout(bytes(c))['templates'][[t['name'] for t in lay['templates']].index(tn)]['tabs'][[x['name'] for x in lay['templates'][[t['name'] for t in lay['templates']].index(tn)]['tabs']].index(tb)]['blocks'])
    check('copy added exactly one button to that tab', nb == len([x for x in lay['templates'] if x['name'] == tn][0]['tabs'][[y['name'] for y in [x for x in lay['templates'] if x['name'] == tn][0]['tabs']].index(tb)]['blocks']) + 1)
    d = delete_button(bytes(c), tn, tb, nb)
    tix = [t['name'] for t in lay['templates']].index(tn)
    check('delete the copy and put the id counter back: the original exactly (a copy is a new button: new id, counter + 1)', bytes(set_template_counter(bytes(d), tix, template_counter(raw, tix))) == raw and button_id(block_bytes(bytes(c), tix, [x['name'] for x in layout(bytes(c))['templates'][tix]['tabs']].index(tb), nb - 1)) == template_counter(raw, tix) + 1)
    # delete then copy back to same place
    s, e = [x for x in lay['templates'] if x['name'] == tn][0]['tabs'][[y['name'] for y in [x for x in lay['templates'] if x['name'] == tn][0]['tabs']].index(ta)]['blocks'][1]
    d2 = delete_button(raw, tn, ta, 2)
    check('delete button 2: structure consistent and file shorter by that button', not validate(bytes(d2)) and len(raw) - len(d2) == e - s)
    # bulk edits and their exact reversal
    nb = len(tab_blocks_bytes(raw, [i for i, x in enumerate(lay['templates']) if x['name'] == tn][0], [y['name'] for y in [x for x in lay['templates'] if x['name'] == tn][0]['tabs']].index(ta)))
    tidx = [i for i, x in enumerate(lay['templates']) if x['name'] == tn][0]
    taidx = [y['name'] for y in lay['templates'][tidx]['tabs']].index(ta)
    tbidx = [y['name'] for y in lay['templates'][tidx]['tabs']].index(tb)
    orig_a, orig_b = tab_blocks_bytes(raw, tidx, taidx), tab_blocks_bytes(raw, tidx, tbidx)
    dm = bytes(delete_many(raw, tidx, taidx, [1, 3, 5]))
    check('delete 3 buttons at once: structure consistent, 3 fewer buttons', not validate(dm) and len(tab_blocks_bytes(dm, tidx, taidx)) == len(orig_a) - 3)
    check('...and restoring the tab gives back the original exactly', bytes(set_tab_blocks(dm, tidx, taidx, orig_a)) == raw)
    mm = bytes(move_many(raw, tidx, taidx, [0, 2, 4], 1))
    check('move 3 buttons as a group: same size, consistent, order changed', not validate(mm) and len(mm) == len(raw) and mm != raw)
    check('...and restoring gives back the original exactly', bytes(set_tab_blocks(mm, tidx, taidx, orig_a)) == raw)
    cm = bytes(copy_many(raw, tidx, taidx, [0, 1, 2], tbidx))
    check('copy 3 buttons into another tab: consistent, 3 more there', not validate(cm) and len(tab_blocks_bytes(cm, tidx, tbidx)) == len(orig_b) + 3)
    ids3 = [button_id(b) for b in tab_blocks_bytes(cm, tidx, tbidx)[-3:]]
    check('...the 3 copies got the next 3 ids (counter + 1..3) and the counter moved up by 3', ids3 == [template_counter(raw, tidx) + k for k in (1, 2, 3)] and template_counter(cm, tidx) == template_counter(raw, tidx) + 3)
    check('...restoring that tab and the counter gives back the original exactly', bytes(set_template_counter(set_tab_blocks(cm, tidx, tbidx, orig_b), tidx, template_counter(raw, tidx))) == raw)
    cs = bytes(copy_many(raw, tidx, taidx, [0, 1], taidx, 2))
    check('duplicate 2 buttons inside the same tab: consistent', not validate(cs) and len(tab_blocks_bytes(cs, tidx, taidx)) == len(orig_a) + 2)
    # tab exclusion = tab name with ' [E]'
    ex, was, now = toggle_tab_exclude(raw, tidx, taidx); ex = bytes(ex)
    check('exclude a tab: name %r -> %r, structure consistent, file grew by %d bytes' % (was, now, len(ex) - len(raw)), not validate(ex) and now == was + ' [E]' and len(ex) - len(raw) == 8)
    lay2 = layout(ex); check('...the new name is read back from the file', lay2['templates'][tidx]['tabs'][taidx]['name'] == now)
    back, w2, n2 = toggle_tab_exclude(ex, tidx, taidx)
    check('include it again: identical to the original, byte for byte', bytes(back) == raw)
    ren = bytes(rename_tab(raw, tidx, taidx, 'A much longer tab name than before')); check('any rename keeps the file consistent and reverses exactly', not validate(ren) and bytes(rename_tab(ren, tidx, taidx, ta)) == raw)
    # Exclude From Fill: buttons and tabs
    bb = tab_blocks_bytes(raw, tidx, taidx)
    ex1, olds = set_button_flags(raw, tidx, taidx, [1, 3], True); ex1 = bytes(ex1)
    diff = [i for i in range(len(raw)) if raw[i] != ex1[i]]
    check('exclude 2 buttons: same file size, only flag bytes change (%d byte(s) differ), structure consistent' % len(diff), len(ex1) == len(raw) and len(diff) <= 2 and not validate(ex1))
    lay3 = layout(ex1); s1, e1 = lay3['templates'][tidx]['tabs'][taidx]['blocks'][1]
    check('...the two buttons now read back as excluded and a neighbour does not', btn_excluded(ex1, s1, e1) is True and btn_excluded(ex1, *lay3['templates'][tidx]['tabs'][taidx]['blocks'][2]) is False)
    check('...and putting the old values back gives the original exactly', bytes(restore_button_flags(ex1, tidx, taidx, olds)) == raw)
    tf, oldf = set_tab_flag(raw, tidx, taidx, True); tf = bytes(tf)
    check('exclude a tab: only its 4-byte mode value changes, structure consistent', len(tf) == len(raw) and sum(1 for i in range(len(raw)) if raw[i] != tf[i]) <= 4 and not validate(tf) and tab_excluded(tf, layout(tf)['templates'][tidx]['tabs'][taidx]))
    check('...and putting it back gives the original exactly', bytes(set_tab_flag(tf, tidx, taidx, oldf)[0]) == raw)
    # size restrictions
    tplx = lay['templates'][tidx]
    conds = template_conditions(raw, tplx)
    check('the template\'s conditions read back: %s' % ['%s %s..%s' % (c['name'], c['gt'], c['lte']) for c in conds][:4], len(conds) >= 1 and conds[0]['gt'] == -1.0 and sorted(c['id'] for c in conds) == [c['id'] for c in conds])
    s0, e0 = layout(raw)['templates'][tidx]['tabs'][taidx]['blocks'][0]
    its = btn_items(raw, s0, e0)
    check('a button\'s item and its condition id read back (%d item, id %s)' % (len(its), its[0]['cid'] if its else '-'), len(its) >= 1 and its[0]['cid'] in [c['id'] for c in conds])
    if len(conds) > 1:
        other = [c['id'] for c in conds if c['id'] != its[0]['cid']][0]
        n1, old = set_item_condition(raw, tidx, taidx, 0, 0, other); n1 = bytes(n1)
        check('change an item\'s condition: 4 bytes change, file valid, reads back', len(n1) == len(raw) and sum(1 for i in range(len(raw)) if raw[i] != n1[i]) <= 4 and not validate(n1) and btn_items(n1, *layout(n1)['templates'][tidx]['tabs'][taidx]['blocks'][0])[0]['cid'] == other)
        check('...and putting it back gives the original exactly', bytes(set_item_condition(n1, tidx, taidx, 0, 0, old)[0]) == raw)
        n2, olds = set_condition_limits(raw, tidx, conds[-1]['id'], 3.0, None); n2 = bytes(n2)
        check('change a condition\'s limits: file valid, reads back, 16 bytes at most differ', len(n2) == len(raw) and not validate(n2) and template_conditions(n2, layout(n2)['templates'][tidx])[-1]['gt'] == 3.0 and sum(1 for i in range(len(raw)) if raw[i] != n2[i]) <= 16)
        check('...and putting it back gives the original exactly', bytes(set_condition_limits(n2, tidx, conds[-1]['id'], None if olds[0] == -1.0 else olds[0], None if olds[1] == -1.0 else olds[1])[0]) == raw)
    try:
        set_item_condition(raw, tidx, taidx, 0, 0, 987654); check('an unknown condition number is refused', False)
    except ValueError:
        check('an unknown condition number is refused', True)
    # the three enclosing size fields follow every size change (the stock editor changes them too)
    for nm, fnc in (('rename a tab', lambda: rename_tab(raw, tidx, taidx, 'A much longer tab name than before')),
                    ('copy a button', lambda: copy_idx(raw, tidx, taidx, 0, tidx, tbidx)),
                    ('delete 3 buttons', lambda: delete_many(raw, tidx, taidx, [0, 1, 2])),
                    ('delete a tab', lambda: delete_tab(raw, tidx, taidx))):
        r2 = bytes(fnc())
        check('%s: stream size, record around the templates and the template inner record all follow' % nm, not enclosing_problems(r2) and len(r2) != len(raw), str(enclosing_problems(r2)[:2]))
    stale = bytearray(raw); struct.pack_into('<I', stale, 20, 12345)
    check('an out-of-date stream size is repaired when writing, and nothing else changes', len(fix_enclosing_sizes(bytes(stale))) == len(raw) and bytes(fix_enclosing_sizes(bytes(stale))) == raw)
    # the template's list of size restrictions
    tpl_c = layout(raw)['templates'][tidx]
    c0 = template_conditions(raw, tpl_c)
    ac, nid = add_condition(raw, tidx, 'Self-test limit', 3.0, 6.0); ac = bytes(ac)
    c1 = template_conditions(ac, layout(ac)['templates'][tidx])
    check('add a restriction: the file stays valid (all the enclosing sizes too), %d -> %d restrictions, new id %d, limits read back' % (len(c0), len(c1), nid), not validate(ac) and len(c1) == len(c0) + 1 and c1[-1]['name'] == 'Self-test limit' and (c1[-1]['gt'], c1[-1]['lte']) == (3.0, 6.0) and nid == max(c['id'] for c in c0) + 1)
    dc, rec, st, pt = delete_condition(ac, tidx, nid)
    check('deleting it again gives the original file exactly', bytes(dc) == raw)
    rc, oldn = rename_condition(ac, tidx, nid, 'A much longer restriction name here'); rc = bytes(rc)
    check('rename it: valid, new name read back, and renaming back is exact', not validate(rc) and template_conditions(rc, layout(rc)['templates'][tidx])[-1]['name'] == 'A much longer restriction name here' and bytes(rename_condition(rc, tidx, nid, oldn)[0]) == ac)
    for bad_name, why in (('', 'empty'), (c0[0]['name'], 'duplicate')):
        try:
            add_condition(raw, tidx, bad_name, None, None); check('a %s name is refused' % why, False)
        except ValueError:
            check('a %s name is refused' % why, True)
    if len(c0) > 1:
        s0, e0 = layout(raw)['templates'][tidx]['tabs'][taidx]['blocks'][0]
        usecid = c0[-1]['id']
        un, uo = set_item_condition(raw, tidx, taidx, 0, 0, usecid); un = bytes(un)
        dd, drec, dst, dpt = delete_condition(un, tidx, usecid); dd = bytes(dd)
        check('deleting a restriction that buttons use sends every one of them back to Unrestricted (%d items changed)' % len(dpt), not validate(dd) and len(dpt) >= 1 and all(it['cid'] != usecid for t_ in layout(dd)['templates'][tidx]['tabs'] for (s_, e_) in t_['blocks'] for it in btn_items(dd, s_, e_)) and btn_items(dd, *layout(dd)['templates'][tidx]['tabs'][taidx]['blocks'][0])[0]['cid'] == 0)
        check('...and undoing the delete restores the file and the button exactly', bytes(restore_condition(dd, tidx, dst, drec, dpt)) == un)
    try:
        delete_condition(raw, tidx, 0); check('the Unrestricted entry cannot be deleted', False)
    except ValueError:
        check('the Unrestricted entry cannot be deleted', True)
    mn, mo, sk = set_item_conditions_many(raw, tidx, taidx, [0, 1, 2], c0[-1]['id']); mn = bytes(mn)
    check('apply one restriction to 3 selected buttons at once (%d changed, %d skipped): valid' % (len(mo), sk), not validate(mn) and len(mo) + sk <= 3)
    nb = bytearray(mn)
    for p, v in mo: struct.pack_into('<I', nb, p, v)
    check('...and restoring the old values gives the original file exactly', bytes(nb) == raw)
    # copying buttons from one template into another
    names_ = [t['name'] for t in layout(raw)['templates']]
    srcT = next((i for i, t in enumerate(layout(raw)['templates']) if i != tidx and t['tabs'] and t['tabs'][0]['blocks'] and len(t['tabs'][0]['blocks']) >= 4), None)
    if srcT is not None:
        cx, info = copy_across(raw, srcT, 0, [0, 1, 2], tidx, taidx, 0); cx = bytes(cx)
        d_ = layout(cx)['templates'][tidx]['tabs'][taidx]['blocks']
        check('copy 3 buttons from another template into this one: valid, %d copied, %d already there, restrictions added: %s' % (info['copied'], info['skipped'], info['added'] or 'none'),
              not validate(cx) and not enclosing_problems(cx) and info['copied'] + info['skipped'] == 3 and len(d_) == len(layout(raw)['templates'][tidx]['tabs'][taidx]['blocks']) + info['copied'])
        dt_ = layout(cx)['templates'][tidx]
        check('every restriction on the copies points at a restriction that exists in the destination, and the ids are unique',
              ids_look_right(cx, dt_, counter_pos(cx, dt_)) and all(it['cid'] in [c['id'] for c in template_conditions(cx, dt_)] for s_, e_ in d_[:info['copied']] for it in btn_items(cx, s_, e_)))
        try:
            copy_across(cx, srcT, 0, [0, 1, 2], tidx, taidx, 0); check('copying the same buttons again is refused (they are already there)', False)
        except ValueError:
            check('copying the same buttons again is refused (they are already there)', True)
        check('the source template is untouched', tab_blocks_bytes(cx, srcT, 0) == tab_blocks_bytes(raw, srcT, 0))
    # empty tabs (real files have 99 of them) and deleting a whole tab
    allb = tab_blocks_bytes(raw, tidx, taidx)
    em = bytes(delete_many(raw, tidx, taidx, list(range(len(allb)))))
    check('delete every button of a tab: file stays valid and the tab is empty (%d buttons removed)' % len(allb), not validate(em) and not layout(em)['templates'][tidx]['tabs'][taidx]['blocks'])
    check('...putting the buttons back gives the original exactly', bytes(set_tab_blocks(em, tidx, taidx, allb)) == raw)
    one = bytes(insert_block(em, tidx, taidx, None, allb[0]))
    check('a button can be copied into an empty tab', not validate(one) and len(layout(one)['templates'][tidx]['tabs'][taidx]['blocks']) == 1)
    lone = bytes(delete_idx(one, tidx, taidx, 0))
    check('deleting the only button leaves a valid empty tab again, identical to before', lone == em)
    tdat = tab_bytes(raw, tidx, taidx)
    dt = bytes(delete_tab(raw, tidx, taidx))
    ntabs = len(lay['templates'][tidx]['tabs'])
    check('delete a whole tab: valid, one tab fewer (%d -> %d), template and collection sizes follow' % (ntabs, len(layout(dt)['templates'][tidx]['tabs'])), not validate(dt) and len(layout(dt)['templates'][tidx]['tabs']) == ntabs - 1 and len(dt) == len(raw) - len(tdat))
    check('...putting it back at the same place gives the original exactly', bytes(insert_tab(dt, tidx, taidx, tdat)) == raw)
    lt = bytes(delete_tab(raw, tidx, 0)); check('deleting the first tab works the same way and undoes exactly', not validate(lt) and bytes(insert_tab(lt, tidx, 0, tab_bytes(raw, tidx, 0))) == raw)
    try:
        delete_tab(dt if ntabs == 2 else raw, tidx, 0) if ntabs == 2 else (_ for _ in ()).throw(ValueError('skip'))
        check('the last remaining tab of a template cannot be deleted', False)
    except ValueError:
        check('the last remaining tab of a template cannot be deleted (or the template has more than two tabs)', True)
    # full file round trip through the container on an edited copy
    out = build_container(c, data[:HEADER_LEN])
    back = decompress_bytes(out)
    check('edited file survives write + read through the container', back == bytes(c))
    # the same edited bytes parse with the viewer's reader
    tp, chk = parse_templates(back)
    check('viewer reader parses the edited file with 0 problems', not chk['problems'] and chk['buttons'] == sum(len(tb['blocks']) for t in lay['templates'] for tb in t['tabs']) + 1, 'buttons=%d' % chk['buttons'])
    print('\nALL PASSED' if ok else '\nSOME TESTS FAILED')
    return 0 if ok else 1


# ---------------------------------------------------------------- CLI
def write_main(argv):
    if len(argv) < 3 or argv[1] in ('-h', '--help'):
        print(__doc__)
        return 1
    src = argv[1]
    a = argv[2:]
    if '--selftest' in a:
        return selftest(src)
    data, raw = read_map(src)
    try:
        return _edit(src, a, data, raw)
    except ValueError as ex:
        print('Cannot do that:', ex)
        return 2


def _edit(src, a, data, raw):
    new = None
    if '--move' in a:
        i = a.index('--move'); new = move_button(raw, a[i + 1], a[i + 2], int(a[i + 3]), int(a[i + 4]))
    elif '--copy' in a:
        i = a.index('--copy')
        pos = int(a[i + 5]) if len(a) > i + 5 and not a[i + 5].startswith('--') else None
        new = copy_button(raw, a[i + 1], a[i + 2], int(a[i + 3]), a[i + 4], pos)
    elif '--delete' in a:
        i = a.index('--delete'); new = delete_button(raw, a[i + 1], a[i + 2], int(a[i + 3]))
    else:
        print('nothing to do; give --move, --copy, --delete or --selftest')
        return 1
    new = bytes(new)
    probs = validate(new)
    if probs:
        print('REFUSING to write: result is not consistent:')
        for p in probs[:10]:
            print('  ', p)
        return 2
    if '--compare' in a:
        _, theirs = read_map(a[a.index('--compare') + 1])
        for line in compare(new, theirs):
            print(line)
    if '--out' in a:
        out = a[a.index('--out') + 1]
        if os.path.abspath(out) == os.path.abspath(src):
            print('--out must be a different file than the source')
            return 2
        with open(out, 'wb') as f:
            f.write(build_container(new, data[:HEADER_LEN]))
        print('wrote', out, '(', len(new), 'bytes uncompressed )')
    return 0



# ======================================================================================
#  PALETTE EDITOR  (profile picker + drag and drop + in-place save)
# ======================================================================================
import http.server, socketserver, socket, threading, urllib.parse, secrets, mimetypes, collections, hashlib


def read_profiles(profiles_map_path):
    """profiles.map -> [{'name':..., 'folder':...}] (name followed by './folder/')."""
    with open(profiles_map_path, 'rb') as f:
        raw = decompress_bytes(f.read())
    end = len(raw.rstrip(b'\0'))
    names, o = [], 0
    while o < end:
        s, e = _str_at(raw, o, end)
        if s:
            names.append(s)
            o = e
        else:
            o += 1
    out, i = [], 0
    while i + 1 < len(names):
        if names[i + 1].startswith('./') and names[i + 1].endswith('/'):
            out.append({'name': names[i], 'folder': names[i + 1][2:-1]})
            i += 2
        else:
            i += 1
    return out


def find_profiles_map(env, cfg):
    import os
    if cfg.get('profiles_map_path') and os.path.isfile(cfg['profiles_map_path']):
        return cfg['profiles_map_path']
    cands = []
    pdir = env['ini_paths'].get('PROFILES') or os.path.join(env['db'], 'Profiles')
    for folder in (pdir, env.get('database') or '', env['db']):
        if folder and os.path.isdir(folder):
            p = _find_ci(folder, 'profiles.map')
            if p:
                cands.append(p)
    return cands[0] if cands else ''


def find_profile_service(profiles_dir, folder):
    """service.map inside a profile folder (directly, in a Database subfolder, or up to 3 levels down)."""
    import os
    base = os.path.join(profiles_dir, folder)
    if not os.path.isdir(base):
        return ''
    for sub in ('', 'Database'):
        d = os.path.join(base, sub) if sub else base
        p = _find_ci(d, 'service.map') if os.path.isdir(d) else ''
        if p:
            return p
    for root, dirs, files in os.walk(base):
        if root[len(base):].count(os.sep) >= 3:
            dirs[:] = []
        for n in files:
            if n.lower() == 'service.map':
                return os.path.join(root, n)
    return ''


_IMG_TYPES = {'.png': 'image/png', '.bmp': 'image/bmp', '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg', '.gif': 'image/gif'}


class PictureIndex(object):
    """Remembers between runs where each button's picture is, so nothing has to be searched for twice, plus what was
    learned by opening .ITM files. Paths are stored relative to the ITEMS and Images folders, so moving the database or
    changing a drive letter does not matter. Answers are re-checked cheaply whenever a template is viewed."""

    def __init__(self, path, items, images):
        self.path, self.items, self.images = path, items or '', images or ''
        self.memo = {}                  # (png, itm) -> absolute path, or '' when no picture was found
        self.stamp = {}                 # (png, itm) -> folder timestamps from the last unsuccessful search
        self.itm = {'d': {}, 'dirty': False}
        self.tpl_keys, self.tk_seen = {}, set()     # sha1 of a template's bytes -> its (png, itm) pairs; skips re-reading unchanged templates
        self.saved_len, self.dirty, self.last_save, self.loaded = 0, False, 0.0, 0
        self.lock = threading.RLock()
        self._load()

    def _enc(self, p):
        if not p:
            return ''
        for tag, root in (('I', self.items), ('M', self.images)):
            if root:
                try:
                    rel = os.path.relpath(p, root)
                except ValueError:
                    continue
                if not rel.startswith('..') and not os.path.isabs(rel):
                    return tag + ':' + rel.replace('\\', '/')
        return 'A:' + p

    def _dec(self, s):
        if not s:
            return ''
        tag, rel = s[0], s[2:]
        if tag == 'I' and self.items:
            return os.path.normpath(os.path.join(self.items, rel)).replace('\\', '/')
        if tag == 'M' and self.images:
            return os.path.normpath(os.path.join(self.images, rel)).replace('\\', '/')
        if tag == 'A':
            return rel
        return None

    def _load(self):
        if not self.path:
            return
        d = _read_json(self.path)
        if d.get('v') != 1:
            return
        for k, v in (d.get('memo') or {}).items():
            png, _, itm = k.partition('\t')
            p = self._dec(v)
            if p is not None:
                self.memo[(png, itm)] = p
        for k, v in (d.get('stamps') or {}).items():
            png, _, itm = k.partition('\t')
            self.stamp[(png, itm)] = v
        self.itm['d'] = d.get('itm') or {}
        self.tpl_keys = dict((h, [tuple(k) for k in v]) for h, v in (d.get('tk') or {}).items())
        self.saved_len = self.loaded = len(self.memo)

    def save(self, force=False):
        with self.lock:
            if not self.path or not (force or self.dirty or self.itm.get('dirty') or len(self.memo) != self.saved_len):
                return False
            d = {'v': 1,
                 'memo': dict((k[0] + '\t' + k[1], self._enc(p)) for k, p in list(self.memo.items())),
                 'stamps': dict((k[0] + '\t' + k[1], s) for k, s in list(self.stamp.items()) if self.memo.get(k) == ''),
                 'itm': self.itm['d'],
                 'tk': dict((h, [list(k) for k in v]) for h, v in list(self.tpl_keys.items()))}
            tmp = self.path + '.~t'
            try:
                os.makedirs(os.path.dirname(self.path), exist_ok=True)
                with open(tmp, 'w', encoding='utf-8') as f:
                    json.dump(d, f, separators=(',', ':'))
                os.replace(tmp, self.path)
            except Exception:
                try:
                    os.remove(tmp)
                except Exception:
                    pass
                return False
            self.saved_len, self.dirty, self.last_save = len(self.memo), False, time.time()
            self.itm['dirty'] = False
            return True

    def maybe_save(self):
        if time.time() - self.last_save > 10:
            self.save()

    @staticmethod
    def _mt(p):
        try:
            return os.stat(p).st_mtime_ns
        except Exception:
            return 0

    def stamp_of(self, key, cache=None):
        itm = key[1]
        d = ''
        if itm:
            rel = itm.replace('\\', '/')
            rel = rel[2:] if rel.startswith('./') else rel
            d = os.path.dirname(os.path.join(self.items, rel))
        out = []
        for p in (d, self.images):
            if cache is not None and p in cache:
                out.append(cache[p])
                continue
            v = self._mt(p) if p else 0
            if cache is not None:
                cache[p] = v
            out.append(v)
        return out

    @staticmethod
    def key_of(b):
        return (b['png'], b['items'][0] if b['items'] else '')

    def keys_for(self, raw, tpl):
        """The distinct (png, itm) pairs of one template. Unchanged templates are recognised by a fingerprint of their
        bytes, so they are not parsed again on later runs."""
        h = hashlib.sha1(memoryview(raw)[tpl['t']:tpl['te']]).hexdigest()
        self.tk_seen.add(h)
        v = self.tpl_keys.get(h)
        if v is None:
            v, seen = [], set()
            for tab in tpl['tabs']:
                for s, e in tab['blocks']:
                    k = self.key_of(_button_info(raw, s, e))
                    if k not in seen:
                        seen.add(k)
                        v.append(k)
            self.tpl_keys[h] = v
            self.dirty = True
        return v

    def verify(self, tmp):
        """Forget answers that are no longer true: a picture that vanished, or a missing picture whose folder changed."""
        dropped, cache = 0, {}
        for tb in tmp['tabs']:
            for b in tb['buttons']:
                key = self.key_of(b)
                v = self.memo.get(key)
                if v is None:
                    continue
                if v:
                    if not os.path.isfile(v):
                        self.memo.pop(key, None)
                        dropped += 1
                elif self.stamp.get(key) != self.stamp_of(key, cache):
                    self.memo.pop(key, None)
                    self.stamp.pop(key, None)
                    dropped += 1
        if dropped:
            self.dirty = True
        return dropped

    def note(self, tmp):
        """After a search: remember when (by folder timestamps) each unsuccessful search was made."""
        cache = {}
        for tb in tmp['tabs']:
            for b in tb['buttons']:
                key = self.key_of(b)
                if self.memo.get(key) == '' and key not in self.stamp:
                    self.stamp[key] = self.stamp_of(key, cache)
                    self.dirty = True


class _Stale(ValueError):
    """The page sent an edit based on an out-of-date view of the file."""


class Session(object):
    def __init__(self, env, cfg=None, json_path='', progress=None):
        import os
        _p = progress or (lambda text: None)
        self.busy, self.slow_log = {}, []
        self.health = {'damaged': 0, 'lost': 0, 'lost_names': [], 'checked': False}
        self.env = env
        self.cfg = cfg if cfg is not None else {}
        self.json_path = json_path
        self.lock = threading.RLock()
        self.pdir = env['ini_paths'].get('PROFILES') or os.path.join(env['db'], 'Profiles')
        _p('Looking for the profile list (profiles.map)')
        self.pmap = find_profiles_map(env, self.cfg)
        _p('Reading the profile list (%s)' % (self.pmap or 'not found'))
        self.profiles = read_profiles(self.pmap) if self.pmap else []
        self.cur = None
        self.img_by_id, self.id_by_path = [], {}
        self.index, self.index_ready, self.index_gen, self.index_timer, self.lib_cache = {}, False, 0, None, {}
        self._starts_for, self._starts_v, self._used_for, self._used_v, self._root_ok, self._fseed_done, self._fseed = None, None, None, None, False, False, []
        self.index_building = False
        self.lib_files, self.lib_ready, self.lib_building = [], False, False
        self.lib_scanned_at, self.lib_prog, self._lib_maps_v, self._lib_maps_for = 0, {'done': 0, 'total': 0}, None, None
        self.samples = {'state': 'idle', 'done': 0, 'total': 0, 'path': '', 'message': ''}
        self.lib_path = self.cfg.get('library_index') or (os.path.join(os.path.dirname(json_path), 'fabmap_library_index.json') if json_path else '')
        self.lib_loaded, self.lib_dirs, self._subs_for, self._subs_v = False, {}, None, None
        self.lib_jobs = {}
        self.save_lock = threading.Lock()
        self.rev = 0          # bumps on every open / edit / undo / discard; the page must quote it when editing
        idx_path = self.cfg.get('picture_index') or (os.path.join(os.path.dirname(json_path), 'fabmap_picture_index.json') if json_path else '')
        _p('Loading the picture index')
        self.pidx = PictureIndex(idx_path, env.get('items') or '', env.get('images') or '')
        self.img_memo, self.res_shared = self.pidx.memo, {}
        self.gen, self.pics_rev = 0, 0
        self.scan_info = {'running': False, 'seconds': 0.0, 'found': 0, 'missing': 0, 'total': 0, 'itm_read': 0,
                          'from_index': 0, 'done': 0, 'of': 0}
        self.todo_lock, self.todo, self.active = threading.Lock(), collections.deque(), 0
        self.img_lock, self.recheck_at = threading.Lock(), {}
        self.tpl_rev, self.order_log, self.cur_ti = {}, [], -1
        self.scan_keys, self.scan_start_keys, self.scan_before, self.scan_itm_read = set(), set(), 0, 0
        self.scan_t0, self.scan_final = 0.0, True

    def _auto_repair(self):
        """Put right what earlier editors damaged (a wrong id, a lost picture path), as the first undoable change of this session.
        Returns (buttons repaired, buttons it could not repair). Quick: about 0.1 s even for the biggest profile."""
        c = self.cur
        raw = c['raw']
        fixed, lost = 0, []
        try:
            dam = find_damaged_ids(raw)
            rep_, lost = find_damaged_items(raw)
            if dam or rep_:
                new, olds = repair_damaged_ids(raw, dam) if dam else (bytearray(raw), [])
                ob = [(p, struct.pack('<I', v)) for p, v in olds]
                new, po = repair_damaged_items(new, rep_)
                ob += [(p, bytes(o)) for p, o in po]
                new = bytes(new)
                if not validate(new, deep=False):
                    c['raw'] = new
                    c['hist'].append((('bytes', ob), 0))
                    fixed = len(dam) + len(rep_)
        except Exception:
            pass
        self.health = {'damaged': 0, 'lost': len(lost), 'lost_names': [x[1] for x in lost[:6]], 'checked': True}
        return fixed, len(lost)

    def _check_health_async(self):
        c = self.cur
        if not c:
            return
        raw = c['raw']
        self.health = {'damaged': 0, 'lost': 0, 'lost_names': [], 'checked': False}

        def run():
            try:
                n = len(find_damaged_ids(raw))
                rep, lost = find_damaged_items(raw)
                n += len(rep)
            except Exception:
                n, lost = 0, []
            if self.cur is c and self.cur['raw'] is raw:
                self.health = {'damaged': n, 'lost': len(lost), 'lost_names': [x[1] for x in lost[:6]], 'checked': True}
        th = threading.Thread(target=run, name='fabmap-health')
        th.daemon = True
        th.start()

    # ---- what is waiting on a drive right now (shown by the page when something takes more than a couple of seconds)
    def _busy(self, what):
        """Marks a wait on the drive that you asked for: it is shown while it lasts, background work steps aside for it, and
        any wait of 3 s or more is remembered (the small bar at the bottom of the page lists them)."""
        import contextlib
        @contextlib.contextmanager
        def cm():
            key = object()
            t0 = time.time()
            self.busy[id(key)] = (what, t0)
            _FG['n'] += 1
            try:
                yield
            finally:
                _FG['n'] = max(0, _FG['n'] - 1)
                self.busy.pop(id(key), None)
                dur = time.time() - t0
                if dur >= 3.0:
                    self.slow_log.append({'what': what, 'seconds': int(round(dur)), 'at': time.strftime('%H:%M:%S')})
                    del self.slow_log[:-30]
        return cm()

    def busy_list(self):
        now = time.time()
        return [{'what': w, 'seconds': int(now - t)} for w, t in sorted(list(self.busy.values()), key=lambda x: x[1])]

    # ---- profile
    def open_profile(self, idx):
        import os
        with self.lock:
            if not 0 <= idx < len(self.profiles):
                raise ValueError('no such profile')
            p = self.profiles[idx]
            with self._busy('Looking for the service.map of %s' % p['name']):
                path = find_profile_service(self.pdir, p['folder'])
            if not path:
                raise ValueError('No service.map found under %s' % os.path.join(self.pdir, p['folder']))
            with self._busy('Reading %s' % path):
                data, raw = read_map(path)
            probs = validate(raw, deep=False)
            if probs:
                raise ValueError('This service.map did not pass the structure check (%s). Nothing was opened.' % probs[0])
            self.cur = {'idx': idx, 'name': p['name'], 'path': path, 'header': data[:HEADER_LEN], 'raw': raw,
                        'saved_raw': raw, 'hist': [], 'saved_at': 0}
            self.index_ready = self.index_building = False
            self.rev += 1
            self.gen += 1
            fixed, lost = self._auto_repair()
            raw = self.cur['raw']
            last = (self.cfg.get('last_templates') or {}).get(p['name'])
            first = next((i for i, tp in enumerate(layout(raw)['templates']) if tp['name'] == last), None)
            self.tpl_rev, self.order_log, self.cur_ti = {}, [], -1
            self._start_scan(self.gen, raw, first)
            if self.json_path:
                self.cfg['last_profile'] = p['name']
                _write_json(self.json_path, self.cfg)
            return dict({'profile': p['name'], 'path': path, 'palettes': len(layout(raw)['templates']), 'rev': self.rev, 'repaired': fixed}, **self.state())

    def state(self):
        c = self.cur
        if not c:
            return {'dirty': False, 'n_unsaved': 0, 'can_undo': False, 'n_hist': 0, 'saved_at': 0}
        return {'dirty': c['raw'] != c['saved_raw'], 'n_unsaved': abs(len(c['hist']) - c['saved_at']), 'can_undo': bool(c['hist']), 'n_hist': len(c['hist']), 'saved_at': c['saved_at']}

    def _need(self):
        if not self.cur:
            raise ValueError('Pick a profile first')
        return self.cur

    # ---- data for the page
    def palettes(self):
        with self.lock:
            raw = self._need()['raw']
            return [{'i': i, 'name': t['name'], 'tabs': len(t['tabs']),
                     'buttons': sum(len(x['blocks']) for x in t['tabs'])} for i, t in enumerate(layout(raw)['templates'])]

    def _img_id(self, path):
        if not path:
            return -1
        key = os.path.normcase(os.path.normpath(path))
        with self.img_lock:
            if key not in self.id_by_path:
                self.id_by_path[key] = len(self.img_by_id)
                self.img_by_id.append(path)
            return self.id_by_path[key]

    def palette(self, ti, preview=False):
        """The button list of one template, built from memory only: no disk or network access at all, so it cannot be slowed
        by a slow drive. Pictures already known are included as-is; the rest are flagged 'pend' and are found (and
        checked) in the background right behind this request, then the page refreshes itself."""
        t0 = time.time()
        c = self.cur                         # a plain read of one reference: never waits for an edit or a save in progress
        if not c:
            raise ValueError('Pick a profile first')
        raw = c['raw']
        t_lock = time.time()
        lay = layout(raw)
        if not 0 <= ti < len(lay['templates']):
            raise ValueError('no such palette')
        tpl = lay['templates'][ti]
        t_lay = time.time()
        srcs = [(tab['name'], [_button_info(raw, s, e) for s, e in tab['blocks']], [btn_excluded(raw, s, e) for s, e in tab['blocks']], tab_excluded(raw, tab),
                 [[i['cid'] for i in btn_items(raw, s, e)] for s, e in tab['blocks']])
                for tab in tpl['tabs']]
        t_btn = time.time()
        memo, tabs, unresolved = self.img_memo, [], False
        for name, btns, exs, tex, ics in srcs:
            out = []
            for bi, b in enumerate(btns):
                k = (b['png'], b['items'][0] if b['items'] else '')
                path = memo.get(k)           # trusted here; checked in the background (a vanished file simply fails to load)
                pend = path is None and bool(b['items'])
                unresolved = unresolved or pend
                out.append({'name': b['name'], 'items': b['items'], 'img': self._img_id(path) if path else -1, 'pend': pend, 'ex': bool(exs[bi]), 'ic': ics[bi]})
            tabs.append({'name': name, 'ex': bool(tex), 'buttons': out})
        if not preview:
            self.cur_ti = ti
        self.prioritize(ti)                  # the template you opened goes to the front of the picture queue
        if not preview:
            if not self.scan_info['running']:
                self._recheck_async(ti, raw, lay)    # idle: re-verify this template's pictures in the background
            self._remember_template(tpl['name'])
        t1 = time.time()
        ms = {'total': int((t1 - t0) * 1000), 'lock': int((t_lock - t0) * 1000), 'layout': int((t_lay - t_lock) * 1000),
              'buttons': int((t_btn - t_lay) * 1000)}
        if t1 - t0 > 1.0:
            _log('slow open of %r: %.1f s (waited for the editor %.1f s, file layout %.1f s, button list %.1f s)'
                  % (tpl['name'], t1 - t0, t_lock - t0, t_lay - t_lock, t_btn - t_lay))
        conds = [{'id': c['id'], 'name': c['name'], 'gt': c['gt'], 'lte': c['lte']} for c in template_conditions(raw, tpl)]
        return {'i': ti, 'name': tpl['name'], 'tabs': tabs, 'conds': conds, 'ms': ms}

    def _recheck_async(self, ti, raw, lay):
        now = time.time()
        with self.todo_lock:
            if now - self.recheck_at.get(ti, 0) < 3.0:
                return
            self.recheck_at[ti] = now
        gen = self.gen
        items = self.env.get('items') or ''
        images = self.env.get('images') or ''

        def run():
            try:
                if items and os.path.isdir(items):
                    self._scan_template(gen, raw, lay, ti, items, images if os.path.isdir(images or '') else '', {}, False)
                    self.pidx.maybe_save()
            except Exception as ex:
                _log('pictures: re-check skipped: %s: %s' % (type(ex).__name__, ex))
        th = threading.Thread(target=run, name='fabmap-recheck')
        th.daemon = True
        th.start()

    # ---- pictures: find out where every picture is (remembered between runs), the template you need first
    def _remember_template(self, name):
        c = self.cur
        if not c or not self.json_path:
            return
        lt = self.cfg.setdefault('last_templates', {})
        if lt.get(c['name']) != name:
            lt[c['name']] = name
            snap = json.loads(json.dumps(self.cfg))
            th = threading.Thread(target=_write_json, args=(self.json_path, snap), name='fabmap-settings')
            th.daemon = True
            th.start()

    def _start_scan(self, gen, raw, first=None):
        lay = _layout(raw)                   # uncached on purpose: must not disturb the editor's own layout cache
        order = list(range(len(lay['templates'])))
        if first is not None and first in order:
            order.remove(first)
            order.insert(0, first)
        with self.todo_lock:
            self.todo = collections.deque(order)
            self.active, self.scan_final = 0, False
            self.scan_keys, self.scan_itm_read = set(), 0
            self.scan_start_keys, self.scan_before = set(self.img_memo), len(self.img_memo)
            self.scan_t0 = time.time()
        self.scan_info.update(running=True, done=0, of=len(order))
        shared = {}
        for n in range(2):
            th = threading.Thread(target=self._scan_worker, args=(gen, raw, lay, shared), name='fabmap-pictures-%d' % n)
            th.daemon = True
            th.start()

    def prioritize(self, ti):
        with self.todo_lock:
            if ti in self.todo:
                self.todo.remove(ti)
                self.todo.appendleft(ti)

    def _scan_worker(self, gen, raw, lay, shared):
        items = self.env.get('items') or ''
        images = self.env.get('images') or ''
        images = images if os.path.isdir(images or '') else ''
        usable = bool(items) and os.path.isdir(items)
        while True:
            _yield_to_user()
            with self.todo_lock:
                if gen != self.gen or not self.todo:
                    break
                ti = self.todo.popleft()
                self.active += 1
            try:
                if usable:
                    self._scan_template(gen, raw, lay, ti, items, images, shared)
            except Exception as ex:
                _log('pictures: template %d skipped: %s: %s' % (ti, type(ex).__name__, ex))
            finally:
                with self.todo_lock:
                    self.active -= 1
                    self.scan_info['done'] += 1
            time.sleep(0.002)               # lets edits and page requests through between templates
        self._scan_finish(gen)

    def _scan_template(self, gen, raw, lay, ti, items, images, shared, track=True):
        tpl = lay['templates'][ti]
        keys = self.pidx.keys_for(raw, tpl)
        tmp = {'name': tpl['name'], 'tabs': [{'name': '', 'buttons': [
            {'name': '', 'items': [k[1]] if k[1] else [], 'png': k[0], 'png_missing': not k[0]} for k in keys]}]}
        before = dict((k, self.img_memo.get(k)) for k in keys)
        self.pidx.verify(tmp)
        resolve_images([tmp], items, images, False, '', -1, self.img_memo, shared, self.pidx.itm)       # by name first
        if any(k not in self.img_memo and k[1] for k in keys):                                           # then inside .ITM files
            left = (60.0 - (time.time() - self.scan_t0)) if track else 10.0
            if left > 0:
                st = resolve_images([tmp], items, images, False, '', min(10.0, left), self.img_memo, shared, self.pidx.itm)
                if track:
                    self.scan_itm_read += st.get('itm_files_read', 0)
            if not track:                                   # a re-check always settles: nothing may stay 'pending' forever
                for k in keys:
                    if k not in self.img_memo:
                        self.img_memo[k] = ''
        self.pidx.note(tmp)
        changed = any(before[k] != self.img_memo.get(k) for k in keys)
        if track:
            with self.todo_lock:
                self.scan_keys.update(keys)
                self.order_log.append(ti)
            self.pidx.maybe_save()
        if changed:
            self.tpl_rev[ti] = self.tpl_rev.get(ti, 0) + 1

    def _scan_finish(self, gen):
        with self.todo_lock:
            if gen != self.gen or self.todo or self.active or self.scan_final:
                return
            self.scan_final = True
            keys = set(self.scan_keys)
        found = sum(1 for k in keys if self.img_memo.get(k))
        known = len(keys & self.scan_start_keys)
        changed = len(self.img_memo) != self.scan_before or self.scan_itm_read > 0
        self.pidx.tpl_keys = dict((h, v) for h, v in self.pidx.tpl_keys.items() if h in self.pidx.tk_seen)
        self.pidx.tk_seen = set()
        self.pidx.save()
        secs = round(time.time() - self.scan_t0, 1)
        self.scan_info.update(found=found, missing=len(keys) - found, total=len(keys), itm_read=self.scan_itm_read,
                              from_index=known, seconds=secs)
        if changed:
            self.pics_rev += 1
        _log('pictures: %d distinct, %d found, %d not found | %d already in the saved index, %d .ITM files opened | %.1f s'
              % (len(keys), found, len(keys) - found, known, self.scan_itm_read, secs))
        self.scan_info['running'] = False

    def image_status(self, ti=None):
        d = dict(self.scan_info)
        d['scanning'] = d.pop('running')
        d['pics_rev'] = self.pics_rev          # NOT 'rev': that name is the page's edit version
        d['tpl_rev'] = self.tpl_rev.get(ti, 0) if ti is not None else 0
        d['slow'] = self.slow_log[-8:]
        d['damaged'] = self.health.get('damaged', 0)
        d['lost'] = self.health.get('lost', 0)
        d['lost_names'] = self.health.get('lost_names', [])
        return d

    def rescan(self):
        with self.lock:
            c = self._need()
            self.pidx.memo.clear()
            self.pidx.stamp.clear()
            self.pidx.dirty = True
            self.gen += 1
            self.pics_rev += 1
            self.tpl_rev, self.order_log = {}, []
            self._start_scan(self.gen, c['raw'], self.cur_ti if self.cur_ti >= 0 else None)
        return self.image_status()

    # ---- services (what the stock editor lists), each pointing at a palette by template id
    def services(self):
        with self.lock:
            c = self._need()
            if 'svc' not in c:
                c['svc'] = self._parse_services(c['raw'])
            return c['svc']

    @staticmethod
    def _parse_services(raw):
        lay = layout(raw)

        def first_str(a, b):
            o = a
            while o < b:
                s, e = _str_at(raw, o, b)
                if s:
                    return s, o, e
                o += 1
            return None

        def lp(o, lim):
            """length-prefixed UTF-16 string that may be empty -> (text, next offset) or None"""
            if o + 4 > lim:
                return None
            n = U(raw, o)
            if not (1 <= n < 500) or o + 4 + 2 * n > lim:
                return None
            b = raw[o + 4:o + 4 + 2 * n]
            if b[-2:] != b'\x00\x00':
                return None
            try:
                s = b[:-2].decode('utf-16le')
            except Exception:
                return None
            if any(ord(ch) < 32 for ch in s):
                return None
            return s, o + 4 + 2 * n

        tid = {}
        for ti, t in enumerate(lay['templates']):
            r = first_str(t['t'] + 8, t['te'])
            if r:
                tid[U(raw, r[2] + 4)] = ti          # a template's own id sits right after its name
        out = []
        c0 = raw.find(TAG_COLLECTION)
        if c0 < 0:
            return out
        coll_end = c0 + 8 + U(raw, c0 + 4)
        starts, i = [], c0 + 4
        while True:
            i = raw.find(TAG_SERVICE, i, coll_end)
            if i < 0:
                break
            starts.append(i)
            i += 4
        for k, s in enumerate(starts):
            end = s + 8 + U(raw, s + 4)
            r = first_str(s + 8, end)
            if not r:
                continue
            name, o = r[0], r[2]
            fields = []
            for _ in range(3):                      # two unnamed fields, then the group
                x = lp(o, end)
                if not x:
                    break
                fields.append(x[0])
                o = x[1]
            group = fields[2] if len(fields) == 3 else ''
            ref = U(raw, o) if o + 4 <= end else 0
            ti = tid.get(ref, -1)                   # the service's palette, by template id
            out.append({'i': k, 'name': name, 'group': group, 'ti': ti})
        return out

    # ---- library (items already used by a button somewhere in this file)
    def _lib_entries(self, raw, tpl):
        """(tab, position, name, item, png) of every button that has an item; remembered by a fingerprint of the template."""
        h = hashlib.sha1(memoryview(raw)[tpl['t']:tpl['te']]).hexdigest()
        ent = self.lib_cache.get(h)
        if ent is None:
            ent = []
            for tabi, tab in enumerate(tpl['tabs']):
                for n, (s, e) in enumerate(tab['blocks']):
                    info = _button_info(raw, s, e)
                    if info['items']:
                        ent.append((tabi, n, info['name'], info['items'][0], info['png']))
                time.sleep(0.0005)          # be polite: edits and clicks come first
            self.lib_cache[h] = ent
        return ent

    def _build_index_now(self, raw, gen):
        try:
            lay = _layout(raw)               # uncached on purpose: must not disturb the editor's own layout cache
            idx = {}
            for ti, tpl in enumerate(lay['templates']):
                if gen != self.index_gen:
                    return
                for tabi, n, name, item, png in self._lib_entries(raw, tpl):
                    k = item.lower()
                    ent = idx.get(k)
                    if ent:
                        ent['uses'] += 1
                    else:
                        idx[k] = {'item': item, 'name': name, 'png': png, 'loc': (ti, tabi, n), 'uses': 1}
                time.sleep(0)
            if gen == self.index_gen:
                self.index, self.index_ready, self.index_building = idx, True, False
        except Exception as ex:
            self.index_building = False
            _log('library list: stopped: %s: %s' % (type(ex).__name__, ex))

    def _index_soon(self, delay=0.8, force=False):
        """(Re)build the library list shortly after the last change, in the background. It is only started when the machine
        is idle (picture scan finished) or when someone first searches; once built it is kept fresh after edits."""
        c = self.cur
        if not c or not (force or self.index_ready):
            return
        self.index_building = True
        with self.todo_lock:
            self.index_gen += 1
            gen = self.index_gen
            if self.index_timer:
                self.index_timer.cancel()
            th = threading.Timer(delay, self._build_index_now, args=(c['raw'], gen))
            th.daemon = True
            self.index_timer = th
        th.start()

    def _library_part(self, item):
        """(part file bytes, picture path) of a library part such as './Folder/Part.ITM'; raises ValueError with a plain reason."""
        root = self._lib_root()
        itm_path = os.path.join(root, *item[2:].split('/')) if item.startswith('./') else ''
        if not itm_path or not os.path.isfile(itm_path):
            raise ValueError('That part file was not found in the Libraries folder.')
        with self._busy('Reading %s' % item.rsplit('/', 1)[-1]):
            with open(itm_path, 'rb') as fh:
                data = fh.read()
            png = library_png(itm_path, item)
        return data, png

    def _locate_item(self, raw, lay, item):
        """The first button whose first item is this part: a text search of the file, then one button is read (milliseconds)."""
        import bisect
        k = item.lower()
        if self._starts_for is not lay:
            st = sorted((s, ti, tabi, n, e) for ti, tp in enumerate(lay['templates']) for tabi, tab in enumerate(tp['tabs'])
                        for n, (s, e) in enumerate(tab['blocks']))
            self._starts_for, self._starts_v = lay, ([x[0] for x in st], st)
        keys, st = self._starts_v
        for m in re.finditer(re.escape(item.encode('utf-16le')), raw, re.I):
            i = bisect.bisect_right(keys, m.start()) - 1
            if i < 0:
                continue
            s, ti, tabi, n, e = st[i]
            if m.start() < e:
                info = _button_info(raw, s, e)
                if info['items'] and info['items'][0].lower() == k:
                    return (ti, tabi, n), info['name']
        return None, ''

    # ---- the Libraries folder: browse it, search it.
    # The disk is the slow part (a network-style drive answers one folder at a time slowly), so the part list is read once with
    # many folders at the same time, starting from the folder list in FOLDERS.MAP, shown while it is still being listed, saved
    # to disk, and then served from memory (browsing and searching do not touch the disk at all).
    def _lib_root(self):
        root = self.env.get('items') or ''
        if self._root_ok:
            return root
        if not root or not os.path.isdir(root):
            raise ValueError('The Libraries folder was not found. Check Items= in MAP.INI.')
        self._root_ok = True
        return root

    def _lib_folder_seed(self):
        """Every folder of the Libraries tree from FOLDERS.MAP, read once per run (no disk walking); [] if it cannot be read."""
        if self._fseed_done:
            return self._fseed
        out = []
        try:
            with self._busy('Reading FOLDERS.MAP'):
                data, raw = read_map(os.path.join(self.env.get('db') or '', 'FOLDERS.MAP'))
            for m in re.finditer(b'\\.\\x00/\\x00', raw):
                o = m.start()
                if o < 4:
                    continue
                n = struct.unpack_from('<I', raw, o - 4)[0]
                if 2 <= n <= 400 and o + 2 * n <= len(raw) and raw[o + 2 * (n - 1):o + 2 * n] == b'\x00\x00':
                    s = raw[o:o + 2 * (n - 1)].decode('utf-16le', 'ignore')
                    if s.startswith('./') and s.endswith('/') and s.isprintable():
                        out.append(s[2:-1])
        except Exception:
            out = []
        self._fseed, self._fseed_done = out, True
        return out

    def _lib_load(self):
        """The remembered part list from disk: (files, complete, saved-at) or None."""
        try:
            if not self.lib_path or not os.path.isfile(self.lib_path):
                return None
            with self._busy('Reading the saved part list'):
                with open(self.lib_path, 'rb') as f:
                    d = json.loads(zlib.decompress(f.read()).decode('utf-8'))
            root = self.env.get('items') or ''
            if os.path.normcase(os.path.normpath(d.get('root', ''))) != os.path.normcase(os.path.normpath(root)):
                return None
            return [(r, r.lower(), p) for r, p in d['files']], bool(d.get('complete', True)), float(d.get('at', 0))
        except Exception:
            return None

    def _lib_save(self, files, complete):
        if not self.lib_path:
            return
        try:
            os.makedirs(os.path.dirname(self.lib_path), exist_ok=True)
            tmp = self.lib_path + '.~t'
            body = {'root': self.env.get('items') or '', 'complete': complete, 'at': time.time(), 'files': [[r, p] for r, _, p in list(files)]}
            with open(tmp, 'wb') as f:
                f.write(zlib.compress(json.dumps(body).encode('utf-8'), 6))
            os.replace(tmp, self.lib_path)
        except Exception:
            pass

    def _lib_kick(self, force=False):
        """Start (or refresh) the listing in the background when it is missing, incomplete, old, or asked for."""
        if self.lib_building:
            return
        if force or not self.lib_ready or (time.time() - self.lib_scanned_at > 12 * 3600):
            self._lib_walk_async()

    def _lib_walk_async(self):
        if self.lib_building:
            return
        self.lib_building = True
        self.lib_prog = {'done': 0, 'total': 0}
        th = threading.Thread(target=self._lib_walk, name='fabmap-library')
        th.daemon = True
        th.start()

    def _lib_walk(self):
        import queue
        try:
            root = self._lib_root()
            q = queue.Queue()
            lock = threading.Lock()
            seen, files, prog = set(), [], self.lib_prog
            if not self.lib_ready:
                self.lib_files = files                 # first listing: searchable while it is being built
            def put(rel):
                k = os.path.normcase(rel)               # the operating system's own idea of 'the same folder'
                with lock:
                    if k in seen:
                        return
                    seen.add(k)
                    prog['total'] += 1
                q.put(rel)
            def scan(rel):
                _yield_to_user()
                path = os.path.join(root, *rel.split('/')) if rel else root
                try:
                    with os.scandir(path) as it:
                        ents = list(it)
                except OSError:
                    return
                low, subs, itms = {}, [], []
                for e in ents:
                    nm = e.name
                    low[nm.lower()] = nm
                    try:
                        isd = e.is_dir()
                    except OSError:
                        isd = False
                    if isd:
                        subs.append(nm)
                    elif nm.lower().endswith('.itm'):
                        itms.append(nm)
                rb = (rel + '/') if rel else ''
                chunk = []
                for f in itms:
                    pn = low.get(f[:-4].lower() + '.png')
                    r = './' + rb + f
                    chunk.append((r, r.lower(), ('./' + rb + pn) if pn else ''))
                if chunk:
                    with lock:
                        files.extend(chunk)
                for s in subs:
                    put(rb + s)
            def work():
                while True:
                    rel = q.get()
                    if rel is None:
                        q.task_done()
                        return
                    try:
                        scan(rel)
                    except Exception:
                        pass
                    finally:
                        with lock:
                            prog['done'] += 1
                        q.task_done()
            put('')
            for rel in self._lib_folder_seed():
                put(rel)
            threads = [threading.Thread(target=work, name='fabmap-lib-%d' % n) for n in range(_LIB_THREADS)]
            for th in threads:
                th.daemon = True
                th.start()
            last = time.time()
            while q.unfinished_tasks:
                time.sleep(0.25)
                if time.time() - last > 30:             # keep what was found so far in case the editor is closed early
                    self._lib_save(files, False)
                    last = time.time()
            for _ in threads:
                q.put(None)
            self.lib_files, self._lib_maps_for = files, None
            self._lib_maps()                              # folder maps are built here, not while somebody is waiting
            self.lib_ready, self.lib_scanned_at = True, time.time()
            self._lib_save(files, True)
        except Exception as ex:
            _log('libraries: listing stopped: %s: %s' % (type(ex).__name__, ex))
        finally:
            self.lib_building = False

    def _lib_listing(self):
        """Progress while the part list is not complete yet, else None."""
        if self.lib_ready or not self.lib_building:
            return None
        return {'done': self.lib_prog['done'], 'total': self.lib_prog['total']}

    def _lib_maps(self):
        """folder -> parts and folder -> sub-folders from the finished list (instant browsing)."""
        if getattr(self, '_lib_maps_for', None) is self.lib_files and self._lib_maps_v:
            return self._lib_maps_v
        by, subs = {}, {}
        def add_folder(rel):
            parts = rel.split('/')
            for i in range(len(parts)):
                parent = '/'.join(parts[:i]).lower()
                subs.setdefault(parent, {})[parts[i].lower()] = parts[i]
        for rel in self._lib_folder_seed():
            if rel:
                add_folder(rel)
        for rel, low, png in self.lib_files:
            parent = rel[2:].rsplit('/', 1)[0] if '/' in rel[2:] else ''
            by.setdefault(parent.lower(), []).append((rel, png))
            if parent:
                add_folder(parent)
        self._lib_maps_for, self._lib_maps_v = self.lib_files, (by, subs)
        return self._lib_maps_v

    def _lib_ensure(self):
        """Read the remembered part list, once, in the background the first time the pane is used (never at start-up, never
        while you wait: until it is ready the pane works from FOLDERS.MAP and one folder at a time)."""
        if not self.lib_loaded:
            self.lib_loaded = True
            th = threading.Thread(target=self._lib_load_bg, name='fabmap-lib-load')
            th.daemon = True
            th.start()

    def _lib_load_bg(self):
        got = self._lib_load()
        if got and not self.lib_building and not self.lib_files:
            files, complete, at = got
            self.lib_files = files
            self._lib_maps_for = None
            if complete:
                self._lib_maps()                          # build the folder maps first, then say the list is ready
            self.lib_ready, self.lib_scanned_at = complete, at

    def _seed_subs(self):
        """folder -> its sub-folders, from FOLDERS.MAP (memory only)."""
        seed = self._lib_folder_seed()
        if self._subs_v is not None and self._subs_for is seed:
            return self._subs_v
        subs = {}
        for rel in seed:
            if rel:
                parts = rel.split('/')
                for i in range(len(parts)):
                    subs.setdefault('/'.join(parts[:i]).lower(), {})[parts[i].lower()] = parts[i]
        self._subs_for, self._subs_v = seed, subs
        return subs

    # ---- 'Collect samples': one zip with the .ITM files of ~100 different kinds of parts used in the open file, and everything the
    # file stores for each of them, so the part-specific values a new button needs can be worked out. Read-only: nothing is changed.
    def samples_status(self):
        return dict(self.samples)

    def samples_start(self):
        c = self.cur
        if not c:
            raise ValueError('Open a profile first.')
        if self.samples.get('state') == 'running':
            return self.samples_status()
        self.samples = {'state': 'running', 'done': 0, 'total': 0, 'path': '', 'message': ''}
        th = threading.Thread(target=self._samples_run, args=(c['raw'],), name='fabmap-samples')
        th.daemon = True
        th.start()
        return self.samples_status()

    def _samples_run(self, raw):
        import zipfile
        try:
            lay = layout(raw)
            root = self._lib_root()
            facts, seen = [], set()
            for ti, tp in enumerate(lay['templates']):
                for tab in tp['tabs']:
                    for s, e in tab['blocks']:
                        blk = bytes(raw[s:e])
                        its = btn_items(raw, s, e)
                        if len(its) != 1 or its[0]['path'].lower() in seen:
                            continue
                        f = _part_facts(blk)
                        if f:
                            seen.add(its[0]['path'].lower())
                            f['path'], f['template'], f['tab'], f['name'] = its[0]['path'], tp['name'], tab['name'], _button_info(raw, s, e)['name']
                            f['blk'] = blk
                            facts.append(f)
            by_key = {}
            for f in facts:
                by_key.setdefault((f['a'], f['b'], f['rowbyte'], min(f['ndim'], 3)), []).append(f)
            pick = [v[0] for v in by_key.values()]
            folders = {}
            for f in facts:
                folders.setdefault(f['path'].rsplit('/', 1)[0], []).append(f)
            extra = [f for fl in folders.values() for f in fl[:1] if f not in pick]
            pick = (pick + extra)[:130]
            self.samples['total'] = len(pick)
            out = os.path.join(os.path.dirname(self.json_path) if self.json_path else os.getcwd(), 'fabmap_samples.zip')
            tmp = out + '.~t'
            used = 0
            rows = ['path,template,tab,a,b,cid,rowbyte,ndim,nrows,block8,timestamp16,itm_bytes,itm_mtime_utc,itm_mtime_local,labels']
            with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as z:
                for n, f in enumerate(pick):
                    src = os.path.join(root, *f['path'][2:].split('/'))
                    size, mt, mtl = '', '', ''
                    try:
                        st = os.stat(src)
                        size, mt = st.st_size, time.strftime('%Y-%m-%d %H:%M:%S', time.gmtime(st.st_mtime))
                        mtl = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(st.st_mtime))
                        if used + st.st_size < 24000000:
                            z.write(src, 'parts/%03d_%s' % (n, os.path.basename(src)))
                            used += st.st_size
                    except OSError:
                        pass
                    z.writestr('buttons/%03d.bin' % n, f['blk'])
                    rows.append(','.join('"%s"' % str(x).replace('"', "'") for x in
                                         (f['path'], f['template'], f['tab'], f['a'], f['b'], f['cid'], f['rowbyte'], f['ndim'], len(f['labels']), f['block8'],
                                          f['ts'], size, mt, mtl, ' | '.join(f['labels'])[:300])))
                    self.samples['done'] = n + 1
                z.writestr('samples.csv', '\n'.join(rows))
            os.replace(tmp, out)
            self.samples.update(state='done', path=out, message='%d parts' % len(pick))
        except Exception as ex:
            self.samples.update(state='error', message='%s: %s' % (type(ex).__name__, ex))

    def lib_refresh(self):
        """'List all parts': read the whole Libraries folder once, in the background (only when asked for)."""
        self._lib_ensure()
        self._fseed_done = False                          # read FOLDERS.MAP again too
        self._lib_kick(True)
        return {'refreshing': True}

    def lib_used(self, rels):
        """Which of these parts does the open file already use? One text search of the file (about 0.05 s), remembered until it changes."""
        c = self.cur
        if not c:
            return {'ready': True, 'used': dict((r, False) for r in rels)}
        raw = c['raw']
        if self._used_for is not raw:
            self._used_v = set(m.group().decode('utf-16le', 'ignore').lower() for m in _ITM16.finditer(raw))
            self._used_for = raw
        return {'ready': True, 'used': dict((r, r.lower() in self._used_v) for r in rels)}

    def lib_folders(self, rel):
        """The sub-folders of a folder: from FOLDERS.MAP, no disk access at all."""
        self._lib_ensure()
        parts = [p for p in rel.replace('\\', '/').split('/') if p not in ('', '.')]
        if '..' in parts:
            raise ValueError('not a Libraries path')
        key = '/'.join(parts).lower()
        names = dict(self._seed_subs().get(key, {}))
        if self.lib_ready:
            for k, v in self._lib_maps()[1].get(key, {}).items():
                names.setdefault(k, v)
        return {'prefix': './' + '/'.join(parts) + ('/' if parts else ''),
                'folders': sorted(names.values(), key=lambda s: s.lower()), 'complete': self.lib_ready,
                'listing': self._lib_listing(), 'refreshing': self.lib_ready and self.lib_building}

    def lib_children(self, rel):
        """The parts in ONE folder (one disk call, remembered), plus any folder the disk has that FOLDERS.MAP lacks."""
        self._lib_ensure()
        root = self._lib_root()
        parts = [p for p in rel.replace('\\', '/').split('/') if p not in ('', '.')]
        if '..' in parts:
            raise ValueError('not a Libraries path')
        key = '/'.join(parts).lower()
        prefix = './' + '/'.join(parts) + ('/' if parts else '')
        if self.lib_ready:                                      # finished list in memory: no disk access
            rows = sorted(self._lib_maps()[0].get(key, []), key=lambda r: r[0].lower())
            extra = []
        else:
            hit = self.lib_dirs.get(key)
            if hit is None:
                # not read yet: start reading it in the background and answer at once ('pending'), so a slow drive never holds a
                # browser connection open; the page asks again every half second
                started = self.lib_jobs.get(key)
                if started is None:
                    started = self.lib_jobs[key] = time.time()
                    th = threading.Thread(target=self._lib_dir_job, args=(key, root, parts, prefix), name='fabmap-folder')
                    th.daemon = True
                    th.start()
                return {'prefix': prefix, 'pending': True, 'seconds': int(time.time() - started), 'items': [], 'more': False, 'extra_folders': []}
            rows, extra = hit
        items = []
        for r, png in rows[:3000]:
            items.append({'name': r.rsplit('/', 1)[-1][:-4], 'rel': r, 'used': None,
                          'img': self._img_id(os.path.join(root, png[2:])) if png else -1})
        return {'prefix': prefix, 'items': items, 'more': len(rows) > 3000, 'extra_folders': extra}

    def _lib_dir_job(self, key, root, parts, prefix):
        try:
            base = os.path.join(root, *parts)
            try:
                with self._busy('Listing the folder %s' % base):
                    entries = list(os.scandir(base))
            except OSError:
                entries = []
            low = dict((e.name.lower(), e.name) for e in entries)
            rows = []
            for fn in sorted((e.name for e in entries if e.is_file() and e.name.lower().endswith('.itm')), key=lambda s: s.lower()):
                pn = low.get(fn[:-4].lower() + '.png')
                rows.append((prefix + fn, ('./' + '/'.join(parts + [pn])) if pn else ''))
            extra = sorted((e.name for e in entries if e.is_dir()), key=lambda s: s.lower())
            self.lib_dirs[key] = (rows, extra)
        finally:
            self.lib_jobs.pop(key, None)

    def lib_search(self, q):
        """Folder names (instant, FOLDERS.MAP) and part names (the complete list if it has been read, else the folders opened so far)."""
        self._lib_ensure()
        toks = [x for x in q.lower().split() if x]
        listing = self._lib_listing()
        base = {'listing': listing, 'complete': self.lib_ready, 'refreshing': self.lib_ready and self.lib_building}
        if not toks:
            return dict(base, results=[], folders=[], total=0, indexing=False)
        folders = []
        for rel in self._lib_folder_seed():
            lo = rel.lower()
            if rel and all(x in lo for x in toks):
                folders.append(rel)
        folders.sort(key=lambda r: (0 if all(x in r.rsplit('/', 1)[-1].lower() for x in toks) else 1, r.count('/'), r.lower()))
        fout = [{'name': r.rsplit('/', 1)[-1], 'rel': './' + r, 'where': r} for r in folders[:30]]
        if self.lib_files:
            files = self.lib_files[:]
        else:
            files = [(r, r.lower(), p) for rows, _ in list(self.lib_dirs.values()) for r, p in rows]
        root = self._lib_root()
        t0, rest = toks[0], toks[1:]
        hits = [f for f in files if t0 in f[1] and all(x in f[1] for x in rest)]
        def rank(f):
            nm = f[1].rsplit('/', 1)[-1]
            return (0 if all(x in nm for x in toks) else 1, len(f[0]), f[1])
        hits.sort(key=rank)
        out = []
        for rel, low, png in hits[:100]:
            out.append({'name': rel.rsplit('/', 1)[-1][:-4], 'rel': rel, 'folder': rel.rsplit('/', 1)[0], 'used': None,
                        'img': self._img_id(os.path.join(root, png[2:])) if png else -1})
        return dict(base, results=out, folders=fout, total=len(hits), indexing=False)

    def search(self, q):
        """Never waits: answers from the latest library list that exists, and says when it is still being built."""
        self._need()
        if not self.index_ready and not self.index_building:
            self._index_soon(0.0, True)
        toks = [x for x in q.lower().split() if x]
        idx = self.index
        if not toks:
            return {'results': [], 'indexing': not self.index_ready}
        hits = [v for v in list(idx.values()) if all(x in (v['name'] + ' ' + v['item']).lower() for x in toks)]
        hits.sort(key=lambda v: (-v['uses'], v['name'].lower()))
        out = []
        for v in hits[:60]:
            path = self.img_memo.get((v['png'], v['item']))
            out.append({'name': v['name'], 'item': v['item'], 'uses': v['uses'], 'img': self._img_id(path) if path else -1})
        return {'results': out, 'indexing': not self.index_ready}

    # ---- edits: kept in memory until Apply. Each edit remembers the tiny step that reverses it.
    def _commit(self, new, inverse, ti):
        c = self.cur
        new = bytes(new)
        probs = validate(new, deep=False)
        if probs:
            raise ValueError('The result failed its own structure check, so nothing was changed: %s' % probs[0])
        c['hist'].append((inverse, ti))
        c['raw'] = new
        self.rev += 1

    def apply(self, body):
        with self.lock:
            t0 = time.time()
            c = self._need()
            if body.get('rev') is not None and int(body['rev']) != self.rev:
                raise _Stale('The page was out of date, so nothing was changed. It has been refreshed - please repeat the change.')
            raw, op = c['raw'], body.get('op')
            ti, tb = int(body['ti']), int(body['tab'])
            lay = layout(raw)
            if op == 'move':
                frm, to = int(body['from']), int(body['to'])
                new = move_idx(raw, ti, tb, frm, to)
                inv = ('move', ti, tb, to, frm)
                msg = 'Moved button %d to position %d' % (frm + 1, to + 1)
            elif op == 'delete':
                idx = int(body['idx'])
                blk = block_bytes(raw, ti, tb, idx)
                new = delete_idx(raw, ti, tb, idx)
                inv = ('insert', ti, tb, idx, blk)
                msg = 'Deleted button %d' % (idx + 1)
            elif op in ('copy', 'add'):
                newpart = None
                if op == 'copy':
                    dtab, pos = int(body['dtab']), body.get('pos')
                    src = (ti, tb, int(body['idx']))
                    msg = 'Copied button %d to another tab' % (int(body['idx']) + 1)
                else:
                    item = str(body['item'])
                    ent = None
                    src, nm = None, ''
                    if ent:                                      # the list may be slightly behind: check it still points at that item
                        li, ltab, ln = ent['loc']
                        try:
                            s0, e0 = lay['templates'][li]['tabs'][ltab]['blocks'][ln]
                            info = _button_info(raw, s0, e0)
                            if info['items'] and info['items'][0].lower() == item.lower():
                                src, nm = ent['loc'], info['name']
                        except (IndexError, KeyError):
                            pass
                    if src is None:
                        src, nm = self._locate_item(raw, lay, item)
                    dtab, pos = tb, body.get('pos')
                    if any(it['items'] and it['items'][0].lower() == item.lower() for it in (_button_info(raw, s_, e_) for s_, e_ in lay['templates'][ti]['tabs'][dtab]['blocks'])):
                        raise ValueError('That part is already in this tab.')
                    if src is None:                              # no button of this profile uses the part yet: build one from its part file
                        root = self._lib_root()
                        itm_path = os.path.join(root, *item[2:].split('/')) if item.startswith('./') else ''
                        if not itm_path or not os.path.isfile(itm_path):
                            raise ValueError('That part file was not found in the Libraries folder.')
                        with self._busy('Reading %s' % item.rsplit('/', 1)[-1]):
                            with open(itm_path, 'rb') as fh:
                                data = fh.read()
                            png = library_png(itm_path, item)
                        try:
                            nb, info = add_new_part(raw, ti, dtab, None if pos is None else int(pos), item, data, png)
                        except (zlib.error, struct.error, IndexError) as ex:
                            raise ValueError('That part file could not be read (%s), so nothing was added.' % ex)
                        newpart = (nb, 'Added %s (new to this profile: %d size%s read from its part file; Fabrication fills in the rest the first time the button is used)' % (info['name'], info['sizes'], '' if info['sizes'] == 1 else 's'))
                    else:
                        msg = 'Added %s' % nm
                old_counter = template_counter(raw, ti)
                n_dst = len(lay['templates'][ti]['tabs'][dtab]['blocks']) if 0 <= ti < len(lay['templates']) and 0 <= dtab < len(lay['templates'][ti]['tabs']) else 0
                fpos = n_dst if pos is None else int(pos)
                if newpart is not None:
                    new, msg = newpart
                else:
                    new = copy_idx(raw, src[0], src[1], src[2], ti, dtab, None if pos is None else int(pos))
                inv = ('countered', ('delete', ti, dtab, fpos), ti, old_counter)
            elif op in ('delete_many', 'move_many', 'copy_many'):
                idxs = sorted(set(int(i) for i in body['idxs']))
                if op == 'copy_many':
                    dtab = int(body['dtab'])
                    before = tab_blocks_bytes(raw, ti, dtab)
                    pos = body.get('pos')
                    old_counter = template_counter(raw, ti)
                    new = copy_many(raw, ti, tb, idxs, dtab, None if pos is None else int(pos))
                    inv, msg, tb = ('countered', ('setblocks', ti, dtab, before), ti, old_counter), 'Copied %d buttons' % len(idxs), dtab
                else:
                    before = tab_blocks_bytes(raw, ti, tb)
                    if op == 'delete_many':
                        new, msg = delete_many(raw, ti, tb, idxs), 'Deleted %d buttons' % len(idxs)
                    else:
                        new, msg = move_many(raw, ti, tb, idxs, int(body['to'])), 'Moved %d buttons' % len(idxs)
                    inv = ('setblocks', ti, tb, before)
            elif op == 'btn_exclude':
                idxs = sorted(set(int(i) for i in body['idxs']))
                val = bool(body.get('value', True))
                new, olds = set_button_flags(raw, ti, tb, idxs, val)
                inv = ('btnflags', ti, tb, olds)
                msg = ('Excluded %d button%s from fill' if val else 'Included %d button%s in fill') % (len(idxs), '' if len(idxs) == 1 else 's')
            elif op == 'tab_exclude':
                cur = tab_excluded(raw, lay['templates'][ti]['tabs'][tb])
                val = bool(body['value']) if body.get('value') is not None else (not cur)
                new, old = set_tab_flag(raw, ti, tb, val)
                inv = ('tabflag', ti, tb, old)
                msg = ('Excluded tab from fill: %s' if val else 'Included tab in fill: %s') % lay['templates'][ti]['tabs'][tb]['name']
            elif op == 'tab_fix':
                tabd = lay['templates'][ti]['tabs'][tb]
                oldname, oldflag = tabd['name'], int(tab_excluded(raw, tabd))
                if not tab_is_excluded(oldname):
                    raise ValueError('That tab name has no [E] to repair.')
                keep_excl = bool(body.get('exclude', True))
                mid = rename_tab(raw, ti, tb, _E_TAIL.sub('', oldname))
                new, _ = set_tab_flag(bytes(mid), ti, tb, keep_excl)
                inv = ('tabfix', ti, tb, oldname, oldflag)
                msg = ('Repaired tab: now really excluded, name %s' if keep_excl else 'Removed [E] from the tab name, tab stays included: %s') % _E_TAIL.sub('', oldname)
            elif op == 'repair_ids':
                dam = find_damaged_ids(raw)
                rep_, lost = find_damaged_items(raw)
                if not dam and not rep_:
                    raise ValueError('Nothing to repair.' + (' %d button%s lost %s picture and cannot be repaired automatically.' % (len(lost), '' if len(lost) == 1 else 's', 'its' if len(lost) == 1 else 'their') if lost else ''))
                new, olds = repair_damaged_ids(raw, dam) if dam else (bytearray(raw), [])
                olds_b = [(p, struct.pack('<I', v)) for p, v in olds]
                new, pic_olds = repair_damaged_items(new, rep_)
                olds_b += [(p, bytes(o)) for p, o in pic_olds]
                inv = ('bytes', olds_b)
                msg = 'Repaired %d button%s' % (len(dam) + len(rep_), '' if len(dam) + len(rep_) == 1 else 's') + (' (%d could not be repaired automatically)' % len(lost) if lost else '')
            elif op == 'copy_ext':
                sti, stb = int(body['src_ti']), int(body['src_tab'])
                idxs = [int(i) for i in body['idxs']]
                pos = None if body.get('pos') is None else int(body['pos'])
                new, info = copy_across(raw, sti, stb, idxs, ti, tb, pos)
                inv = ('whole', bytes(raw))
                msg = 'Copied %d button%s from another template' % (info['copied'], '' if info['copied'] == 1 else 's')
                if info['skipped']:
                    msg += ' (%d already there)' % info['skipped']
                if info['added']:
                    msg += ' - added restriction%s: %s' % ('' if len(info['added']) == 1 else 's', ', '.join(info['added']))
            elif op == 'stack_add':
                idx, item = int(body['idx']), str(body['item'])
                before = tab_blocks_bytes(raw, ti, tb)
                if not 0 <= idx < len(before):
                    raise ValueError('button position must be 0..%d' % (len(before) - 1))
                destb = before[idx]
                if any(it['path'].lower() == item.lower() for it in btn_items(destb, 0, len(destb))):
                    raise ValueError('That part is already on this button.')
                dfmt = button_format(destb) or template_format(raw, lay['templates'][ti])
                src, nm = self._locate_item(raw, lay, item)
                rec = existing_item_record(raw, src, item, dfmt) if src is not None else None
                if rec is None:                                  # no button of this profile has it (in this record version): build it from its part file
                    data, png = self._library_part(item)
                    try:
                        rec, info = new_part_item(raw, ti, dfmt, item, data, png)
                    except (zlib.error, struct.error, IndexError) as ex:
                        raise ValueError('That part file could not be read (%s), so nothing was added.' % ex)
                new = add_item_to_button(raw, ti, tb, idx, rec)
                inv = ('setblocks', ti, tb, before)
                msg = 'Stacked %s on button %d' % (item.rsplit('/', 1)[-1].rsplit('.', 1)[0], idx + 1)
            elif op == 'stack_remove':
                idx, k = int(body['idx']), int(body['k'])
                before = tab_blocks_bytes(raw, ti, tb)
                new = remove_item_from_button(raw, ti, tb, idx, k)
                inv = ('setblocks', ti, tb, before)
                msg = 'Took part %d off button %d' % (k + 1, idx + 1)
            elif op == 'cond_add':
                new, nid = add_condition(raw, ti, str(body.get('name', '')), None if body.get('gt') is None else float(body['gt']),
                                         None if body.get('lte') is None else float(body['lte']))
                inv = ('conddel', ti, nid)
                msg = 'Restriction added'
            elif op == 'cond_delete':
                cid = int(body['cid'])
                new, rec, start, patches = delete_condition(raw, ti, cid)
                inv = ('condrestore', ti, start, rec, patches)
                msg = 'Restriction deleted' + (' (%d item%s back to Unrestricted)' % (len(patches), '' if len(patches) == 1 else 's') if patches else '')
            elif op == 'cond_rename':
                cid = int(body['cid'])
                new, old = rename_condition(raw, ti, cid, str(body.get('name', '')))
                inv = ('condname', ti, cid, old)
                msg = 'Restriction renamed'
            elif op == 'item_cond_many':
                idxs = [int(i) for i in body['idxs']]
                new, olds, skipped = set_item_conditions_many(raw, ti, tb, idxs, int(body['value']), None if body.get('k') is None else int(body['k']))
                inv = ('dwords', olds)
                msg = 'Size restriction set on %d button%s' % (len(idxs) - skipped, '' if len(idxs) - skipped == 1 else 's')
                if skipped:
                    msg += (' (%d with several items left alone: choose which item to set)' if body.get('k') is None else ' (%d had no such item)') % skipped
            elif op == 'item_cond':
                idx, k, cid = int(body['idx']), int(body['item']), int(body['value'])
                new, old = set_item_condition(raw, ti, tb, idx, k, cid)
                inv = ('itemcond', ti, tb, idx, k, old)
                msg = 'Size restriction changed'
            elif op == 'cond_limits':
                cid = int(body['cid'])
                gt = None if body.get('gt') is None else float(body['gt'])
                lte = None if body.get('lte') is None else float(body['lte'])
                new, olds = set_condition_limits(raw, ti, cid, gt, lte)
                inv = ('condlim', ti, cid, olds)
                msg = 'Condition changed'
            elif op == 'tab_delete':
                tdat = tab_bytes(raw, ti, tb)
                tname = lay['templates'][ti]['tabs'][tb]['name']
                new = delete_tab(raw, ti, tb)
                inv = ('tabins', ti, tb, tdat)
                msg = 'Deleted tab %s' % tname
            else:
                raise ValueError('unknown operation')
            self._commit(new, inv, ti)
            out = {'msg': msg, 'rev': self.rev, 'ms': int((time.time() - t0) * 1000)}
            if body.get('echo', 1):                    # the page normally shows the change itself and does not need the list back
                out['palette'] = self.palette(ti)
            return dict(out, **self.state())

    def _inverse(self, raw, inv):
        kind = inv[0]
        if kind == 'move':
            return move_idx(raw, inv[1], inv[2], inv[3], inv[4])
        if kind == 'insert':
            return insert_block(raw, inv[1], inv[2], inv[3], inv[4])
        if kind == 'setblocks':
            return set_tab_blocks(raw, inv[1], inv[2], inv[3])
        if kind == 'rename':
            return rename_tab(raw, inv[1], inv[2], inv[3])
        if kind == 'whole':
            return bytearray(inv[1])
        if kind == 'conddel':
            return delete_condition(raw, inv[1], inv[2])[0]
        if kind == 'condrestore':
            return restore_condition(raw, inv[1], inv[2], inv[3], inv[4])
        if kind == 'condname':
            return rename_condition(raw, inv[1], inv[2], inv[3])[0]
        if kind == 'bytes':
            nb = bytearray(raw)
            for pos, b in reversed(inv[1]):
                nb[pos:pos + len(b)] = b
            return nb
        if kind in ('idfix', 'dwords'):
            nb = bytearray(raw)
            for pos, val in inv[1]:
                struct.pack_into('<I', nb, pos, val)
            return nb
        if kind == 'countered':
            return set_template_counter(self._inverse(raw, inv[1]), inv[2], inv[3])
        if kind == 'itemcond':
            return set_item_condition(raw, inv[1], inv[2], inv[3], inv[4], inv[5])[0]
        if kind == 'condlim':
            g, l = inv[3]
            return set_condition_limits(raw, inv[1], inv[2], None if g == -1.0 else g, None if l == -1.0 else l, False)[0]
        if kind == 'tabins':
            return insert_tab(raw, inv[1], inv[2], inv[3])
        if kind == 'btnflags':
            return restore_button_flags(raw, inv[1], inv[2], inv[3])
        if kind == 'tabflag':
            return set_tab_flag(raw, inv[1], inv[2], inv[3])[0]
        if kind == 'tabfix':
            mid = rename_tab(raw, inv[1], inv[2], inv[3])
            return set_tab_flag(bytes(mid), inv[1], inv[2], inv[4])[0]
        return delete_idx(raw, inv[1], inv[2], inv[3])

    def undo(self):
        with self.lock:
            c = self._need()
            if not c['hist']:
                raise ValueError('Nothing to undo')
            inv, ti = c['hist'][-1]
            new = bytes(self._inverse(c['raw'], inv))
            probs = validate(new, deep=False)
            if probs:
                raise ValueError('Undo failed its own structure check, so nothing was changed: %s' % probs[0])
            c['hist'].pop()
            c['raw'] = new
            self.rev += 1
            nt = len(layout(new)['templates'])
            return dict({'msg': 'Undone', 'palette': self.palette(min(ti, nt - 1)), 'rev': self.rev}, **self.state())

    def save(self):
        """Apply: write what the working copy holds right now into the profile's service.map. The write happens outside the
        editor's lock, so edits, undo and page requests carry on while the file is being written."""
        with self.save_lock:                              # one write at a time
            with self.lock:
                c = self._need()
                raw, path, header = c['raw'], c['path'], c['header']
                hist_len = len(c['hist'])
                n = abs(hist_len - c['saved_at'])
            t0 = time.time()
            write_inplace(path, raw, header)
            with self.lock:
                if self.cur is c:                         # not switched to another profile meanwhile
                    c['saved_raw'] = raw
                    c['saved_at'] = hist_len
            return dict({'msg': 'Saved %d change%s.' % (n, '' if n == 1 else 's'),
                         'seconds': round(time.time() - t0, 2), 'time': time.strftime('%H:%M:%S'), 'rev': self.rev}, **self.state())

    def discard(self, ti=-1):
        with self.lock:
            c = self._need()
            c['raw'] = c['saved_raw']
            if len(c['hist']) >= c['saved_at']:
                del c['hist'][c['saved_at']:]      # keep the undo steps that are already in the file
            else:
                c['hist'] = []                      # undone past the saved point: those steps are gone
                c['saved_at'] = 0
            self.rev += 1
            pal = None
            try:
                if 0 <= ti < len(layout(c['raw'])['templates']):
                    pal = self.palette(ti)
            except Exception:
                pal = None
            return dict({'msg': 'Discarded unsaved changes', 'palette': pal, 'palettes': self.palettes(), 'rev': self.rev}, **self.state())


EDITOR_HTML = r"""<!DOCTYPE html><html lang="en"><head><meta charset="utf-8"><title>Fabrication palette editor</title>
<style>
:root{--y:#FFE716}*{box-sizing:border-box}
body{margin:0;font-family:Arial,sans-serif;color:#000;background:#fff;font-size:13px}
header{background:var(--y);border-bottom:2px solid #000;padding:8px 12px;display:flex;gap:12px;align-items:center;flex-wrap:wrap;min-height:46px}
header button{font:inherit;border:2px solid #000;background:#fff;padding:3px 10px;cursor:pointer}header button:disabled{opacity:.4;cursor:default}
#svc{font-size:11px;overflow-wrap:anywhere}#stat{font-weight:bold}#stat.bad{background:#000;color:var(--y);padding:2px 6px}
#pend{background:#000;color:var(--y);padding:2px 8px;font-weight:bold;display:none}#apply:not(:disabled){background:#000;color:var(--y)}#imgs{font-size:11px}
#wrap{display:flex;height:calc(100vh - 46px)}
#left{width:400px;min-width:240px;display:flex;flex-direction:column;flex:0 0 auto}
#split{width:6px;flex:0 0 6px;background:#000;cursor:col-resize}
.sec{display:flex;flex-direction:column;min-height:0}#s1{flex:0 0 36%;border-bottom:2px solid #000}#s2{flex:1}
.hd{background:var(--y);border-bottom:1px solid #000;padding:4px 10px;font-weight:bold;display:flex;justify-content:space-between}
.hd small{font-weight:normal}.sec input{margin:6px 8px;padding:5px;font:inherit;border:2px solid #000}
.lst{overflow-y:auto;overflow-x:hidden;flex:1}
.row{padding:5px 10px;border-bottom:1px solid #000;cursor:pointer;overflow-wrap:anywhere;word-break:break-word}.row small{display:block;font-size:11px}
.row.svc{padding-left:22px}.row.sel{background:var(--y);font-weight:bold}.row:hover{background:var(--y)}.none{padding:10px;font-style:italic}
.grp{background:#000;color:var(--y);padding:4px 10px;font-weight:bold;cursor:pointer;position:sticky;top:0;z-index:1;overflow-wrap:anywhere}
#main{flex:1;overflow:auto;padding:10px 14px;min-width:0}
#crumb{font-size:11px;overflow-wrap:anywhere}#rel{font-size:12px;border:1px solid #000;border-left:6px solid var(--y);padding:5px 10px;margin:4px 0 2px;overflow-wrap:anywhere}#rel:empty{display:none}
#rel a{color:#000;font-weight:bold;text-decoration:underline;cursor:pointer}#rel a:hover{background:var(--y)}#rel summary{cursor:pointer;margin-top:3px}#rel .ul{max-height:150px;overflow:auto;padding:4px 0;line-height:1.6}#rel .g{font-weight:normal}
.seg button{font:inherit;border:2px solid #000;background:#fff;padding:1px 10px;cursor:pointer;margin-right:2px}.seg button.on{background:#000;color:var(--y)}h2{margin:0 0 2px;font-size:17px;overflow-wrap:anywhere}
#tabs{display:flex;flex-wrap:wrap;margin-top:8px}#tabs button{font:inherit;font-size:12px;border:1px solid #000;border-bottom:0;background:#fff;padding:4px 14px;cursor:pointer}
#tabs button.on{background:var(--y);font-weight:bold}#tabs button.over{outline:3px solid #000}
#lib .hd .seg button{font:inherit;font-weight:bold;font-size:13px;padding:2px 12px;border:2px solid #000;background:#fff;cursor:pointer}#lib .hd .seg button.on{background:#000;color:var(--y)}
.xhd{padding:6px 8px;border-bottom:1px solid #000;display:flex;gap:8px;align-items:center;justify-content:space-between}.xhd button{font:inherit;font-size:11px;border:2px solid #000;background:var(--y);cursor:pointer;padding:1px 8px}
.xtabs{display:flex;flex-wrap:wrap;gap:3px;padding:6px 8px}.xtabs button{font:inherit;font-size:11px;border:1px solid #000;background:#fff;cursor:pointer;padding:1px 7px}.xtabs button.on{background:var(--y);font-weight:bold}
.xgrid{display:flex;flex-wrap:wrap;gap:4px;padding:6px 8px}.xtile{width:86px;border:2px solid #000;background:#fff;cursor:grab;user-select:none;font-size:10px;text-align:center;position:relative}
.xtile.picked{outline:3px solid #000;background:var(--y)}.xtile .xim{height:58px;display:flex;align-items:center;justify-content:center;overflow:hidden}.xtile .xim img{max-width:80px;max-height:56px}
.xtile .xl{padding:2px;height:26px;overflow:hidden;line-height:1.1}.xnote{padding:2px 8px 6px;font-size:11px;color:#444}
#stackzone{border:2px dashed #000;padding:8px;margin:6px 0;text-align:center;font-size:12px;background:#fffbe0}#stackzone.over{background:var(--y)}#props button.rm{font:inherit;font-size:11px;border:2px solid #000;background:var(--y);cursor:pointer;padding:0 8px}
#llist .lrow.located{outline:3px solid #000;outline-offset:-3px;background:var(--y)}
#conds{display:none;border-top:2px solid #000;margin-top:6px;padding:6px 4px;font-size:12px}#conds h3{margin:0 0 4px;font-size:13px}#conds table{border-collapse:collapse;margin:4px 0}#conds td,#conds th{border:1px solid #000;padding:2px 6px;text-align:left;font-weight:normal}
#conds th{background:var(--y);font-weight:bold}#conds select,#conds input[type=text]{font:inherit;font-size:12px;border:1px solid #000;padding:1px 3px}#conds input.num{width:60px}#conds input.nm{width:220px}#conds button{font:inherit;font-size:12px;border:2px solid #000;background:var(--y);cursor:pointer;padding:1px 10px}
#conds .applyrow{background:var(--y);border:2px solid #000;padding:5px 8px;margin:4px 0}#conds .hint{color:#444;margin:2px 0}#conds tr.newrow td{background:#fffbe0}
#props{display:none;border-top:2px solid #000;margin-top:6px;padding:6px 4px;font-size:12px}
#props h3{margin:0 0 4px;font-size:13px}#props table{border-collapse:collapse;margin:4px 0}#props td,#props th{border:1px solid #000;padding:2px 6px;text-align:left;font-weight:normal}
#props th{background:var(--y);font-weight:bold}#props select,#props input[type=text]{font:inherit;font-size:12px;border:1px solid #000;padding:1px 3px}#props input.num{width:60px}
#grid{display:flex;flex-wrap:wrap;gap:2px;border-top:2px solid #000;padding-top:2px;min-height:140px}
.tile{width:112px;height:130px;border:2px solid #000;background:#fff;display:flex;flex-direction:column;align-items:center;cursor:grab;position:relative}
.tile:hover{background:var(--y)}.tile.drag{opacity:.35}.tile.over{outline:4px solid #000}
.tile .im{height:76px;width:100%;display:flex;align-items:center;justify-content:center;padding:4px}.tile img{max-width:100%;max-height:100%;pointer-events:none}
.tile .no{font-size:10px;border:1px dashed #000;padding:14px 6px}.tile .wait{font-size:10px;opacity:.5}
.tile .lb{font-size:10px;text-transform:uppercase;text-align:center;line-height:1.15;overflow:hidden;padding:0 3px;flex:1;width:100%;word-break:break-word}
.tile .n{position:absolute;left:2px;top:2px;background:#000;color:var(--y);font-size:9px;padding:0 3px}
.tile .exb{position:absolute;right:5px;bottom:14px;font:bold 26px Arial,sans-serif;line-height:1;color:#000;pointer-events:none;text-shadow:0 0 3px #fff,0 0 3px #fff}
.tile .stk{position:absolute;right:2px;top:2px;min-width:16px;height:16px;line-height:16px;padding:0 3px;box-sizing:border-box;background:#000;color:var(--y);font:bold 12px/16px Arial,sans-serif;text-align:center;pointer-events:none}
.xtile .stk{position:absolute;right:2px;top:2px;min-width:14px;height:14px;line-height:14px;padding:0 2px;box-sizing:border-box;background:#000;color:var(--y);font:bold 11px/14px Arial,sans-serif;text-align:center;pointer-events:none}
.tile.picked{background:var(--y);outline:4px solid #000;outline-offset:-4px}.tile.picked .n{background:#fff;color:#000}
#selinfo{font-size:12px;font-weight:bold;min-height:16px;margin:2px 0}
#add{margin-top:14px;border-top:2px solid #000;padding-top:8px}#aq{width:380px;max-width:100%;font:inherit;padding:4px;border:1px solid #000}
.res{display:flex;gap:8px;align-items:center;border-bottom:1px solid #000;padding:3px 0;font-size:12px}.res button{font:inherit;border:2px solid #000;background:var(--y);cursor:pointer}
.res span.p{font-size:10px}
#lib{display:none;width:340px;flex:0 0 340px;flex-direction:column;min-height:0;border-right:2px solid #000}
#lib input{margin:6px 8px;padding:5px;font:inherit;border:2px solid #000}#libbtn.on{background:#000;color:var(--y)}
.lrow{display:flex;gap:6px;align-items:center;padding:3px 6px;border-bottom:1px solid #000;cursor:grab}.lrow:hover{background:var(--y)}
.lim{width:40px;height:40px;flex:0 0 40px;display:flex;align-items:center;justify-content:center}.lim img{max-width:100%;max-height:100%;pointer-events:none}
.ltx{min-width:0;overflow-wrap:anywhere;font-size:12px}.ltx small{display:block;font-size:10px}
.ub{font-size:10px}.inuse,.newp{font-size:10px;border:1px solid #000;padding:0 3px}.newp{background:var(--y)}
.lfold{cursor:pointer;font-weight:bold;padding:3px 6px;border-bottom:1px solid #000;background:#fff;overflow-wrap:anywhere;font-size:12px}.lfold:hover{background:var(--y)}
#stat.quiet,#imgs,#svc{display:none}
#boot{position:fixed;inset:0;z-index:100;background:var(--y);display:flex;align-items:center;justify-content:center}
#boot .card{max-width:640px;padding:22px 26px;border:3px solid #000;background:#fff;font-size:14px}#boot h2{margin:0 0 10px;font-size:20px}
#boot .st{font-weight:bold;margin:8px 0}#boot .hint{margin-top:8px;font-size:12px}#boot .err{color:#000;background:var(--y);border:2px solid #000;padding:8px;margin:10px 0;white-space:pre-wrap}
#boot button{font:inherit;padding:6px 14px;margin:6px 8px 0 0;border:2px solid #000;background:#fff;cursor:pointer}
#busy{display:none;background:#000;color:var(--y);padding:6px 12px;font-size:12px}
#notice{display:none;background:var(--y);border:3px solid #000;padding:8px 12px;margin:6px;font-size:13px}#notice button{font:inherit;margin-left:12px;padding:3px 12px;border:2px solid #000;background:#000;color:var(--y);cursor:pointer}
#slowlog{display:none;position:fixed;left:0;bottom:0;z-index:50;max-width:70%;background:#000;color:var(--y);padding:3px 10px;font-size:11px;cursor:pointer;line-height:1.5}
.tile.ins1{box-shadow:-7px 0 0 #000}.tile.ins2{box-shadow:7px 0 0 #000}
#lib .hd button{font:inherit;font-size:11px;padding:1px 6px;margin-left:8px;border:2px solid #000;background:#fff;cursor:pointer}
#ctx{position:fixed;display:none;background:#fff;border:2px solid #000;z-index:9;min-width:180px}
#ctx div{padding:5px 10px;cursor:pointer;border-bottom:1px solid #000}#ctx div:hover{background:var(--y)}#ctx .h{font-size:11px;font-weight:bold;cursor:default;background:#000;color:var(--y)}#ctx .h:hover{background:#000}
</style></head><body>
<div id="boot"><div class="card"><h2>Palette editor is starting</h2><div class="st" id="bootst">Connecting...</div><div class="hint" id="boothint"></div><div class="hint" id="bootsteps" style="white-space:pre-wrap"></div>
<div class="err" id="booterr" style="display:none"></div><div id="bootbtns" style="display:none"><button id="bootretry">Try again</button><button id="bootchoose">Choose the database folder...</button></div></div></div>
<div id="busy"></div><div id="slowlog"></div>
<header><b>Palette editor</b><button id="apply" disabled title="Write all changes to the profile's service.map (Ctrl+S)">Apply to file</button><button id="discard" disabled>Discard</button><button id="undo" disabled>Undo</button><button id="libbtn" title="Browse and search the Libraries folder; drag a part onto a tab">Libraries</button><span id="pend"></span><span id="stat"></span><span id="imgs"></span><span id="svc"></span></header>
<div id="wrap"><div id="left">
 <div class="sec" id="s1"><div class="hd"><span>1. Profile</span><small id="pcount"></small></div>
  <input id="pq" placeholder="type any part of a profile name" autocomplete="off"><div class="lst" id="plist"></div></div>
 <div class="sec" id="s2"><div class="hd"><span>2. Services</span><span class="seg" style="display:none"><button id="mSvc" class="on">Services</button><button id="mTpl">Templates</button></span><small id="scount"></small></div>
  <input id="sq" placeholder="type any part of a service name" autocomplete="off"><div class="lst" id="slist"></div></div>
</div><div id="split" title="drag to resize"></div>
<div id="lib"><div class="hd"><span class="seg"><button id="lmParts" class="on">Libraries</button><button id="lmTpl">Templates</button></span><small id="lcount"></small></div>
 <input id="lq" placeholder="search the libraries: any part of a name or folder" autocomplete="off"><div class="lst" id="llist"></div></div>
<div id="main">
<div id="notice"></div><div id="crumb"></div><h2 id="title"></h2><div id="rel"></div><div id="tabs"></div><div id="selinfo"></div><div id="grid"></div><div id="props"></div><div id="conds"></div>
</div></div>
<div id="ctx"></div>
<script>
const K='__TOKEN__';const $=id=>document.getElementById(id);
let P=[],L=[],S=[],pal=null,tab=0,drag=null,sel=new Set(),anchor=-1,curP=-1,curKey='',curSvc='',rev=0,dirty=false,nun=0,serverCanUndo=false,collapsed={},navSeq=0,profSeq=0,mode='svc';
const ORPHAN='(palettes with no service)';
let serverDirty=false,nHist=0,savedAt=0;
function qdelta(){return opq.reduce((a,it)=>a+(it.kind==='op'?1:(it.kind==='undo'&&!it.server?-1:0)),0)}
function applyCounter(){const dl=qdelta(),h=Math.max(0,nHist+dl);nun=Math.abs(h-savedAt);dirty=dl===0?serverDirty:nun>0;   /* distance from the last saved point, as the editor counts it */
  $('apply').disabled=!dirty;$('discard').disabled=!dirty;
  const p=$('pend');p.style.display=dirty?'inline':'none';p.textContent=dirty?(nun+' unsaved change'+(nun===1?'':'s')):'';
  window.onbeforeunload=dirty?(e=>{e.preventDefault();e.returnValue='';return ''}):null}
function setState(d){if(d.dirty!==undefined){serverDirty=d.dirty;if(d.n_hist!==undefined){nHist=d.n_hist;savedAt=d.saved_at||0}else{nHist=(d.n_unsaved||0)+savedAt}applyCounter()}
  if(d.can_undo!==undefined){serverCanUndo=d.can_undo;refreshUndo()}}
let inflight=0,infT=0;
function hideBusy(){$('busy').style.display='none'}
function track(pr){inflight++;if(inflight===1)infT=Date.now();const done=()=>{inflight--;if(inflight<=0){inflight=0;hideBusy()}};return pr.then(v=>{done();return v},e=>{done();throw e})}
let slowOpen=false;
function fmtSec(s){return s<120?s+' s':Math.round(s/60)+' min'}
function showNotice(n,lost,names){const el=$('notice');if(!n&&!lost){el.style.display='none';return}
  el.style.display='block';let h='';
  if(n)h+='<div><b>'+n+' button'+(n===1?'':'s')+' in this profile ha'+(n===1?'s':'ve')+' been damaged by an earlier version of this editor</b> (a lost picture or a wrong internal value). They can be repaired. <button id="fixids">Repair '+(n===1?'it':'them')+'</button></div>';
  if(lost)h+='<div><b>'+lost+' button'+(lost===1?'':'s')+' lost '+(lost===1?'its':'their')+' picture and cannot be repaired automatically</b>'+(names&&names.length?' ('+names.map(esc).join(', ')+(lost>names.length?', ...':'')+')':'')+'. Restore this profile from your server copy, or drag the part in again.</div>';
  el.innerHTML=h;const fb=$('fixids');if(fb)fb.onclick=repairIds}
function repairIds(){stat('Repairing...',false,true);whenIdle().then(()=>post('/api/op',{op:'repair_ids',ti:0,tab:0,rev:rev,echo:0})).then(d=>{if(!d||!d.ok){stat((d&&d.error)||'Could not repair',true);return}
  showNotice(0,0,[]);serverCanUndo=true;refreshUndo();applyCounter();stat(d.msg,false,true);setTimeout(refreshSlow,800)})}
function showSlow(l){const el=$('slowlog');if(!l||!l.length){el.style.display='none';return}
  const mx=Math.max.apply(null,l.map(x=>x.seconds));el.style.display='block';
  el.innerHTML=slowOpen?'Waits on the drive this session (newest last):<br>'+l.map(x=>esc(x.at+'   '+x.what+'   '+fmtSec(x.seconds))).join('<br>')+'<br>(click to close)'
    :'Slow drive: '+l.length+' wait'+(l.length===1?'':'s')+' of 3 s or more, longest '+fmtSec(mx)+'  -  click for details'}
$('slowlog').onclick=()=>{slowOpen=!slowOpen;refreshSlow()};
function refreshSlow(){fetch('/api/busy?k='+K).then(r=>r.json()).then(d=>{if(d&&d.ok){showSlow(d.slow);showNotice(d.damaged||0,d.lost||0,d.lost_names||[])}}).catch(()=>{})}
setInterval(refreshSlow,15000);
setInterval(()=>{if($('notice').style.display==='block'||pal)refreshSlow()},4000);
function showBusy(){fetch('/api/busy?k='+K).then(r=>r.json()).then(d=>{if(!inflight){hideBusy();return}const l=(d.busy||[]).map(x=>x.what+' ('+x.seconds+' s)'),el=$('busy');
  showSlow(d.slow);el.textContent=l.length?'Waiting for the drive: '+l.join('  |  ')+'.  The editor is fine - the disk or network is answering slowly.':'Working...';el.style.display='block'}).catch(()=>{})}
setInterval(()=>{if(inflight>0&&Date.now()-infT>2500)showBusy()},1500);
function api(p,o){return track(fetch(p+(p.includes('?')?'&':'?')+'k='+K,o).then(r=>r.json()).then(d=>{if(d){if(d.rev!==undefined)rev=d.rev;setState(d)}return d}))}
function post(p,b){return api(p,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(b)})}
function stat(m,bad,show){$('stat').textContent=m;$('stat').className=bad?'bad':(show?'':'quiet')}
function esc(s){return String(s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]))}
/* ---------- pictures: remembered in this page, fetched 6 at a time, current tab first ---------- */
const IMG=new Map(),PEND=new Map(),BAD=new Set(),FLY=new Set(),Q=[];let active=0;
function loadImg(el,id,holder){if(IMG.has(id)){el.src=IMG.get(id);return}
  if(BAD.has(id)){holder.innerHTML='<div class="no">no image</div>';return}
  if(PEND.has(id)){PEND.get(id).push({el,holder});return}PEND.set(id,[{el,holder}]);Q.push(id);pump()}
function pump(){while(active<3&&Q.length){        /* at most 3 pictures at a time: the browser allows 6 connections in all, and slow pictures must never block a click */const id=Q.shift();active++;FLY.add(id);
  fetch('/img/'+id+'?k='+K).then(r=>{if(r.status===404){const e=new Error('missing');e.nf=true;throw e}if(!r.ok||!(r.headers.get('Content-Type')||'').startsWith('image/'))throw new Error('bad reply');return r.blob()})
   .then(b=>{const u=URL.createObjectURL(b);IMG.set(id,u);(PEND.get(id)||[]).forEach(w=>{w.el.src=u});PEND.delete(id)})
   .catch(e=>{if(e&&e.nf)BAD.add(id);(PEND.get(id)||[]).forEach(w=>{w.holder.innerHTML='<div class="no">no image</div>'});PEND.delete(id)})
   .finally(()=>{FLY.delete(id);active--;pump()})}}
function prio(ids){const s=new Set(ids);for(let k=Q.length-1;k>=0;k--)if(s.has(Q[k]))Q.splice(k,1);Q.unshift(...ids.filter(i=>PEND.has(i)&&!FLY.has(i)))}
let imgTimer=null,lastPicRev=null,lastTpl={ti:-1,rev:-1};
function pendingCount(){return pal?pal.tabs.reduce((a,t)=>a+t.buttons.filter(b=>b.pend).length,0):0}
function pollImgs(){clearTimeout(imgTimer);const ti=pal?pal.i:-1;api('/api/imgstatus?ti='+ti).then(d=>{if(!d.ok)return;const el=$('imgs');
  let refresh=false;if(lastPicRev!==null&&d.pics_rev!==lastPicRev)refresh=true;if(lastTpl.ti===ti&&lastTpl.rev!==d.tpl_rev)refresh=true;
  lastPicRev=d.pics_rev;lastTpl={ti:ti,rev:d.tpl_rev};
  if(refresh&&pal&&!opq.length&&!sending){BAD.clear();api('/api/palette?i='+ti).then(x=>{if(x.ok&&pal&&pal.i===ti&&!opq.length&&!sending){pal=x.palette;palCache.set(ti,pal);draw()}})}
  const pend=pendingCount();
  if(d.scanning){el.textContent='Pictures: checking '+d.done+' of '+d.of+' templates...'+(pend?' ('+pend+' on this template still being found)':'');imgTimer=setTimeout(pollImgs,pend?400:800)}
  else{el.textContent=d.total?('Pictures: '+d.found.toLocaleString()+' found, '+d.missing.toLocaleString()+' not found ('+d.seconds+' s)'):'';if(pend)imgTimer=setTimeout(pollImgs,500)}}).catch(()=>{})}
/* ---------- search: names containing exactly what you typed first, then names containing every word ---------- */
function pick(names,q){const s=q.trim().toLowerCase();if(!s)return names.map((n,i)=>i);const w=s.split(/\s+/),a=[],b=[];
  names.forEach((n,i)=>{const h=n.toLowerCase();if(h.includes(s))a.push(i);else if(w.every(x=>h.includes(x)))b.push(i)});return a.concat(b)}
/* ---------- 1. profiles ---------- */
function drawProfiles(){const q=$('pq').value,el=$('plist');el.innerHTML='';let n=0;
  pick(P.map(p=>p.name),q).forEach(i=>{const p=P[i];n++;const r=document.createElement('div');r.className='row'+(i===curP?' sel':'');r.dataset.i=i;r.title=p.name;r.textContent=p.name;r.onclick=()=>openProfile(i);el.appendChild(r)});
  $('pcount').textContent=n+' of '+P.length;if(!n)el.innerHTML='<div class="none">No profile contains "'+esc(q)+'"</div>'}
function markP(){document.querySelectorAll('#plist .row').forEach(r=>r.classList.toggle('sel',+r.dataset.i===curP))}
$('pq').oninput=drawProfiles;
$('pq').onkeydown=e=>{if(e.key==='Enter'){const r=$('plist').querySelector('.row');if(r)r.click()}};
function openProfile(i){if(dirty&&i!==curP&&!confirm('Discard '+nun+' unsaved change'+(nun===1?'':'s')+' in the current profile?'))return;
  const my=++profSeq;navSeq++;stat('Opening...');whenIdle().then(()=>post('/api/open',{idx:i})).then(d=>{if(my!==profSeq)return;if(!d.ok){stat(d.error,true);return}
  curP=i;lastPicRev=null;pal=null;tab=0;curKey='';curSvc='';ustack.length=0;barrier=false;palCache.clear();xCache.clear();xView=null;if(lmode==='tpl')setTimeout(xList,300);repairedNote=d.repaired||0;refreshUndo();$('sq').value='';$('svc').textContent=d.path;stat('Opened '+d.profile);markP();draw();loadServices(my);pollImgs()})}
/* ---------- 2. services and templates, and the relationship between them ---------- */
let repairedNote=0;
function loadServices(my){api('/api/services').then(d=>{if(my!==undefined&&my!==profSeq)return;if(!d.ok){stat(d.error,true);return}S=d.services;L=d.palettes;drawServices();stat('Opened '+(P[curP]?P[curP].name:'')+' - '+S.length+' services, '+L.length+' templates'+(repairedNote?'. Repaired '+repairedNote+' button'+(repairedNote===1?'':'s')+' that an earlier version of this editor had damaged.':''));if(lmode==='tpl')xList()})}
function palLine(ti){const p=L[ti];return p?p.name+' ('+p.tabs+' tabs, '+p.buttons+' buttons)':''}
function usersOf(ti){return S.filter(s=>s.ti===ti)}
function setMode(m){mode=m;$('mSvc').classList.toggle('on',m==='svc');$('mTpl').classList.toggle('on',m==='tpl');
  $('sq').placeholder=m==='svc'?'type any part of a service, group or template name':'type any part of a template or service name';drawServices()}
$('mSvc').onclick=()=>setMode('svc');$('mTpl').onclick=()=>setMode('tpl');
function drawServices(){mode==='svc'?drawSvcList():drawTplList()}
function drawSvcList(){const q=$('sq').value.trim().toLowerCase(),el=$('slist'),top=el.scrollTop;el.innerHTML='';
  if(curP<0){el.innerHTML='<div class="none">Pick a profile first</div>';$('scount').textContent='';return}
  const items=S.map(s=>({key:'s'+s.i,name:s.name,group:s.group||'(no group)',ti:s.ti,pal:s.ti>=0&&L[s.ti]?L[s.ti].name:''}));
  const words=q.split(/\s+/).filter(Boolean);
  const rows=items.filter(it=>{if(!q)return true;const h=(it.name+' '+it.group+' '+it.pal).toLowerCase();return h.includes(q)||words.every(w=>h.includes(w))});
  const groups={};rows.forEach(r=>{(groups[r.group]=groups[r.group]||[]).push(r)});
  const names=Object.keys(groups).sort((a,b)=>a.toLowerCase().localeCompare(b.toLowerCase()));
  names.forEach(g=>{const open=q?true:!collapsed[g];const h=document.createElement('div');h.className='grp';h.textContent=(open?'[-] ':'[+] ')+g+' ('+groups[g].length+')';
    h.onclick=()=>{collapsed[g]=!collapsed[g];drawServices()};el.appendChild(h);if(!open)return;
    groups[g].sort((a,b)=>a.name.toLowerCase().localeCompare(b.name.toLowerCase())).forEach(it=>{const r=document.createElement('div');r.className='row svc'+(it.key===curKey?' sel':'');r.dataset.key=it.key;
      r.title=it.name+(it.pal?'  ->  template: '+it.pal:'');r.innerHTML=esc(it.name)+'<small>'+(it.ti>=0?'template: '+esc(it.pal):'no template assigned')+'</small>';
      r.onclick=()=>openService(it);el.appendChild(r)})});
  $('scount').textContent=rows.length+' of '+items.length;if(!rows.length)el.innerHTML='<div class="none">No service contains "'+esc($('sq').value)+'"</div>';el.scrollTop=top}
function drawTplList(){const q=$('sq').value.trim().toLowerCase(),el=$('slist'),top=el.scrollTop;el.innerHTML='';
  if(curP<0){el.innerHTML='<div class="none">Pick a profile first</div>';$('scount').textContent='';return}
  const byT={};S.forEach(s=>{if(s.ti>=0)(byT[s.ti]=byT[s.ti]||[]).push(s)});
  const items=L.map(p=>({ti:p.i,name:p.name,p:p,users:byT[p.i]||[]}));const words=q.split(/\s+/).filter(Boolean);
  const rows=items.filter(it=>{if(!q)return true;const h=(it.name+' '+it.users.map(u=>u.name).join(' ')).toLowerCase();return h.includes(q)||words.every(w=>h.includes(w))});
  const rank=it=>q&&it.name.toLowerCase().includes(q)?0:1;
  rows.sort((a,b)=>rank(a)-rank(b)||a.name.toLowerCase().localeCompare(b.name.toLowerCase()));
  rows.forEach(it=>{const r=document.createElement('div');r.className='row'+((pal&&pal.i===it.ti)?' sel':'');r.dataset.key='t'+it.ti;
    const u=it.users.length;r.title=it.name+(u?'  |  used by: '+it.users.slice(0,8).map(x=>x.name).join(', ')+(u>8?', ...':''):'');
    r.innerHTML=esc(it.name)+'<small>'+it.p.tabs+' tabs, '+it.p.buttons+' buttons | '+(u?'used by '+u+' service'+(u===1?'':'s'):'not used by any service')+'</small>';
    r.onclick=()=>openTemplate(it.ti);el.appendChild(r)});
  $('scount').textContent=rows.length+' of '+items.length;if(!rows.length)el.innerHTML='<div class="none">No template or service contains "'+esc($('sq').value)+'"</div>';el.scrollTop=top}
$('sq').oninput=drawServices;
$('sq').onkeydown=e=>{if(e.key==='Enter'){const r=$('slist').querySelector('.row');if(r)r.click()}};
function showPalette(ti,key,name){curKey=key;curSvc=name||'';document.querySelectorAll('#slist .row').forEach(r=>r.classList.toggle('sel',r.dataset.key===curKey));
  if(ti<0){navSeq++;pal=null;draw();$('title').textContent=name||'';const sv=S.find(s=>'s'+s.i===key);
    $('rel').innerHTML='<b>Service:</b> '+esc(name||'')+(sv&&sv.group?' <span class="g">('+esc(sv.group)+')</span>':'')+'  &rarr;  <b>Template:</b> none assigned';return}
  const my=++navSeq,c=palCache.get(ti);
  if(c){pal=c;tab=0;sel.clear();draw();pollImgs()}                       // revisiting: on screen at once
  api('/api/palette?i='+ti).then(d=>{if(my!==navSeq)return;if(!d.ok){stat(d.error,true);return}
    if(opq.length||sending)return;                                          // local edits in flight: the screen is newer than this reply
    if(c&&pal===c&&sameLayout(c,d.palette)){let ch=false;c.tabs.forEach((t,k)=>t.buttons.forEach((b,n)=>{const s=d.palette.tabs[k].buttons[n];if(b.img!==s.img||b.pend!==s.pend){b.img=s.img;b.pend=s.pend;ch=true}}));if(ch)draw();pollImgs();return}
    pal=d.palette;if(!c)tab=0;palCache.set(ti,pal);draw();pollImgs()})}
function openService(it){showPalette(it.ti,it.key,it.name)}
function openTemplate(ti){showPalette(ti,'t'+ti,'')}
function goService(i){const s=S.find(x=>x.i===i);if(!s)return;$('sq').value='';collapsed[s.group||'(no group)']=false;setMode('svc');
  showPalette(s.ti,'s'+s.i,s.name);const r=document.querySelector('#slist .row[data-key="s'+s.i+'"]');if(r&&r.scrollIntoView)r.scrollIntoView({block:'center'})}
function relHTML(){const us=usersOf(pal.i);const link=s=>'<a data-s="'+s.i+'">'+esc(s.name)+'</a>';
  if(curKey[0]==='s'){const cur=S.find(s=>'s'+s.i===curKey);const others=us.filter(s=>!cur||s.i!==cur.i);
    return '<b>Service:</b> '+esc(cur?cur.name:'')+(cur&&cur.group?' <span class="g">('+esc(cur.group)+')</span>':'')+'  &rarr;  <b>Template:</b> '+esc(pal.name)+'<br>'
      +(others.length?'<details><summary>Also used by '+others.length+' other service'+(others.length===1?'':'s')+' - an edit here changes all of them</summary><div class="ul">'+others.map(link).join(' &middot; ')+'</div></details>':'Only this service uses this template.')}
  return '<b>Template:</b> '+esc(pal.name)+'<br>'+(us.length?'<details'+(us.length<=12?' open':'')+'><summary>Used by '+us.length+' service'+(us.length===1?'':'s')+(us.length>1?' - an edit here changes all of them':'')+'</summary><div class="ul">'+us.map(link).join(' &middot; ')+'</div></details>':'No service uses this template.')}
function setPal(p){pal=p;palCache.set(p.i,p);if(tab>=pal.tabs.length)tab=pal.tabs.length-1;if(pendingCount())setTimeout(pollImgs,200);const e=L.find(x=>x.i===p.i);
  if(e){e.tabs=p.tabs.length;e.buttons=p.tabs.reduce((a,t)=>a+t.buttons.length,0)}draw();drawServices()}
/* ---------- edits: shown on screen at once, sent to the editor in order in the background ---------- */
const opq=[],ustack=[],palCache=new Map();let sending=false,barrier=false,localSeq=0,vTimer=null,saving=false,msgSeq=0;
function cl(x){return JSON.parse(JSON.stringify(x))}
function refreshUndo(){$('undo').disabled=!(ustack.length||serverCanUndo)}
function whenIdle(){return new Promise(res=>{const f=()=>{if(!opq.length&&!sending)res();else setTimeout(f,25)};f()})}
function localMsg(b){const n=x=>x+1;const extra=b._note?'  ['+b._note+']':'';return localMsg0(b)+extra}
function localMsg0(b){const n=x=>x+1;switch(b.op){
  case 'move':return 'Moved button '+n(b.from)+' to position '+n(b.to);case 'delete':return 'Deleted button '+n(b.idx);
  case 'copy':return 'Copied button '+n(b.idx)+' to another tab';case 'add':return 'Added '+(b._name||b.item);
  case 'delete_many':return 'Deleted '+b.idxs.length+' buttons';case 'move_many':return 'Moved '+b.idxs.length+' buttons';
  case 'copy_many':return 'Copied '+b.idxs.length+' buttons';case 'tab_exclude':return b.value?'Excluding tab from fill':'Including tab in fill';
  case 'repair_ids':return 'Repairing buttons';case 'tab_delete':return 'Deleting tab';case 'item_cond':return 'Changing size restriction';case 'item_cond_many':return 'Setting size restriction on '+b.idxs.length+' button'+(b.idxs.length===1?'':'s');case 'stack_add':return 'Stacking a part on the button';case 'stack_remove':return 'Taking a part off the button';case 'cond_add':return 'Adding restriction';case 'cond_delete':return 'Deleting restriction';case 'cond_rename':return 'Renaming restriction';case 'cond_limits':return 'Changing condition';case 'tab_fix':return b.exclude===false?'Removing [E] from the tab name':'Repairing tab';case 'btn_exclude':return (b.value?'Excluding ':'Including ')+b.idxs.length+' button'+(b.idxs.length===1?'':'s')+(b.value?' from fill':' in fill');default:return b.op}}
/* the same list changes the editor makes, applied to what is on screen */
function applyLocal(b){const T=pal.tabs[b.tab];
  if(b.op==='move'){const x=T.buttons.splice(b.from,1)[0];T.buttons.splice(b.to,0,x)}
  else if(b.op==='delete'){if(b.idx<0||b.idx>=T.buttons.length)throw new Error('range');T.buttons.splice(b.idx,1)}
  else if(b.op==='copy'){const D=pal.tabs[b.dtab],c=cl(T.buttons[b.idx]);D.buttons.splice(b.pos==null?D.buttons.length:b.pos,0,c)}
  else if(b.op==='add'){T.buttons.splice(b.pos==null?T.buttons.length:b.pos,0,{name:b._name||b.item,items:[b.item],img:b._img>=0?b._img:-1,pend:!(b._img>=0)})}
  else if(b.op==='delete_many'){const s=new Set(b.idxs);T.buttons=T.buttons.filter((x,i)=>!s.has(i));}
  else if(b.op==='move_many'){const s=new Set(b.idxs),pick=b.idxs.map(i=>T.buttons[i]),rest=T.buttons.filter((x,i)=>!s.has(i));T.buttons=rest.slice(0,b.to).concat(pick,rest.slice(b.to))}
  else if(b.op==='copy_many'){const D=pal.tabs[b.dtab],cp=b.idxs.map(i=>cl(T.buttons[i])),pos=b.pos==null?D.buttons.length:b.pos;D.buttons.splice(pos,0,...cp)}
  else if(b.op==='item_cond'){const ic=T.buttons[b.idx].ic||(T.buttons[b.idx].ic=[]);ic[b.item]=b.value}
  else if(b.op==='cond_limits'){const c=(pal.conds||[]).find(x=>x.id===b.cid);if(!c)throw new Error('cond');c.gt=b.gt==null?-1:b.gt;c.lte=b.lte==null?-1:b.lte}
  else if(b.op==='stack_add'){const x=T.buttons[b.idx];if(!x)throw new Error('b');x.items.push(b.item);x.ic=(x.ic||[]).slice();while(x.ic.length<x.items.length)x.ic.push(0)}
  else if(b.op==='stack_remove'){const x=T.buttons[b.idx];if(!x||x.items.length<2||b.k<0||b.k>=x.items.length)throw new Error('b');x.items.splice(b.k,1);if(x.ic)x.ic.splice(b.k,1)}
  else if(b.op==='cond_add'){const cs=pal.conds||(pal.conds=[]);cs.push({id:Math.max(...cs.map(c=>c.id))+1,name:b.name.trim(),gt:b.gt==null?-1:b.gt,lte:b.lte==null?-1:b.lte})}
  else if(b.op==='cond_delete'){pal.conds=(pal.conds||[]).filter(c=>c.id!==b.cid);pal.tabs.forEach(t=>t.buttons.forEach(x=>{if(x.ic)x.ic=x.ic.map(v=>v===b.cid?0:v)}))}
  else if(b.op==='cond_rename'){const c=(pal.conds||[]).find(x=>x.id===b.cid);if(!c)throw new Error('cond');c.name=b.name.trim()}
  else if(b.op==='item_cond_many'){b.idxs.forEach(i=>{const x=T.buttons[i];if(b.k!=null){if(x&&x.items.length>b.k){(x.ic||(x.ic=x.items.map(()=>0)))[b.k]=b.value}}else if(x&&x.items.length===1){(x.ic||(x.ic=[]))[0]=b.value}})}
  else if(b.op==='repair_ids'){}
  else if(b.op==='tab_delete'){if(pal.tabs.length<=1)throw new Error('last');pal.tabs.splice(b.tab,1);if(tab>=pal.tabs.length)tab=pal.tabs.length-1;else if(tab>b.tab)tab--}
  else if(b.op==='btn_exclude'){b.idxs.forEach(i=>{T.buttons[i].ex=!!b.value})}
  else if(b.op==='tab_exclude'){const tb=pal.tabs[b.tab];tb.ex=b.value!=null?!!b.value:!tb.ex}
  else if(b.op==='tab_fix'){const tb=pal.tabs[b.tab];tb.name=tb.name.replace(/\s*\[E\]\s*$/,'')||'(unnamed)';tb.ex=b.exclude!==false}
  else throw new Error('unknown')}
function paintLocal(){if(!pal)return;const e=L.find(x=>x.i===pal.i);if(e){e.tabs=pal.tabs.length;e.buttons=pal.tabs.reduce((a,t)=>a+t.buttons.length,0)}
  palCache.set(pal.i,pal);if(xCache.has(pal.i))xCache.set(pal.i,pal);draw();drawServices();if(lmode==='tpl'&&xView&&xView.ti===pal.i&&!xdrag)xDraw();if(pendingCount())setTimeout(pollImgs,200)}
function doOp(b){if(!pal)return Promise.resolve();if(barrier){stat('Undoing - one moment...',true);return Promise.resolve()}
  b.ti=pal.i;const snap={ti:pal.i,name:pal.name,tabs:cl(pal.tabs),conds:cl(pal.conds||[])};
  try{applyLocal(b)}catch(e){pal.tabs=snap.tabs;pal.conds=snap.conds;stat('That change is not possible here.',true);return Promise.resolve()}
  ustack.push(snap);if(ustack.length>150)ustack.shift();
  localSeq++;if(!b._keep){sel.clear();anchor=-1}const mid=++msgSeq;stat(localMsg(b)+'  (saving...)');
  paintLocal();opq.push({kind:'op',body:b,mid:mid});applyCounter();refreshUndo();pumpQ();return Promise.resolve()}
function postQ(p,b){return track(fetch(p+'?k='+K,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(b)}).then(r=>r.json()).then(d=>{if(d&&d.rev!==undefined)rev=d.rev;return d}))}
function pumpQ(){if(sending||!opq.length)return;sending=true;const it=opq[0];let p;
  if(it.kind==='op'){const body=Object.assign({},it.body,{rev:rev,echo:0});delete body._name;delete body._img;delete body._note;delete body._keep;
    p=postQ('/api/op',body).then(d=>{if(!d.ok){opq.length=0;ustack.length=0;refreshUndo();return resyncHard(d.error)}
      opq.shift();setState(d);if(msgSeq===it.mid)stat(d.msg+(it.body._note?'  ['+it.body._note+']':''))})}
  else if(it.kind==='undo'){p=postQ('/api/undo',{}).then(d=>{if(!d.ok){opq.length=0;ustack.length=0;barrier=false;refreshUndo();return resyncHard(d.error)}
      opq.shift();setState(d);if(it.server){barrier=false;ustack.length=0;refreshUndo();if(d.palette){tab=0;pal=d.palette;paintLocal()}}if(msgSeq===it.mid||it.server)stat(d.msg)})}
  else if(it.kind==='discard'){p=postQ('/api/discard',{ti:it.ti}).then(d=>{opq.shift();if(!d.ok){stat(d.error,true);applyCounter();return}setState(d);ustack.length=0;refreshUndo();if(d.palettes)L=d.palettes;
      stat(d.msg);if(d.palette){pal=d.palette;paintLocal()}else{pal=null;draw();drawServices()}})}
  else if(it.kind==='apply'){opq.shift();saving=true;stat('Writing to the file...');      // the write is not waited for: editing carries on
    post('/api/apply',{}).then(d=>{if(!d.ok)stat(d.error,true);else{stat(d.msg,false,true);const sm=d.msg;setTimeout(()=>{if($('stat').textContent===sm)stat('')},3000)}}).catch(()=>stat('No reply from the editor - check the file before trusting it.',true)).finally(()=>{saving=false});
    sending=false;pumpQ();return}
  p.catch(()=>{const later=Math.max(0,opq.length-1);opq.length=0;ustack.length=0;barrier=false;refreshUndo();
    return resyncHard('No reply from the editor'+(later?' - '+later+' later change'+(later===1?' was':'s were')+' not applied, please repeat '+(later===1?'it':'them'):'')+'.')}).finally(()=>{sending=false;if(opq.length)pumpQ();else scheduleVerify()})}
function sameLayout(a,b){if(a.tabs.length!==b.tabs.length)return false;
  for(let k=0;k<a.tabs.length;k++){const x=a.tabs[k],y=b.tabs[k];if(x.name!==y.name||x.buttons.length!==y.buttons.length)return false;
    for(let n=0;n<x.buttons.length;n++)if(x.buttons[n].name!==y.buttons[n].name||(x.buttons[n].items[0]||'')!==(y.buttons[n].items[0]||''))return false}return true}
/* once the queue is empty, quietly check the screen against the editor's data and correct it if they ever differ */
function scheduleVerify(){clearTimeout(vTimer);if(!pal)return;const seq=localSeq,ti=pal.i;
  vTimer=setTimeout(()=>{if(opq.length||sending||seq!==localSeq||!pal||pal.i!==ti)return;
    api('/api/palette?i='+ti).then(d=>{if(!d.ok||opq.length||sending||seq!==localSeq||!pal||pal.i!==ti)return;
      if(sameLayout(pal,d.palette)){let ch=false;pal.tabs.forEach((t,k)=>t.buttons.forEach((b,n)=>{const s=d.palette.tabs[k].buttons[n];if(b.img!==s.img||b.pend!==s.pend){b.img=s.img;b.pend=s.pend;ch=true}}));if(ch){palCache.set(ti,pal);draw()}return}
      ustack.length=0;refreshUndo();pal=d.palette;paintLocal();stat('Corrected from the editor.',true)}).catch(()=>{})},700)}
function resyncHard(why){clearTimeout(vTimer);ustack.length=0;refreshUndo();applyCounter();if(!pal)return Promise.resolve();
  return api('/api/palette?i='+pal.i).then(d=>{if(d.ok){pal=d.palette;paintLocal();stat((why?why+'  ':'')+'(The view was reloaded to match the editor.)',true)}else stat(d.error,true)}).catch(()=>stat('Lost contact with the editor. Reload this page.',true))}
$('undo').onclick=()=>{
  if(ustack.length){const e=ustack.pop();
    if(pal&&e.ti===pal.i){pal.tabs=e.tabs;pal.conds=e.conds}else{pal={i:e.ti,name:e.name,tabs:e.tabs,conds:e.conds};tab=0;curKey='t'+e.ti;curSvc=''}
    if(tab>=pal.tabs.length)tab=pal.tabs.length-1;
    localSeq++;sel.clear();anchor=-1;const mid=++msgSeq;opq.push({kind:'undo',mid:mid});stat('Undone');applyCounter();refreshUndo();paintLocal();pumpQ()}
  else{if(barrier||!serverCanUndo)return;barrier=true;stat('Undoing...');opq.push({kind:'undo',server:true});pumpQ()}};
function applyAll(){if(saving||!dirty)return;opq.push({kind:'apply'});pumpQ()}
$('apply').onclick=applyAll;
document.addEventListener('keydown',e=>{if((e.ctrlKey||e.metaKey)&&e.key==='s'){e.preventDefault();applyAll();return}
  const tg=e.target&&e.target.tagName,typing=tg==='INPUT'||tg==='TEXTAREA'||tg==='SELECT';if(typing||!pal)return;
  if((e.ctrlKey||e.metaKey)&&(e.key==='a'||e.key==='A')){e.preventDefault();sel=new Set([...Array(nTiles()).keys()]);anchor=0;markSel();return}
  if(e.key==='Escape'){clearSel();hideCtx();return}
  if((e.key==='Delete'||e.key==='Backspace')&&sel.size){e.preventDefault();deleteGroup(selList())}});
$('discard').onclick=()=>{if(!dirty)return;if(!confirm('Discard '+nun+' unsaved change'+(nun===1?'':'s')+'? The file on disk is left as it is.'))return;
  opq.push({kind:'discard',ti:pal?pal.i:-1});pumpQ()};
/* ---------- 3. buttons ---------- */
function tabLabel(t){return t.name+(t.ex&&!/\[E\]\s*$/.test(t.name)?' [E]':'')}
function nTiles(){const T=pal&&pal.tabs[tab];return T?T.buttons.length:0}
function condText(c){const g=c.gt<0?null:c.gt,l=c.lte<0?null:c.lte;return c.name+(g==null&&l==null?'':'   ('+(g!=null?'greater than '+g:'')+(g!=null&&l!=null?', ':'')+(l!=null?'up to '+l:'')+')')}
function renderProps(){const el=$('props');if(!pal||sel.size!==1||!pal.tabs[tab]){el.style.display='none';el.innerHTML='';return}
  const n=[...sel][0],b=pal.tabs[tab].buttons[n];if(!b){el.style.display='none';el.innerHTML='';return}
  const cs=pal.conds||[],ic=b.ic||[];el.style.display='block';
  let h='<h3>Button: '+esc(b.name)+'</h3><label><input type="checkbox" id="pex"'+(b.ex?' checked':'')+'> Exclude From Fill</label>';
  h+='<table><tr><th>Item</th><th>Size restriction (condition)</th><th></th></tr>';
  b.items.forEach((p,k)=>{h+='<tr><td title="'+esc(p)+'">'+esc(p.split('/').pop())+'</td><td><select data-k="'+k+'">'+cs.map(c=>'<option value="'+c.id+'"'+(c.id===ic[k]?' selected':'')+'>'+esc(condText(c))+'</option>').join('')+'</select></td><td>'+(b.items.length>1?'<button class="rm" data-rm="'+k+'" title="Take this part off the button">Remove</button>':'')+'</td></tr>'});
  h+='</table><div id="stackzone">Drop a part from the Libraries here to stack it on this button</div>';
  el.innerHTML=h;
  const px=el.querySelector('#pex');if(px)px.onchange=()=>doOp({op:'btn_exclude',ti:pal.i,tab:tab,idxs:[n],value:px.checked,_keep:1});
  el.querySelectorAll('select[data-k]').forEach(s=>{s.onchange=()=>doOp({op:'item_cond',ti:pal.i,tab:tab,idx:n,item:+s.dataset.k,value:+s.value,_keep:1})});
  el.querySelectorAll('button[data-rm]').forEach(x=>{x.onclick=()=>doOp({op:'stack_remove',ti:pal.i,tab:tab,idx:n,k:+x.dataset.rm,_keep:1})});
  const z=el.querySelector('#stackzone');
  if(z){z.ondragover=e=>{if(libDrag){e.preventDefault();e.stopPropagation();hideIns();z.classList.add('over')}};z.ondragleave=()=>z.classList.remove('over');
    z.ondrop=e=>{e.preventDefault();e.stopPropagation();z.classList.remove('over');if(libDrag){const it=libDrag;libDrag=null;stackFromLib(it,n)}}}}
/* stacking: another part on the same button (a button holds one or more parts; each has its own size restriction) */
function stackFromLib(it,n){const b=pal.tabs[tab].buttons[n];if(!b)return;
  if(b.items.some(x=>x.toLowerCase()===it.rel.toLowerCase())){msgSeq++;stat('That part is already on this button.',true);return}
  doOp({op:'stack_add',ti:pal.i,tab:tab,idx:n,item:it.rel,_keep:1})}
/* the template's list of size restrictions: always on screen while a template is open, applied to any number of selected buttons */
let newCond={n:'',g:'',l:''};
function numOrNull(s){s=String(s).trim();if(s==='')return null;const v=parseFloat(s);return isNaN(v)||v<0?NaN:v}
function renderConds(){const el=$('conds');if(!pal||!pal.conds||!pal.tabs[tab]){el.style.display='none';el.innerHTML='';return}
  const cs=pal.conds,n=sel.size;el.style.display='block';
  let h='<h3>Size restrictions of this template</h3>';
  const mi=Math.max(1,...[...sel].map(i=>(pal.tabs[tab].buttons[i]||{items:[]}).items.length)),ord=['1st','2nd','3rd'];
  h+=n?'<div class="applyrow">Apply to the '+n+' selected button'+(n===1?'':'s')+': <select id="cap">'+cs.map(c=>'<option value="'+c.id+'">'+esc(condText(c))+'</option>').join('')+'</select>'+(mi>1?' on the <select id="capk">'+Array.from({length:mi},(_,k)=>'<option value="'+k+'">'+(ord[k]||(k+1)+'th')+' item</option>').join('')+'</select> of each button':'')+' <button id="capb">Apply</button></div>'
     :'<div class="hint">Select one or more buttons, then choose a restriction to apply to all of them.</div>';
  h+='<table><tr><th>Name</th><th>Greater than</th><th>Up to and including</th><th></th></tr>';
  cs.forEach(c=>{h+='<tr><td><input type="text" class="nm" data-r="'+c.id+'" value="'+esc(c.name)+'"'+(c.id===0?' readonly title="This entry always exists"':'')+'></td>'
    +'<td><input type="text" class="num" data-c="'+c.id+'" data-f="gt" value="'+(c.gt<0?'':c.gt)+'"></td><td><input type="text" class="num" data-c="'+c.id+'" data-f="lte" value="'+(c.lte<0?'':c.lte)+'"></td>'
    +'<td>'+(c.id===0?'':'<button class="del" data-d="'+c.id+'">Delete</button>')+'</td></tr>'});
  h+='<tr class="newrow"><td><input type="text" id="nn" placeholder="New restriction name" value="'+esc(newCond.n)+'"></td><td><input type="text" class="num" id="ng" value="'+esc(newCond.g)+'"></td><td><input type="text" class="num" id="nl" value="'+esc(newCond.l)+'"></td><td><button id="nadd">Add</button></td></tr></table>';
  h+='<div class="hint">A restriction applies to sizes greater than the first number and up to the second; empty means no limit. Each item of a button uses one restriction.</div>';
  el.innerHTML=h;
  const cap=el.querySelector('#capb');if(cap)cap.onclick=()=>{const ck=el.querySelector('#capk');doOp({op:'item_cond_many',ti:pal.i,tab:tab,idxs:selList(),value:+el.querySelector('#cap').value,k:ck?+ck.value:undefined,_keep:1})};
  el.querySelectorAll('input[data-c]').forEach(i=>{i.onchange=()=>{const c=cs.find(x=>x.id===+i.dataset.c);const v=numOrNull(i.value);if(Number.isNaN(v)){stat('A size limit is a number of inches, or empty for none.',true);renderConds();return}
    const gt=i.dataset.f==='gt'?v:(c.gt<0?null:c.gt),lte=i.dataset.f==='lte'?v:(c.lte<0?null:c.lte);doOp({op:'cond_limits',ti:pal.i,tab:tab,cid:c.id,gt:gt,lte:lte,_keep:1})}});
  el.querySelectorAll('input[data-r]').forEach(i=>{i.onchange=()=>{const c=cs.find(x=>x.id===+i.dataset.r),nm=i.value.trim();
    if(!nm||nm.length>40){stat('A restriction needs a name of 1 to 40 characters.',true);renderConds();return}
    if(cs.some(x=>x.id!==c.id&&x.name.toLowerCase()===nm.toLowerCase())){stat('This template already has a restriction called "'+nm+'".',true);renderConds();return}
    if(nm!==c.name)doOp({op:'cond_rename',ti:pal.i,tab:tab,cid:c.id,name:nm,_keep:1})}});
  el.querySelectorAll('button[data-d]').forEach(b=>{b.onclick=()=>{const id=+b.dataset.d,c=cs.find(x=>x.id===id);let used=0;
    pal.tabs.forEach(t=>t.buttons.forEach(x=>(x.ic||[]).forEach(v=>{if(v===id)used++})));
    if(used&&!confirm(used+' item'+(used===1?'':'s')+' use "'+c.name+'". They will go back to Unrestricted. Delete it?'))return;
    doOp({op:'cond_delete',ti:pal.i,tab:tab,cid:id,_keep:1})}});
  const keep=()=>{newCond={n:$('nn').value,g:$('ng').value,l:$('nl').value}};['nn','ng','nl'].forEach(id=>{$(id).oninput=keep});
  $('nadd').onclick=()=>{keep();const nm=newCond.n.trim(),g=numOrNull(newCond.g),l=numOrNull(newCond.l);
    if(!nm||nm.length>40){stat('A restriction needs a name of 1 to 40 characters.',true);return}
    if(cs.some(x=>x.name.toLowerCase()===nm.toLowerCase())){stat('This template already has a restriction called "'+nm+'".',true);return}
    if(Number.isNaN(g)||Number.isNaN(l)){stat('A size limit is a number of inches, or empty for none.',true);return}
    if(g!==null&&l!==null&&g>=l){stat('"Greater than" must be smaller than "up to and including".',true);return}
    newCond={n:'',g:'',l:''};doOp({op:'cond_add',ti:pal.i,tab:tab,name:nm,gt:g,lte:l,_keep:1})}}
function markSel(){renderProps();renderConds();document.querySelectorAll('#grid .tile').forEach((el,i)=>el.classList.toggle('picked',sel.has(i)));
  $('selinfo').textContent=sel.size>1?(sel.size+' buttons selected - drag to move them, Delete to remove them, right-click for more'):''}
function clearSel(){sel.clear();anchor=-1;markSel()}
function pickTile(n,e){if(e.ctrlKey||e.metaKey){if(sel.has(n))sel.delete(n);else sel.add(n);anchor=n}
  else if(e.shiftKey&&anchor>=0){sel.clear();for(let i=Math.min(anchor,n);i<=Math.max(anchor,n);i++)sel.add(i)}
  else{sel.clear();sel.add(n);anchor=n}markSel()}
function selList(){return [...sel].sort((a,b)=>a-b)}
/* where a dragged group lands: same rule as a single tile (forward = after the target, backward = before it) */
function landing(idxs,n){const set=new Set(idxs),rest=[];for(let i=0;i<nTiles();i++)if(!set.has(i))rest.push(i);
  const r=rest.indexOf(n);return idxs.every(i=>i<n)?r+1:r}
function moveGroup(idxs,n){if(idxs.length===1)return doOp({op:'move',ti:pal.i,tab:tab,from:idxs[0],to:n});return doOp({op:'move_many',ti:pal.i,tab:tab,idxs:idxs,to:landing(idxs,n)})}
function moveGroupToEnd(idxs){const N=nTiles();if(idxs.length===1){if(idxs[0]!==N-1)doOp({op:'move',ti:pal.i,tab:tab,from:idxs[0],to:N-1});return}
  return doOp({op:'move_many',ti:pal.i,tab:tab,idxs:idxs,to:N-idxs.length})}
function hasItem(k,item){const l=(item||'').toLowerCase();return !!l&&pal.tabs[k].buttons.some(b=>b.items.length&&b.items[0].toLowerCase()===l)}
/* copying to another tab (drag onto its name): a button that is already in that tab is skipped - you never want it twice */
function copyGroup(idxs,dtab){const T=pal.tabs[tab];const keep=idxs.filter(i=>!hasItem(dtab,T.buttons[i].items[0]||''));const skipped=idxs.length-keep.length;
  if(!keep.length){msgSeq++;stat(idxs.length===1?'That button is already in that tab - not copied.':'All '+idxs.length+' buttons are already in that tab - nothing copied.',true);return Promise.resolve()}
  const note=skipped?skipped+' already in that tab':'';
  if(keep.length===1)return doOp({op:'copy',ti:pal.i,tab:tab,idx:keep[0],dtab:dtab,_note:note});
  return doOp({op:'copy_many',ti:pal.i,tab:tab,idxs:keep,dtab:dtab,_note:note})}
function deleteGroup(idxs){if(idxs.length===1)return doOp({op:'delete',ti:pal.i,tab:tab,idx:idxs[0]});return doOp({op:'delete_many',ti:pal.i,tab:tab,idxs:idxs})}
function draw(){hideCtx();
  if(!pal){$('title').textContent=curP<0?'':'Pick a service';$('crumb').textContent='';$('rel').innerHTML='';$('tabs').innerHTML='';$('grid').innerHTML='';sel.clear();$('selinfo').textContent='';return}
  $('crumb').textContent=(P[curP]?P[curP].name:'')+'  >  '+(curSvc?curSvc+'  >  ':'')+pal.name+(pal.ms?'   (opened in '+pal.ms.total+' ms)':'');$('title').textContent=pal.name;
  $('rel').innerHTML=relHTML();$('rel').querySelectorAll('a[data-s]').forEach(a=>{a.onclick=()=>goService(+a.dataset.s)});
  const tb=$('tabs');tb.innerHTML='';
  pal.tabs.forEach((t,k)=>{const b=document.createElement('button');b.textContent=tabLabel(t)+' ('+t.buttons.length+')';if(k===tab)b.className='on';
    b.onclick=()=>{if(tab!==k)clearSel();tab=k;draw()};
    b.oncontextmenu=e=>{e.preventDefault();showTabCtx(e.clientX,e.clientY,k)};
    b.ondragover=e=>{if((drag&&k!==tab)||libDrag||xdrag){e.preventDefault();b.classList.add('over')}};b.ondragleave=()=>b.classList.remove('over');
    b.ondrop=e=>{e.preventDefault();b.classList.remove('over');if(xdrag){addFromTemplate(k,null);return}if(libDrag){addFromLib(libDrag,k,null);libDrag=null;return}if(drag&&k!==tab)copyGroup(drag.idxs,k);drag=null};
    tb.appendChild(b)});
  const g=$('grid');g.innerHTML='';const T=pal.tabs[tab];if(!T)return;const ids=[];
  sel=new Set([...sel].filter(i=>i<T.buttons.length));
  g.onclick=e=>{if(e.target===g)clearSel()};
  g.ondragover=e=>{if(drag||libDrag||xdrag)e.preventDefault()};
  g.ondrop=e=>{e.preventDefault();if(xdrag){e.stopPropagation();const i=dropIndex(e);hideIns();addFromTemplate(tab,i);return}if(libDrag){e.stopPropagation();const i=dropIndex(e);hideIns();addFromLib(libDrag,tab,i);libDrag=null;return}if(drag)moveGroupToEnd(drag.idxs);drag=null};
  T.buttons.forEach((b,n)=>{const d=document.createElement('div');d.className='tile'+(sel.has(n)?' picked':'');d.draggable=true;d.title=b.items.join('\n');
    const num=document.createElement('span');num.className='n';num.textContent=n+1;d.appendChild(num);
    const im=document.createElement('div');im.className='im';
    if(b.img>=0){const i=document.createElement('img');loadImg(i,b.img,im);ids.push(b.img);im.appendChild(i)}
    else if(b.pend)im.innerHTML='<div class="wait">looking for picture...</div>';
    else im.innerHTML='<div class="no">no image</div>';d.appendChild(im);
    const l=document.createElement('div');l.className='lb';l.textContent=b.name;d.appendChild(l);
    if(b.ex||T.ex){const x=document.createElement('span');x.className='exb';x.textContent='E';x.title=T.ex?'Excluded from fill because its whole tab is excluded':'Excluded from fill';d.appendChild(x)}
    if(b.items&&b.items.length>1){const s=document.createElement('span');s.className='stk';s.textContent=b.items.length;s.title=b.items.length+' parts are stacked on this button';d.appendChild(s)}
    d.onclick=e=>{e.stopPropagation();pickTile(n,e)};
    d.ondragstart=e=>{if(!sel.has(n)){sel.clear();sel.add(n);anchor=n}markSel();drag={idxs:selList()};d.classList.add('drag');try{e.dataTransfer.setData('text/plain',String(n))}catch(x){}};
    d.ondragend=()=>{d.classList.remove('drag');drag=null};
    d.ondragover=e=>{if(drag){e.preventDefault();d.classList.add('over')}else if(libDrag||xdrag)e.preventDefault()};d.ondragleave=()=>d.classList.remove('over');
    d.ondrop=e=>{e.preventDefault();e.stopPropagation();d.classList.remove('over');if(xdrag){const r=d.getBoundingClientRect(),after=r.width>0&&e.clientX>=r.left+r.width/2,ix=after?n+1:n;hideIns();addFromTemplate(tab,ix>=nTiles()?null:ix);return}if(libDrag){const r=d.getBoundingClientRect(),after=r.width>0&&e.clientX>=r.left+r.width/2,ix=after?n+1:n;hideIns();addFromLib(libDrag,tab,ix>=nTiles()?null:ix);libDrag=null;return}if(drag&&!drag.idxs.includes(n))moveGroup(drag.idxs,n);drag=null};
    d.oncontextmenu=e=>{e.preventDefault();if(!sel.has(n)){sel.clear();sel.add(n);anchor=n;markSel()}showCtx(e.clientX,e.clientY,n,b)};g.appendChild(d)});
  if(!T.buttons.length){const e=document.createElement('div');e.className='none';e.style.cssText='padding:14px';e.textContent='This tab has no buttons. Drag parts here from Libraries, or drop buttons on this tab\'s name.';g.appendChild(e)}
  markSel();prio(ids)}
function showCtx(x,y,n,b){const c=$('ctx');c.innerHTML='';const idxs=selList(),many=idxs.length>1,N=nTiles(),last=idxs[idxs.length-1];
  const add=(t,f,h)=>{const d=document.createElement('div');d.textContent=t;if(h)d.className='h';else d.onclick=()=>{hideCtx();f()};c.appendChild(d)};
  add(many?idxs.length+' buttons selected':b.name,null,true);
  add(many?'Delete '+idxs.length+' buttons':'Delete',()=>deleteGroup(idxs));
  if(idxs[0]>0)add('Move to first',()=>many?doOp({op:'move_many',ti:pal.i,tab:tab,idxs:idxs,to:0}):doOp({op:'move',ti:pal.i,tab:tab,from:n,to:0}));
  if(last<N-1)add('Move to last',()=>moveGroupToEnd(idxs));
  const allEx=idxs.every(i=>pal.tabs[tab].buttons[i].ex);
  add(allEx?(many?'Include '+idxs.length+' buttons in fill':'Include in fill'):(many?'Exclude '+idxs.length+' buttons from fill':'Exclude from fill'),()=>doOp({op:'btn_exclude',ti:pal.i,tab:tab,idxs:idxs,value:!allEx}));
  if(!many&&b&&b.items&&b.items.length===1)add('Locate in Library',()=>locateInLibrary(b.items[0]));
  else if(!many&&b&&b.items&&b.items.length>1){add('Locate in Library:',null,true);b.items.forEach(p=>add('   '+p.split('/').pop().replace(/\.itm$/i,''),()=>locateInLibrary(p)))}
  if(many)add('Clear selection',()=>clearSel());
  c.style.left=Math.min(x,innerWidth-200)+'px';c.style.top=Math.min(y,innerHeight-30*c.children.length-10)+'px';c.style.display='block'}
function showTabCtx(x,y,k){const c=$('ctx');c.innerHTML='';const T=pal.tabs[k],lit=/\s*\[E\]\s*$/.test(T.name);
  const add=(t,f,h)=>{const d=document.createElement('div');d.textContent=t;if(h)d.className='h';else d.onclick=()=>{hideCtx();f()};c.appendChild(d)};
  add('Tab: '+tabLabel(T),null,true);
  add(T.ex?'Include tab in fill':'Exclude tab from fill',()=>doOp({op:'tab_exclude',ti:pal.i,tab:k,value:!T.ex}));
  if(lit&&!T.ex){add('Repair: make this a real exclusion (set the mode, remove "[E]" from the name)',()=>doOp({op:'tab_fix',ti:pal.i,tab:k}));
    add('Remove "[E]" from the name only (the tab stays included)',()=>doOp({op:'tab_fix',ti:pal.i,tab:k,exclude:false}))}
  if(pal.tabs.length>1)add('Delete tab'+(T.buttons.length?' ('+T.buttons.length+' button'+(T.buttons.length===1?'':'s')+')':''),()=>doOp({op:'tab_delete',ti:pal.i,tab:k}));
  c.style.left=Math.min(x,innerWidth-340)+'px';c.style.top=Math.min(y,innerHeight-100)+'px';c.style.display='block'}
function hideCtx(){$('ctx').style.display='none'}
document.addEventListener('click',hideCtx);
/* ---------- Libraries fly-out: browse the folder tree or search, then drag a part onto a tab ---------- */
let libDrag=null,libOpen=false,libPollOn=false,lmode='parts',xdrag=null,xView=null,lgen=0;const xCache=new Map();
function setLmode(m){lgen++;lmode=m;$('lmParts').classList.toggle('on',m==='parts');$('lmTpl').classList.toggle('on',m==='tpl');$('lq').value='';xView=null;
  $('lq').placeholder=m==='tpl'?'search templates: any part of a name':'search the libraries: any part of a name or folder';
  if(m==='tpl')xList();else{$('lcount').textContent='';libShowTree()}}
$('lmParts').onclick=()=>setLmode('parts');$('lmTpl').onclick=()=>setLmode('tpl');
/* ---------- Templates mode: pick a template, see its buttons, drag them into the template you are editing ---------- */
function xList(){const q=$('lq').value.trim().toLowerCase(),el=$('llist');xView=null;el.innerHTML='';$('lcount').textContent='';
  if(curP<0||!L.length){el.innerHTML='<div class="none">Pick a profile first.</div>';return}
  const words=q.split(/\s+/).filter(Boolean),rows=L.filter(p=>!q||words.every(w=>p.name.toLowerCase().includes(w))).sort((a,b)=>a.name.toLowerCase().localeCompare(b.name.toLowerCase()));
  rows.slice(0,400).forEach(p=>{const r=document.createElement('div');r.className='lfold';r.innerHTML=esc(p.name)+'<small style="display:block;font-weight:normal;font-size:10px">'+p.tabs+' tabs, '+p.buttons+' buttons'+(pal&&pal.i===p.i?' - the template you are editing':'')+'</small>';r.onclick=()=>xOpen(p.i);el.appendChild(r)});
  if(!rows.length)el.innerHTML='<div class="none">No template matches "'+esc(q)+'".</div>';
  else if(rows.length>400){const n=document.createElement('div');n.className='none';n.textContent='Showing 400 of '+rows.length+' templates - type to narrow the list.';el.appendChild(n)}}
function xOpen(ti,keepTab){const have=xCache.get(ti);
  if(have){xView={ti:ti,tab:Math.min(keepTab||0,have.tabs.length-1),sel:new Set(),anchor:-1};xDraw();xPend(ti,0);return}
  $('llist').innerHTML='<div class="none">Opening...</div>';
  api('/api/palette?i='+ti+'&preview=1').then(d=>{if(lmode!=='tpl')return;if(!d.ok){stat(d.error,true);xList();return}
    xCache.set(ti,d.palette);xView={ti:ti,tab:0,sel:new Set(),anchor:-1};xDraw();xPend(ti,0)}).catch(()=>{stat('No reply from the editor.',true);xList()})}
function xPend(ti,n){const p=xCache.get(ti);if(!p||!xView||xView.ti!==ti||n>10)return;
  if(!p.tabs.some(t=>t.buttons.some(b=>b.pend)))return;
  setTimeout(()=>{if(!xView||xView.ti!==ti||lmode!=='tpl'||libDrag||xdrag)return;
    api('/api/palette?i='+ti+'&preview=1').then(d=>{if(!d.ok||!xView||xView.ti!==ti||xdrag)return;xCache.set(ti,d.palette);xDraw();xPend(ti,n+1)}).catch(()=>{})},1500)}
function xDraw(){const v=xView,p=v&&xCache.get(v.ti),el=$('llist');if(!p){xList();return}el.innerHTML='';
  const back=document.createElement('div');back.className='lfold';back.textContent='< back to the template list';back.onclick=()=>{xView=null;xList()};el.appendChild(back);
  const h=document.createElement('div');h.className='xhd';const nm=document.createElement('b');nm.textContent=p.name;h.appendChild(nm);
  const ed=document.createElement('button');ed.textContent=pal&&pal.i===v.ti?'(editing it now)':'Edit this template';ed.disabled=!!(pal&&pal.i===v.ti);ed.onclick=()=>openTemplate(v.ti);h.appendChild(ed);el.appendChild(h);
  const tb=document.createElement('div');tb.className='xtabs';
  p.tabs.forEach((t,k)=>{const b=document.createElement('button');b.textContent=tabLabel(t)+' ('+t.buttons.length+')';if(k===v.tab)b.className='on';b.onclick=()=>{v.tab=k;v.sel.clear();v.anchor=-1;xDraw()};tb.appendChild(b)});el.appendChild(tb);
  const nt=document.createElement('div');nt.className='xnote';nt.textContent=pal?'Click a button (Ctrl+click for several, Shift+click for a range), then drag it onto the template on the right or onto one of its tab names.':'Open a template on the right first.';el.appendChild(nt);
  const g=document.createElement('div');g.className='xgrid';const T=p.tabs[v.tab]||{buttons:[]};
  T.buttons.forEach((b,n)=>{const d=document.createElement('div');d.className='xtile'+(v.sel.has(n)?' picked':'');d.draggable=true;d.title=b.items.join('\n');
    const im=document.createElement('div');im.className='xim';if(b.img>=0){const i=document.createElement('img');loadImg(i,b.img,im);im.appendChild(i)}else im.textContent=b.pend?'...':'';d.appendChild(im);
    const l=document.createElement('div');l.className='xl';l.textContent=b.name;d.appendChild(l);
    if(b.items&&b.items.length>1){const s=document.createElement('span');s.className='stk';s.textContent=b.items.length;s.title=b.items.length+' parts are stacked on this button';d.appendChild(s)}
    d.onclick=e=>{e.stopPropagation();if(e.ctrlKey||e.metaKey){if(v.sel.has(n))v.sel.delete(n);else v.sel.add(n);v.anchor=n}
      else if(e.shiftKey&&v.anchor>=0){v.sel.clear();for(let i=Math.min(v.anchor,n);i<=Math.max(v.anchor,n);i++)v.sel.add(i)}else{v.sel.clear();v.sel.add(n);v.anchor=n}xMark()};
    d.ondragstart=e=>{if(!v.sel.has(n)){v.sel.clear();v.sel.add(n);v.anchor=n}xMark();xdrag={ti:v.ti,tab:v.tab,idxs:[...v.sel].sort((a,b)=>a-b)};drag=null;libDrag=null;try{e.dataTransfer.setData('text/plain','tpl:'+v.ti)}catch(x){}};
    d.ondragend=()=>{xdrag=null;hideIns()};g.appendChild(d)});
  if(!T.buttons.length){const e=document.createElement('div');e.className='none';e.textContent='This tab has no buttons.';g.appendChild(e)}
  el.appendChild(g);$('lcount').textContent=v.sel.size?v.sel.size+' selected':''}
function xMark(){const v=xView;if(!v)return;document.querySelectorAll('#llist .xtile').forEach((e,i)=>e.classList.toggle('picked',v.sel.has(i)));$('lcount').textContent=v.sel.size?v.sel.size+' selected':''}
/* drop: the buttons are copied into the template being edited (another template: the editor remaps the size restrictions) */
function addFromTemplate(k,pos){const x=xdrag;xdrag=null;if(!x||!pal)return;const sp=xCache.get(x.ti);if(!sp||!sp.tabs[x.tab])return;const T=sp.tabs[x.tab];
  const have=new Set(pal.tabs[k].buttons.map(b=>b.items.join('|')));const keep=x.idxs.filter(i=>T.buttons[i]&&!have.has(T.buttons[i].items.join('|')));
  if(!keep.length){msgSeq++;stat(x.idxs.length>1?'Those buttons are already in this tab.':'That button is already in this tab.',true);return}
  if(x.ti===pal.i)return doOp({op:'copy_many',ti:pal.i,tab:x.tab,idxs:keep,dtab:k,pos:pos});
  return doExt({op:'copy_ext',ti:pal.i,tab:k,src_ti:x.ti,src_tab:x.tab,idxs:keep,pos:pos})}
function doExt(b){if(!pal||barrier)return Promise.resolve();ustack.length=0;refreshUndo();msgSeq++;stat('Copying from another template...');
  return whenIdle().then(()=>post('/api/op',Object.assign({rev:rev,echo:0},b))).then(d=>{if(!d.ok){return resyncHard(d.error)}
    serverCanUndo=true;refreshUndo();sel.clear();anchor=-1;xCache.delete(b.ti);
    return api('/api/palette?i='+b.ti).then(r=>{if(r.ok&&pal&&pal.i===b.ti){pal=r.palette;paintLocal()}stat(d.msg,false,true);if(xView&&xView.ti===b.ti)xOpen(b.ti,xView.tab)})}).catch(()=>stat('No reply from the editor.',true))}
function toggleLib(){libOpen=!libOpen;$('lib').style.display=libOpen?'flex':'none';$('libbtn').className=libOpen?'on':'';try{localStorage.setItem('fablib',libOpen?'1':'0')}catch(e){}
  if(libOpen&&!$('llist').children.length)(lmode==='tpl'?xList:libShowTree)()}
$('libbtn').onclick=toggleLib;
function libRow(it,showFolder){const r=document.createElement('div');r.className='lrow';r.draggable=true;r.title=it.rel;r._it=it;
  const im=document.createElement('div');im.className='lim';if(it.img>=0){const i=document.createElement('img');loadImg(i,it.img,im);im.appendChild(i)}r.appendChild(im);
  const tx=document.createElement('div');tx.className='ltx';
  tx.innerHTML='<b>'+esc(it.name)+'</b>'+(showFolder?'<small>'+esc((it.folder||'').replace(/^\.\//,''))+'</small>':'');r.appendChild(tx);
  r.ondragstart=e=>{libDrag=it;drag=null;try{e.dataTransfer.setData('text/plain','lib:'+it.rel)}catch(x){}};r.ondragend=()=>{libDrag=null;hideIns()};return r}
/* which listed parts does this file use? (the editor prepares that list in the background: ask, and fill the badges in when it is ready) */
let usedT=null;
function libResolveUsed(){clearTimeout(usedT);const pend=[...document.querySelectorAll('#llist .lrow')].filter(r=>r._it&&r._it.used==null);if(!pend.length)return;
  post('/api/lib/used',{rels:pend.map(r=>r._it.rel)}).then(d=>{if(!d.ok)return;pend.forEach(r=>{const v=d.used[r._it.rel];if(v!==null&&v!==undefined)r._it.used=v});
    if(!d.ready)usedT=setTimeout(libResolveUsed,600)})}
function libFolder(rel,name,host,depth){const row=document.createElement('div');row.className='lfold';row.style.paddingLeft=(6+depth*14)+'px';row.textContent='[+] '+name;
  const box=document.createElement('div');let open=false,loaded=false;box.style.display='none';
  row.onclick=()=>{open=!open;row.textContent=(open?'[-] ':'[+] ')+name;box.style.display=open?'block':'none';if(open&&!loaded){loaded=true;libLoad(rel,box,depth+1)}};
  host.appendChild(row);host.appendChild(box)}
function libLoad(rel,host,depth){const g=lgen;host.innerHTML='<div class="none">loading...</div>';const q=encodeURIComponent(rel);
  api('/api/lib/folders?p='+q).then(f=>{if(g!==lgen)return;host.innerHTML='';if(!f.ok){host.innerHTML='<div class="none">'+esc(f.error)+'</div>';return}
    const have=new Set(f.folders.map(x=>x.toLowerCase()));f.folders.forEach(n=>libFolder(f.prefix+n,n,host,depth));
    const ih=document.createElement('div');host.appendChild(ih);
    const ask=()=>api('/api/lib/children?p='+q).then(d=>{if(g!==lgen)return;if(!d.ok){ih.innerHTML='<div class="none">'+esc(d.error)+'</div>';return}
      if(d.pending){ih.innerHTML='<div class="none" style="padding-left:'+(6+depth*14)+'px">'+(d.seconds>=3?'still reading this folder from the drive ('+d.seconds+' s)...':'reading this folder...')+'</div>';return new Promise(r=>setTimeout(r,500)).then(ask)}
      ih.innerHTML='';(d.extra_folders||[]).forEach(n=>{if(!have.has(n.toLowerCase()))libFolder(d.prefix+n,n,ih,depth)});
      d.items.forEach(it=>{const r=libRow(it,false);r.style.paddingLeft=(6+depth*14)+'px';ih.appendChild(r)});
      if(d.more)ih.insertAdjacentHTML('beforeend','<div class="none">more parts here than can be listed - use the search box</div>');
      if(!f.folders.length&&!d.items.length&&!(d.extra_folders||[]).length)host.innerHTML='<div class="none" style="padding-left:'+(6+depth*14)+'px">(empty)</div>';
      libResolveUsed()});
    return ask()}).catch(()=>{})}
function libNote(d){$('lcount').textContent=d.listing?('listing '+d.listing.done.toLocaleString()+' of '+d.listing.total.toLocaleString()+' folders...'):(d.refreshing?'refreshing...':'')}
function libShowTree(){$('lcount').textContent='';libLoad('',$('llist'),0)}
function libShowFolder(rel){const el=$('llist');el.innerHTML='';const back=document.createElement('div');back.className='lfold';back.textContent='< back to the folder list';back.onclick=()=>{$('lq').value='';libShowTree()};el.appendChild(back);
  const h=document.createElement('div');h.className='none';h.textContent=rel.replace(/^\.\//,'');el.appendChild(h);const box=document.createElement('div');el.appendChild(box);libLoad(rel,box,0)}
function locateInLibrary(item){const k=item.lastIndexOf('/');if(k<0)return;const folder=item.slice(0,k);
  if(!libOpen)toggleLib();if(lmode!=='parts')setLmode('parts');lgen++;$('lq').value='';libShowFolder(folder);   // lgen++: a tree still loading must not overwrite the folder
  const g=lgen,t0=Date.now(),want=item.toLowerCase();
  const find=()=>{if(g!==lgen||!libOpen)return;
    const r=[...document.querySelectorAll('#llist .lrow')].find(x=>x._it&&x._it.rel&&x._it.rel.toLowerCase()===want);
    if(r){r.classList.add('located');if(r.scrollIntoView)r.scrollIntoView({block:'center'});return}
    if(Date.now()-t0<20000)setTimeout(find,150)};find()}
function libPoll(){api('/api/lib/search?q=zzzzzzzz').then(d=>{if(!d.ok)return;libNote(d);if(d.refreshing||d.listing)setTimeout(libPoll,1500);else{$('lcount').textContent='';if($('lq').value.trim())libSearch();else libShowTree()}})}
let lsT=null;$('lq').oninput=()=>{clearTimeout(lsT);lsT=setTimeout(()=>lmode==='tpl'?xList():libSearch(),200)};
function libSearch(){const q=$('lq').value.trim(),el=$('llist');if(!q){libShowTree();return}
  api('/api/lib/search?q='+encodeURIComponent(q)).then(d=>{if($('lq').value.trim()!==q||lmode!=='parts')return;if(!d.ok){el.innerHTML='<div class="none">'+esc(d.error)+'</div>';return}
    const again=()=>{if(d.listing||d.refreshing)setTimeout(()=>{if($('lq').value.trim()===q&&!libDrag)libSearch()},1500)};
    if(libDrag){again();return}
    el.innerHTML='';libNote(d);
    (d.folders||[]).forEach(f=>{const r=document.createElement('div');r.className='lfold';r.title=f.where;r.innerHTML='[+] '+esc(f.name)+'<small style="display:block;font-weight:normal;font-size:10px">'+esc(f.where)+'</small>';r.onclick=()=>libShowFolder(f.rel);el.appendChild(r)});
    d.results.forEach(it=>el.appendChild(libRow(it,true)));
    if(!d.folders.length&&!d.results.length)el.innerHTML='<div class="none">Nothing named "'+esc(q)+'" '+(d.complete?'in the Libraries.':'in the folder names or in the folders you have opened.')+'</div>';
    if(!d.complete&&!d.listing){const n=document.createElement('div');n.className='none';n.textContent='Part names are only searched in folders you have opened.';el.appendChild(n)}
    libResolveUsed();again()})}
function gridTiles(){return [...document.querySelectorAll('#grid .tile')]}
/* where would a dropped part land? in front of the first button the pointer is before; past the last one = the end (null) */
function dropIndex(e){const ts=gridTiles();for(let i=0;i<ts.length;i++){const r=ts[i].getBoundingClientRect();
  if(e.clientY<r.top)return i;if(e.clientY<=r.bottom&&e.clientX<r.left+r.width/2)return i}return null}
let insAt=-2;
function clearIns(){document.querySelectorAll('#grid .ins1,#grid .ins2').forEach(x=>x.classList.remove('ins1','ins2'))}
function hideIns(){insAt=-2;clearIns()}
function showIns(i){if(i===insAt)return;insAt=i;clearIns();const ts=gridTiles();if(!ts.length)return;if(i===null)ts[ts.length-1].classList.add('ins2');else ts[i].classList.add('ins1')}
$('main').ondragover=e=>{if(!libDrag&&!xdrag)return;e.preventDefault();showIns(dropIndex(e))};
$('main').ondragleave=e=>{if(!e.relatedTarget)hideIns()};
$('main').ondrop=e=>{if(xdrag){e.preventDefault();const i=dropIndex(e);hideIns();addFromTemplate(tab,i);return}if(!libDrag)return;e.preventDefault();const i=dropIndex(e);hideIns();addFromLib(libDrag,tab,i);libDrag=null};
function addFromLib(it,k,pos){const item=it.rel;
  if(hasItem(k,item)){msgSeq++;stat('That part is already in this tab.',true);return}
  if(it.used==null){   // not known yet: ask the editor, wait for the answer, then finish the drop by itself (never guess)
    const mine=++msgSeq;stat('Checking whether this file uses "'+it.name+'"...');let tries=0;
    const ask=()=>post('/api/lib/used',{rels:[item]}).then(d=>{if(mine!==msgSeq)return;
      if(d.ok&&d.used[item]!=null){it.used=d.used[item];addFromLib(it,k,pos)}else if(++tries<30)setTimeout(ask,400);else stat('Could not find out which parts this file uses.',true)});
    return ask()}
  return doOp({op:'add',ti:pal.i,tab:k,item:item,pos:pos,_name:it.name,_img:it.img})}
/* ---------- sidebar width ---------- */
(function(){const sp=$('split'),lf=$('left');let on=false;try{const w=localStorage.getItem('fabw');if(w)lf.style.width=w}catch(e){}
  sp.onmousedown=e=>{on=true;e.preventDefault()};document.addEventListener('mousemove',e=>{if(on)lf.style.width=Math.max(240,Math.min(innerWidth*0.7,e.clientX))+'px'});
  document.addEventListener('mouseup',()=>{if(on){on=false;try{localStorage.setItem('fabw',lf.style.width)}catch(e){}}})})();
function bootGate(cb){const ov=$('boot');let started=false;
  function tick(){fetch('/api/boot?k='+K).then(r=>r.json()).then(d=>{
    if(d.ready){ov.style.display='none';if(!started){started=true;cb()}return}
    $('bootst').textContent=d.state==='error'?'The editor could not start':(d.status||'Starting')+(d.step_seconds>=2?'   ('+d.step_seconds+' s)':'');
    $('bootsteps').textContent=(d.steps||[]).map(s=>'done: '+s.what+' - '+s.seconds+' s').join('\n');
    $('boothint').textContent=d.state==='starting'&&d.step_seconds>=6?'This step is waiting on a drive or network folder. The editor itself is running; it carries on as soon as the drive answers.':'';
    const er=$('booterr');if(d.state==='error'){er.style.display='block';er.textContent=d.error+(d.db?'\n\nDatabase folder: '+d.db:'')}else er.style.display='none';
    $('bootbtns').style.display=d.state==='error'?'block':'none';setTimeout(tick,d.state==='error'?1500:500)}).catch(()=>setTimeout(tick,1000))}
  $('bootretry').onclick=()=>{fetch('/api/boot/retry?k='+K)};$('bootchoose').onclick=()=>{fetch('/api/boot/choose?k='+K).then(r=>r.json()).then(d=>{if(d&&d.need_path){const p=prompt('Type or paste the Fabrication database folder (the one that contains MAP.INI):','');if(p)fetch('/api/boot/choose?k='+K+'&path='+encodeURIComponent(p))}})};tick()}
function startUp(){drawServices();refreshSlow();
api('/api/profiles').then(d=>{if(!d.ok){stat(d.error,true);return}P=d.profiles;drawProfiles();stat(P.length?'Pick a profile':'No profiles found (profiles.map missing?)',!P.length);
  if(d.last){const i=P.findIndex(p=>p.name===d.last);if(i>=0){openProfile(i);const r=document.querySelector('#plist .row[data-i="'+i+'"]');if(r&&r.scrollIntoView)r.scrollIntoView({block:'center'})}}})}
bootGate(startUp);
</script></body></html>"""


class _Handler(http.server.BaseHTTPRequestHandler):
    protocol_version = 'HTTP/1.1'      # keep connections open: many small picture requests reuse them
    timeout = 120

    def setup(self):
        http.server.BaseHTTPRequestHandler.setup(self)
        try:
            self.request.setsockopt(socket.IPPROTO_TCP, socket.TCP_NODELAY, 1)     # headers and body go out without a 40-200 ms wait
        except Exception:
            pass
    session = None
    token = ''
    server_ref = None

    def log_message(self, *a):
        pass

    def _send(self, code, body, ctype='application/json', cache='no-store'):
        if not isinstance(body, bytes):
            body = json.dumps(body).encode('utf-8')
        try:
            self.send_response(code)
            self.send_header('Content-Type', ctype)
            self.send_header('Content-Length', str(len(body)))
            self.send_header('Cache-Control', cache)
            self.end_headers()
            self.wfile.write(body)
        except (ConnectionError, OSError):
            self.close_connection = True        # the browser hung up (page reloaded, request cancelled); nothing to do

    def _route(self, method):
        type(self).last_request = time.time()
        u = urllib.parse.urlparse(self.path)
        q = urllib.parse.parse_qs(u.query)
        body = {}
        if method == 'POST':                       # always read the body first so the connection can be reused
            n = int(self.headers.get('Content-Length') or 0)
            raw_body = self.rfile.read(n) if n else b''
            try:
                body = json.loads(raw_body or b'{}')
            except Exception:
                body = {}
        if q.get('k', [''])[0] != self.token:
            return self._send(403, {'ok': False, 'error': 'forbidden'})
        s = self.session
        try:
            p = u.path
            if p == '/':
                return self._send(200, EDITOR_HTML.replace('__TOKEN__', self.token).encode('utf-8'), 'text/html; charset=utf-8')
            if p == '/api/ping':
                return self._send(200, self.ping())
            if p == '/api/boot':
                return self._send(200, self.boot_state())
            if p == '/api/boot/retry':
                return self._send(200, self.boot_retry(False))
            if p == '/api/boot/choose':
                typed = q.get('path', [''])[0].strip().strip('"')
                if not typed and getattr(type(self), 'mode', 'revit') == 'process':
                    return self._send(200, {'ok': True, 'need_path': True})      # no dialog in the separate process: the page asks for the path
                return self._send(200, self.boot_retry(True, typed))
            if s is None:
                return self._send(200, {'ok': False, 'booting': True, 'error': 'The editor is still starting.'})
            if p == '/api/busy':
                return self._send(200, {'ok': True, 'busy': s.busy_list(), 'slow': s.slow_log[-8:], 'damaged': s.health.get('damaged', 0), 'lost': s.health.get('lost', 0), 'lost_names': s.health.get('lost_names', [])})
            if p == '/api/profiles':
                return self._send(200, {'ok': True, 'profiles': s.profiles, 'last': s.cfg.get('last_profile', ''), 'map': s.pmap})
            if p == '/api/open':
                return self._send(200, dict(ok=True, **s.open_profile(int(body['idx']))))
            if p == '/api/services':
                return self._send(200, dict({'ok': True, 'services': s.services(), 'palettes': s.palettes(), 'rev': s.rev}, **s.state()))
            if p == '/api/rescan':
                return self._send(200, dict(ok=True, **s.rescan()))
            if p == '/api/samples/start':
                return self._send(200, dict(ok=True, **s.samples_start()))
            if p == '/api/samples/status':
                return self._send(200, dict(ok=True, **s.samples_status()))
            if p == '/api/lib/folders':
                return self._send(200, dict(ok=True, **s.lib_folders(q.get('p', [''])[0])))
            if p == '/api/lib/children':
                return self._send(200, dict(ok=True, **s.lib_children(q.get('p', [''])[0])))
            if p == '/api/lib/refresh':
                return self._send(200, dict(ok=True, **s.lib_refresh()))
            if p == '/api/lib/used':
                return self._send(200, dict(ok=True, **s.lib_used([str(x) for x in body.get('rels', [])][:400])))
            if p == '/api/lib/search':
                return self._send(200, dict(ok=True, **s.lib_search(q.get('q', [''])[0])))
            if p == '/api/imgstatus':
                return self._send(200, dict(ok=True, **s.image_status(int(q.get('ti', ['-1'])[0]))))
            if p == '/api/palettes':
                return self._send(200, dict({'ok': True, 'palettes': s.palettes(), 'rev': s.rev}, **s.state()))
            if p == '/api/palette':
                return self._send(200, dict({'ok': True, 'palette': s.palette(int(q['i'][0]), preview=('preview' in q)), 'rev': s.rev}, **s.state()))
            if p == '/api/op':
                return self._send(200, dict(ok=True, **s.apply(body)))
            if p == '/api/undo':
                return self._send(200, dict(ok=True, **s.undo()))
            if p == '/api/apply':
                return self._send(200, dict(ok=True, **s.save()))
            if p == '/api/discard':
                return self._send(200, dict(ok=True, **s.discard(int(body.get('ti', -1)))))
            if p == '/api/search':
                return self._send(200, dict({'ok': True}, **s.search(q.get('q', [''])[0])))
            if p.startswith('/img/'):
                path = s.img_by_id[int(p[5:])]
                with s._busy('Reading the picture %s' % os.path.basename(path)):
                    how, data = _read_with_deadline(path, _IMG_DEADLINE)
                if how == 'slow':
                    return self._send(504, {'ok': False, 'error': 'the drive did not answer in time'})
                if how != 'ok':
                    return self._send(404, {'ok': False, 'error': 'no image'})
                return self._send(200, data, _IMG_TYPES.get(os.path.splitext(path)[1].lower(), 'application/octet-stream'), 'private, max-age=86400')
            if p == '/api/quit':
                self._send(200, {'ok': True})
                threading.Thread(target=self.server_ref.shutdown).start()
                return
            return self._send(404, {'ok': False, 'error': 'not found'})
        except _Stale as ex:
            fresh = None
            try:
                fresh = s.palette(int(body['ti']))
            except Exception:
                pass
            return self._send(200, dict({'ok': False, 'stale': True, 'error': str(ex.args[0]), 'rev': s.rev, 'palette': fresh}, **s.state()))
        except (ValueError, KeyError, IndexError) as ex:
            return self._send(200, {'ok': False, 'error': str(ex.args[0]) if ex.args else str(ex)})
        except Exception as ex:
            return self._send(200, {'ok': False, 'error': '%s: %s' % (type(ex).__name__, ex)})

    def do_GET(self):
        self._route('GET')

    def do_POST(self):
        self._route('POST')


class _Server(http.server.ThreadingHTTPServer):
    daemon_threads = True
    request_queue_size = 128           # the default of 5 can refuse a burst of picture requests on Windows

    def server_bind(self):
        # the standard server asks DNS for its own name here (socket.getfqdn), which can stall for minutes on a bad network
        import socketserver
        socketserver.TCPServer.server_bind(self)
        host, port = self.server_address[:2]
        self.server_name, self.server_port = host, port

    def handle_error(self, request, client_address):
        ex = sys.exc_info()[1]
        if isinstance(ex, (ConnectionError, TimeoutError, OSError)):
            return              # dropped connection: harmless
        _log('editor: unexpected error: %s: %s' % (type(ex).__name__, ex))


_LIB_THREADS = 6
_FG = {'n': 0}                           # how many things you are waiting for on the drive right now


def _yield_to_user():
    """Background workers (finding pictures, 'List all parts') call this before each piece of drive work: while you are waiting
    for something on the drive they stand aside (at most 20 s at a time, so they cannot starve)."""
    nm = threading.current_thread().name
    if _FG['n'] > 0 and (nm.startswith('fabmap-pictures') or nm.startswith('fabmap-lib-') or nm.startswith('fabmap-recheck')):
        t0 = time.time()
        while _FG['n'] > 0 and time.time() - t0 < 20:
            time.sleep(0.05)
_IMG_DEADLINE = 8.0


def _read_with_deadline(path, limit):
    """Read a small file, but give up waiting after 'limit' seconds (a hung drive then cannot hold a browser connection).
    Returns ('ok', bytes), ('missing', None) or ('slow', None)."""
    box = {}

    def work():
        try:
            with open(path, 'rb') as f:
                box['data'] = f.read()
        except OSError:
            box['gone'] = True
    th = threading.Thread(target=work, name='fabmap-img')
    th.daemon = True
    th.start()
    th.join(limit)
    if th.is_alive():
        return 'slow', None
    return ('ok', box['data']) if 'data' in box else ('missing', None)


def _part_facts(blk):
    """What a button stores that depends on the part: two ids in the item header, the byte in each size record, the date stamp
    and the fingerprint block. Read-only helper for the sample collection."""
    try:
        m = _E0.search(blk, 8)
        top = m.start()
        sz = U(blk, top + 4)
        inner = _E0.search(blk, top + 8, top + 8 + sz)
        rs = inner.start()
        labels, rowbyte, ndim, o = [], None, 1, rs
        while o < top + 8 + sz - 16:
            rsz = U(blk, o + 4)
            row = blk[o:o + 8 + rsz]
            n = U(row, 110)
            lab = row[114:114 + 2 * (n - 1)].decode('utf-16le')
            labels.append(lab)
            rowbyte = row[114 + 2 * n + 78 - 4]
            ndim = max(ndim, lab.count('x') + 1)
            o += 8 + rsz
        return {'a': U(blk, rs - 148), 'b': U(blk, rs - 144), 'cid': U(blk, rs - 140), 'ts': blk[rs - 136:rs - 120].hex(),
                'block8': blk[rs - 54:rs - 46].hex(), 'rowbyte': rowbyte, 'ndim': ndim, 'labels': labels}
    except Exception:
        return None
_ITM16 = re.compile(b'\\.\\x00/\\x00(?:[^\\x00]\\x00)+?\\.\\x00[iI]\\x00[tT]\\x00[mM]\\x00')


def _log(*a):
    """Console text is off (pyRevit would show it in its output window). Set FABMAP_VERBOSE=1 to see it."""
    import os
    if os.environ.get('FABMAP_VERBOSE'):
        print(*a)


def start_server(session, port=0, boot=None):
    """Start listening at once. 'session' may be None: the editor is then still starting (see boot_run) and the page shows
    its start-up card until the session is attached."""
    token = secrets.token_hex(8)
    boot = boot if boot is not None else {'state': 'ready' if session is not None else 'starting', 'status': '', 'error': '', 'db': '',
                                          't0': time.time(), 't1': time.time(), 'steps': []}
    handler = type('H', (_Handler,), {'session': session, 'token': token, 'boot': boot})
    srv = _Server(('127.0.0.1', port), handler)
    handler.server_ref = srv
    return srv, token


def _boot_state(self):
    b = self.boot
    now = time.time()
    return {'ok': True, 'ready': b['state'] == 'ready' and self.session is not None, 'state': b['state'], 'status': b['status'],
            'error': b['error'], 'db': b.get('db', ''), 'step_seconds': int(now - b['t1']), 'seconds': int(now - b['t0']),
            'steps': [s for s in b.get('steps', []) if s['seconds'] >= 1][-6:]}


def boot_run(handler, picker=None):
    """Do the slow work (settings, MAP.INI, the folders on your drive, the profile list) in the background, reporting each step.
    Nothing here can open a dialog by itself: a missing database folder is reported, with buttons to retry or choose another."""
    b = handler.boot

    def step(text):
        now = time.time()
        if b.get('status') and b['state'] == 'starting':
            b['steps'].append({'what': b['status'], 'seconds': round(now - b['t1'], 1)})
        b['status'], b['t1'] = text, now
    b.update(state='starting', error='', t0=time.time(), t1=time.time(), steps=[], status='')
    try:
        step('Reading the settings file')
        env = resolve_environment(pick_folder=None, progress=step)
        b['db'] = env.get('db', '')
        if not env['ini']:
            b.update(state='error', error=(env['problems'][0] if env['problems'] else 'MAP.INI was not found.'))
            return
        cfg = _read_json(env['json'])
        sess = Session(env, cfg, env['json'], progress=step)
        step('Ready')
        for st in b['steps']:
            if st['seconds'] >= 3:
                sess.slow_log.append({'what': 'Start-up: ' + st['what'], 'seconds': int(round(st['seconds'])), 'at': time.strftime('%H:%M:%S')})
        handler.session = sess
        b.update(state='ready', status='Ready')
    except Exception as ex:
        b.update(state='error', error='%s: %s' % (type(ex).__name__, ex))


def _boot_retry(self, choose, typed=''):
    b = self.boot
    cls = type(self)                      # the handler class: the session is attached there, not to this one request
    if b['state'] == 'starting':
        return {'ok': True, 'again': False}
    if choose:
        def pick_then_boot():
            try:
                folder = typed or _pick_folder('Select the Fabrication database folder (the one that contains MAP.INI)')
                if folder:
                    j = RIBBON_JSON
                    cfg = _read_json(j)
                    cfg['database_path'] = folder
                    _write_json(j, cfg)
            except Exception:
                pass
            boot_run(cls)
        threading.Thread(target=pick_then_boot, name='fabmap-pick', daemon=True).start()
    else:
        threading.Thread(target=boot_run, args=(cls,), name='fabmap-boot', daemon=True).start()
    return {'ok': True, 'again': True}


_Handler.boot_state = _boot_state
_Handler.boot_retry = _boot_retry


def _build_id():
    """A short fingerprint of this script, so a newer version can replace an older running one."""
    try:
        import hashlib
        with open(globals().get('__file__') or sys.argv[0], 'rb') as f:
            return hashlib.md5(f.read()).hexdigest()[:10]
    except Exception:
        return ''


def _run_file():
    import os
    return os.path.join(os.path.dirname(RIBBON_JSON), 'fabmap_editor.run')


def _running_editor():
    """(url, ping answer) of an editor that is already running on this computer, else None."""
    import urllib.request
    try:
        with open(_run_file(), 'r') as f:
            info = json.load(f)
        url = 'http://127.0.0.1:%d/?k=%s' % (int(info['port']), info['token'])
        with urllib.request.urlopen('http://127.0.0.1:%d/api/ping?k=%s' % (int(info['port']), info['token']), timeout=2) as r:
            d = json.loads(r.read())
        return (url, d) if d.get('ok') else None
    except Exception:
        return None


def _ping(self):
    s = self.session
    return {'ok': True, 'build': getattr(type(self), 'build', ''), 'mode': getattr(type(self), 'mode', 'revit'), 'ready': s is not None, 'dirty': bool(s and s.state().get('dirty')), 'pid': __import__('os').getpid()}


_Handler.ping = _ping


def _find_python(cfg_path=''):
    """A Python 3 that can run this script on its own (it needs only the standard library). Looks beside the interpreter this
    script is running on (pyRevit's CPython engine folder normally has one), then the usual places; every candidate is tried
    for real, so the Microsoft Store stub is never picked. The one that works is remembered in the settings file."""
    import os, shutil, subprocess
    cfg = _read_json(cfg_path or RIBBON_JSON)
    cands = [cfg.get('python_path') or cfg.get('python_exe') or '']
    try:                                           # ask pyRevit where its bundled CPython is (bin\cengines\CPY3xx\python.exe)
        from pyrevit import BIN_DIR
        root = os.path.dirname(BIN_DIR)
        for eng in (os.path.join(root, 'cengines'), os.path.join(root, 'engines'), os.path.join(BIN_DIR, 'cengines'), os.path.join(BIN_DIR, 'engines')):
            if os.path.isdir(eng):
                for folder in sorted(os.listdir(eng), reverse=True):
                    if folder.upper().startswith('CPY'):
                        cands += [os.path.join(eng, folder, 'python.exe'), os.path.join(eng, folder, 'pythonw.exe')]
    except Exception:
        pass
    try:
        import ctypes
        buf = ctypes.create_unicode_buffer(520)
        ctypes.windll.kernel32.GetModuleFileNameW(ctypes.c_void_p(sys.dllhandle), buf, 520)
        d = os.path.dirname(buf.value)
        cands += [os.path.join(d, 'python.exe'), os.path.join(d, 'pythonw.exe')]
    except Exception:
        pass
    for base in (getattr(sys, 'base_prefix', ''), sys.prefix, os.path.dirname(sys.executable)):
        if base:
            cands += [os.path.join(base, 'python.exe'), os.path.join(base, 'pythonw.exe')]
    if 'revit' not in os.path.basename(sys.executable).lower():
        cands.append(sys.executable)
    for n in ('pythonw', 'python', 'python3'):
        w = shutil.which(n)
        if w:
            cands.append(w)
    for c in cands:
        if not c or not os.path.isfile(c):
            continue
        try:
            flags = 0x08000000 if os.name == 'nt' else 0
            r = subprocess.run([c, '-c', 'import http.server, json, zlib, struct, threading; print("fabmap-ok")'], capture_output=True,
                               timeout=15, creationflags=flags)
            if b'fabmap-ok' in (r.stdout or b''):
                if (cfg.get('python_exe') or '') != c and cfg_path != 'nosave':
                    try:
                        cfg['python_exe'] = c
                        _write_json(cfg_path or RIBBON_JSON, cfg)
                    except Exception:
                        pass
                return c
        except Exception:
            continue
    return ''


def _start_listening(mode):
    """Start the web server, record where it is (so the next press of the button finds it) and return what the caller needs."""
    import os
    srv, token = start_server(None)
    handler = srv.RequestHandlerClass
    handler.build, handler.mode, handler.last_request = _build_id(), mode, time.time()
    url = 'http://127.0.0.1:%d/?k=%s' % (srv.server_address[1], token)
    try:
        os.makedirs(os.path.dirname(_run_file()), exist_ok=True)
        with open(_run_file(), 'w') as f:
            json.dump({'port': srv.server_address[1], 'token': token, 'pid': os.getpid()}, f)
    except Exception:
        pass
    return srv, token, url, handler


def serve_main():
    """The editor as its own process (started by the button inside Revit). Nothing in Revit's Python can stall it, and it
    ends by itself after a long idle time when nothing is unsaved."""
    import os
    srv, token, url, handler = _start_listening('process')
    threading.Thread(target=boot_run, args=(handler,), name='fabmap-boot', daemon=True).start()
    limit = float(os.environ.get('FABMAP_IDLE_SECONDS') or 6 * 3600)

    def watchdog():
        while True:
            time.sleep(min(30.0, max(0.5, limit / 4)))
            s = handler.session
            if time.time() - handler.last_request > limit and not (s and s.state().get('dirty')):
                srv.shutdown()
                return
    threading.Thread(target=watchdog, name='fabmap-idle', daemon=True).start()
    try:
        srv.serve_forever()
    finally:
        try:
            with open(_run_file(), 'r') as f:
                if json.load(f).get('pid') == os.getpid():
                    os.remove(_run_file())
        except Exception:
            pass


def _spawn_editor(exe, script):
    """Start serve_main in its own process; returns the page address once it answers, else ''."""
    import os, subprocess
    err = os.path.join(os.path.dirname(_run_file()), 'fabmap_editor.err')
    try:
        os.makedirs(os.path.dirname(err), exist_ok=True)
        ef = open(err, 'wb')
    except Exception:
        ef = subprocess.DEVNULL
    flags = (0x08000000 | 0x00000008) if os.name == 'nt' else 0          # no console window, detached from Revit
    subprocess.Popen([exe, script, '--serve'], stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=ef, creationflags=flags,
                     close_fds=True, cwd=os.path.dirname(os.path.abspath(script)))
    t0 = time.time()
    while time.time() - t0 < 25:
        found = _running_editor()
        if found:
            return found[0]
        time.sleep(0.2)
    return ''


def editor_main(wait=True, separate=None):
    import os, webbrowser
    g = globals()
    in_revit = 'revit' in os.path.basename(sys.executable).lower()
    if bool(g.get('__shiftclick__', False)):
        resolve_environment(reselect=True, pick_folder=_pick_folder)      # you asked for the folder picker: it opens now
    # one editor at a time: running the button again brings the running one forward instead of starting another
    # (several copies would all read the same slow drive at once)
    found = _running_editor()
    if found:
        url0, d0 = found
        mine = _build_id()
        if d0.get('dirty') or not mine or d0.get('build') == mine:
            try:
                webbrowser.open(url0)
            except Exception:
                pass
            return 0
        try:                                                      # an older build with nothing unsaved: replace it
            import urllib.request
            urllib.request.urlopen(url0.replace('/?k=', '/api/quit?k='), timeout=2).read()
            time.sleep(1.0)
        except Exception:
            pass
    # Inside Revit the editor runs as its own process: Revit's embedded Python only lets background threads run while it is
    # executing a script, so a server living in there answers only when the button is pressed again.
    if separate is None:
        separate = in_revit
    if separate:
        script = g.get('__file__') or (sys.argv[0] if sys.argv else '')
        exe = _find_python() if script and os.path.isfile(script) else ''
        url = _spawn_editor(exe, script) if exe else ''
        if url:
            try:
                webbrowser.open(url)
            except Exception:
                pass
            return 0
        print('The editor could not be started in its own process (no Python was found next to the pyRevit engine). It will run '
              'inside Revit instead, and may only react while you press the button again. To fix this, put the full path of a '
              'python.exe in C:\\temp\\Ribbon_FABCADMEP-Path.json as "python_path".')
    srv, token, url, handler = _start_listening('revit')
    if not in_revit:
        print('Editor:', url)
    t = threading.Thread(target=srv.serve_forever, name='fabmap-editor')
    t.daemon = True
    t.start()
    threading.Thread(target=boot_run, args=(handler,), name='fabmap-boot', daemon=True).start()
    try:
        webbrowser.open(url)                                   # the page opens at once; it shows what the editor is waiting for
    except Exception:
        pass
    if not in_revit and wait:
        print('Close this window or press Ctrl+C to stop the editor.')
        try:
            while t.is_alive():
                t.join(1)
        except KeyboardInterrupt:
            srv.shutdown()
    return 0


if __name__ == '__main__':
    argv = sys.argv
    if len(argv) >= 2 and argv[1] == '--serve':
        serve_main()
        sys.exit(0)
    if len(argv) >= 2 and argv[1].lower().endswith('.map'):
        if any(f in argv for f in ('--move', '--copy', '--delete', '--selftest')):
            sys.exit(write_main(argv))
        sys.exit(main(argv))
    editor_main()      # no sys.exit: pyRevit reports SystemExit as an error