# -*- coding: utf-8 -*-
"""The 3D preview as messages. No Revit imports.

1. scene_table: the parts the ticked exports include, with what the page needs to name, color, filter and pick them. No
   geometry, so it can be sent at once.
2. geometry messages, in build order (exports.BUILD_ORDER): single-line centerlines for every part with connectors, then
   straight pipe as cylinders, then the real geometry of fittings, then valves, then hangers. Positions are made relative to
   one fixed centre so float32 keeps millimetre precision far from the origin.
"""
import base64
from array import array

from . import exports


def _b64(a):
    return base64.b64encode(a.tobytes()).decode('ascii')


def choose_parts(*args, **kwargs):
    """choose_parts_iter run to the end."""
    g = choose_parts_iter(*args, **kwargs)
    while True:
        try:
            next(g)
        except StopIteration as stop:
            return stop.value


def choose_parts_iter(sources, keys, levels=None, hidden_worksets=None, selection_ids=None, chunk=1500):
    """The parts the ticked exports include, once each, as (part, group label, row, kind), and the group labels.
    Same rule as the MR forms (exports.select_parts). Pure Python. A generator that yields after every `chunk` parts, so a model
    of tens of thousands of parts is never one long step (an export clicked meanwhile waits a fraction of a second at most)."""
    wanted = {exports._level_key(l): l for l in (levels or [])}
    hidden = set(hidden_worksets or [])
    sel = set(selection_ids) if selection_ids else None
    hanger_ids = set(p.id for p in sources.get('hangers', []))
    chosen, seen, groups = [], set(), []
    for key in keys:
        d = exports.type_by_key(key)
        if not d:
            continue
        groups.append(d['label'])
        src = list(sources.get(d['source'], []))
        for i in range(0, len(src), chunk):
            _, kept, _, _, _, _ = exports.select_parts(d, src[i:i + chunk], wanted, hidden, sel)
            for it in kept:
                part = it['part']
                if part.id in seen:
                    continue
                seen.add(part.id)
                chosen.append((part, d['label'], it['row'], exports.part_kind(part, part.id in hanger_ids)))
            yield
    return chosen, groups


def scene_table(chosen, groups):
    """The scene message: the parts table only. Geometry follows in messages of type 'geom'."""
    by_kind = {}
    table = []
    for part, label, row, kind in chosen:
        by_kind[kind] = by_kind.get(kind, 0) + 1
        table.append({
            'id': part.id, 'group': label, 'kind': kind,
            'name': exports.description(part),
            'size': row.get('Size') or '',
            'length': row.get('Length') or '',
            'material': row.get('Material') or '',
            'level': exports._s(exports.first_value(part.props, exports.LEVEL_PROPS)),
            'service': exports.service_text(part),
        })
    return {'type': 'scene', 'groups': groups, 'parts': table,
            'counts': {'shown': len(table), 'included': len(table), 'byKind': by_kind}}


def centre_of(verts):
    """Centre of the bounding box of a flat x,y,z list (the fixed origin for everything sent afterwards)."""
    xs, ys, zs = verts[0::3], verts[1::3], verts[2::3]
    return [(min(xs) + max(xs)) / 2.0, (min(ys) + max(ys)) / 2.0, (min(zs) + max(zs)) / 2.0]


def lines_message(entries, center):
    """Centerlines for some parts: entries is a list of (part id, flat segment end points). One vertex pair per segment."""
    pos = array('f')
    table = []
    cx, cy, cz = center
    for pid, verts in entries:
        v0 = len(pos) // 3
        for k in range(0, len(verts), 3):
            pos.extend((verts[k] - cx, verts[k + 1] - cy, verts[k + 2] - cz))
        table.append({'id': pid, 'v': [v0, len(pos) // 3 - v0]})
    return {'type': 'geom', 'kind': 'lines', 'stage': 'lines', 'center': list(center), 'parts': table, 'pos': _b64(pos)}


def solid_message(stage, lod, entries, center):
    """Solid geometry for some parts: entries is a list of (part id, verts, tris). lod is 'cyl' (an exact cylinder for straight
    pipe) or 'detail' (the real tessellated geometry). Indices are local to this message (the page offsets them)."""
    pos = array('f')
    idx = array('I')
    table = []
    cx, cy, cz = center
    for pid, verts, tris in entries:
        v0 = len(pos) // 3
        i0 = len(idx)
        for k in range(0, len(verts), 3):
            pos.extend((verts[k] - cx, verts[k + 1] - cy, verts[k + 2] - cz))
        idx.extend([v0 + i for i in tris])
        table.append({'id': pid, 'v': [v0, len(pos) // 3 - v0], 'i': [i0, len(idx) - i0]})
    return {'type': 'geom', 'kind': 'solid', 'stage': stage, 'lod': lod, 'center': list(center),
            'parts': table, 'pos': _b64(pos), 'idx': _b64(idx)}
