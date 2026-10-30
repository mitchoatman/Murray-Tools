# -*- coding: utf-8 -*-
"""Reads MEP Fabrication parts from the open Revit model into exports.Part objects.
Revit API only, read only (no transactions). Run from a pyRevit script while a Revit API context is active,
which is the case inside a modal window opened by a pushbutton.

Property names are the Revit parameter names (the same names the web app read from the published model data).
Names that do not exist on a given model come back blank; use 'Save sample data' in the window to see what the
model really carries, then edit exports.py. Items marked [CONFIRM] have not been checked against a real model."""
import re

from Autodesk.Revit.DB import (FilteredElementCollector, BuiltInCategory, BuiltInParameter, StorageType, Level,
                               FilteredWorksetCollector, WorksetKind)

from .exports import Part, EXPORT_TYPES, LEVEL_PROPS

CATEGORY_BY_SOURCE = {
    'pipe': BuiltInCategory.OST_FabricationPipework,
    'hangers': BuiltInCategory.OST_FabricationHangers,
}

# Length parameters that must be read in exact internal units even if the data type check fails on this Revit version.
LENGTH_NAMES = set(['Length', 'FP_Rod Length A', 'FP_Rod Length B', 'FP_Rod Size'])


def eid_int(eid):
    try:
        return int(eid.Value)            # Revit 2024 and later
    except Exception:
        return int(eid.IntegerValue)     # earlier versions


def names_for(source):
    """The parameter names worth reading for a source. Pipework has no rods, so its parts skip those lookups."""
    names = needed_names()
    return [n for n in names if not n.startswith('FP_Rod')] if source == 'pipe' else names


def needed_names():
    names = set(LEVEL_PROPS)
    names.add('Length Option')
    for d in EXPORT_TYPES:
        for c in d['columns']:
            for n in c.get('props', []):
                names.add(n)
    return sorted(names)


def _is_length(p, name):
    try:
        from Autodesk.Revit.DB import SpecTypeId
        return p.Definition.GetDataType() == SpecTypeId.Length
    except Exception:
        pass
    try:
        from Autodesk.Revit.DB import ParameterType
        return p.Definition.ParameterType == ParameterType.Length
    except Exception:
        pass
    return name in LENGTH_NAMES


def param_value(p, name):
    """Text for a text parameter; '12.500000 in' for a length (exact, from internal feet); the displayed
    value for anything else (a level parameter shows the level name; a material parameter its material name)."""
    if p is None:
        return ''
    try:
        st = p.StorageType
        if st == StorageType.String:
            v = p.AsString()
            if v:
                return v
            return p.AsValueString() or ''
        if st == StorageType.Double and _is_length(p, name):
            return '%.6f in' % (p.AsDouble() * 12.0)
        v = p.AsValueString()
        if v:
            return v
        if st == StorageType.ElementId:
            eid = p.AsElementId()
            if eid_int(eid) > 0:
                el = p.Element.Document.GetElement(eid)
                return (el.Name if el is not None else '') or ''
            return ''
        if st == StorageType.Integer:
            return str(p.AsInteger())
        if st == StorageType.Double:
            return repr(p.AsDouble())
        return ''
    except Exception as ex:
        if len(SCAN['errors']) < 5:
            SCAN['errors'].append('%s: %s' % (name, ex))
        return ''


def workset_lookup(doc):
    if not doc.IsWorkshared:
        return lambda e: ''
    table = doc.GetWorksetTable()
    cache = {}

    def of(e):
        wid = e.WorksetId
        k = wid.IntegerValue
        if k not in cache:
            try:
                cache[k] = table.GetWorkset(wid).Name
            except Exception:
                cache[k] = ''
        return cache[k]
    return of


SCAN = {'material_like': set(), 'errors': [], 'name_src': {}, 'mat_src': {}}


def name_source_note():
    """Which source supplied the description (family name) for the parts read, for the export notes."""
    d = SCAN['name_src']
    if not d:
        return ''
    txt = ', '.join('%s (%d)' % (k, v) for k, v in sorted(d.items(), key=lambda kv: -kv[1]))
    if d.get('none') == sum(d.values()):
        return ('No family name could be read, so descriptions use Element.Name, which is "Default" on fabrication parts. '
                'Use Save sample data and send the name candidates.')
    return 'Description (family name) read from: ' + txt


def material_source_note():
    d = SCAN['mat_src']
    if not d:
        return ''
    if d.get('none') == sum(d.values()):
        return 'Part Material was not found on any part by any method.'
    return 'Part Material read from: ' + ', '.join('%s (%d)' % (k, v) for k, v in sorted(d.items(), key=lambda kv: -kv[1]))


def material_hint():
    """Names of material-like parameters seen on the first part of each category, for the export notes."""
    return ', '.join(sorted(SCAN['material_like']))


def _name_map(e):
    """Parameter names on one element: lowercase, trimmed name -> exact name."""
    m = {}
    for p in e.Parameters:
        try:
            n = p.Definition.Name
        except Exception:
            continue
        m.setdefault(n.strip().lower(), n)
        if 'material' in n.lower():
            SCAN['material_like'].add(n)
    return m


def _resolve(names, amap):
    """Needed name -> the name this model really uses (case, spaces, FP_ prefix), or the name itself."""
    out = {}
    for n in names:
        k = n.strip().lower()
        if k in amap:
            out[n] = amap[k]
        elif 'fp_' + k in amap:
            out[n] = amap['fp_' + k]
        elif k.startswith('fp_') and k[3:] in amap:
            out[n] = amap[k[3:]]
        else:
            out[n] = n
    return out


# Properties the exports cannot do without. When the direct lookup comes back empty (a duplicate parameter of the same name,
# a different spelling), every parameter of the part is scanned for them. Dropped for a category after many misses in a row.
CRITICAL = {
    'Part Material': ['part material', 'fp_part material', 'material'],
    'Service Type': ['service type', 'fp_service type', 'fabrication service type'],
    'Fabrication Service': ['fabrication service'],
}
MISS_LIMIT = 50


def scan_value(e, wanted):
    """First non-empty value among all parameters whose name (any case) is in wanted (a list of lowercase names, in order)."""
    found = {}
    for p in e.Parameters:
        try:
            n = p.Definition.Name.strip().lower()
        except Exception:
            continue
        if n in wanted and n not in found:
            v = param_value(p, n)
            if v:
                found[n] = v
    for n in wanted:
        if n in found:
            return found[n]
    return ''


_API_ATTRS = {}
_SKIP_PREFIX = ('Set', 'Delete', 'Remove', 'Add', 'Dispose', 'Copy', 'Move', 'Rotate', 'Mirror')


def _to_text(doc, v):
    """Text from a property value: a string, an element or material (its Name), an ElementId (that element's Name),
    or the first of a list of those."""
    if v is None:
        return ''
    if isinstance(v, str):
        return v.strip()
    try:
        if hasattr(v, 'Name') and isinstance(v.Name, str):
            return v.Name.strip()
    except Exception:
        pass
    try:
        if v.GetType().Name == 'ElementId':
            if eid_int(v) > 0:
                el = doc.GetElement(v)
                return (el.Name or '').strip() if el is not None else ''
            return ''
    except Exception:
        pass
    if hasattr(v, '__iter__'):
        try:
            for x in v:
                t = _to_text(doc, x)
                if t:
                    return t
        except Exception:
            pass
    return ''


def api_material(doc, e):
    """Part Material as a property or getter on the part itself, found at run time: any public member whose name
    contains 'material' (properties are read; methods are called with no arguments). Returns (text, member)."""
    names = _API_ATTRS.get('material')
    if names is None:
        names = [a for a in dir(e) if 'material' in a.lower() and not a.startswith('_') and not a.startswith(_SKIP_PREFIX)]
        _API_ATTRS['material'] = names
    for a in names:
        try:
            v = getattr(e, a)
            if callable(v):
                v = v()
        except Exception:
            continue
        t = _to_text(doc, v)
        if t:
            return t, a
    return '', ''


def get_parameters_value(e, name):
    """Every parameter called `name` (Element.GetParameters), first non-empty value."""
    try:
        for p in e.GetParameters(name):
            v = param_value(p, name)
            if v:
                return v
    except Exception:
        pass
    return ''


GENERIC_NAMES = ('', 'default')


_NEEDED = []


def _type_info(doc, tid, cache):
    """What a part's type says, as plain values (family name, name, and the needed parameters' text). Cached by type id.
    The type element itself is never kept between slices."""
    k = eid_int(tid)
    if k in cache:
        return cache[k]
    info = {'family': '', 'name': '', 'params': {}}
    if k > 0:
        try:
            te = doc.GetElement(tid)
            if te is not None:
                info['family'] = (te.FamilyName or '').strip()
                info['name'] = (te.Name or '').strip()
                if not _NEEDED:
                    _NEEDED.extend(needed_names())
                for n in _NEEDED:
                    p = te.LookupParameter(n)
                    v = param_value(p, n) if p is not None else ''
                    if v:
                        info['params'][n] = v
        except Exception:
            pass
    cache[k] = info
    return info


def _left_of_type(v):
    """'TUBE TYPE K COPPER SMLS CFOS: Default' -> 'TUBE TYPE K COPPER SMLS CFOS'."""
    v = (v or '').strip()
    i = v.rfind(':')
    return v[:i].strip() if i > 0 else v


def iter_family_candidates(doc, e, type_cache, tid=None):
    """Every place a family name could come from, as (source, value), one at a time and in the order they are trusted, so the
    caller can stop at the first that works (most parts answer on the first or second). [CONFIRM] on a real model which one
    carries the fabrication part's family name; 'Save sample data' lists all of them."""
    try:
        if tid is None:
            tid = e.GetTypeId()
        ti = _type_info(doc, tid, type_cache)
        if ti['family']:
            yield ('type FamilyName', ti['family'])
    except Exception:
        ti = None

    def bip(label, left=False):
        b = getattr(BuiltInParameter, label, None)
        if b is None:
            return None
        try:
            p = e.get_Parameter(b)
            v = param_value(p, label).strip() if p is not None else ''
            return (label + (' (left of colon)' if left else ''), _left_of_type(v) if left else v)
        except Exception:
            return None
    for item in (lambda: bip('ELEM_FAMILY_AND_TYPE_PARAM', True), lambda: bip('ELEM_FAMILY_PARAM')):
        c = item()
        if c:
            yield c
    for nm in ('Family and Type', 'Family'):
        try:
            p = e.LookupParameter(nm)
            v = param_value(p, nm).strip() if p is not None else ''
            yield ('parameter ' + nm + (' (left of colon)' if nm == 'Family and Type' else ''),
                   _left_of_type(v) if nm == 'Family and Type' else v)
        except Exception:
            pass
    c = bip('ALL_MODEL_DESCRIPTION')
    if c:
        yield c
    try:
        yield ('type Name', _type_info(doc, tid if tid is not None else e.GetTypeId(), type_cache)['name'])
    except Exception:
        pass


def family_candidates(doc, e, type_cache):
    """All candidates at once (for the sample data file)."""
    return list(iter_family_candidates(doc, e, type_cache))


def family_name(doc, e, type_cache, tid=None):
    """(family name, source), stopping at the first usable candidate. 'Default' (what Element.Name returns on a fabrication
    part) is never accepted."""
    for src, v in iter_family_candidates(doc, e, type_cache, tid):
        if v and v.strip().lower() not in GENERIC_NAMES:
            return v.strip(), src
    return '', ''


def read_parts(doc, source, names=None):
    """All parts of one source, read in one go (see read_parts_iter for the time-sliced version)."""
    g = read_parts_iter(doc, source, names)
    while True:
        try:
            next(g)
        except StopIteration as stop:
            return stop.value


def read_parts_iter(doc, source, names=None):
    """A generator: one step reads one part (yielding a progress dict), so the caller can stop between parts and give
    Revit back its main thread. Returns the list of exports.Part when it finishes."""
    names = names or names_for(source)
    ws_of = workset_lookup(doc)
    type_cache = {}
    out = []
    lookup = None
    misses = dict((k, 0) for k in CRITICAL)
    stage = 'Reading ' + ('pipework' if source == 'pipe' else 'hangers')
    # A snapshot of element ids only. No collector and no element is kept between slices: Revit invalidates those when the
    # user picks, edits or changes the view while a read is in progress. Each part is fetched fresh in its own step.
    ids = list(FilteredElementCollector(doc).OfCategory(CATEGORY_BY_SOURCE[source]).WhereElementIsNotElementType().ToElementIds())
    total = len(ids)
    state = {'lookup': None}

    def one(e):
        props = {}
        if state['lookup'] is None:              # parameter names are the same on every part of a category
            state['lookup'] = _resolve(names, _name_map(e))
        lookup = state['lookup']
        tid = None                               # the type id is asked of Revit at most once per part
        for n in names:
            p = e.LookupParameter(lookup[n])
            if p is None:
                if tid is None:
                    tid = e.GetTypeId()
                v = _type_info(doc, tid, type_cache)['params'].get(n, '')
            else:
                v = param_value(p, n)
            if n == 'Part Material':
                src = 'parameter lookup' if v else ''
                if v == '':
                    v = get_parameters_value(e, n)
                    src = 'GetParameters' if v else ''
                if v == '' and misses[n] < MISS_LIMIT:
                    v = scan_value(e, CRITICAL[n])
                    src = 'parameter scan' if v else ''
                if v == '':
                    v, member = api_material(doc, e)
                    src = ('part.' + member) if v else ''
                misses[n] = 0 if v else misses[n] + 1
                SCAN['mat_src'][src or 'none'] = SCAN['mat_src'].get(src or 'none', 0) + 1
            elif v == '' and n in CRITICAL and misses[n] < MISS_LIMIT:
                v = scan_value(e, CRITICAL[n])
                misses[n] = 0 if v else misses[n] + 1
            if v != '':
                props[n] = v
        fam, fam_src = family_name(doc, e, type_cache, tid)
        SCAN['name_src'][fam_src or 'none'] = SCAN['name_src'].get(fam_src or 'none', 0) + 1
        if fam:
            props['Family'] = fam
        straight = None
        try:
            straight = bool(e.IsAStraight())      # FabricationPart.IsAStraight()
        except Exception:
            pass
        name = fam                               # Element.Name is only read when there is no family name
        if not name:
            try:
                name = e.Name
            except Exception:
                name = ''
        ws = ws_of(e)
        if ws:
            props['Workset'] = ws
        part = Part(eid_int(e.Id), name, props, straight, ws)
        return part

    for n, eid in enumerate(ids):
        try:
            e = doc.GetElement(eid)
            if e is not None:
                part = one(e)
                part.eid = eid                       # geometry is read later from the id, not from a kept element
                part.doc = doc
                out.append(part)
        except Exception as ex:                      # one part that cannot be read is skipped; the read carries on
            SCAN['skipped'] = SCAN.get('skipped', 0) + 1
            if len(SCAN['errors']) < 5:
                SCAN['errors'].append('part %s skipped: %s' % (n, ex))
        yield {'stage': stage, 'done': n + 1, 'total': total}
    return out


def level_names(doc, parts):
    """Levels used by the parts, lowest elevation first."""
    used = set()
    for p in parts:
        for n in LEVEL_PROPS:
            v = p.props.get(n)
            if v:
                used.add(v.strip())
                break
    elev = {}
    for lv in FilteredElementCollector(doc).OfClass(Level):
        try:
            elev[lv.Name.strip()] = lv.Elevation
        except Exception:
            pass
    return sorted(used, key=lambda n: (elev.get(n, 1e12), n.lower()))


def workset_names(doc, parts=None):
    """User worksets of the model (empty when the model is not workshared)."""
    if not doc.IsWorkshared:
        return []
    names = []
    for w in FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset):
        names.append(w.Name)
    return sorted(names, key=lambda s: s.lower())


CHANGE_CAP = 3000      # beyond this many added or modified elements in one change, do not check each one


def reset_scan():
    """Clears the counters shown in the export notes (a refresh reads everything again)."""
    SCAN['errors'] = []
    SCAN['name_src'] = {}
    SCAN['mat_src'] = {}
    SCAN['skipped'] = 0


def fab_changes(doc, args, known):
    """From a DocumentChanged event: (added ids, modified ids, deleted ids, many) for fabrication pipework and hangers only.
    `known` is the set of part ids already read (a deleted element can no longer be looked up, so it is matched by id).
    Runs inside the event, so it stays cheap: past CHANGE_CAP added or modified elements it just reports many=True.
    [CONFIRM] DocumentChanged firing and these three Get...ElementIds calls on your Revit versions."""
    cats = set(int(b) for b in CATEGORY_BY_SOURCE.values())

    def ids(coll):
        return list(coll) if coll is not None else []

    def is_fab(eid):
        try:
            e = doc.GetElement(eid)
            c = e.Category if e is not None else None
            return c is not None and eid_int(c.Id) in cats
        except Exception:
            return False
    added, modified, deleted = ids(args.GetAddedElementIds()), ids(args.GetModifiedElementIds()), ids(args.GetDeletedElementIds())
    many = len(added) + len(modified) > CHANGE_CAP
    out_a, out_m = [], []
    if not many:
        out_a = [eid_int(e) for e in added if is_fab(e)]
        out_m = [eid_int(e) for e in modified if eid_int(e) in known or is_fab(e)]     # a known part needs no lookup
    out_d = [eid_int(e) for e in deleted if eid_int(e) in known]
    return out_a, out_m, out_d, many


def on_document_changed(doc, args, svc):
    """The whole DocumentChanged handler, as a function that never raises. Returns the page message, or None.
    Revit can hand this event a document object that is already invalid (seen when an edit lands while a read is
    running). That must not stop the check: the part ids in the event are enough. Only a document that is valid and
    clearly a different one is ignored."""
    try:
        try:
            same = args.GetDocument().Equals(doc)
        except Exception:
            same = None                              # invalid or unreadable document object: decide by element ids instead
        if same is False:
            return None
        added, modified, deleted, many = fab_changes(doc, args, svc.known_ids())
        return svc.note_changes(added, modified, deleted, many)
    except Exception:
        SCAN['change_errors'] = SCAN.get('change_errors', 0) + 1
        return None


def selected_ids(uidoc):
    try:
        return [eid_int(i) for i in uidoc.Selection.GetElementIds()]
    except Exception:
        return []


def dump_sample(doc, per_source=3):
    """Text with every parameter of the first few parts of each source. Contains model data."""
    lines = ['Murray bulk BOM sample data. Contains model data: share only with whoever is fixing the export.', '']
    for source, bic in CATEGORY_BY_SOURCE.items():
        lines.append('=== %s (%s) ===' % (source, bic))
        col = FilteredElementCollector(doc).OfCategory(bic).WhereElementIsNotElementType()
        n = 0
        for e in col:
            if n >= per_source:
                break
            n += 1
            try:
                lines.append('-- element %s name=%r' % (eid_int(e.Id), e.Name))
            except Exception:
                lines.append('-- element %s' % eid_int(e.Id))
            try:
                for src, v in family_candidates(doc, e, {}):
                    lines.append('   NAME CANDIDATE %s = %r' % (src, v))
            except Exception as ex:
                lines.append('   name candidates unavailable: %s' % ex)
            try:
                mt, mem = api_material(doc, e)
                lines.append('   API MATERIAL members %s -> %r (via %s)' % (_API_ATTRS.get('material'), mt, mem or 'none'))
                lines.append('   GetParameters("Part Material") = %r' % get_parameters_value(e, 'Part Material'))
                for a in sorted(dir(e)):
                    if a.startswith('_') or a.startswith(_SKIP_PREFIX):
                        continue
                    try:
                        v = getattr(e, a)
                    except Exception:
                        continue
                    if callable(v):
                        continue
                    t = v if isinstance(v, (int, float, bool)) else _to_text(doc, v)
                    if t != '' and t is not None:
                        lines.append('   API %s = %r' % (a, t))
            except Exception as ex:
                lines.append('   API members unavailable: %s' % ex)
            try:
                lines.append('   IsAStraight() = %s' % e.IsAStraight())
            except Exception as ex:
                lines.append('   IsAStraight() unavailable: %s' % ex)
            rows = []
            for p in e.Parameters:
                try:
                    rows.append((p.Definition.Name, param_value(p, p.Definition.Name)))
                except Exception:
                    pass
            for k, v in sorted(rows, key=lambda kv: kv[0].lower()):
                lines.append('   %s = %s' % (k, v))
            lines.append('')
        if n == 0:
            lines.append('(no elements in this category)')
            lines.append('')
    return '\n'.join(lines)


def _fetch(part):
    """The Revit element of a part, fetched fresh from its id (never a kept element), or None."""
    try:
        if getattr(part, 'eid', None) is not None:
            return part.doc.GetElement(part.eid)
        return part.element
    except Exception:
        return None


def read_connectors(part):
    """The connectors of a part as ([(x, y, z), ...] origins in feet, radius or None), or None when it has none that can be read.
    The only Revit calls the centerline and straight-pipe stages need. [CONFIRM] FabricationPart.ConnectorManager,
    Connector.Origin and Connector.Radius on your Revit versions."""
    e = _fetch(part)
    if e is None:
        return None
    try:
        origins, radius = [], None
        for c in e.ConnectorManager.Connectors:
            o = c.Origin
            origins.append((o.X, o.Y, o.Z))
            if radius is None:
                try:
                    radius = c.Radius
                except Exception:
                    try:
                        radius = c.Width / 2.0
                    except Exception:
                        radius = None
        return (origins, radius) if origins else None
    except Exception:
        return None


def read_mesh(part, level_of_detail=0.2):
    """Triangles of one part for the 3D preview: (flat x,y,z list in feet, flat vertex index list), or None.
    Coarse detail, no references. Read only. [CONFIRM] triangle counts and speed on a real model."""
    from Autodesk.Revit.DB import Options, ViewDetailLevel, Solid, Mesh, GeometryInstance
    e = _fetch(part)
    if e is None:
        return None
    opts = Options()
    opts.DetailLevel = ViewDetailLevel.Coarse
    opts.ComputeReferences = False
    opts.IncludeNonVisibleObjects = False
    verts, tris = [], []

    def add_mesh(m):
        if m is None:
            return
        base = len(verts) // 3
        for v in m.Vertices:
            verts.extend((v.X, v.Y, v.Z))
        for i in range(m.NumTriangles):
            t = m.get_Triangle(i)
            tris.extend((base + t.get_Index(0), base + t.get_Index(1), base + t.get_Index(2)))

    def walk(geom):
        for g in geom:
            if isinstance(g, Solid):
                if g.Faces.Size == 0:
                    continue
                for face in g.Faces:
                    add_mesh(face.Triangulate(level_of_detail))
            elif isinstance(g, Mesh):
                add_mesh(g)
            elif isinstance(g, GeometryInstance):
                walk(g.GetInstanceGeometry())

    try:
        geom = e.get_Geometry(opts)
        if geom is not None:
            walk(geom)
    except Exception:
        return None
    return (verts, tris) if tris else None
