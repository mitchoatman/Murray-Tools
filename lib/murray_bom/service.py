# -*- coding: utf-8 -*-
"""Message handling between the web page and Revit. No UI or Revit imports here, so it can be tested:
the Revit reader is passed in (revit_read in Revit, a stub in tests)."""
import os
import re
import datetime
import time

from . import exports, mrform, scene, shapes

TEMPLATE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'templates', 'MR-BOM.xlsx')


def safe_name(s):
    return re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', s).strip().rstrip('.') or 'model'


def model_base_name(title):
    t = re.sub(r'\.rvt$', '', title or '', flags=re.I)
    return t


class Service(object):
    """Talks to the page. Everything that needs the Revit API (reading parts, geometry, selection, sample data) is a
    job that runs only inside pump(), which the pushbutton calls from a Revit ExternalEvent, a few hundredths of a
    second at a time. handle() answers page messages without touching the Revit API, so it is safe to call from the
    window's own event handlers (outside a Revit API context)."""

    def __init__(self, reader, doc, uidoc, project_name='', project_number=''):
        # Built while a Revit API context is active (the command itself), so these reads are valid.
        self.reader = reader
        self.doc = doc
        self.uidoc = uidoc
        self.title = doc.Title
        self.workshared = bool(getattr(doc, 'IsWorkshared', False))
        self.project_name = project_name
        self.project_number = project_number
        self.sel = []                    # empty until the user clicks Use current selection (nothing is read at launch)
        self.cache = {}
        self.mesh_cache = {}             # part id -> (verts, tris) read in full, or False when the part has none
        self.conn_cache = {}             # part id -> (connector origins, radius) or False: feeds centerlines and cylinders
        self.od_cache = {}               # (family, size, connector radius) -> outside radius measured from real geometry
        self.default_folder = ''         # where exports go until the user picks another folder (set by the pushbutton)
        self.build_gen = None            # the staged 3D build (centerlines, straights, fittings, valves, hangers)
        self.chosen = []                 # (part, group, row, kind) of the current view
        self.sent_lines = set()          # ids whose centerline has been sent to the page for the current view
        self.sent_solid = set()          # ids whose solid geometry has been sent
        self.stream_center = None        # fixed origin of everything sent, set by the first geometry (kept across views)
        self.line_batch = 500            # parts per centerline message
        self.solid_batch = 250           # parts per solid geometry message
        self.worksets = []
        self.levels = []
        self.loaded = False
        self.sources_ready = set()       # 'pipe' and 'hangers' once each has been read: an export needs only its own source
        self.read_seconds = 0.0
        self.active = {}                 # seconds of Revit time spent per kind of job
        self.load_gen = None
        self.misc = []                   # [(kind, generator)] sample and selection jobs
        self.scene_gen = None
        self.scene_msg = None
        self.tracking = False            # set by the pushbutton once it is listening for model changes
        self._known = set()              # ids of the parts read, so a deleted part can be recognised
        self.changed = {'added': 0, 'modified': 0, 'deleted': 0, 'many': False}
        self.dirty = set()               # ids whose geometry must be read again on refresh
        self.dirty_all = False

    # ---- what the page may ask for (no Revit API here) ----

    def handle(self, msg):
        """Answers one page message. Returns the messages to send back now; longer work is queued for pump()."""
        cmd = msg.get('cmd')
        if cmd == 'ready':
            if self.load_gen is None and not self.loaded:
                self.load_gen = self._load_iter()
            return [self._basic_init()]
        if cmd == 'scene':
            self.scene_msg = msg
            self._close(self.scene_gen)
            self._close(self.build_gen)
            self.scene_gen, self.build_gen = None, None
            self.scene_gen = self._scene_iter(msg) if self.loaded else None
            return []
        if cmd == 'refreshModel':
            return self._refresh_model()
        if cmd == 'cancel':
            stage = 'scene' if self.scene_gen is not None else 'build'
            self._close(self.scene_gen)
            self._close(self.build_gen)
            self.scene_gen, self.build_gen = None, None
            if stage == 'scene':
                self.scene_msg = None                    # no view was started; the last one stays on the page
            return [{'type': 'cancelled', 'stage': stage, 'remaining': self._remaining()}]
        if cmd == 'refine':                              # continue the build where it stopped
            self._start_build()
            return []
        if cmd == 'export':
            return [self.export(msg)]
        if cmd == 'sample':
            self.misc.append(('sample', self._sample_iter(msg)))
            return []
        if cmd == 'refreshSelection':
            if self.loaded:                              # nothing to match a selection against before the model is read
                self.misc.append(('selection', self._selection_iter()))
            return []
        return []

    def known_ids(self):
        return self._known

    def note_changes(self, added, modified, deleted, many=False):
        """Called from the pushbutton's DocumentChanged handler with the fabrication parts that changed. Pure bookkeeping
        (no Revit API): returns the message for the page, or None when nothing relevant changed."""
        if not (added or modified or deleted or many):
            return None
        c = self.changed
        c['added'] += len(added)
        c['modified'] += len(modified)
        c['deleted'] += len(deleted)
        c['many'] = c['many'] or bool(many)
        self.dirty.update(modified)
        self.dirty.update(deleted)
        self.dirty_all = self.dirty_all or bool(many)
        self._known.update(added)
        self._known.difference_update(deleted)
        return dict(c, type='modelChanged')

    def _refresh_model(self):
        """Reads the parts again. Geometry already read is kept for parts that did not change, so a refresh after a small
        edit only re-reads what the edit touched (everything when changes were not tracked or too many to check)."""
        if not self.loaded:
            return []                                    # a read is already running
        self._close(self.scene_gen)
        self.scene_gen = None
        self._close(self.build_gen)
        self.build_gen = None
        if self.dirty_all or not self.tracking:
            self.mesh_cache.clear()
            self.conn_cache.clear()
            self.od_cache.clear()
        else:
            for i in self.dirty:                         # a changed part may have a new shape: read it again
                self.mesh_cache.pop(i, None)
                self.conn_cache.pop(i, None)
        self.dirty = set()
        self.dirty_all = False
        self.changed = {'added': 0, 'modified': 0, 'deleted': 0, 'many': False}
        self.loaded = False
        self.sources_ready = set()
        self.cache = {}
        self.load_gen = self._load_iter()
        return [{'type': 'refreshing'}]

    @property
    def busy(self):
        return bool(self.load_gen or self.misc or self.scene_gen or self.build_gen)

    def _basic_init(self):
        job = mrform.job_from_project(self.project_name or self.title)
        number = self.project_number if re.match(r'^\d{5}$', (self.project_number or '').strip()) else job['jobNumber']
        name = self.project_name if self.project_name and self.project_name.strip().lower() != 'project name' else job['jobName']
        if number and name:
            name = mrform.job_from_project(name)['jobName'] or name
        return {'type': 'init', 'model': self.title, 'jobName': name, 'jobNumber': number,
                'defaultFolder': self.default_folder if os.path.isdir(self.default_folder or '') else '',
                'exports': [{'key': t['key'], 'label': t['label'], 'source': t['source']} for t in exports.EXPORT_TYPES]}

    @staticmethod
    def _close(gen):
        if gen is not None:
            try:
                gen.close()
            except Exception:
                pass

    # ---- the time-sliced part (Revit API context) ----

    def pump(self, budget=0.05):
        """Runs jobs for about `budget` seconds, then returns the messages to send (at most one progress message).
        One step of a job is one part, so a slice never overruns by more than one part."""
        out = []
        t_end = time.time() + budget
        last = None
        while time.time() < t_end:
            kind, gen = self._next_job()
            if gen is None:
                break
            t0 = time.time()
            try:
                p = next(gen)
                if p and p.get('type'):                  # a finished message (a geometry patch), sent as it comes
                    out.append(p)
                elif p:
                    last = p
            except StopIteration as stop:
                last = None
                self._finish(kind, stop.value, out)
            except Exception as ex:                      # a failing job reports and stops; the window stays usable
                last = None
                self._drop(kind)
                out.append({'type': 'error', 'message': '%s: %s' % (kind, ex)})
            self.active[kind] = self.active.get(kind, 0.0) + (time.time() - t0)
        if last:
            out.append({'type': 'progress', 'stage': last.get('stage', ''), 'done': last.get('done', 0), 'total': last.get('total', 0),
                        'cancel': kind_is_scene(last)})
        return out

    def run_sync(self, budget=0.05, limit=100000):
        """Pumps until nothing is left to do (tests, and anywhere time slicing is not needed)."""
        msgs = []
        for _ in range(limit):
            if not self.busy:
                break
            msgs.extend(m for m in self.pump(budget) if m['type'] != 'progress')
        return msgs

    def _next_job(self):
        if self.load_gen is not None:
            return 'load', self.load_gen
        if self.misc:
            return self.misc[0]
        if self.scene_gen is not None:
            return 'scene', self.scene_gen
        if self.build_gen is not None:
            return 'build', self.build_gen
        return None, None

    def _drop(self, kind):
        if kind == 'load':
            self.load_gen = None
        elif kind == 'scene':
            self.scene_gen = None
        elif kind == 'build':
            self.build_gen = None
        elif self.misc:
            self.misc.pop(0)

    def _finish(self, kind, value, out):
        self._drop(kind)
        if kind == 'load':
            self.loaded = True
            out.append(value)
            c = self.changed
            if c['added'] or c['modified'] or c['deleted'] or c['many']:     # edits made while this read was running
                out.append(dict(c, type='modelChanged'))
            if self.scene_msg is not None:               # the page asked for a view while parts were still loading
                self.scene_gen = self._scene_iter(self.scene_msg)
        elif kind == 'scene':
            out.append(value)
            self._start_build()                          # the parts table is up: build the view in order
        elif value:
            out.append(value)

    def _iter_parts(self, source):
        if hasattr(self.reader, 'read_parts_iter'):
            parts = yield from self.reader.read_parts_iter(self.doc, source)
        else:
            parts = self.reader.read_parts(self.doc, source)
            yield {'stage': 'Reading ' + source, 'done': len(parts), 'total': len(parts)}
        return parts

    def _load_iter(self):
        reset = getattr(self.reader, 'reset_scan', None)
        if reset:
            reset()
        t0 = time.time()
        for source in ('pipe', 'hangers'):
            self.cache[source] = yield from self._iter_parts(source)
            self.sources_ready.add(source)
            if source == 'pipe':
                # Copper, Non Copper and Valve exports (valves are pipework parts) need only this, so they can run while the
                # hangers are still being read. Levels and worksets are listed now and refreshed when everything is in.
                pipe = self.cache['pipe']
                self.worksets = self.reader.workset_names(self.doc, pipe)
                self.levels = self.reader.level_names(self.doc, pipe)
                self._known = set(p.id for p in pipe)
                yield {'type': 'sourceReady', 'source': 'pipe', 'levels': self.levels, 'worksets': self.worksets,
                       'workshared': self.workshared,
                       'counts': '%d fabrication pipework parts read. Copper, Non Copper and Valve exports are ready; reading the hangers.' % len(pipe)}
        self.read_seconds = time.time() - t0
        yield {'stage': 'Reading levels and worksets', 'done': 0, 'total': 0}
        pipe, hang = self.cache['pipe'], self.cache['hangers']
        self.worksets = self.reader.workset_names(self.doc, pipe + hang)
        self.levels = self.reader.level_names(self.doc, pipe + hang)
        self._known = set(p.id for p in pipe + hang)
        self.sel = [i for i in self.sel if i in self._known]     # a selected part may have been deleted
        return {
            'type': 'partsReady', 'levels': self.levels, 'worksets': self.worksets, 'workshared': self.workshared,
            'counts': ('%d fabrication pipework parts and %d fabrication hangers in this model.' % (len(pipe), len(hang))
                       + self._skipped_note()),
        }

    def _skipped_note(self):
        n = getattr(self.reader, 'SCAN', {}).get('skipped', 0) if hasattr(self.reader, 'SCAN') else 0
        return (' %d parts could not be read (changed in Revit while reading) and were skipped. Reopen the window to read them.' % n) if n else ''

    def _apply_selection(self, ids, parts):
        """Only fabrication parts count as a selection; other selected elements are ignored, so a selection with no
        fabrication parts behaves like no selection."""
        fab_ids = set(p.id for p in parts)
        other = len([i for i in ids if i not in fab_ids])
        self.sel = [i for i in ids if i in fab_ids]
        return other

    def _selection_iter(self):
        yield
        ids = self.reader.selected_ids(self.uidoc)
        other = self._apply_selection(ids, self.cache['pipe'] + self.cache['hangers'])
        return {'type': 'selection', 'selection': len(self.sel), 'selectionOther': other}

    def _hidden_worksets(self, msg):
        """The page sends the worksets ticked in the list; none ticked means all (like levels)."""
        ticked = set(msg.get('worksets') or [])
        if not ticked:
            return []
        return [w for w in self.worksets if w not in ticked]

    def _scene_iter(self, msg):
        """The parts table of the view (no Revit API, no geometry), sent at once; the staged build follows."""
        sources = {'pipe': self.cache['pipe'], 'hangers': self.cache['hangers']}
        self.chosen, groups = yield from scene.choose_parts_iter(
            sources, msg.get('keys') or [], levels=msg.get('levels') or [],
            hidden_worksets=self._hidden_worksets(msg),
            selection_ids=self.sel if msg.get('selectionOnly') else None)
        self.sent_lines, self.sent_solid = set(), set()
        out = scene.scene_table(self.chosen, groups)
        out['counts']['readSeconds'] = round(self.read_seconds, 1)
        return out

    def _remaining(self):
        """Parts of the current view that do not have solid geometry yet (and might get it)."""
        return len([1 for p, _, _, _ in self.chosen if p.id not in self.sent_solid and self.mesh_cache.get(p.id) is not False])

    def _start_build(self):
        """Starts (or continues) the staged build of the current view. Whatever the page already has is skipped."""
        if not self.loaded or self.build_gen is not None or self.scene_gen is not None or not self.chosen:
            return
        self.build_gen = self._build_iter()

    def _build_iter(self):
        """The 3D view in order: single-line centerlines for every part with connectors, then straight pipe as cylinders, then
        the real geometry of fittings, then valves, then hangers. One part per step (so a slice never overruns by more than a
        part); a message is sent for every batch and a stageDone message at the end of each stage. Everything read is kept in
        conn_cache / mesh_cache, so a changed view re-sends from memory instead of asking Revit again."""
        kinds = {'straight': [], 'fitting': [], 'valve': [], 'hanger': []}
        for part, _, _, kind in self.chosen:
            kinds[kind].append(part)

        # Stage 1: centerlines (everything that can have connectors; hangers have none)
        line_parts = [p for p, _, _, k in self.chosen if k != 'hanger']
        batch, total = [], len(line_parts)
        for n, part in enumerate(line_parts):
            if part.id not in self.sent_lines:
                c = self.conn_cache.get(part.id)
                if c is None:
                    c = self.reader.read_connectors(part) or False
                    self.conn_cache[part.id] = c
                segs = shapes.centerline_segments(c[0]) if c else []
                if segs:
                    batch.append((part.id, segs))
                else:
                    self.sent_lines.add(part.id)             # nothing to draw: do not try again
            if len(batch) >= self.line_batch:
                yield self._lines_out(batch)
                batch = []
            yield {'stage': 'Centerlines', 'done': n + 1, 'total': total}
        if batch:
            yield self._lines_out(batch)
        yield {'type': 'stageDone', 'stage': 'lines', 'label': exports.STAGE_LABEL['lines'], 'parts': total}

        # Stage 2: straight pipe as cylinders from the connectors already read. The radius is the outside radius measured from the
        # real geometry of one pipe of each family, size and connector radius (a connector's own radius is the nominal size, not the
        # outside diameter), so a handful of reads covers thousands of pipes. A straight whose connectors cannot be read gets its
        # real geometry here, in the same stage.
        batch, real = [], []
        for n, part in enumerate(kinds['straight']):
            if part.id not in self.sent_solid:
                c = self.conn_cache.get(part.id)
                m = None
                if c and c[1] and len(c[0]) == 2:
                    r = self._outside_radius(part, c)
                    m = shapes.cylinder_mesh(c[0][0], c[0][1], r)
                if m:
                    batch.append((part.id, m[0], m[1]))
                else:
                    mesh = self.mesh_cache.get(part.id)
                    if mesh is None:
                        mesh = self.reader.read_mesh(part) or False
                        self.mesh_cache[part.id] = mesh
                    if mesh:
                        real.append((part.id, mesh[0], mesh[1]))
            if len(batch) >= self.solid_batch:
                yield self._solid_out('straights', 'cyl', batch)
                batch = []
            if len(real) >= self.solid_batch:
                yield self._solid_out('straights', 'detail', real)
                real = []
            yield {'stage': 'Straights', 'done': n + 1, 'total': len(kinds['straight'])}
        if batch:
            yield self._solid_out('straights', 'cyl', batch)
        if real:
            yield self._solid_out('straights', 'detail', real)
        if kinds['straight']:
            yield {'type': 'stageDone', 'stage': 'straights', 'label': exports.STAGE_LABEL['straights'], 'parts': len(kinds['straight'])}

        # Stages 3 to 5: the real geometry, valves then hangers then fittings
        for stage, parts in (('valves', kinds['valve']), ('hangers', kinds['hanger']), ('fittings', kinds['fitting'])):
            if not parts:
                continue
            label = exports.STAGE_LABEL[stage]
            batch = []
            for n, part in enumerate(parts):
                if part.id not in self.sent_solid:
                    m = self.mesh_cache.get(part.id)
                    if m is None:
                        m = self.reader.read_mesh(part) or False
                        self.mesh_cache[part.id] = m
                    if m:
                        batch.append((part.id, m[0], m[1]))
                if len(batch) >= self.solid_batch:
                    yield self._solid_out(stage, 'detail', batch)
                    batch = []
                yield {'stage': label, 'done': n + 1, 'total': len(parts)}
            if batch:
                yield self._solid_out(stage, 'detail', batch)
            yield {'type': 'stageDone', 'stage': stage, 'label': label, 'parts': len(parts)}
        return {'type': 'buildDone', 'seconds': round(self.active.get('build', 0.0), 1), 'solids': len(self.sent_solid),
                'total': len(self.chosen), 'missing': self._remaining()}

    def _outside_radius(self, part, conn):
        """The outside radius for a straight's cylinder: measured once from the real geometry of the first pipe of its family, size
        and connector radius, then reused. Falls back to the connector radius when that geometry cannot be read or measured."""
        key = (exports.description(part), exports._s(exports.first_value(part.props, ['Product Entry', 'Size'])), round(conn[1] or 0.0, 4))
        if key not in self.od_cache:
            r = None
            mesh = self.mesh_cache.get(part.id)
            if mesh is None:
                mesh = self.reader.read_mesh(part) or False
                self.mesh_cache[part.id] = mesh
            if mesh:
                r = shapes.barrel_radius(mesh[0], conn[0][0], conn[0][1])
            self.od_cache[key] = r or conn[1]
        return self.od_cache[key]

    def _centre(self, verts):
        if self.stream_center is None:
            self.stream_center = scene.centre_of(verts)
        return self.stream_center

    def _lines_out(self, batch):
        center = self._centre([x for _, v in batch for x in v])
        self.sent_lines.update(b[0] for b in batch)
        return scene.lines_message(batch, center)

    def _solid_out(self, stage, lod, batch):
        center = self._centre(batch[0][1])
        self.sent_solid.update(b[0] for b in batch)
        return scene.solid_message(stage, lod, batch, center)

    def _sample_iter(self, msg):
        yield {'stage': 'Writing sample data', 'done': 0, 'total': 0}
        return self.sample(msg)

    # ---- exports (pure Python on the parts already read) ----

    def export(self, msg):
        keys = msg.get('keys') or []
        folder = msg.get('folder') or ''
        if not folder or not os.path.isdir(folder):
            return {'type': 'error', 'message': 'Choose an output folder first.'}
        if not keys:
            return {'type': 'error', 'message': 'Tick at least one export.'}
        needed = set(exports.type_by_key(k)['source'] for k in keys if exports.type_by_key(k))
        waiting = sorted(needed - self.sources_ready)
        if waiting:
            names = ' and '.join('the pipework' if w == 'pipe' else 'the hangers' for w in waiting)
            return {'type': 'error', 'message': 'Still reading %s. That export will be ready in a moment.' % names}
        sources = {'pipe': self.cache.get('pipe', []), 'hangers': self.cache.get('hangers', [])}
        levels = msg.get('levels') or []
        res = exports.run_exports(
            sources, keys, levels=levels, hidden_worksets=self._hidden_worksets(msg),
            excluded_services=msg.get('excludedServices') or [],
            selection_ids=self.sel if msg.get('selectionOnly') else None)
        base = safe_name(model_base_name(self.title))
        if levels:
            base += ' - ' + safe_name(', '.join(levels))
        files, reports = [], []
        for label in [t['label'] for t in exports.EXPORT_TYPES if t['key'] in keys]:
            r = res.get(label)
            if r is None:
                continue
            notes = list(r['warnings'])
            mn = getattr(self.reader, 'material_source_note', lambda: '')()
            if mn:
                notes.insert(0, mn)
            nn = getattr(self.reader, 'name_source_note', lambda: '')()
            if nn and not any(x.startswith('Description (family') or x.startswith('No family name') for x in notes):
                notes.insert(0, nn)
            if (not r['rows'] or any('Part Material' in w and 'blank on every' in w for w in notes)):
                hint = getattr(self.reader, 'material_hint', lambda: '')()
                notes.append('Part Material was not read from the parts. Material-like parameters on a part: %s' % (hint or 'none found. Use Save sample data'))
                errs = getattr(self.reader, 'SCAN', {}).get('errors') if hasattr(self.reader, 'SCAN') else None
                if errs:
                    notes.append('Read errors: ' + '; '.join(errs))
            if not r['rows']:
                notes.insert(0, 'No rows, so no file was written for this export.')
                reports.append({'label': label, 'notes': notes})
                continue
            data = mrform.build(TEMPLATE, r['rows'], {
                'date': datetime.date.today(),
                'jobName': (msg.get('jobName') or '').strip(),
                'jobNumber': (msg.get('jobNumber') or '').strip(),
                'levels': levels or None,
                'exportName': label,
                'selectionCount': len(self.sel) if msg.get('selectionOnly') else 0,
                'servicesExcluded': list(msg.get('excludedServices') or []),
                'rodRows': mrform.rod_rows(r['rods']),
            })
            path = os.path.join(folder, '%s - %s MR.xlsx' % (base, safe_name(label)))
            with open(path, 'wb') as f:
                f.write(data)
            files.append(path)
            reports.append({'label': label + ' (' + str(len(r['rows'])) + ' rows)', 'notes': notes})
        return {'type': 'done', 'files': files, 'reports': reports}

    def sample(self, msg):
        folder = msg.get('folder') or os.environ.get('TEMP') or os.getcwd()
        if not os.path.isdir(folder):
            folder = os.environ.get('TEMP') or os.getcwd()
        path = os.path.join(folder, '%s - sample data.txt' % safe_name(model_base_name(self.title)))
        with open(path, 'w', encoding='utf-8') as f:
            f.write(self.reader.dump_sample(self.doc))
        return {'type': 'sample', 'path': path}


def kind_is_scene(progress):
    """Only geometry reading can be cancelled from the page."""
    return progress.get('stage') in ('Centerlines', 'Straights', 'Fittings', 'Valves', 'Hangers')
