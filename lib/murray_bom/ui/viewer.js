// Minimal WebGL2 viewer for the Bulk BOM window. No libraries, works offline.
// A view arrives as a parts table (message scene) followed by geometry in build order (messages geom): single-line
// centerlines, then straight pipe as cylinders, then the real geometry of fittings, valves and hangers. Positions are float32
// xyz in feet relative to one fixed centre chosen by Python.
// Controls: left drag orbit, right or middle drag (or Shift + left drag) pan, wheel zoom, click pick,
// double click fit (to the part under the cursor, else everything).
var BomViewer = (function () {
  'use strict';
  var gl, canvas, prog, loc = {};
  // Two growing layers. LL holds centerlines (drawn as GL lines), LS holds solids (cylinders, then real geometry). A part that gets
  // a solid has its centerline hidden (alpha 0).
  var LL = null, LS = null;
  var parts = [], groups = [], indexById = {};
  var autoFit = true, framed = false;   // frame the model as geometry arrives, until the user touches the camera
  var cam = { az: -0.8, el: 0.6, d: 50, t: [0, 0, 0] };
  var radius = 10, mode = 'export', hiddenCats = {}, picked = -1, ui = {}, dirty = true, hasScene = false;
  var excluded = {};      // service filter: services unticked in the legend are left out of the export and hidden here
  var center = null;      // Revit coordinates of the scene origin; the camera is kept through updates (see load)

  var GROUP_COLORS = { 'Bulk Fab Copper': '#E07B00', 'Bulk Fab Non Copper': '#1F77B4', 'Bulk Fab Valve': '#D62728', 'Bulk Fab Hanger': '#2CA02C' };
  var NO_COLOR = [160, 160, 160];     // a service with no entry in systemColors.js
  var YELLOW = [255, 231, 22];

  var VS = '#version 300 es\nin vec3 aPos; in vec4 aCol; uniform mat4 uV; uniform mat4 uP; out vec3 vV; out vec4 vC;\n' +
    'void main(){ vec4 v = uV * vec4(aPos, 1.0); vV = v.xyz; vC = aCol; gl_Position = uP * v; }';
  var FS = '#version 300 es\nprecision highp float; uniform float uFlat; in vec3 vV; in vec4 vC; out vec4 o;\n' +
    'void main(){ if (vC.a < 0.5) discard; if (uFlat > 0.5) { o = vec4(vC.rgb, 1.0); return; } vec3 n = normalize(cross(dFdx(vV), dFdy(vV)));\n' +
    ' float h = abs(n.z); float k = abs(dot(n, normalize(vec3(-0.4, 0.6, 0.7)))); float l = 0.35 + 0.4 * h + 0.25 * k;\n' +
    ' o = vec4(vC.rgb * l, 1.0); }';

  // ---------- small matrix helpers (column major) ----------
  function perspective(fovy, aspect, n, f) {
    var t = 1 / Math.tan(fovy / 2), m = new Float32Array(16);
    m[0] = t / aspect; m[5] = t; m[10] = (f + n) / (n - f); m[11] = -1; m[14] = 2 * f * n / (n - f);
    return m;
  }
  function cross(a, b) { return [a[1] * b[2] - a[2] * b[1], a[2] * b[0] - a[0] * b[2], a[0] * b[1] - a[1] * b[0]]; }
  function norm(a) { var l = Math.hypot(a[0], a[1], a[2]) || 1; return [a[0] / l, a[1] / l, a[2] / l]; }
  function sub(a, b) { return [a[0] - b[0], a[1] - b[1], a[2] - b[2]]; }
  function dot(a, b) { return a[0] * b[0] + a[1] * b[1] + a[2] * b[2]; }
  function lookAt(e, c, up) {
    var f = norm(sub(c, e)), s = norm(cross(f, up)), u = cross(s, f), m = new Float32Array(16);
    m[0] = s[0]; m[4] = s[1]; m[8] = s[2]; m[1] = u[0]; m[5] = u[1]; m[9] = u[2];
    m[2] = -f[0]; m[6] = -f[1]; m[10] = -f[2]; m[12] = -dot(s, e); m[13] = -dot(u, e); m[14] = dot(f, e); m[15] = 1;
    return m;
  }
  function mul(a, b) {
    var o = new Float32Array(16);
    for (var c = 0; c < 4; c++) for (var r = 0; r < 4; r++) {
      var s = 0; for (var k = 0; k < 4; k++) s += a[k * 4 + r] * b[c * 4 + k]; o[c * 4 + r] = s;
    }
    return o;
  }
  function invert(m) {
    var a = m, inv = new Float64Array(16), det, i;
    inv[0] = a[5]*a[10]*a[15]-a[5]*a[11]*a[14]-a[9]*a[6]*a[15]+a[9]*a[7]*a[14]+a[13]*a[6]*a[11]-a[13]*a[7]*a[10];
    inv[4] = -a[4]*a[10]*a[15]+a[4]*a[11]*a[14]+a[8]*a[6]*a[15]-a[8]*a[7]*a[14]-a[12]*a[6]*a[11]+a[12]*a[7]*a[10];
    inv[8] = a[4]*a[9]*a[15]-a[4]*a[11]*a[13]-a[8]*a[5]*a[15]+a[8]*a[7]*a[13]+a[12]*a[5]*a[11]-a[12]*a[7]*a[9];
    inv[12] = -a[4]*a[9]*a[14]+a[4]*a[10]*a[13]+a[8]*a[5]*a[14]-a[8]*a[6]*a[13]-a[12]*a[5]*a[10]+a[12]*a[6]*a[9];
    inv[1] = -a[1]*a[10]*a[15]+a[1]*a[11]*a[14]+a[9]*a[2]*a[15]-a[9]*a[3]*a[14]-a[13]*a[2]*a[11]+a[13]*a[3]*a[10];
    inv[5] = a[0]*a[10]*a[15]-a[0]*a[11]*a[14]-a[8]*a[2]*a[15]+a[8]*a[3]*a[14]+a[12]*a[2]*a[11]-a[12]*a[3]*a[10];
    inv[9] = -a[0]*a[9]*a[15]+a[0]*a[11]*a[13]+a[8]*a[1]*a[15]-a[8]*a[3]*a[13]-a[12]*a[1]*a[11]+a[12]*a[3]*a[9];
    inv[13] = a[0]*a[9]*a[14]-a[0]*a[10]*a[13]-a[8]*a[1]*a[14]+a[8]*a[2]*a[13]+a[12]*a[1]*a[10]-a[12]*a[2]*a[9];
    inv[2] = a[1]*a[6]*a[15]-a[1]*a[7]*a[14]-a[5]*a[2]*a[15]+a[5]*a[3]*a[14]+a[13]*a[2]*a[7]-a[13]*a[3]*a[6];
    inv[6] = -a[0]*a[6]*a[15]+a[0]*a[7]*a[14]+a[4]*a[2]*a[15]-a[4]*a[3]*a[14]-a[12]*a[2]*a[7]+a[12]*a[3]*a[6];
    inv[10] = a[0]*a[5]*a[15]-a[0]*a[7]*a[13]-a[4]*a[1]*a[15]+a[4]*a[3]*a[13]+a[12]*a[1]*a[7]-a[12]*a[3]*a[5];
    inv[14] = -a[0]*a[5]*a[14]+a[0]*a[6]*a[13]+a[4]*a[1]*a[14]-a[4]*a[2]*a[13]-a[12]*a[1]*a[6]+a[12]*a[2]*a[5];
    inv[3] = -a[1]*a[6]*a[11]+a[1]*a[7]*a[10]+a[5]*a[2]*a[11]-a[5]*a[3]*a[10]-a[9]*a[2]*a[7]+a[9]*a[3]*a[6];
    inv[7] = a[0]*a[6]*a[11]-a[0]*a[7]*a[10]-a[4]*a[2]*a[11]+a[4]*a[3]*a[10]+a[8]*a[2]*a[7]-a[8]*a[3]*a[6];
    inv[11] = -a[0]*a[5]*a[11]+a[0]*a[7]*a[9]+a[4]*a[1]*a[11]-a[4]*a[3]*a[9]-a[8]*a[1]*a[7]+a[8]*a[3]*a[5];
    inv[15] = a[0]*a[5]*a[10]-a[0]*a[6]*a[9]-a[4]*a[1]*a[10]+a[4]*a[2]*a[9]+a[8]*a[1]*a[6]-a[8]*a[2]*a[5];
    det = a[0]*inv[0]+a[1]*inv[4]+a[2]*inv[8]+a[3]*inv[12];
    if (!det) return null;
    det = 1 / det; for (i = 0; i < 16; i++) inv[i] *= det;
    return inv;
  }
  function hex(h) { return [parseInt(h.substr(1, 2), 16), parseInt(h.substr(3, 2), 16), parseInt(h.substr(5, 2), 16)]; }
  // Service colors come from systemColors.js (the list from the web app): the service name, any case, to [r, g, b].
  function serviceColor(name) {
    var key = (window.normalizeService ? window.normalizeService(name) : String(name || '').toUpperCase());
    return (window.SYSTEM_COLORS_BY_NAME && window.SYSTEM_COLORS_BY_NAME[key]) || null;
  }

  // ---------- GL setup ----------
  function shader(type, src) {
    var s = gl.createShader(type); gl.shaderSource(s, src); gl.compileShader(s);
    if (!gl.getShaderParameter(s, gl.COMPILE_STATUS)) throw new Error(gl.getShaderInfoLog(s));
    return s;
  }
  function init(opts) {
    ui = opts; canvas = opts.canvas;
    gl = canvas.getContext('webgl2', { antialias: true });
    if (!gl) { ui.status('3D preview needs WebGL2, which this WebView2 did not provide.'); return false; }
    prog = gl.createProgram();
    gl.attachShader(prog, shader(gl.VERTEX_SHADER, VS)); gl.attachShader(prog, shader(gl.FRAGMENT_SHADER, FS));
    gl.linkProgram(prog);
    if (!gl.getProgramParameter(prog, gl.LINK_STATUS)) throw new Error(gl.getProgramInfoLog(prog));
    loc = { pos: gl.getAttribLocation(prog, 'aPos'), col: gl.getAttribLocation(prog, 'aCol'),
            v: gl.getUniformLocation(prog, 'uV'), p: gl.getUniformLocation(prog, 'uP'), flat: gl.getUniformLocation(prog, 'uFlat') };
    LL = makeLayer(); LS = makeLayer();
    gl.enable(gl.DEPTH_TEST);
    wire();
    new ResizeObserver(function () { dirty = true; }).observe(canvas);
    (function loop() { if (dirty) draw(); requestAnimationFrame(loop); })();
    return true;
  }

  function makeLayer() {
    return { pos: new Float32Array(0), idx: new Uint32Array(0), col: new Uint8Array(0), nv: 0, ni: 0,
             vbo: gl.createBuffer(), ibo: gl.createBuffer(), cbo: gl.createBuffer() };
  }
  function freeLayer(L) { gl.deleteBuffer(L.vbo); gl.deleteBuffer(L.ibo); gl.deleteBuffer(L.cbo); }
  function uploadFull(L, usage) {
    gl.bindBuffer(gl.ARRAY_BUFFER, L.vbo); gl.bufferData(gl.ARRAY_BUFFER, L.pos, usage);
    gl.bindBuffer(gl.ARRAY_BUFFER, L.cbo); gl.bufferData(gl.ARRAY_BUFFER, L.col, gl.DYNAMIC_DRAW);
    gl.bindBuffer(gl.ELEMENT_ARRAY_BUFFER, L.ibo); gl.bufferData(gl.ELEMENT_ARRAY_BUFFER, L.idx, usage);
  }
  // Makes room in a layer for more vertices and indices. Returns true when the arrays were replaced (a full upload is needed).
  function ensureCap(L, needV, needI) {
    var capV = L.pos.length / 3, capI = L.idx.length, grew = false;
    if (needV > capV) {
      var nv = Math.max(needV, capV * 2, 1 << 15), p = new Float32Array(nv * 3), c = new Uint8Array(nv * 4);
      p.set(L.pos); c.set(L.col); L.pos = p; L.col = c; grew = true;
    }
    if (needI > capI) {
      var ni = Math.max(needI, capI * 2, 1 << 16), x = new Uint32Array(ni);
      x.set(L.idx); L.idx = x; grew = true;
    }
    return grew;
  }
  function matrices() {
    var w = canvas.clientWidth || 1, h = canvas.clientHeight || 1;
    var e = [cam.t[0] + cam.d * Math.cos(cam.el) * Math.cos(cam.az), cam.t[1] + cam.d * Math.cos(cam.el) * Math.sin(cam.az), cam.t[2] + cam.d * Math.sin(cam.el)];
    var V = lookAt(e, cam.t, [0, 0, 1]);
    var P = perspective(0.8, w / h, Math.max(cam.d * 0.02, 0.01), cam.d + radius * 4);
    return { V: V, P: P, eye: e };
  }

  function draw() {
    dirty = false;
    var dpr = window.devicePixelRatio || 1, w = Math.round(canvas.clientWidth * dpr), h = Math.round(canvas.clientHeight * dpr);
    if (canvas.width !== w || canvas.height !== h) { canvas.width = w; canvas.height = h; }
    gl.viewport(0, 0, w, h);
    gl.clearColor(1, 1, 1, 1); gl.clear(gl.COLOR_BUFFER_BIT | gl.DEPTH_BUFFER_BIT);
    if (!hasScene) return;
    var m = matrices();
    gl.useProgram(prog);
    gl.uniformMatrix4fv(loc.v, false, m.V); gl.uniformMatrix4fv(loc.p, false, m.P);
    function bind(L, withIdx) {
      gl.bindBuffer(gl.ARRAY_BUFFER, L.vbo); gl.enableVertexAttribArray(loc.pos); gl.vertexAttribPointer(loc.pos, 3, gl.FLOAT, false, 0, 0);
      gl.bindBuffer(gl.ARRAY_BUFFER, L.cbo); gl.enableVertexAttribArray(loc.col); gl.vertexAttribPointer(loc.col, 4, gl.UNSIGNED_BYTE, true, 0, 0);
      if (withIdx) gl.bindBuffer(gl.ELEMENT_ARRAY_BUFFER, L.ibo);
    }
    if (LS.ni) { gl.uniform1f(loc.flat, 0); bind(LS, true); gl.drawElements(gl.TRIANGLES, LS.ni, gl.UNSIGNED_INT, 0); }
    if (LL.nv) { gl.uniform1f(loc.flat, 1); bind(LL, false); gl.drawArrays(gl.LINES, 0, LL.nv); }
  }

  // ---------- scene ----------
  function b64ToBuffer(s) {
    var bin = atob(s), n = bin.length, u = new Uint8Array(n);
    for (var i = 0; i < n; i++) u[i] = bin.charCodeAt(i);
    return u.buffer;
  }
  function serviceLabel(p) { return p.service || '(no service)'; }
  // Hidden = left out by the service filter, or (when colored by export) hidden with its export's legend checkbox.
  function isHidden(p) { return !!excluded[serviceLabel(p)] || (mode === 'export' && !!hiddenCats[p.group]); }
  function excludedServices() { return Object.keys(excluded).filter(function (k) { return excluded[k]; }); }
  function filterChanged() {
    if (picked >= 0 && parts[picked] && isHidden(parts[picked])) { picked = -1; info(null); }     // a hidden part is no longer picked
    recolor(); legend(); if (ui.onFilter) ui.onFilter(excludedServices());
  }
  function clearExcluded() { excluded = {}; filterChanged(); }
  function category(p) {
    if (mode === 'service') return p.service || '(no service)';
    return p.group;
  }
  function baseColor(p) {
    if (mode === 'service') return serviceColor(p.service) || NO_COLOR;
    return hex(GROUP_COLORS[p.group] || '#7F7F7F');
  }
  function paintRange(L, v, rgb, alpha) {
    for (var k = 0; k < v[1]; k++) {
      var o = (v[0] + k) * 4; L.col[o] = rgb[0]; L.col[o + 1] = rgb[1]; L.col[o + 2] = rgb[2]; L.col[o + 3] = alpha;
    }
  }
  // A part with a solid is drawn from LS and its centerline in LL is hidden.
  function paintPart(i, rgb, alpha) {
    var p = parts[i];
    if (p.ln) paintRange(LL, p.ln.v, rgb, p.sd ? 0 : alpha);
    if (p.sd) paintRange(LS, p.sd.v, rgb, alpha);
  }
  function paintOne(n) { var p = parts[n]; paintPart(n, n === picked ? YELLOW : baseColor(p), isHidden(p) ? 0 : 255); }
  function recolor() {
    if (!parts.length) return;
    for (var i = 0; i < parts.length; i++) paintOne(i);
    if (LL.nv) { gl.bindBuffer(gl.ARRAY_BUFFER, LL.cbo); gl.bufferData(gl.ARRAY_BUFFER, LL.col, gl.DYNAMIC_DRAW); }
    if (LS.nv) { gl.bindBuffer(gl.ARRAY_BUFFER, LS.cbo); gl.bufferData(gl.ARRAY_BUFFER, LS.col, gl.DYNAMIC_DRAW); }
    dirty = true;
  }
  function bboxOf(L, v) {
    var lo = [1e30, 1e30, 1e30], hi = [-1e30, -1e30, -1e30];
    for (var k = 0; k < v[1]; k++) {
      var o = (v[0] + k) * 3;
      for (var a = 0; a < 3; a++) { var x = L.pos[o + a]; if (x < lo[a]) lo[a] = x; if (x > hi[a]) hi[a] = x; }
    }
    return [lo, hi];
  }

  function load(scene) {
    // A new view: the parts table only (geometry follows in geom messages). The camera, the hidden legend items, the service
    // filter and the picked part are kept.
    var keepId = picked >= 0 && parts[picked] ? parts[picked].id : null;
    parts = scene.parts; groups = scene.groups; picked = -1;
    freeLayer(LL); freeLayer(LS);
    LL = makeLayer(); LS = makeLayer();
    indexById = {};
    parts.forEach(function (p, n) {
      p.ln = null; p.sd = null; p.lo = null; p.hi = null; indexById[p.id] = n;
      if (keepId !== null && p.id === keepId) picked = n;
    });
    hasScene = parts.length > 0;
    legend(); info(picked >= 0 ? picked : null); dirty = true;
  }

  function adoptCentre(c) {
    if (!c) return;
    if (center !== null && (center[0] !== c[0] || center[1] !== c[1] || center[2] !== c[2])) {
      for (var a = 0; a < 3; a++) cam.t[a] += center[a] - c[a];     // the origin moved: move the camera target with it
    }
    center = c.slice();
  }

  // Geometry for some parts, in build order: kind 'lines' (centerlines) or 'solid' (cylinders, then real geometry). Appends to a
  // layer and leaves the camera, legend choices and picked part alone. Unknown ids (a view that has since changed) are ignored.
  function geom(m) {
    if (!hasScene || !m.parts || !m.parts.length) return 0;
    adoptCentre(m.center);
    var n = m.kind === 'lines' ? appendLines(m) : appendSolids(m);
    if (autoFit && n) { if (!framed) { iso(); framed = true; } fit(); }
    dirty = true;
    return n;
  }
  function appendLines(m) {
    var pp = new Float32Array(b64ToBuffer(m.pos)), add = pp.length / 3, v0 = LL.nv;
    var grew = ensureCap(LL, v0 + add, 0);
    LL.pos.set(pp, v0 * 3); LL.nv = v0 + add;
    var changed = [];
    m.parts.forEach(function (e) {
      var n = indexById[e.id];
      if (n === undefined || parts[n].ln) return;
      var p = parts[n]; p.ln = { v: [v0 + e.v[0], e.v[1]] };
      if (!p.sd) { var bb = bboxOf(LL, p.ln.v); p.lo = bb[0]; p.hi = bb[1]; }
      changed.push(n);
    });
    changed.forEach(paintOne);
    if (grew) {
      uploadFull(LL, gl.DYNAMIC_DRAW);
    } else {
      gl.bindBuffer(gl.ARRAY_BUFFER, LL.vbo); gl.bufferSubData(gl.ARRAY_BUFFER, v0 * 12, LL.pos.subarray(v0 * 3, LL.nv * 3));
      gl.bindBuffer(gl.ARRAY_BUFFER, LL.cbo); gl.bufferSubData(gl.ARRAY_BUFFER, v0 * 4, LL.col.subarray(v0 * 4, LL.nv * 4));
    }
    return changed.length;
  }
  function appendSolids(m) {
    var pp = new Float32Array(b64ToBuffer(m.pos)), pi = new Uint32Array(b64ToBuffer(m.idx));
    var add = pp.length / 3, niAdd = pi.length, v0 = LS.nv, i0 = LS.ni;
    var grew = ensureCap(LS, v0 + add, i0 + niAdd);
    LS.pos.set(pp, v0 * 3);
    for (var k = 0; k < niAdd; k++) LS.idx[i0 + k] = pi[k] + v0;
    LS.nv = v0 + add; LS.ni = i0 + niAdd;
    var changed = [];
    m.parts.forEach(function (e) {
      var n = indexById[e.id];
      if (n === undefined || parts[n].sd) return;
      var p = parts[n];
      p.sd = { v: [v0 + e.v[0], e.v[1]], i: [i0 + e.i[0], e.i[1]], lod: m.lod };
      var bb = bboxOf(LS, p.sd.v); p.lo = bb[0]; p.hi = bb[1];
      changed.push(n);
    });
    changed.forEach(paintOne);
    if (grew) {
      uploadFull(LS, gl.DYNAMIC_DRAW);
    } else {
      gl.bindBuffer(gl.ARRAY_BUFFER, LS.vbo); gl.bufferSubData(gl.ARRAY_BUFFER, v0 * 12, LS.pos.subarray(v0 * 3, LS.nv * 3));
      gl.bindBuffer(gl.ARRAY_BUFFER, LS.cbo); gl.bufferSubData(gl.ARRAY_BUFFER, v0 * 4, LS.col.subarray(v0 * 4, LS.nv * 4));
      gl.bindBuffer(gl.ELEMENT_ARRAY_BUFFER, LS.ibo); gl.bufferSubData(gl.ELEMENT_ARRAY_BUFFER, i0 * 4, LS.idx.subarray(i0, LS.ni));
    }
    gl.bindBuffer(gl.ARRAY_BUFFER, LL.cbo);                     // centerlines hidden under their new solids
    changed.forEach(function (n) {
      var p = parts[n]; if (!p.ln) return;
      var v = p.ln.v; gl.bufferSubData(gl.ARRAY_BUFFER, v[0] * 4, LL.col.subarray(v[0] * 4, (v[0] + v[1]) * 4));
    });
    if (picked >= 0 && changed.indexOf(picked) >= 0) info(picked);
    return changed.length;
  }

  function fitBox(mn, mx) {
    cam.t = [(mn[0] + mx[0]) / 2, (mn[1] + mx[1]) / 2, (mn[2] + mx[2]) / 2];
    radius = Math.max(Math.hypot(mx[0] - mn[0], mx[1] - mn[1], mx[2] - mn[2]) / 2, 0.5);
    cam.d = radius * 2.4; dirty = true;
  }
  function allBox() {
    var mn = [1e30, 1e30, 1e30], mx = [-1e30, -1e30, -1e30];
    parts.forEach(function (p) { if (!p.lo || isHidden(p)) return; for (var a = 0; a < 3; a++) { if (p.lo[a] < mn[a]) mn[a] = p.lo[a]; if (p.hi[a] > mx[a]) mx[a] = p.hi[a]; } });
    return mn[0] > mx[0] ? null : [mn, mx];
  }
  function fit() { var b = allBox(); if (b) fitBox(b[0], b[1]); }
  function iso() { cam.az = -0.8; cam.el = 0.6; dirty = true; }
  function top() { cam.az = -Math.PI / 2; cam.el = 1.5707; dirty = true; }
  function front() { cam.az = -Math.PI / 2; cam.el = 0.0; dirty = true; }

  // ---------- legend and info panel ----------
  function legend() {
    var box = ui.legend; box.innerHTML = '';
    var counts = {}, order = [];
    parts.forEach(function (p) { var c = category(p); if (!(c in counts)) { counts[c] = 0; order.push(c); } counts[c]++; });
    order.sort();
    var byService = mode === 'service';
    order.forEach(function (c) {
      var l = document.createElement('label'), cb = document.createElement('input'), sw = document.createElement('span');
      cb.type = 'checkbox'; cb.checked = byService ? !excluded[c] : !hiddenCats[c];
      // By service the checkboxes are the export filter; by export they only hide a group in the view
      // (the Exports list on the left decides what is exported).
      cb.onchange = function () {
        if (byService) { excluded[c] = !cb.checked; filterChanged(); } else { hiddenCats[c] = !cb.checked; recolor(); }
      };
      var rep = parts.filter(function (p) { return category(p) === c; })[0];
      var rgb = baseColor(rep); sw.className = 'sw'; sw.style.background = 'rgb(' + rgb.join(',') + ')';
      var missing = byService && !serviceColor(rep.service);
      l.appendChild(cb); l.appendChild(sw); l.appendChild(document.createTextNode(c + ' (' + counts[c] + ')' + (missing ? ' - no color set' : '')));
      box.appendChild(l);
    });
    if (!order.length) return;
    var n = excludedServices().length, note = document.createElement('div');
    note.className = 'lnote';
    if (byService) {
      note.textContent = 'Unticked services are left out of the export.';
    } else if (n) {
      note.textContent = n + ' service' + (n === 1 ? '' : 's') + ' left out of the export. Color by Service to change.';
    } else { return; }
    box.appendChild(note);
    if (n) {
      var b = document.createElement('button'); b.type = 'button'; b.textContent = 'Include all services'; b.onclick = clearExcluded;
      box.appendChild(b);
    }
  }
  function info(i) {
    var el = ui.info;
    if (i === null || i < 0) { el.hidden = true; return; }
    var p = parts[i], rows = [['Part', p.name], ['Export', p.group], ['Size', p.size], ['Length', p.length], ['Material', p.material], ['Level', p.level], ['Service', p.service], ['Geometry', p.sd ? (p.sd.lod === 'cyl' ? 'cylinder' : 'detailed') : (p.ln ? 'centerline only, solid coming' : 'not built yet')], ['Element id', p.id]];
    el.innerHTML = '';
    rows.forEach(function (r) {
      if (r[1] === '' || r[1] === undefined) return;
      var d = document.createElement('div'), b = document.createElement('b');
      b.textContent = r[0] + ': '; d.appendChild(b); d.appendChild(document.createTextNode(String(r[1]))); el.appendChild(d);
    });
    el.hidden = false;
  }

  // ---------- picking ----------
  function rayAt(x, y) {
    var m = matrices(), inv = invert(mul(m.P, m.V)); if (!inv) return null;
    var nx = (x / canvas.clientWidth) * 2 - 1, ny = 1 - (y / canvas.clientHeight) * 2;
    function un(z) {
      var v = [nx, ny, z, 1], o = [0, 0, 0, 0];
      for (var r = 0; r < 4; r++) o[r] = inv[r] * v[0] + inv[4 + r] * v[1] + inv[8 + r] * v[2] + inv[12 + r] * v[3];
      return [o[0] / o[3], o[1] / o[3], o[2] / o[3]];
    }
    var a = un(-1), b = un(1);
    return { o: m.eye, d: norm(sub(b, a)) };
  }
  function hitBox(r, lo, hi) {
    var tmin = 0, tmax = 1e30;
    for (var a = 0; a < 3; a++) {
      if (Math.abs(r.d[a]) < 1e-12) { if (r.o[a] < lo[a] || r.o[a] > hi[a]) return -1; continue; }
      var t1 = (lo[a] - r.o[a]) / r.d[a], t2 = (hi[a] - r.o[a]) / r.d[a];
      if (t1 > t2) { var s = t1; t1 = t2; t2 = s; }
      if (t1 > tmin) tmin = t1; if (t2 < tmax) tmax = t2; if (tmin > tmax) return -1;
    }
    return tmin;
  }
  function hitPart(r, p) {
    var pos = LS.pos, idx = LS.idx;
    var best = -1, i0 = p.sd.i[0], n = p.sd.i[1];
    for (var k = 0; k < n; k += 3) {
      var a = idx[i0 + k] * 3, b = idx[i0 + k + 1] * 3, c = idx[i0 + k + 2] * 3;
      var e1 = [pos[b] - pos[a], pos[b + 1] - pos[a + 1], pos[b + 2] - pos[a + 2]];
      var e2 = [pos[c] - pos[a], pos[c + 1] - pos[a + 1], pos[c + 2] - pos[a + 2]];
      var h = cross(r.d, e2), det = dot(e1, h); if (Math.abs(det) < 1e-12) continue;
      var f = 1 / det, s = [r.o[0] - pos[a], r.o[1] - pos[a + 1], r.o[2] - pos[a + 2]];
      var u = f * dot(s, h); if (u < 0 || u > 1) continue;
      var q = cross(s, e1), v = f * dot(r.d, q); if (v < 0 || u + v > 1) continue;
      var t = f * dot(e2, q); if (t > 0 && (best < 0 || t < best)) best = t;
    }
    return best;
  }
  // A part that only has its centerline so far: hit when the ray passes within tol of one of its segments.
  function hitLine(r, p, tol) {
    var pos = LL.pos, best = -1, v0 = p.ln.v[0], n = p.ln.v[1];
    for (var k = 0; k + 1 < n; k += 2) {
      var a = (v0 + k) * 3, b = (v0 + k + 1) * 3;
      var A = [pos[a], pos[a + 1], pos[a + 2]], v = [pos[b] - A[0], pos[b + 1] - A[1], pos[b + 2] - A[2]];
      var w = [r.o[0] - A[0], r.o[1] - A[1], r.o[2] - A[2]];
      var bb = dot(r.d, v), c = dot(v, v), d0 = dot(r.d, w), e0 = dot(v, w), den = c - bb * bb;
      var tc = den < 1e-12 ? (c > 0 ? e0 / c : 0) : (e0 - bb * d0) / den;      // along the segment
      tc = Math.max(0, Math.min(1, tc));
      var sc = bb * tc - d0;                                                   // along the ray
      if (sc <= 0) continue;
      var dx = r.o[0] + sc * r.d[0] - (A[0] + tc * v[0]), dy = r.o[1] + sc * r.d[1] - (A[1] + tc * v[1]), dz = r.o[2] + sc * r.d[2] - (A[2] + tc * v[2]);
      if (Math.hypot(dx, dy, dz) < tol && (best < 0 || sc < best)) best = sc;
    }
    return best;
  }
  function pick(x, y) {
    var r = rayAt(x, y); if (!r) return -1;
    var tol = cam.d * 0.006, cands = [];
    for (var i = 0; i < parts.length; i++) {
      var p = parts[i];
      if (!p.lo || isHidden(p)) continue;
      var lo = p.lo, hi = p.hi;
      if (!p.sd) { lo = [lo[0] - tol, lo[1] - tol, lo[2] - tol]; hi = [hi[0] + tol, hi[1] + tol, hi[2] + tol]; }   // a line has no thickness
      var t = hitBox(r, lo, hi); if (t >= 0) cands.push([t, i]);
    }
    cands.sort(function (a, b) { return a[0] - b[0]; });
    var bestT = 1e30, bestI = -1;
    for (var k = 0; k < cands.length; k++) {
      if (cands[k][0] > bestT) break;
      var q = parts[cands[k][1]], t2 = q.sd ? hitPart(r, q) : hitLine(r, q, tol);
      if (t2 >= 0 && t2 < bestT) { bestT = t2; bestI = cands[k][1]; }
    }
    return bestI;
  }
  function select(i) {
    if (i === picked) i = -1;
    picked = i; recolor(); info(i);
  }

  // ---------- mouse ----------
  function wire() {
    var drag = null;
    canvas.addEventListener('contextmenu', function (e) { e.preventDefault(); });
    canvas.addEventListener('pointerdown', function (e) {
      autoFit = false;
      canvas.setPointerCapture(e.pointerId);
      drag = { x: e.clientX, y: e.clientY, sx: e.clientX, sy: e.clientY, pan: e.button === 1 || e.button === 2 || e.shiftKey, btn: e.button };
    });
    canvas.addEventListener('pointermove', function (e) {
      if (!drag) return;
      var dx = e.clientX - drag.x, dy = e.clientY - drag.y; drag.x = e.clientX; drag.y = e.clientY;
      if (drag.pan) {
        var m = matrices(), f = norm(sub(cam.t, m.eye)), s = norm(cross(f, [0, 0, 1])), u = cross(s, f);
        var k = cam.d * 0.0016;
        for (var a = 0; a < 3; a++) cam.t[a] += (-dx * s[a] + dy * u[a]) * k;
      } else {
        cam.az -= dx * 0.008; cam.el = Math.max(-1.5707, Math.min(1.5707, cam.el + dy * 0.008));
      }
      dirty = true;
    });
    canvas.addEventListener('pointerup', function (e) {
      if (drag && drag.btn === 0 && Math.hypot(e.clientX - drag.sx, e.clientY - drag.sy) < 4) {
        var r = canvas.getBoundingClientRect(); select(pick(e.clientX - r.left, e.clientY - r.top));
      }
      drag = null;
    });
    canvas.addEventListener('wheel', function (e) {
      autoFit = false; e.preventDefault(); cam.d = Math.max(radius * 0.02, Math.min(radius * 20, cam.d * Math.exp(e.deltaY * 0.0012))); dirty = true;
    }, { passive: false });
    canvas.addEventListener('dblclick', function (e) {
      var r = canvas.getBoundingClientRect(), i = pick(e.clientX - r.left, e.clientY - r.top);
      if (i >= 0) { var p = parts[i]; fitBox(p.lo, p.hi); cam.d = Math.max(cam.d, radius * 3); } else fit();
    });
  }

  function ready() { return !!gl; }
  function setMode(m) { mode = m; hiddenCats = {}; recolor(); legend(); if (ui.onFilter) ui.onFilter(excludedServices()); }
  function clear() { parts = []; hasScene = false; if (LL) { LL.nv = 0; LS.ni = 0; LS.nv = 0; } picked = -1; ui.legend.innerHTML = ''; info(null); dirty = true; }

  return { init: init, ready: ready, geom: geom, stopAutoFit: function () { autoFit = false; }, excludedServices: excludedServices, clearExcluded: clearExcluded, load: load, fit: fit, iso: iso, top: top, front: front, setMode: setMode, clear: clear,
           // for tests
           _debug: function () { return { parts: parts, picked: picked, cam: cam, pick: pick, visible: parts.filter(function (p) { return !isHidden(p); }).length,
      lines: parts.filter(function (p) { return p.ln; }).length,
      solids: parts.filter(function (p) { return p.sd; }).length,
      cyl: parts.filter(function (p) { return p.sd && p.sd.lod === 'cyl'; }).length,
      detailed: parts.filter(function (p) { return p.sd && p.sd.lod === 'detail'; }).length,
      onlyLine: parts.filter(function (p) { return p.ln && !p.sd; }).length,
      none: parts.filter(function (p) { return !p.ln && !p.sd; }).length,
      nvL: LL.nv, nvS: LS.nv, autoFit: autoFit }; } };
})();
