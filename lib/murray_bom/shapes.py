# -*- coding: utf-8 -*-
"""Simple shapes for the first stages of the 3D view. Pure Python, no Revit: single-line centerlines for every part and a
cylinder for straight pipe. Points are flat x,y,z lists in the units of the points given (feet)."""
import math


def centerline_segments(origins):
    """The single-line centerline of a part from its connector origins [(x, y, z), ...], as a flat list of segment end points
    (two points per segment). Two connectors: one line between them. Three or more (a tee, a cross): a line from each
    connector to their average, so the branches meet. Fewer than two: nothing to draw."""
    n = len(origins)
    if n < 2:
        return []
    if n == 2:
        return list(origins[0]) + list(origins[1])
    cx = sum(o[0] for o in origins) / n
    cy = sum(o[1] for o in origins) / n
    cz = sum(o[2] for o in origins) / n
    out = []
    for o in origins:
        out += [o[0], o[1], o[2], cx, cy, cz]
    return out


def cylinder_mesh(p0, p1, radius, seg=10):
    """A capped cylinder from p0 to p1. None when the two points are the same."""
    (x0, y0, z0), (x1, y1, z1) = p0, p1
    d = [x1 - x0, y1 - y0, z1 - z0]
    length = math.sqrt(sum(c * c for c in d))
    if length < 1e-9 or radius <= 0:
        return None
    d = [c / length for c in d]
    a = [0, 0, 1] if abs(d[2]) < 0.9 else [1, 0, 0]
    u = [d[1] * a[2] - d[2] * a[1], d[2] * a[0] - d[0] * a[2], d[0] * a[1] - d[1] * a[0]]
    ul = math.sqrt(sum(c * c for c in u))
    u = [c / ul for c in u]
    v = [d[1] * u[2] - d[2] * u[1], d[2] * u[0] - d[0] * u[2], d[0] * u[1] - d[1] * u[0]]
    verts = []
    for k in range(seg):
        t = 2 * math.pi * k / seg
        o = [radius * (math.cos(t) * u[i] + math.sin(t) * v[i]) for i in range(3)]
        verts += [x0 + o[0], y0 + o[1], z0 + o[2], x1 + o[0], y1 + o[1], z1 + o[2]]
    c0, c1 = 2 * seg, 2 * seg + 1                    # cap centres
    verts += [x0, y0, z0, x1, y1, z1]
    tris = []
    for k in range(seg):
        a0, a1 = 2 * k, 2 * ((k + 1) % seg)
        tris += [a0, a1, a0 + 1, a1, a1 + 1, a0 + 1]
        tris += [c0, a1, a0, c1, a0 + 1, a1 + 1]
    return verts, tris


def barrel_radius(verts, p0, p1):
    """The outside radius of a straight pipe, measured from its real geometry: its vertices' distances from the axis (the line
    from p0 to p1). A tessellated straight has vertices only on its end rings, so every vertex is used. Distances are grouped
    into bins of 0.0005 ft and the largest bin that holds at least a ring's worth of vertices (6) wins, so a stray vertex is
    ignored and a hollow pipe gives its outside radius, not its bore. This is the outside diameter as modelled (insulation
    included when the part carries it), which a connector's nominal radius is not. A hubbed or flanged end would read slightly
    large. Returns None when there is nothing to measure."""
    ax, ay, az = p1[0] - p0[0], p1[1] - p0[1], p1[2] - p0[2]
    length = math.sqrt(ax * ax + ay * ay + az * az)
    if length < 1e-9:
        return None
    ux, uy, uz = ax / length, ay / length, az / length
    bins = {}
    for k in range(0, len(verts), 3):
        dx, dy, dz = verts[k] - p0[0], verts[k + 1] - p0[1], verts[k + 2] - p0[2]
        cx, cy, cz = dy * uz - dz * uy, dz * ux - dx * uz, dx * uy - dy * ux
        r = math.sqrt(cx * cx + cy * cy + cz * cz)               # distance from the axis
        if r < 1e-6:
            continue                                             # a cap centre on the axis
        b = bins.setdefault(int(r / 0.0005 + 0.5), [0, 0.0])
        b[0] += 1
        b[1] += r
    if not bins:
        return None
    ring = [k for k, b in bins.items() if b[0] >= 6]
    k = max(ring) if ring else max(bins)
    return bins[k][1] / bins[k][0]
