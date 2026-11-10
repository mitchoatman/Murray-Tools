# coding: utf8
# Re-host a fabrication tap onto another fabrication straight (duct or pipe).
# Pick 1: the tap to move (click near the edge that mates to the host).
# Pick 2: the straight to move it to (click on the face where the tap should land).
# DRAFT. Items marked [CONFIRM] need a test in your Fabrication config before field use.
import os
from math import pi, atan2, sqrt, cos, sin
from Autodesk.Revit.DB import (Transaction, SubTransaction, FabricationPart, XYZ,
                               ElementTransformUtils, ConnectorProfileType)
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

TOL = 1e-4  # feet


def dbg(msg):
    """Append a line to the log file right away, so the last line shows the call that crashed Revit."""
    try:
        with open(LOG_PATH, 'a') as f:
            f.write(msg + '\n')
    except Exception:
        pass
SEARCH_STEP = 0.1  # feet. Step used to slide the tap inward when a click is too close to an end.
SEARCH_MAX = 0.0  # feet. Sliding is OFF at 0. Try 1.0 once a run completes cleanly with the log on.
LOG_PATH = os.path.join(os.environ.get('TEMP', '.'), 'RehostTap_log.txt')  # last line survives a crash
SNAP_DEG = 15.0  # round hosts: snap the tap direction around the pipe to this step. 0 = no snap.


class TapFilter(ISelectionFilter):
    def AllowElement(self, elem):
        return isinstance(elem, FabricationPart) and elem.IsATap()

    def AllowReference(self, reference, position):
        return True


class StraightFilter(ISelectionFilter):
    def AllowElement(self, elem):
        return isinstance(elem, FabricationPart) and elem.IsAStraight()

    def AllowReference(self, reference, position):
        return True


def connectors_of(part):
    return [c for c in part.ConnectorManager.Connectors]


def straight_refs(conn, tap):
    """Connectors on other straights that this connector is mated to."""
    out = []
    for ref in conn.AllRefs:
        owner = ref.Owner
        if owner.Id != tap.Id and isinstance(owner, FabricationPart) and owner.IsAStraight():
            out.append(ref)
    return out


def tap_connector_candidates(tap, pick_pt):
    """Order of connectors to try as the tap's primary connector:
    1) any connector currently mated to a straight, 2) the rest, closest to the tap pick first.
    The API rejects a non-primary connector, so each is tried until one is accepted."""
    conns = connectors_of(tap)
    mated = [c for c in conns if c.IsConnected and straight_refs(c, tap)]
    others = [c for c in conns if c not in mated]
    others.sort(key=lambda c: c.Origin.DistanceTo(pick_pt))
    return mated + others


def nearest_quarter_turn(angle):
    """Rect and oval taps need a multiple of PI/2."""
    return round(angle / (pi / 2.0)) * (pi / 2.0)


def host_shape(conn):
    """'rect', 'round' or 'other' for the host connector profile."""
    try:
        if conn.Shape == ConnectorProfileType.Rectangular:
            return "rect"
        if conn.Shape == ConnectorProfileType.Round:
            return "round"
        return "other"
    except Exception:
        return "rect"  # [CONFIRM] assume rectangular if the shape cannot be read


def rehost():
    try:
        r1 = uidoc.Selection.PickObject(ObjectType.Element, TapFilter(),
                                        "Pick the tap to re-host. [ESC] to cancel")
    except:
        return False
    try:
        r2 = uidoc.Selection.PickObject(ObjectType.Element, StraightFilter(),
                                        "Pick the straight and click on the face where the tap should land")
    except:
        return False

    tap = doc.GetElement(r1)
    dbg('--- run: tap id {} host picked'.format(tap.Id.IntegerValue))
    tap_pt = r1.GlobalPoint
    host = doc.GetElement(r2)
    pt = r2.GlobalPoint

    host_conns = connectors_of(host)
    if len(host_conns) < 2:
        TaskDialog.Show("Re-host Tap", "Target straight does not have two connectors.")
        return True

    # Host axis, measured from one host connector
    hc = host_conns[0]
    other = host_conns[1]
    axis = (other.Origin - hc.Origin)
    length = axis.GetLength()
    axis = axis.Normalize()
    cs = hc.CoordinateSystem
    shape = host_shape(hc)
    host_r = None
    if shape == "round":
        try:
            host_r = hc.Radius
        except Exception:
            host_r = None
    try:
        vd = uidoc.ActiveView.ViewDirection.Normalize()
    except Exception:
        vd = None

    def solve_on_cylinder():
        """Round host: intersect the view ray through the click with the pipe surface.
        This does not rely on the click's depth, so a click anywhere on the visible pipe,
        edges included, gives a real surface point. Returns (distance along axis, outward direction)."""
        if vd is None or host_r is None:
            return None
        u = pt - hc.Origin
        u_p = u - axis.Multiply(u.DotProduct(axis))
        v_p = vd - axis.Multiply(vd.DotProduct(axis))
        a2 = v_p.DotProduct(v_p)
        if a2 < 1e-6:
            return None  # looking along the pipe
        b2 = 2.0 * u_p.DotProduct(v_p)
        c2 = u_p.DotProduct(u_p) - host_r * host_r
        disc = b2 * b2 - 4.0 * a2 * c2
        if disc >= 0:
            sp = (-b2 - sqrt(disc)) / (2.0 * a2)  # near hit, the face toward the viewer
        else:
            sp = -b2 / (2.0 * a2)  # click just outside the silhouette: closest approach
        hit = pt + vd.Multiply(sp)
        dist = axis.DotProduct(hit - hc.Origin)
        w = (hit - hc.Origin) - axis.Multiply(dist)
        if w.GetLength() < 1e-9:
            return None
        return dist, w.Normalize()

    def snap_about_axis(v):
        """Snap direction v (perpendicular to the axis) to SNAP_DEG steps about the axis, from 'up'."""
        if SNAP_DEG <= 0:
            return v
        ref = XYZ.BasisZ - axis.Multiply(XYZ.BasisZ.DotProduct(axis))
        if ref.GetLength() < 0.1:
            ref = XYZ.BasisX - axis.Multiply(XYZ.BasisX.DotProduct(axis))
        ref = ref.Normalize()
        side = axis.CrossProduct(ref)
        ang = atan2(v.DotProduct(side), v.DotProduct(ref))
        step = SNAP_DEG * pi / 180.0
        ang = round(ang / step) * step
        return ref.Multiply(cos(ang)) + side.Multiply(sin(ang))

    solved = solve_on_cylinder() if shape == "round" else None
    if solved:
        distance, n = solved
        offset = host_r
    else:
        vec = pt - hc.Origin
        distance = axis.DotProduct(vec)
        radial = vec - axis.Multiply(distance)
        R = radial.GetLength()
        if R < 1e-6:
            TaskDialog.Show("Re-host Tap", "Click on the surface of the straight, not on its centerline.")
            return True
        if shape == "rect":
            # n = direction from the centerline out through the clicked face.
            # offset = centerline-to-face distance along n (any point on the face gives the same value).
            cands = [cs.BasisX, cs.BasisX.Negate(), cs.BasisY, cs.BasisY.Negate()]
            n = max(cands, key=lambda v: v.DotProduct(radial))
            offset = radial.DotProduct(n)
        else:
            n = radial.Normalize()
            offset = host_r if host_r else R
    distance = min(max(distance, 0.0), length)  # clamp instead of rejecting a click near an end
    if shape != "rect":
        n = snap_about_axis(n)

    # First guess for the rotation, refined below by testing the placed result. [CONFIRM: zero reference]
    base_angle = atan2(n.DotProduct(cs.BasisY), n.DotProduct(cs.BasisX))
    snapped = nearest_quarter_turn(base_angle)
    quarter_list = [snapped + k * pi / 2.0 for k in (0, 2, 1, 3)]
    if shape == "rect":
        angles = quarter_list
    else:
        # Round taps accept any angle; fall back to quarter turns for oval/rect taps. [CONFIRM]
        angles = [base_angle + d for d in (0.0, pi, pi / 2.0, -pi / 2.0)] + quarter_list

    def on_axis_point(p):
        return hc.Origin + axis.Multiply(axis.DotProduct(p - hc.Origin))

    def side_score(tc):
        """+1 when the tap body extends along n (out through the clicked surface), -1 opposite,
        about 0 when it extends elsewhere. The component along the host axis is removed so an
        angled or offset tap body does not skew the result."""
        bb = tap.get_BoundingBox(None)
        if bb is None:
            return 0.0
        center = (bb.Min + bb.Max).Multiply(0.5)
        c = center - tc.Origin
        c = c - axis.Multiply(axis.DotProduct(c))
        if c.GetLength() < 1e-6:
            return 0.0
        return c.Normalize().DotProduct(n)

    def release(tc):
        for ref in straight_refs(tc, tap):
            dbg('DisconnectFrom')
            tc.DisconnectFrom(ref)

    notes = []
    log = []
    t = Transaction(doc, "Re-host tap")
    try:
        t.Start()
        host_origin_before = hc.Origin
        chosen = None
        last_err = None

        # Spots along the straight to try: the click first, then inward toward the middle only
        # (a click that is too close to an end is the likely cause of an out-of-range error).
        step_dir = 1.0 if distance < length / 2.0 else -1.0
        max_shift = min(SEARCH_MAX, length / 2.0)
        dist_list = [distance]
        i = 1
        while i * SEARCH_STEP <= max_shift and i <= 12:
            dist_list.append(distance + step_dir * i * SEARCH_STEP)
            i += 1
        used_distance = distance

        for tc in tap_connector_candidates(tap, tap_pt):
            try:
                release(tc)
                if tc.IsConnected:
                    log.append("connector skipped: still connected to something other than a straight")
                    continue
                # Signature: PlaceAsTap(doc, tapConn, hostConn, distance, axisRotation, secondaryAxisRotation)
                ok = False
                tried_spots = 0
                for dist in dist_list:
                    tried_spots += 1
                    any_placed = False
                    range_fail = 0
                    other_fail = False
                    for rot in angles:
                        label = "{:.0f} deg".format(rot * 180.0 / pi)
                        try:
                            dbg('PlaceAsTap dist={:.4f} rot={:.1f} len={:.4f}'.format(dist, rot * 180.0 / pi, length))
                            FabricationPart.PlaceAsTap(doc, tc, hc, dist, rot, 0.0)
                            doc.Regenerate()
                        except Exception as e:
                            msg = str(e).split("\n")[0]
                            last_err = e
                            release(tc)
                            if "range" in msg.lower():
                                range_fail += 1
                                if range_fail >= 2:
                                    break  # 2 rotations out of range here, try the next spot
                                continue
                            log.append("{}: not placed ({})".format(label, msg))
                            other_fail = True
                            continue
                        any_placed = True
                        sc = side_score(tc)
                        log.append("{}: placed, side score {:.2f}".format(label, sc))
                        if sc > 0.9:
                            ok = True
                            used_distance = dist
                            break
                        release(tc)
                    if ok or any_placed or other_fail:
                        break  # placed (stop sliding) or hit a non-range problem (sliding will not help)
                if tried_spots > 1 and not ok and not any_placed:
                    log.append("tried {} spots within {:.1f} ft of the click: all out of range".format(
                        tried_spots, max_shift))
                if ok:
                    chosen = tc
                    break
                log.append("no rotation matched for this connector")
            except Exception as e:
                log.append("connector failed: {}".format(str(e).split("\n")[0]))
                last_err = e

        if chosen is None:
            raise Exception(str(last_err) if last_err else "No rotation put the tap on the clicked side.")

        if abs(used_distance - distance) > TOL:
            notes.append("Tap was placed {:.1f} in along the straight from the clicked spot, "
                         "the nearest spot where it fits.".format(abs(used_distance - distance) * 12.0))

        # PlaceAsTap leaves the tap's mating connector on the host centerline on the duct tests. [CONFIRM for pipe]
        # Move it out along n so the mating end sits on the surface. If the move drags the host or
        # fails, only the move is undone and the tap stays as placed.
        tc = chosen
        o = tc.Origin
        cur = (o - on_axis_point(o)).DotProduct(n)
        shift = offset - cur
        was_connected = tc.IsConnected
        # Round hosts: pipe olets already sit correctly after PlaceAsTap (tested), and the move
        # drags the host pipe, so the move is skipped. Rect and oval ducts still need it.
        if shape != "round" and abs(shift) > TOL:
            st = SubTransaction(doc)
            st.Start()
            try:
                dbg('MoveElement shift={:.4f}'.format(shift))
                ElementTransformUtils.MoveElement(doc, tap.Id, n.Multiply(shift))
                doc.Regenerate()
                if hc.Origin.DistanceTo(host_origin_before) > TOL:
                    raise Exception("the host straight moved with the tap")
                st.Commit()
                log.append("moved {:.4f} ft toward the clicked surface".format(shift))
                if was_connected and not tc.IsConnected:
                    notes.append("The tap's connection to the host dropped when it was moved to the surface.")
            except Exception as e:
                st.RollBack()
                doc.Regenerate()
                notes.append("Tap was placed but not moved out to the surface ({}). "
                             "Check its position before using it.".format(str(e).split("\n")[0]))
        t.Commit()
        if notes:
            TaskDialog.Show("Re-host Tap", "\n".join(notes))
    except Exception as e:
        if t.HasStarted():
            t.RollBack()
        msg = "Could not re-host tap: {}".format(str(e))
        if log:
            msg += "\n\nAttempts:\n" + "\n".join(log)
        TaskDialog.Show("Re-host Tap", msg)
    return True


while rehost():
    pass