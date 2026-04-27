
# -*- coding: utf-8 -*-
from Autodesk.Revit import DB
from Autodesk.Revit.DB import FabricationPart
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.UI import TaskDialog
import sys

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
active_view = doc.ActiveView


# --------------------------------------------------
# Allowed straight fabrication CIDs only
# Update this list to match your real straight pipe
# and straight duct CID values.
# --------------------------------------------------
ALLOWED_CIDS = set([
    2041,866,40,
])


def get_section_view_type(document):
    for vft in DB.FilteredElementCollector(document).OfClass(DB.ViewFamilyType):
        if vft.ViewFamily == DB.ViewFamily.Section:
            return vft
    return None


def require_plan_view(view):
    if not isinstance(view, DB.ViewPlan):
        raise Exception("Run this command from a floor plan view.")


class StraightFabPipeDuctFilter(ISelectionFilter):
    def is_allowed(self, element):
        try:
            if not isinstance(element, FabricationPart):
                return False

            cid = element.ItemCustomId
            if cid not in ALLOWED_CIDS:
                return False

            loc = element.Location
            if not isinstance(loc, DB.LocationCurve):
                return False

            if not isinstance(loc.Curve, DB.Line):
                return False

            return True
        except:
            return False

    def AllowElement(self, element):
        return self.is_allowed(element)

    def AllowReference(self, reference, point):
        try:
            element = doc.GetElement(reference.ElementId)
            return self.is_allowed(element)
        except:
            return False


def pick_element_and_point():
    ref = uidoc.Selection.PickObject(
        ObjectType.PointOnElement,
        StraightFabPipeDuctFilter(),
        "Pick a straight fabrication pipe/duct on the side where you want the section marker"
    )
    elem = doc.GetElement(ref.ElementId)
    pt = ref.GlobalPoint
    return elem, pt


def get_linear_curve(elem):
    loc = elem.Location
    if isinstance(loc, DB.LocationCurve):
        crv = loc.Curve
        if isinstance(crv, DB.Line):
            return crv
    return None


def bbox_corners(bbox):
    mn = bbox.Min
    mx = bbox.Max
    pts = []
    for x in [mn.X, mx.X]:
        for y in [mn.Y, mx.Y]:
            for z in [mn.Z, mx.Z]:
                pts.append(DB.XYZ(x, y, z))
    return pts


def get_points_for_bounds(elem, curve):
    pts = [curve.GetEndPoint(0), curve.GetEndPoint(1)]
    bbox = elem.get_BoundingBox(None)
    if bbox:
        pts.extend(bbox_corners(bbox))
    return pts


def build_section_transform_from_plan(curve, view, pick_point):
    p0 = curve.GetEndPoint(0)
    p1 = curve.GetEndPoint(1)
    origin = curve.Evaluate(0.5, True)

    line_dir = (p1 - p0)
    if line_dir.GetLength() < 1e-9:
        raise Exception("Selected element is too short.")

    line_dir = line_dir.Normalize()

    plan_view_dir = view.ViewDirection.Normalize()

    x_guess = line_dir - plan_view_dir.Multiply(line_dir.DotProduct(plan_view_dir))
    if x_guess.GetLength() < 1e-6:
        raise Exception("Selected element is vertical or nearly vertical in this plan.")

    x_guess = x_guess.Normalize()
    y_axis = DB.XYZ.BasisZ

    z_guess = x_guess.CrossProduct(y_axis)
    if z_guess.GetLength() < 1e-6:
        raise Exception("Could not determine section direction.")
    z_guess = z_guess.Normalize()

    proj = curve.Project(pick_point)
    if proj:
        on_curve = proj.XYZPoint
    else:
        on_curve = origin

    side_vec = pick_point - on_curve
    side_vec = side_vec - y_axis.Multiply(side_vec.DotProduct(y_axis))

    # reversed marker behavior per your preference
    if side_vec.GetLength() > 1e-6 and side_vec.DotProduct(z_guess) > 0:
        z_axis = z_guess.Negate()
    else:
        z_axis = z_guess

    x_axis = y_axis.CrossProduct(z_axis).Normalize()

    tf = DB.Transform.Identity
    tf.Origin = origin
    tf.BasisX = x_axis
    tf.BasisY = y_axis
    tf.BasisZ = z_axis

    return tf


def get_local_bounds(points, transform):
    inv = transform.Inverse
    xs, ys, zs = [], [], []

    for p in points:
        lp = inv.OfPoint(p)
        xs.append(lp.X)
        ys.append(lp.Y)
        zs.append(lp.Z)

    return DB.XYZ(min(xs), min(ys), min(zs)), DB.XYZ(max(xs), max(ys), max(zs))


try:
    require_plan_view(active_view)

    elem, pick_point = pick_element_and_point()
    curve = get_linear_curve(elem)

    section_type = get_section_view_type(doc)
    if not section_type:
        raise Exception("No Section ViewFamilyType found.")

    tf = build_section_transform_from_plan(curve, active_view, pick_point)
    pts = get_points_for_bounds(elem, curve)
    local_min, local_max = get_local_bounds(pts, tf)

    pad_x = 1.0
    pad_y = 1.0
    pad_z = 0.5

    min_z = local_min.Z - pad_z
    max_z = local_max.Z + pad_z
    if (max_z - min_z) < 0.25:
        max_z = min_z + 0.25

    box = DB.BoundingBoxXYZ()
    box.Transform = tf
    box.Min = DB.XYZ(local_min.X - pad_x, local_min.Y - pad_y, min_z)
    box.Max = DB.XYZ(local_max.X + pad_x, local_max.Y + pad_y, max_z)

    t = DB.Transaction(doc, "Create Section Along Element")
    try:
        t.Start()
        DB.ViewSection.CreateSection(doc, section_type.Id, box)
        t.Commit()
    except:
        if t.HasStarted():
            t.RollBack()
        raise

except Exception as ex:
    TaskDialog.Show("Error", str(ex))
    sys.exit()