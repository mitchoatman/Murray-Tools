# -*- coding: utf-8 -*-
from Autodesk.Revit import DB
from Autodesk.Revit.DB import FabricationPart, FamilyInstance
from Autodesk.Revit.UI.Selection import ObjectType, ObjectSnapTypes
from Autodesk.Revit.UI import TaskDialog
import sys

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
active_view = doc.ActiveView

def get_section_view_type(document):
    for vft in DB.FilteredElementCollector(document).OfClass(DB.ViewFamilyType):
        if vft.ViewFamily == DB.ViewFamily.Section:
            return vft
    return None

def require_plan_view(view):
    if not isinstance(view, DB.ViewPlan):
        raise Exception("Run this command from a floor plan view.")

def bbox_corners(bbox):
    mn = bbox.Min
    mx = bbox.Max
    pts = []
    for x in [mn.X, mx.X]:
        for y in [mn.Y, mx.Y]:
            for z in [mn.Z, mx.Z]:
                pts.append(DB.XYZ(x, y, z))
    return pts

def get_element_points(elem):
    pts = []
    bbox = elem.get_BoundingBox(None)
    if bbox:
        pts.extend(bbox_corners(bbox))
    
    try:
        loc = elem.Location
        if isinstance(loc, DB.LocationCurve) and isinstance(loc.Curve, DB.Line):
            crv = loc.Curve
            pts.append(crv.GetEndPoint(0))
            pts.append(crv.GetEndPoint(1))
    except:
        pass
    return pts

def build_section_transform_for_elements(elements, view, pick_point):
    all_pts = []
    for elem in elements:
        all_pts.extend(get_element_points(elem))
    
    if not all_pts:
        raise Exception("Selected elements have no valid geometry.")
    
    sum_x = sum(p.X for p in all_pts)
    sum_y = sum(p.Y for p in all_pts)
    sum_z = sum(p.Z for p in all_pts)
    n = len(all_pts)
    origin = DB.XYZ(sum_x / n, sum_y / n, sum_z / n)

    # Vector from center of selection to the user's pick point
    click_vec = pick_point - origin
    click_vec = DB.XYZ(click_vec.X, click_vec.Y, 0.0)

    if click_vec.GetLength() < 1e-6:
        raise Exception("Pick point is too close to the center of the selection.")

    # Determine whether the user clicked more horizontally (Left/Right) or vertically (Top/Bottom)
    # Normalized check without aspect-ratio bias
    dx = click_vec.X
    dy = click_vec.Y

    y_axis = DB.XYZ.BasisZ # Vertical height axis remains Z-up

    if abs(dx) > abs(dy):
        # LEFT / RIGHT click: Look along the X axis (cuts a vertical profile looking down the run)
        if dx > 0:
            # Clicked to the RIGHT -> Look right (-X direction)
            z_axis = DB.XYZ(-1, 0, 0)
        else:
            # Clicked to the LEFT -> Look left (+X direction)
            z_axis = DB.XYZ(1, 0, 0)
    else:
        # TOP / BOTTOM click: Look along the Y axis (cuts a horizontal plan looking across the runs)
        if dy > 0:
            # Clicked to the TOP -> Look up (-Y direction in plan view coordinates)
            z_axis = DB.XYZ(0, -1, 0)
        else:
            # Clicked to the BOTTOM -> Look down (+Y direction in plan view coordinates)
            z_axis = DB.XYZ(0, 1, 0)

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

    selected_ids = uidoc.Selection.GetElementIds()
    if not selected_ids:
        refs = uidoc.Selection.PickObjects(
            ObjectType.Element, 
            "Select multiple elements to create a section around"
        )
        selected_ids = [r.ElementId for r in refs]

    elements = [doc.GetElement(eid) for eid in selected_ids if doc.GetElement(eid)]
    if not elements:
        raise Exception("No valid elements selected.")

    pick_point = uidoc.Selection.PickPoint(
        ObjectSnapTypes.None, 
        "Click Left, Right, Top, or Bottom relative to the group center"
    )

    section_type = get_section_view_type(doc)
    if not section_type:
        raise Exception("No Section ViewFamilyType found.")

    tf = build_section_transform_for_elements(elements, active_view, pick_point)
    
    all_pts = []
    for elem in elements:
        all_pts.extend(get_element_points(elem))

    local_min, local_max = get_local_bounds(all_pts, tf)

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

    t = DB.Transaction(doc, "Create Section Around Selection")
    try:
        t.Start()
        section_view = DB.ViewSection.CreateSection(doc, section_type.Id, box)
        
        try:
            section_view.DetailLevel = DB.ViewDetailLevel.Fine
        except:
            pass
            
        t.Commit()
    except:
        if t.HasStarted():
            t.RollBack()
        raise

except Exception as ex:
    TaskDialog.Show("Error", str(ex))
    sys.exit()