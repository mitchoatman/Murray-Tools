# -*- coding: utf-8 -*-
import sys
import traceback
from Autodesk.Revit import DB
from Autodesk.Revit.DB import ViewPlan, ViewSection
from Autodesk.Revit.UI.Selection import ObjectType, ObjectSnapTypes
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.Exceptions import OperationCanceledException

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
active_view = doc.ActiveView

SNAP_NONE = getattr(ObjectSnapTypes, "None")  # "None" is a reserved word in IronPython


def info(title, msg):
    TaskDialog.Show(title, msg)


class WarningCollector(DB.IFailuresPreprocessor):
    """Collects warnings raised during the transaction so they can be shown."""
    def __init__(self):
        self.messages = []

    def PreprocessFailures(self, failures_accessor):
        for f in failures_accessor.GetFailureMessages():
            self.messages.append(f.GetDescriptionText())
        return DB.FailureProcessingResult.Continue


class UserError(Exception):
    pass


def get_section_view_type(document):
    for vft in DB.FilteredElementCollector(document).OfClass(DB.ViewFamilyType):
        if vft.ViewFamily == DB.ViewFamily.Section:
            return vft
    return None


def require_supported_view(view):
    if not isinstance(view, (ViewPlan, ViewSection)):
        raise UserError("Run this command from a floor plan or section view.\n"
                        "Active view is: {} ({})".format(view.Name, view.ViewType))


def bbox_corners(bbox):
    mn, mx = bbox.Min, bbox.Max
    return [DB.XYZ(x, y, z) for x in (mn.X, mx.X) for y in (mn.Y, mx.Y) for z in (mn.Z, mx.Z)]


def get_element_points(elem):
    pts = []
    bbox = elem.get_BoundingBox(None)
    if bbox:
        pts.extend(bbox_corners(bbox))
    loc = elem.Location
    if isinstance(loc, DB.LocationCurve) and isinstance(loc.Curve, DB.Line):
        crv = loc.Curve
        pts.append(crv.GetEndPoint(0))
        pts.append(crv.GetEndPoint(1))
    return pts


def build_section_transform(all_pts, view, pick_point):
    n = len(all_pts)
    origin = DB.XYZ(sum(p.X for p in all_pts) / n,
                    sum(p.Y for p in all_pts) / n,
                    sum(p.Z for p in all_pts) / n)
    click_vec = pick_point - origin
    y_axis = DB.XYZ.BasisZ

    if isinstance(view, ViewSection):
        view_dir = view.ViewDirection.Normalize()
        up_dir = view.UpDirection.Normalize()
        right_dir = view.RightDirection.Normalize()

        click_screen = click_vec - view_dir.Multiply(click_vec.DotProduct(view_dir))
        if click_screen.GetLength() < 1e-6:
            raise UserError("Pick point is too close to the center of the selection.")

        dx = click_screen.DotProduct(right_dir)
        dy = click_screen.DotProduct(up_dir)

        if abs(dx) > abs(dy):
            z_axis = right_dir.Negate() if dx > 0 else right_dir
        else:
            z_axis = view_dir.Negate() if dy > 0 else view_dir
    else:
        click_vec = DB.XYZ(click_vec.X, click_vec.Y, 0.0)
        if click_vec.GetLength() < 1e-6:
            raise UserError("Pick point is too close to the center of the selection.")
        dx, dy = click_vec.X, click_vec.Y
        if abs(dx) > abs(dy):
            z_axis = DB.XYZ(-1, 0, 0) if dx > 0 else DB.XYZ(1, 0, 0)
        else:
            z_axis = DB.XYZ(0, -1, 0) if dy > 0 else DB.XYZ(0, 1, 0)

    x_axis = y_axis.CrossProduct(z_axis).Normalize()
    y_axis = z_axis.CrossProduct(x_axis).Normalize()

    tf = DB.Transform.Identity
    tf.Origin = origin
    tf.BasisX = x_axis
    tf.BasisY = y_axis
    tf.BasisZ = z_axis
    return tf


def get_local_bounds(points, transform):
    inv = transform.Inverse
    local = [inv.OfPoint(p) for p in points]
    return (DB.XYZ(min(p.X for p in local), min(p.Y for p in local), min(p.Z for p in local)),
            DB.XYZ(max(p.X for p in local), max(p.Y for p in local), max(p.Z for p in local)))


def main():
    require_supported_view(active_view)

    # 1. Selection
    selected_ids = list(uidoc.Selection.GetElementIds())
    if not selected_ids:
        refs = uidoc.Selection.PickObjects(
            ObjectType.Element, "Select elements to create a section around, then click Finish")
        selected_ids = [r.ElementId for r in refs]

    elements = []
    for eid in selected_ids:
        e = doc.GetElement(eid)
        if e is not None:
            elements.append(e)
    if not elements:
        raise UserError("No valid elements selected.")

    # 2. Pick point
    pick_point = uidoc.Selection.PickPoint(SNAP_NONE, "Click on the side you want to view from")

    # 3. Section type
    section_type = get_section_view_type(doc)
    if section_type is None:
        raise UserError("No Section ViewFamilyType found in this project.")

    # 4. Geometry
    all_pts = []
    skipped = []
    for e in elements:
        pts = get_element_points(e)
        if pts:
            all_pts.extend(pts)
        else:
            skipped.append(e.Id.IntegerValue)
    if not all_pts:
        raise UserError("Selected elements have no usable geometry.")

    tf = build_section_transform(all_pts, active_view, pick_point)
    local_min, local_max = get_local_bounds(all_pts, tf)

    pad_x, pad_y, pad_z = 1.0, 1.0, 2.0   # feet
    min_z = local_min.Z - pad_z
    max_z = local_max.Z + pad_z
    if (max_z - min_z) < 0.25:
        max_z = min_z + 0.25

    box = DB.BoundingBoxXYZ()
    box.Transform = tf
    box.Min = DB.XYZ(local_min.X - pad_x, local_min.Y - pad_y, min_z)
    box.Max = DB.XYZ(local_max.X + pad_x, local_max.Y + pad_y, max_z)

    # 5. Create
    warnings = WarningCollector()
    notes = []
    t = DB.Transaction(doc, "Create Section Around Selection")
    t.Start()
    try:
        opts = t.GetFailureHandlingOptions()
        opts.SetFailuresPreprocessor(warnings)
        t.SetFailureHandlingOptions(opts)

        section_view = DB.ViewSection.CreateSection(doc, section_type.Id, box)
        if section_view is None:
            raise Exception("ViewSection.CreateSection returned None.")

        try:
            section_view.DetailLevel = DB.ViewDetailLevel.Fine
        except Exception as ex:
            notes.append("Could not set detail level: {}".format(ex))

        section_view.CropBoxActive = True
        section_view.CropBoxVisible = True

        status = t.Commit()
        if status != DB.TransactionStatus.Committed:
            raise Exception("Transaction did not commit. Status: {}".format(status))
    except:
        if t.HasStarted() and not t.HasEnded():
            t.RollBack()
        raise

    # 6. Report
    # msg = "Created section '{}' (Id {}) around {} element(s).".format(
        # section_view.Name, section_view.Id.IntegerValue, len(elements))
    # if skipped:
        # msg += "\n\nElements with no geometry (ignored): {}".format(skipped)
    # if notes:
        # msg += "\n\nNotes:\n" + "\n".join(notes)
    # if warnings.messages:
        # msg += "\n\nRevit warnings:\n" + "\n".join(set(warnings.messages))
    # info("Section Created", msg)

    # try:
        # uidoc.ActiveView = section_view
    # except Exception as ex:
        # info("Section Created", "Section exists but could not be opened: {}".format(ex))


try:
    main()
except OperationCanceledException:
    pass  # user pressed Esc during a pick; not an error
except UserError as ex:
    info("Cannot Create Section", str(ex))
except Exception as ex:
    info("Create Section Failed",
         "{}: {}\n\n{}".format(type(ex).__name__, ex, traceback.format_exc()))