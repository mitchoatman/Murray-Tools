import clr

from Autodesk.Revit.DB import (
    Transaction,
    TransactionGroup,
    FabricationPart,
    ConnectorType,
    XYZ,
    Plane,
    SketchPlane,
    ViewType,
    FilteredElementCollector,
    ReferencePlane,
    LocationPoint,
    LocationCurve,
    FamilyInstance,
    AssemblyInstance
)

from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.UI import (
    TaskDialog,
    TaskDialogCommandLinkId,
    TaskDialogCommonButtons,
    TaskDialogResult
)
from Autodesk.Revit.Exceptions import OperationCanceledException


# -----------------------------------------------------------------------------
# Dialog helpers
# -----------------------------------------------------------------------------
def show_error(msg, title="Create Work Plane"):
    TaskDialog.Show(title, msg)


def choose_two_option_dialog(title, instruction, option1, option2):
    td = TaskDialog(title)
    td.MainInstruction = instruction
    td.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, option1)
    td.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, option2)
    td.CommonButtons = TaskDialogCommonButtons.Cancel
    result = td.Show()

    if result == TaskDialogResult.CommandLink1:
        return 1
    elif result == TaskDialogResult.CommandLink2:
        return 2
    return None


# -----------------------------------------------------------------------------
# Selection filters
# -----------------------------------------------------------------------------
class TopLevelSelectionFilter(ISelectionFilter):
    def AllowElement(self, elem):
        return (
            isinstance(elem, FabricationPart) or
            isinstance(elem, FamilyInstance) or
            isinstance(elem, AssemblyInstance)
        )

    def AllowReference(self, reference, position):
        return True


class AssemblyMemberSelectionFilter(ISelectionFilter):
    def __init__(self, allowed_ids):
        self.allowed_ids = set([eid.IntegerValue for eid in allowed_ids])

    def AllowElement(self, elem):
        if elem.Id.IntegerValue not in self.allowed_ids:
            return False
        return isinstance(elem, FabricationPart) or isinstance(elem, FamilyInstance)

    def AllowReference(self, reference, position):
        return True


# -----------------------------------------------------------------------------
# Geometry helpers
# -----------------------------------------------------------------------------
def midpoint(p1, p2):
    return XYZ(
        (p1.X + p2.X) / 2.0,
        (p1.Y + p2.Y) / 2.0,
        (p1.Z + p2.Z) / 2.0
    )


def is_mostly_vertical(vec):
    return abs(vec.Z) > max(abs(vec.X), abs(vec.Y))


def get_horizontal_direction(vec):
    horiz = XYZ(vec.X, vec.Y, 0.0)
    if horiz.GetLength() == 0:
        return None
    return horiz.Normalize()


# -----------------------------------------------------------------------------
# Selection
# -----------------------------------------------------------------------------
def get_target_element(doc, uidoc):
    ref = uidoc.Selection.PickObject(
        ObjectType.Element,
        TopLevelSelectionFilter(),
        "Select a fabrication part, family instance, or assembly"
    )
    elem = doc.GetElement(ref.ElementId)

    if isinstance(elem, AssemblyInstance):
        member_ids = list(elem.GetMemberIds())
        valid_member_ids = []

        for mid in member_ids:
            member = doc.GetElement(mid)
            if isinstance(member, FabricationPart) or isinstance(member, FamilyInstance):
                valid_member_ids.append(mid)

        if not valid_member_ids:
            raise Exception("The selected assembly has no fabrication parts or family instances.")

        if len(valid_member_ids) == 1:
            return doc.GetElement(valid_member_ids[0])

        member_ref = uidoc.Selection.PickObject(
            ObjectType.Element,
            AssemblyMemberSelectionFilter(valid_member_ids),
            "Assembly selected. Pick the fabrication part or family instance inside the assembly"
        )
        return doc.GetElement(member_ref.ElementId)

    return elem


# -----------------------------------------------------------------------------
# Data extraction
# -----------------------------------------------------------------------------
def get_fabrication_part_data(elem):
    connectors = []
    for conn in elem.ConnectorManager.Connectors:
        if conn.ConnectorType == ConnectorType.End:
            connectors.append(conn)

    if len(connectors) < 2:
        raise Exception("The selected fabrication part does not have at least two end connectors.")

    p1 = connectors[0].Origin
    p2 = connectors[1].Origin
    vec = p2 - p1

    if vec.GetLength() == 0:
        raise Exception("The selected fabrication part has invalid connector geometry.")

    return midpoint(p1, p2), vec.Normalize()


def get_family_instance_data(elem):
    loc = elem.Location

    if isinstance(loc, LocationPoint):
        return "point", loc.Point, None

    if isinstance(loc, LocationCurve):
        curve = loc.Curve
        p1 = curve.GetEndPoint(0)
        p2 = curve.GetEndPoint(1)
        vec = p2 - p1

        if vec.GetLength() == 0:
            raise Exception("The selected line-based family has invalid curve geometry.")

        return "curve", curve.Evaluate(0.5, True), vec.Normalize()

    raise Exception("The selected family instance is not point-based or line-based.")


# -----------------------------------------------------------------------------
# Plane orientation logic
# -----------------------------------------------------------------------------
def get_basis_for_section_view(curview):
    return curview.RightDirection.Normalize(), curview.UpDirection.Normalize()


def get_basis_for_point_family(curview):
    if curview.ViewType == ViewType.Section:
        return get_basis_for_section_view(curview)

    orientation = choose_two_option_dialog(
        "Select Plane Orientation",
        "Choose the plane orientation for the family instance:",
        "Horizontal",
        "Vertical"
    )
    if orientation is None:
        return None, None

    if orientation == 1:
        return XYZ.BasisX, XYZ.BasisY

    axis = choose_two_option_dialog(
        "Select Axis",
        "Choose the axis direction for the vertical plane:",
        "X-axis",
        "Y-axis"
    )
    if axis is None:
        return None, None

    return (XYZ.BasisX, XYZ.BasisZ) if axis == 1 else (XYZ.BasisY, XYZ.BasisZ)


def get_basis_for_linear_element(curview, line_dir, label):
    if curview.ViewType == ViewType.Section:
        return get_basis_for_section_view(curview)

    if is_mostly_vertical(line_dir):
        axis = choose_two_option_dialog(
            "Select Axis",
            "The selected {0} is vertical. Choose which axis you want the plane aligned to:".format(label),
            "Align with X-axis",
            "Align with Y-axis"
        )
        if axis is None:
            return None, None

        return (XYZ.BasisX, line_dir) if axis == 1 else (line_dir, XYZ.BasisY)

    horizontal_dir = get_horizontal_direction(line_dir)
    if horizontal_dir is None:
        raise Exception("Could not determine a horizontal direction from the selected {0}.".format(label))

    orientation = choose_two_option_dialog(
        "Select Plane Orientation",
        "The selected {0} is not vertical. Choose the work plane orientation:".format(label),
        "Vertical",
        "Horizontal"
    )
    if orientation is None:
        return None, None

    if orientation == 1:
        return horizontal_dir, XYZ.BasisZ

    perp = XYZ.BasisZ.CrossProduct(horizontal_dir)
    if perp.GetLength() == 0:
        raise Exception("Could not create a horizontal work plane from the selected {0}.".format(label))

    return horizontal_dir, perp.Normalize()


# -----------------------------------------------------------------------------
# Temporary plane helpers
# -----------------------------------------------------------------------------
def delete_temporary_reference_planes(doc):
    refs_to_delete = []
    collector = FilteredElementCollector(doc).OfClass(ReferencePlane).WhereElementIsNotElementType()

    for rp in collector:
        if rp.Name == "TEMPORARY":
            refs_to_delete.append(rp.Id)

    for rid in refs_to_delete:
        doc.Delete(rid)


def try_name_sketch_plane_temporary(sketch_plane):
    try:
        sketch_plane.Name = "TEMPORARY"
    except:
        pass


def create_temporary_reference_plane(doc, curview, plane):
    bubble_end = plane.Origin
    free_end = plane.Origin + plane.XVec
    cut_vec = plane.YVec

    ref_plane = doc.Create.NewReferencePlane(bubble_end, free_end, cut_vec, curview)
    ref_plane.Name = "TEMPORARY"
    return ref_plane


def create_and_assign_plane(doc, curview, plane):
    tg = TransactionGroup(doc, "Create Work Plane")
    tg.Start()

    try:
        t1 = Transaction(doc, "Set Work Plane")
        t1.Start()
        sketch_plane = SketchPlane.Create(doc, plane)
        try_name_sketch_plane_temporary(sketch_plane)
        curview.SketchPlane = sketch_plane
        curview.ShowActiveWorkPlane()
        t1.Commit()

        t2 = Transaction(doc, "Reset TEMPORARY Reference Plane")
        t2.Start()
        delete_temporary_reference_planes(doc)
        create_temporary_reference_plane(doc, curview, plane)
        t2.Commit()

        tg.Assimilate()

    except Exception:
        tg.RollBack()
        raise


# -----------------------------------------------------------------------------
# Main
# -----------------------------------------------------------------------------
def select_element_and_create_plane():
    doc = __revit__.ActiveUIDocument.Document
    uidoc = __revit__.ActiveUIDocument
    curview = doc.ActiveView

    if curview.ViewType not in (ViewType.ThreeD, ViewType.Section):
        show_error("Run this tool from a 3D view or a Section view.")
        return

    try:
        elem = get_target_element(doc, uidoc)

        plane_origin = None
        x_vector = None
        y_vector = None

        if isinstance(elem, FabricationPart):
            plane_origin, line_dir = get_fabrication_part_data(elem)
            x_vector, y_vector = get_basis_for_linear_element(curview, line_dir, "fabrication part")

        elif isinstance(elem, FamilyInstance):
            family_mode, plane_origin, line_dir = get_family_instance_data(elem)

            if family_mode == "point":
                x_vector, y_vector = get_basis_for_point_family(curview)
            else:
                x_vector, y_vector = get_basis_for_linear_element(curview, line_dir, "line-based family")

        else:
            raise Exception("Only fabrication parts, family instances, and assemblies are supported.")

        if plane_origin is None or x_vector is None or y_vector is None:
            return

        if curview.ViewType == ViewType.Section:
            x_vector, y_vector = get_basis_for_section_view(curview)

        plane = Plane.CreateByOriginAndBasis(plane_origin, x_vector, y_vector)
        create_and_assign_plane(doc, curview, plane)

    except OperationCanceledException:
        return
    except Exception as ex:
        show_error(str(ex))


# Run
select_element_and_create_plane()