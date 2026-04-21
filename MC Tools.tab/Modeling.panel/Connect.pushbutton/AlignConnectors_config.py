# coding: utf8
from System.Collections.Generic import List
from Autodesk.Revit.DB import InsulationLiningBase, Transaction, Element, ConnectorManager, Connector, XYZ, FabricationPart, ElementTransformUtils, ElementId, Line
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# -------------------------------------------------
# Compatibility helper for ElementId across Revit versions 2022-2026
# -------------------------------------------------
def elem_id_to_int(eid):
    """
    Returns the integer representation of an ElementId.
    Compatible with Revit 2022-2024 (.IntegerValue) and 2025-2026 (.Value).
    """
    try:
        return eid.IntegerValue          # Revit 2022-2024
    except AttributeError:
        try:
            return eid.Value             # Revit 2025-2026
        except AttributeError:
            try:
                return long(eid)         # Fallback (rarely needed)
            except:
                return 0                 # Safety net - should not reach here

# -------------------------------------------------
def get_connector_manager(element):
    try:
        return element.ConnectorManager
    except AttributeError:
        try:
            return element.MEPModel.ConnectorManager
        except AttributeError:
            return None

def get_closest_unused_connector(element, click_point):
    cm = get_connector_manager(element)
    if not cm:
        return None
    unused = cm.UnusedConnectors
    if unused.Size == 0:
        return None
    min_dist = float('inf')
    closest = None
    for conn in unused:
        dist = conn.Origin.DistanceTo(click_point)
        if dist < min_dist:
            min_dist = dist
            closest = conn
    # Reject if too far (tolerance in feet - adjust if needed)
    if closest and min_dist < 2.0:
        return closest
    return None

class ValidMEPElementFilter(ISelectionFilter):
    def AllowElement(self, elem):
        if isinstance(elem, InsulationLiningBase):
            return False
        return get_connector_manager(elem) is not None or isinstance(elem, FabricationPart)

    def AllowReference(self, ref, pt):
        return True

# -------------------------------------------------
def execute_move_and_connect():
    # Step 1: Select one or more elements to move
    try:
        refs = uidoc.Selection.PickObjects(
            ObjectType.Element,
            ValidMEPElementFilter(),
            "Select MEP / fabrication element(s) to move (finish with checkmark)"
        )
    except:
        return

    if not refs:
        return

    moved_elements = [doc.GetElement(r) for r in refs]
    moved_ids_set = {elem_id_to_int(el.Id) for el in moved_elements}

    # Step 2: Pick the FROM connector - unrestricted, then validate it's from selection
    try:
        moved_ref = uidoc.Selection.PickObject(
            ObjectType.Element,
            ValidMEPElementFilter(),
            "Click near the open connector on one of the SELECTED elements (FROM)"
        )
        moved_click_pt = moved_ref.GlobalPoint
        picked_elem = doc.GetElement(moved_ref)

        if elem_id_to_int(picked_elem.Id) not in moved_ids_set:
            TaskDialog.Show("Invalid Selection", "Please click on one of the previously selected elements.")
            return

    except:
        TaskDialog.Show("Cancelled", "FROM connector selection cancelled.")
        return

    moved_connector = get_closest_unused_connector(picked_elem, moved_click_pt)
    if not moved_connector:
        TaskDialog.Show("Error", "No suitable unused connector found near your click on the selected element.")
        return

    # Step 3: Pick the TO connector
    try:
        target_ref = uidoc.Selection.PickObject(
            ObjectType.Element,
            ValidMEPElementFilter(),
            "Click near the open connector on the target element / assembly (TO)"
        )
        target_click_pt = target_ref.GlobalPoint
        target_element = doc.GetElement(target_ref)
    except:
        TaskDialog.Show("Cancelled", "TO connector selection cancelled.")
        return

    target_connector = get_closest_unused_connector(target_element, target_click_pt)
    if not target_connector:
        TaskDialog.Show("Error", "No suitable unused connector found near your click on the target.")
        return

    # Domain check
    if moved_connector.Domain != target_connector.Domain:
        TaskDialog.Show("Error", "Connectors belong to different domains (e.g. Duct vs Pipe).")
        return

    # -------------------------------------------------
    # Execute move + connect
    # -------------------------------------------------
    t = Transaction(doc, "Move Selected Elements and Connect")
    t.Start()

    try:
        elem_ids = List[ElementId](el.Id for el in moved_elements)

        # Preferred path: single fabrication part -> use native API
        if (len(moved_elements) == 1 and
            isinstance(moved_connector.Owner, FabricationPart) and
            isinstance(target_connector.Owner, FabricationPart)):

            FabricationPart.AlignPartByConnectors(doc, moved_connector.Owner, moved_connector, target_connector)
            FabricationPart.ConnectAndCouple(doc, moved_connector, target_connector)

        else:
            # Manual rigid transformation (supports groups)
            from_dir = moved_connector.CoordinateSystem.BasisZ
            to_dir   = XYZ.Negate(target_connector.CoordinateSystem.BasisZ)

            angle = from_dir.AngleTo(to_dir)
            if angle > 1e-6:
                axis_vec = from_dir.CrossProduct(to_dir)
                if axis_vec.IsZeroLength():
                    axis_vec = moved_connector.CoordinateSystem.BasisY
                    if axis_vec.IsZeroLength():
                        axis_vec = moved_connector.CoordinateSystem.BasisX

                if not axis_vec.IsZeroLength():
                    axis = Line.CreateUnbound(moved_connector.Origin, axis_vec.Normalize())
                    ElementTransformUtils.RotateElements(doc, elem_ids, axis, angle)

            # Translate entire group
            delta = target_connector.Origin - moved_connector.Origin
            if not delta.IsZeroLength():
                ElementTransformUtils.MoveElements(doc, elem_ids, delta)

            # Final connection
            if isinstance(moved_connector.Owner, FabricationPart) and isinstance(target_connector.Owner, FabricationPart):
                FabricationPart.ConnectAndCouple(doc, moved_connector, target_connector)
            else:
                moved_connector.ConnectTo(target_connector)

        t.Commit()
        # TaskDialog.Show("Success", "Elements moved and connected successfully.")

    except Exception as ex:
        TaskDialog.Show("Operation Failed", "Error:\n{}".format(str(ex)))
        if t.HasStarted():
            t.RollBack()

# -------------------------------------------------
# Run the tool (single execution - no loop)
# -------------------------------------------------
execute_move_and_connect()