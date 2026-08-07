import clr
from Autodesk.Revit.DB import (
    Transaction, FabricationPart, ConnectorType, XYZ, Plane, SketchPlane,
    ViewType, FilteredElementCollector, ReferencePlane,
    TransactionGroup, LocationPoint, FamilyInstance, AssemblyInstance
)
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.UI import TaskDialog, TaskDialogCommandLinkId, TaskDialogCommonButtons, TaskDialogResult

clr.AddReference("RevitServices")
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager

class AssemblyMemberSelectionFilter(ISelectionFilter):
    def __init__(self, allowed_ids):
        self.allowed_ids = set([eid.IntegerValue for eid in allowed_ids])
    def AllowElement(self, elem):
        if elem.Id.IntegerValue not in self.allowed_ids:
            return False
        return isinstance(elem, FabricationPart) or isinstance(elem, FamilyInstance)
    def AllowReference(self, reference, position):
        return True

def get_target_element(doc, uidoc):
    try:
        ref = uidoc.Selection.PickObject(
            ObjectType.Element,
            "Select a fabrication part or family instance"
        )
        elem = doc.GetElement(ref.ElementId)
        
        # If user selects an assembly, drill into it
        if isinstance(elem, AssemblyInstance):
            member_ids = list(elem.GetMemberIds())
            valid_member_ids = []
            for mid in member_ids:
                member = doc.GetElement(mid)
                if isinstance(member, FabricationPart) or isinstance(member, FamilyInstance):
                    valid_member_ids.append(mid)
            if not valid_member_ids:
                return None
            if len(valid_member_ids) == 1:
                return doc.GetElement(valid_member_ids[0])
            member_filter = AssemblyMemberSelectionFilter(valid_member_ids)
            try:
                member_ref = uidoc.Selection.PickObject(
                    ObjectType.Element,
                    member_filter,
                    "Assembly selected. Now pick the specific fabrication part or family instance inside the assembly"
                )
                elem = doc.GetElement(member_ref.ElementId)
            except:
                return None  # User cancelled inner selection
        return elem
    except:
        return None  # User cancelled main selection

def select_fabrication_pipe_and_create_plane():
    doc = __revit__.ActiveUIDocument.Document
    uidoc = __revit__.ActiveUIDocument
    curview = doc.ActiveView

    try:
        if curview.ViewType not in (ViewType.ThreeD, ViewType.Section):
            # Optional: you can remove this print if you want total silence on view type
            return

        # Unified selection prompt
        elem = get_target_element(doc, uidoc)
        if elem is None:
            return  # Quiet exit on cancellation

        x_vector = None
        y_vector = None
        plane_origin = None

        # =================================================================
        # CASE 1: FabricationPart
        # =================================================================
        if isinstance(elem, FabricationPart):
            connectors = list(elem.ConnectorManager.Connectors)
            if len(connectors) < 2:
                return
            connector_points = [conn.Origin for conn in connectors if conn.ConnectorType == ConnectorType.End]
            if len(connector_points) < 2:
                return
            connector_1 = connector_points[0]
            connector_2 = connector_points[1]
            line_vector = (connector_2 - connector_1).Normalize()

            if curview.ViewType == ViewType.Section:
                plane_origin = (connector_1 + connector_2) / 2
                x_vector = curview.RightDirection.Normalize()
                y_vector = curview.UpDirection.Normalize()
            elif abs(line_vector.Z) > max(abs(line_vector.X), abs(line_vector.Y)):
                plane_origin = (connector_1 + connector_2) / 2
                if curview.ViewType != ViewType.Section:
                    task_dialog = TaskDialog("Select Axis")
                    task_dialog.MainInstruction = "Choose which axis you want the plane aligned to:"
                    task_dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Align with X-axis")
                    task_dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Align with Y-axis")
                    task_dialog.CommonButtons = TaskDialogCommonButtons.Cancel
                    result = task_dialog.Show()
                    if result == TaskDialogResult.CommandLink1:
                        x_vector = XYZ.BasisX
                        y_vector = line_vector
                    elif result == TaskDialogResult.CommandLink2:
                        x_vector = line_vector
                        y_vector = XYZ.BasisY
                    else:
                        return  # Cancel
            else:
                task_dialog = TaskDialog("Proceed with Horizontal Pipe")
                task_dialog.MainInstruction = "The pipe is horizontal. Do you want to create a work plane?"
                task_dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Vertical")
                task_dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Horizontal")
                task_dialog.CommonButtons = TaskDialogCommonButtons.Cancel
                result = task_dialog.Show()

                horizontal_dir = XYZ(line_vector.X, line_vector.Y, 0.0)
                if horizontal_dir.GetLength() == 0:
                    return
                horizontal_dir = horizontal_dir.Normalize()

                if result == TaskDialogResult.CommandLink1:   # Vertical
                    x_vector = horizontal_dir
                    y_vector = XYZ.BasisZ

                elif result == TaskDialogResult.CommandLink2:  # Horizontal
                    x_vector = horizontal_dir
                    y_vector = XYZ.BasisZ.CrossProduct(horizontal_dir).Normalize()

                else:
                    return  # Cancel
                plane_origin = (connector_1 + connector_2) / 2

        # =================================================================
        # CASE 2: FamilyInstance
        # =================================================================
        elif isinstance(elem, FamilyInstance):
            loc = elem.Location
            if not isinstance(loc, LocationPoint):
                return
            plane_origin = loc.Point

            if curview.ViewType == ViewType.Section:
                x_vector = curview.RightDirection.Normalize()
                y_vector = curview.UpDirection.Normalize()
            else:
                task_dialog = TaskDialog("Select Plane Orientation")
                task_dialog.MainInstruction = "Choose the plane orientation for the family instance:"
                task_dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Horizontal")
                task_dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Vertical")
                task_dialog.CommonButtons = TaskDialogCommonButtons.Cancel
                result = task_dialog.Show()
                if result == TaskDialogResult.Cancel:
                    return
                is_vertical = (result == TaskDialogResult.CommandLink2)
                align_x = True
                if is_vertical:
                    task_dialog = TaskDialog("Select Axis")
                    task_dialog.MainInstruction = "Choose the axis direction:"
                    task_dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "X-axis")
                    task_dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Y-axis")
                    task_dialog.CommonButtons = TaskDialogCommonButtons.Cancel
                    result = task_dialog.Show()
                    if result == TaskDialogResult.Cancel:
                        return
                    align_x = (result == TaskDialogResult.CommandLink1)
                if is_vertical:
                    x_vector = XYZ.BasisX if align_x else XYZ.BasisY
                    y_vector = XYZ.BasisZ
                else:
                    x_vector = XYZ.BasisX
                    y_vector = XYZ.BasisY
        else:
            return

        # =================================================================
        # Plane creation
        # =================================================================
        if plane_origin is None or x_vector is None or y_vector is None:
            return

        if curview.ViewType == ViewType.Section:
            x_vector = curview.RightDirection.Normalize()
            y_vector = curview.UpDirection.Normalize()

        plane = Plane.CreateByOriginAndBasis(plane_origin, x_vector, y_vector)

        tg = TransactionGroup(doc, "Create Planes")
        tg.Start()
        try:
            t = Transaction(doc, "Set WorkPlane")
            t.Start()
            sketch_plane = SketchPlane.Create(doc, plane)
            curview.SketchPlane = sketch_plane
            curview.ShowActiveWorkPlane()
            t.Commit()

            t = Transaction(doc, "Delete RefPlane")
            t.Start()
            refs_to_delete = [rp.Id for rp in FilteredElementCollector(doc)
                              .OfClass(ReferencePlane).WhereElementIsNotElementType()
                              if rp.Name == "TEMPORARY"]
            for rid in refs_to_delete:
                doc.Delete(rid)
            t.Commit()

            if curview.ViewType == ViewType.ThreeD:
                t = Transaction(doc, "Set RefPlane")
                t.Start()
                bubble_end = plane.Origin
                free_end = plane.Origin + plane.XVec
                cut_vec = plane.YVec
                ref_plane = doc.Create.NewReferencePlane(bubble_end, free_end, cut_vec, curview)
                ref_plane.Name = "TEMPORARY"
                t.Commit()

            tg.Assimilate()
        except:
            tg.RollBack()

    except:
        pass  # Quiet exit on any top-level cancellation or error

# Run the function
select_fabrication_pipe_and_create_plane()