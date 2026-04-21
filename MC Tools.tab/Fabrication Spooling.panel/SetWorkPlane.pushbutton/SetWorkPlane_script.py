import clr
from Autodesk.Revit.DB import (
    Transaction, FabricationPart, ConnectorType, XYZ, Plane, SketchPlane,
    ViewType, FilteredElementCollector, ReferencePlane,
    TransactionGroup, LocationPoint, FamilyInstance
)
from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.UI import TaskDialog, TaskDialogCommandLinkId, TaskDialogCommonButtons, TaskDialogResult
clr.AddReference("RevitServices")
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager

def select_fabrication_pipe_and_create_plane():
    doc = __revit__.ActiveUIDocument.Document
    uidoc = __revit__.ActiveUIDocument
    curview = doc.ActiveView

    try:
        if curview.ViewType != ViewType.ThreeD:
            print("Must be in a 3D view")
            return

        # Unified selection prompt
        ref = uidoc.Selection.PickObject(ObjectType.Element, "Select a fabrication part or family instance")
        elem = doc.GetElement(ref.ElementId)

        x_vector = None
        y_vector = None
        plane_origin = None

        # =================================================================
        # CASE 1: FabricationPart - your original working logic (unchanged)
        # =================================================================
        if isinstance(elem, FabricationPart):
            # Get connectors directly from FabricationPart
            connectors = list(elem.ConnectorManager.Connectors)
            if len(connectors) < 2:
                print("Error: Selected fabrication part does not have enough connectors.")
                return

            connector_points = [conn.Origin for conn in connectors if conn.ConnectorType == ConnectorType.End]
            if len(connector_points) < 2:
                print("Error: Unable to find two valid endpoints for the selected fabrication part.")
                return

            connector_1 = connector_points[0]
            connector_2 = connector_points[1]
            line_vector = (connector_2 - connector_1).Normalize()

            if abs(line_vector.Z) > max(abs(line_vector.X), abs(line_vector.Y)):
                # Vertical - identical to your original
                task_dialog = TaskDialog("Select Axis")
                task_dialog.MainInstruction = "Choose which axis you want the plane aligned to:"
                task_dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Align with X-axis")
                task_dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Align with Y-axis")
                task_dialog.CommonButtons = TaskDialogCommonButtons.Cancel
                result = task_dialog.Show()

                if result == TaskDialogResult.CommandLink1:
                    x_vector = XYZ.BasisX
                    y_vector = line_vector
                    plane_origin = (connector_1 + connector_2) / 2
                elif result == TaskDialogResult.CommandLink2:
                    x_vector = line_vector
                    y_vector = XYZ.BasisY
                    plane_origin = (connector_1 + connector_2) / 2
                else:
                    print("Operation cancelled by the user.")
                    return
            else:
                # Horizontal - identical to your original (fixed minor bugs present in original)
                task_dialog = TaskDialog("Proceed with Horizontal Pipe")
                task_dialog.MainInstruction = "The pipe is horizontal. Do you want to create a work plane?"
                task_dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Vertical")
                task_dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Horizontal")
                task_dialog.CommonButtons = TaskDialogCommonButtons.Cancel
                result = task_dialog.Show()

                if result == TaskDialogResult.CommandLink1:  # Vertical plane
                    x_vector = XYZ.BasisX if abs(line_vector.X) > abs(line_vector.Y) else XYZ.BasisY
                    y_vector = XYZ.BasisZ
                elif result == TaskDialogResult.CommandLink2:  # Horizontal plane
                    x_vector = XYZ.BasisX if abs(line_vector.X) > abs(line_vector.Y) else XYZ.BasisY
                    y_vector = XYZ.BasisY if abs(line_vector.X) > abs(line_vector.Y) else XYZ.BasisX
                else:
                    print("Operation cancelled by the user.")
                    return
                plane_origin = (connector_1 + connector_2) / 2

        # =================================================================
        # CASE 2: FamilyInstance - point-based logic with your original prompts
        # =================================================================
        else:
            if not isinstance(elem, FamilyInstance):
                print("Error: Selected element is neither a FabricationPart nor a FamilyInstance.")
                return

            loc = elem.Location
            if not isinstance(loc, LocationPoint):
                print("Error: Selected family instance has no valid location point.")
                return
            plane_origin = loc.Point

            task_dialog = TaskDialog("Select Plane Orientation")
            task_dialog.MainInstruction = "Choose the plane orientation for the family instance:"
            task_dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Horizontal")
            task_dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Vertical")
            task_dialog.CommonButtons = TaskDialogCommonButtons.Cancel
            result = task_dialog.Show()

            if result == TaskDialogResult.Cancel:
                print("Operation cancelled by the user.")
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
                    print("Operation cancelled by the user.")
                    return
                align_x = (result == TaskDialogResult.CommandLink1)

            if is_vertical:
                x_vector = XYZ.BasisX if align_x else XYZ.BasisY
                y_vector = XYZ.BasisZ
            else:
                x_vector = XYZ.BasisX
                y_vector = XYZ.BasisY

        # =================================================================
        # Plane creation - identical for both cases
        # =================================================================
        if plane_origin is None or x_vector is None or y_vector is None:
            print("Error: Failed to determine plane parameters.")
            return

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

            t = Transaction(doc, "Set RefPlane")
            t.Start()
            bubble_end = plane.Origin
            free_end = plane.Origin + plane.XVec
            cut_vec = plane.YVec
            ref_plane = doc.Create.NewReferencePlane(bubble_end, free_end, cut_vec, curview)
            ref_plane.Name = "TEMPORARY"
            t.Commit()

            tg.Assimilate()
            # print("Success: Work plane created and set.")
        except Exception as ex:
            tg.RollBack()
            print("Error during transaction:", str(ex))

    except Exception as e:
        print("Error:", str(e))

# Run the function
select_fabrication_pipe_and_create_plane()