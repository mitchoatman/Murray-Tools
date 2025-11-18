import clr
from Autodesk.Revit.DB import Transaction, FabricationPart, ConnectorType, XYZ, Plane, SketchPlane, ViewType, FilteredElementCollector, ReferencePlane, ElementId, TransactionGroup, LocationPoint, FamilyInstance
from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.UI import TaskDialog, TaskDialogCommandLinkId, TaskDialogCommonButtons, TaskDialogResult
clr.AddReference("RevitServices")
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager
def select_fabrication_pipe_and_create_plane():
    # Get the current Revit document and UI document
    doc = __revit__.ActiveUIDocument.Document
    uidoc = __revit__.ActiveUIDocument
    curview = doc.ActiveView
    try:
        if str(curview.ViewType) == 'ThreeD':
            # Prompt the user to select a fabrication pipe or family instance
            ref = uidoc.Selection.PickObject(ObjectType.Element, "Select a fabrication pipe or family instance")
            elem = doc.GetElement(ref.ElementId)
            # Check if the selected element is a family instance
            if not isinstance(elem, FamilyInstance):
                print("Error: Selected element is not a family instance.")
                return
            x_vector = None
            y_vector = None
            plane_origin = None
            end_connectors = []
            connectors = []
            try:
                connectors = list(elem.ConnectorManager.Connectors)
                end_connectors = [conn for conn in connectors if conn.ConnectorType == ConnectorType.End]
            except AttributeError:
                end_connectors = []
            if len(end_connectors) >= 2:
                # Treat as pipe: original logic
                connector_points = [conn.Origin for conn in end_connectors]
                if len(connector_points) < 2:
                    print("Error: Unable to find two valid endpoints for the selected pipe.")
                    return
                connector_1 = connector_points[0]
                connector_2 = connector_points[1]
                # Create a vector between the two connectors
                line_vector = connector_2 - connector_1 # Vector from connector 1 to connector 2
                line_vector = line_vector.Normalize()
                # Check if the pipe is vertical or horizontal
                if abs(line_vector.Z) > max(abs(line_vector.X), abs(line_vector.Y)):
                    # Pipe is vertical (connectors differ in Z)
                    # Prompt the user to choose whether to align with X or Y axis
                    task_dialog = TaskDialog("Select Axis")
                    task_dialog.MainInstruction = "Choose which axis you want the plane aligned to:"
                    task_dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Align with X-axis")
                    task_dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Align with Y-axis")
                    task_dialog.CommonButtons = TaskDialogCommonButtons.Cancel
                    result = task_dialog.Show()
                    if result == TaskDialogResult.CommandLink1:
                        # Align with X-axis
                        x_vector = XYZ.BasisX
                        y_vector = line_vector # Use the line vector as Y (already normalized)
                        plane_origin = (connector_1 + connector_2) / 2
                    elif result == TaskDialogResult.CommandLink2:
                        # Align with Y-axis
                        x_vector = line_vector # Use the line vector as X (already normalized)
                        y_vector = XYZ.BasisY
                        plane_origin = (connector_1 + connector_2) / 2
                    else:
                        print("Operation cancelled by the user.")
                        return
                else:
                    # Pipe is horizontal (connectors differ in X or Y)
                    # Ask user if they want to proceed with creating the work plane
                    task_dialog = TaskDialog("Proceed with Horizontal Pipe")
                    task_dialog.MainInstruction = "The pipe is horizontal. Do you want to create a work plane?"
                    task_dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Vertical")
                    task_dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Horizontal")
                    task_dialog.CommonButtons = TaskDialogCommonButtons.Cancel
                    result = task_dialog.Show()
                    if result == TaskDialogResult.CommandLink1:
                        # For horizontal pipe, align with X-axis
                        if abs(line_vector.X) > abs(line_vector.Y):
                            # Pipe is aligned with X-axis
                            x_vector = XYZ.BasisX # Use BasisX
                            y_vector = XYZ.BasisZ # Use BasisZ for the Y vector
                        else:
                            # Pipe is aligned with Y-axis
                            x_vector = XYZ.BasisY # Use BasisY
                            y_vector = XYZ.BasisZ # Use BasisZ for the Y vector
                        plane_origin = (connector_1 + connector_2) / 2 # Midpoint in 3D space
                    elif result == TaskDialogResult.CommandLink2:
                        # Align with Y-axis for horizontal pipe
                        if abs(line_vector.X) > abs(line_vector.Y):
                            # Pipe is aligned with X-axis
                            x_vector = XYZ.BasisX # Use BasisX
                            y_vector = XYZ.BasisY # Use BasisY for the Y vector
                        else:
                            # Pipe is aligned with Y-axis
                            x_vector = XYZ.BasisY # Use BasisY
                            y_vector = XYZ.BasisX # Use BasisX for the Y vector
                        plane_origin = (connector_1 + connector_2) / 2
                    else:
                        print("Operation cancelled by the user.")
                        return
            else:
                # Treat as family instance (e.g., fabrication hanger or other family): new logic
                loc = elem.Location
                if isinstance(loc, LocationPoint):
                    plane_origin = loc.Point
                else:
                    print("Error: Cannot determine center point for the selected family instance.")
                    return
                # Prompt for plane orientation: horizontal or vertical
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
                align_x = True  # Default for horizontal
                if is_vertical:
                    # Prompt for axis alignment: X or Y direction only for vertical
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
                    if align_x:
                        x_vector = XYZ.BasisX
                        y_vector = XYZ.BasisZ
                    else:
                        x_vector = XYZ.BasisY
                        y_vector = XYZ.BasisZ
                else:
                    x_vector = XYZ.BasisX
                    y_vector = XYZ.BasisY
            if x_vector is None or y_vector is None or plane_origin is None:
                print("Error: Failed to determine plane parameters.")
                return
            # Create the plane with the chosen alignment
            plane = Plane.CreateByOriginAndBasis(plane_origin, x_vector, y_vector)
            tg = TransactionGroup(doc, "Create Planes")
            tg.Start()
            try:
                t = Transaction(doc, 'Set WorkPlane')
                t.Start()
                sketch_plane = SketchPlane.Create(doc, plane)
                curview.SketchPlane = sketch_plane
                curview.ShowActiveWorkPlane()
                t.Commit()
                t = Transaction(doc, 'Delete RefPlane')
                t.Start()
                refs_to_delete = []
                existing_refs = FilteredElementCollector(doc).OfClass(ReferencePlane).WhereElementIsNotElementType()
                for rp in existing_refs:
                    if rp.Name == "TEMPORARY":
                        refs_to_delete.append(rp.Id)
                for rid in refs_to_delete:
                    doc.Delete(rid)
                t.Commit()
                t = Transaction(doc, 'Set RefPlane')
                t.Start()
                bubble_end = plane.Origin
                free_end = plane.Origin + plane.XVec
                cut_vec = plane.YVec
                ref_plane = doc.Create.NewReferencePlane(bubble_end, free_end, cut_vec, curview)
                ref_plane.Name = "TEMPORARY"
                t.Commit()
                tg.Assimilate()
            except Exception as e:
                tg.RollBack()
            # Notify the user
            # print("Success: Work plane created and set based on user choice.")
        else:
            print 'Must be in a 3D view'
    except Exception as e:
        print("Error:", str(e))
# Run the function
select_fabrication_pipe_and_create_plane()