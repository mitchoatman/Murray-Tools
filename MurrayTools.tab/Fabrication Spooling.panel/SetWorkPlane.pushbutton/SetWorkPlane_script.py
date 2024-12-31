import clr
from Autodesk.Revit.DB import Transaction, FabricationPart, ConnectorType, XYZ, Plane, SketchPlane, ViewType
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
            # Prompt the user to select a fabrication pipe
            ref = uidoc.Selection.PickObject(ObjectType.Element, "Select a fabrication pipe")
            pipe = doc.GetElement(ref.ElementId)

            # Check if the selected element is a fabrication part
            if not isinstance(pipe, FabricationPart):
                print("Error: Selected element is not a fabrication part.")
                return

            # Get the connectors from the pipe's ConnectorManager
            connectors = list(pipe.ConnectorManager.Connectors)

            # Ensure there are at least 2 connectors
            if len(connectors) < 2:
                print("Error: Selected pipe does not have enough connectors.")
                return

            # Extract two endpoints (origins) from the connectors
            connector_points = [conn.Origin for conn in connectors if conn.ConnectorType == ConnectorType.End]
            if len(connector_points) < 2:
                print("Error: Unable to find two valid endpoints for the selected pipe.")
                return

            connector_1 = connector_points[0]
            connector_2 = connector_points[1]

            # Create a vector between the two connectors
            line_vector = connector_2 - connector_1  # Vector from connector 1 to connector 2
            line_vector.Normalize()

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
                    y_vector = line_vector.Normalize()  # Use the line vector as Y
                    bounding_box = pipe.get_BoundingBox(doc.ActiveView)
                    if bounding_box:
                        plane_origin = (bounding_box.Min + bounding_box.Max) / 2  # Center of the bounding box
                    else:
                        plane_origin = (connector_1 + connector_2) / 2  # Fallback to midpoint

                    # Adjust alignment vectors based on the pipe's direction
                    if abs(line_vector.X) > abs(line_vector.Y):  # X-aligned pipe
                        x_vector = XYZ.BasisX
                        y_vector = XYZ.BasisZ  # Ensure the plane is vertical
                    elif abs(line_vector.Y) > abs(line_vector.X):  # Y-aligned pipe
                        x_vector = XYZ.BasisY
                        y_vector = XYZ.BasisZ  # Ensure the plane is vertical
                    else:
                        x_vector = XYZ.BasisX
                        y_vector = XYZ.BasisY

                elif result == TaskDialogResult.CommandLink2:
                    # Align with Y-axis
                    x_vector = line_vector.Normalize()  # Use the line vector as X
                    y_vector = XYZ.BasisY
                    bounding_box = pipe.get_BoundingBox(doc.ActiveView)
                    if bounding_box:
                        plane_origin = (bounding_box.Min + bounding_box.Max) / 2  # Center of the bounding box
                    else:
                        plane_origin = (connector_1 + connector_2) / 2  # Fallback to midpoint

                    # Adjust alignment vectors based on the pipe's direction
                    if abs(line_vector.X) > abs(line_vector.Y):  # X-aligned pipe
                        x_vector = XYZ.BasisX
                        y_vector = XYZ.BasisZ  # Ensure the plane is vertical
                    elif abs(line_vector.Y) > abs(line_vector.X):  # Y-aligned pipe
                        x_vector = XYZ.BasisY
                        y_vector = XYZ.BasisZ  # Ensure the plane is vertical
                    else:
                        x_vector = XYZ.BasisX
                        y_vector = XYZ.BasisY

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
                        x_vector = XYZ.BasisX  # Use BasisX
                        y_vector = XYZ.BasisZ  # Use BasisZ for the Y vector
                    else:
                        # Pipe is aligned with Y-axis
                        x_vector = XYZ.BasisY  # Use BasisY
                        y_vector = XYZ.BasisZ  # Use BasisZ for the Y vector
                    plane_origin = (connector_1 + connector_2) / 2  # Midpoint in 3D space
                elif result == TaskDialogResult.CommandLink2:
                    # Align with Y-axis for horizontal pipe
                    if abs(line_vector.X) > abs(line_vector.Y):
                        # Pipe is aligned with X-axis
                        x_vector = XYZ.BasisX  # Use BasisX
                        y_vector = XYZ.BasisY  # Use BasisZ for the Y vector
                    else:
                        # Pipe is aligned with Y-axis
                        x_vector = XYZ.BasisX  # Use BasisY
                        y_vector = XYZ.BasisY  # Use BasisZ for the Y vector
                    plane_origin = (connector_1 + connector_2) / 2
                else:
                    print("Operation cancelled by the user.")
                    return

            # Create the plane with the chosen alignment
            plane = Plane.CreateByOriginAndBasis(plane_origin, x_vector, y_vector)

            # Start a transaction to create the SketchPlane and set it active
            t = Transaction(doc, 'Set WorkPlane')
            t.Start()
            sketch_plane = SketchPlane.Create(doc, plane)

            # Set the newly created sketch plane as the active work plane
            curview = doc.ActiveView
            curview.SketchPlane = sketch_plane
            t.Commit()

            # Notify the user
            # print("Success: Work plane created and set based on user choice.")
        else:
            print 'Must be in a 3D view'
    except Exception as e:
        print("Error:", str(e))

# Run the function
select_fabrication_pipe_and_create_plane()
