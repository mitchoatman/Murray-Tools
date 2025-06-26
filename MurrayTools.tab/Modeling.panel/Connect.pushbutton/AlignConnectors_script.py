# coding: utf8
from math import pi
from Autodesk.Revit.DB import Line, InsulationLiningBase, Transaction, Element, ConnectorManager, ConnectorSet, Connector, XYZ, FabricationPart, ElementTransformUtils
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

def get_connector_manager(element):
    # type: (Element) -> ConnectorManager
    """Return element connector manager"""
    try:
        # Return ConnectorManager for pipes, ducts, fabrication parts, etc.
        return element.ConnectorManager
    except AttributeError:
        pass

    try:
        # Return ConnectorManager for family instances etc.
        return element.MEPModel.ConnectorManager
    except AttributeError:
        raise AttributeError("Cannot find connector manager in given element")


def get_connector_closest_to(connectors, xyz):
    # type: (ConnectorSet, XYZ) -> Connector
    """Get connector from connector set or any iterable closest to an XYZ point"""
    min_distance = float("inf")
    closest_connector = None
    for connector in connectors:
        distance = connector.Origin.DistanceTo(xyz)
        if distance < min_distance:
            min_distance = distance
            closest_connector = connector
    return closest_connector


class NoInsulation(ISelectionFilter):
    def AllowElement(self, elem):
        if isinstance(elem, InsulationLiningBase):
            return False
        try:
            # Allow fabrication parts or elements with connector manager
            if isinstance(elem, FabricationPart):
                return True
            get_connector_manager(elem)
            return True
        except AttributeError:
            return False

    def AllowReference(self, reference, position):
        return True


def connect_to():
    # Prompt user to select elements and points to connect
    try:
        reference = uidoc.Selection.PickObject(ObjectType.Element, NoInsulation(), "Pick element to move and connect. [ESC] to Cancel")
    except:
        return False

    try:
        moved_element = doc.GetElement(reference)
        moved_point = reference.GlobalPoint
        reference = uidoc.Selection.PickObject(ObjectType.Element, NoInsulation(), "Pick element to be connected to")
        target_element = doc.GetElement(reference)
        target_point = reference.GlobalPoint
    except:
        TaskDialog.Show("Error", "Oops, it looks like you chose an invalid object. Please try again.")
        return False

    # Check if both elements are fabrication parts
    if isinstance(moved_element, FabricationPart) and isinstance(target_element, FabricationPart):
        t = Transaction(doc, "Connect and couple fabrication parts")
        try:
            t.Start()
            try:
                # Get connector manager for moved fabrication part
                moved_cm = get_connector_manager(moved_element)
                if not moved_cm:
                    print "No connector manager found for the moved fabrication part."
                    t.Commit()
                    return True
                # Get unused connectors for moved element
                moved_connectors = moved_cm.UnusedConnectors
                if not moved_connectors or moved_connectors.Size == 0:
                    TaskDialog.Show("No unused connectors found on the moved fabrication part.")
                    t.Commit()
                    return True
                # Get closest connector for moved element
                moved_connector = get_connector_closest_to(moved_connectors, moved_point)
                if not moved_connector:
                    TaskDialog.Show("Failed to find a valid connector on the moved fabrication part.")
                    t.Commit()
                    return True

                # Get connector manager for target fabrication part
                target_cm = get_connector_manager(target_element)
                if not target_cm:
                    print "No connector manager found for the target fabrication part."
                    t.Commit()
                    return True
                # Get unused connectors for target element
                target_connectors = target_cm.UnusedConnectors
                if not target_connectors or target_connectors.Size == 0:
                    TaskDialog.Show("Error", "No unused connectors found on the target fabrication part.")
                    t.Commit()
                    return True
                # Get closest connector for target element
                target_connector = get_connector_closest_to(target_connectors, target_point)
                if not target_connector:
                    TaskDialog.Show("Failed to find a valid connector on the target fabrication part.")
                    t.Commit()
                    return True

                # Check connector domains
                try:
                    if moved_connector.Domain != target_connector.Domain:
                        print "Selected connectors are from different domains. Please retry."
                        t.Commit()
                        return True
                except AttributeError:
                    print "Unable to verify connector domains."
                    t.Commit()
                    return True

                # Try aligning with AlignPartByConnectors
                try:
                    FabricationPart.AlignPartByConnectors(doc, moved_element, moved_connector, target_connector)
                except Exception as e:
                    # print "AlignPartByConnectors failed: {}. Falling back to manual alignment.".format(str(e))
                    # Manual alignment to make moved_connector's direction opposite to target_connector's
                    try:
                        moved_direction = moved_connector.CoordinateSystem.BasisZ
                        target_direction = target_connector.CoordinateSystem.BasisZ
                        angle = moved_direction.AngleTo(-target_direction)  # Aim for opposite direction
                        if angle > 0.0001:  # Small angle threshold to avoid numerical issues
                            vector = moved_direction.CrossProduct(-target_direction)
                            if vector.IsZeroLength():  # If parallel, use BasisY
                                vector = moved_connector.CoordinateSystem.BasisY
                            # Use ElementTransformUtils for robust rotation
                            axis = Line.CreateBound(moved_point, moved_point + vector)
                            ElementTransformUtils.RotateElement(doc, moved_element.Id, axis, angle)
                        # Move element to match connector position
                        moved_element.Location.Move(target_connector.Origin - moved_connector.Origin)
                    except Exception as e:
                        # print "Manual alignment failed: {}".format(str(e))
                        t.Commit()
                        return True

                # Connect and couple the parts
                FabricationPart.ConnectAndCouple(doc, moved_connector, target_connector)
                t.Commit()
                return True
            except Exception as e:
                TaskDialog.Show("Failed to connect and couple fabrication parts: {}".format(str(e)))
                t.RollBack()
                return True
        except:
            if t.HasStarted():
                t.RollBack()
            return True
    else:
        t = Transaction(doc, "Connect elements")
        try:
            t.Start()
            # Original logic for non-fabrication parts
            # Get associated unused connectors
            try:
                moved_connector = get_connector_closest_to(get_connector_manager(moved_element).UnusedConnectors,
                                                           moved_point)
                target_connector = get_connector_closest_to(get_connector_manager(target_element).UnusedConnectors,
                                                            target_point)
            except AttributeError:
                TaskDialog.Show("It looks like one of the objects have no unused connector")
                t.Commit()
                return True

            try:
                if moved_connector.Domain != target_connector.Domain:
                    print "You picked 2 connectors of different domain. Please retry."
                    t.Commit()
                    return True
            except AttributeError:
                TaskDialog.Show("It looks like one of the objects have no unused connector")
                t.Commit()
                return True

            # Retrieve connectors' direction and catch attribute error
            try:
                moved_direction = moved_connector.CoordinateSystem.BasisZ
                target_direction = target_connector.CoordinateSystem.BasisZ
            except AttributeError:
                TaskDialog.Show("It looks like one of the objects have no unused connector")
                t.Commit()
                return True

            # Move and connect
            angle = moved_direction.AngleTo(target_direction)
            if angle != pi:
                if angle == 0:
                    vector = moved_connector.CoordinateSystem.BasisY
                else:
                    vector = moved_direction.CrossProduct(target_direction)
                try:
                    line = Line.CreateBound(moved_point, moved_point + vector)
                    moved_element.Location.Rotate(line, angle - pi)
                # Revit doesn't like angle and distance too close to 0
                except:
                    pass
            # Move element to match connector position
            moved_element.Location.Move(target_connector.Origin - moved_connector.Origin)
            # Connect connectors
            moved_connector.ConnectTo(target_connector)
            t.Commit()
            return True
        except:
            if t.HasStarted():
                t.RollBack()
            return True


while connect_to():
    pass