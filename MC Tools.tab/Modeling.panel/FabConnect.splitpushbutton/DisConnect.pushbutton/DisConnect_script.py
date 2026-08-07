# coding: utf8
from Autodesk.Revit.DB import Transaction, FabricationPart
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.Exceptions import InvalidOperationException

# Get the active document and UI document
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

# List to collect messages
messages = []


class FabricationPartSelectionFilter(ISelectionFilter):
    """Allow only Fabrication Parts to be selected"""
    def AllowElement(self, element):
        return isinstance(element, FabricationPart)

    def AllowReference(self, reference, position):
        return False


def get_connector_manager(element):
    """Get connector manager for Fabrication parts"""
    if hasattr(element, 'ConnectorManager') and element.ConnectorManager:
        return element.ConnectorManager
    raise AttributeError("No connector manager found")


def disconnect_all_connectors(el):
    """Disconnect all connectors on an element"""
    try:
        connector_manager = get_connector_manager(el)
    except AttributeError:
        raise AttributeError("No connector manager")
    
    for connector in connector_manager.Connectors:
        if not connector.IsConnected:
            continue
        
        connected_connectors = []
        for other_connector in connector.AllRefs:
            if other_connector.Owner.Id != connector.Owner.Id:
                connected_connectors.append(other_connector)
        
        for other_connector in connected_connectors:
            try:
                connector.DisconnectFrom(other_connector)
            except Exception as ex:
                messages.append("Failed to disconnect connector: {}".format(str(ex)))


def disconnect():
    """Main function to disconnect selected fabrication parts"""
    global messages
    messages = []
    
    try:
        selection_filter = FabricationPartSelectionFilter()
        selected_refs = uidoc.Selection.PickObjects(
            ObjectType.Element,
            selection_filter,
            "Select fabrication parts to disconnect (press Finish when done)"
        )
        
        if not selected_refs:
            TaskDialog.Show("Disconnect Fabrication Parts", "No fabrication parts selected")
            return
        
        transaction = Transaction(doc, "Disconnect fabrication parts")
        transaction.Start()
        
        try:
            for ref in selected_refs:
                el = doc.GetElement(ref.ElementId)
                try:
                    disconnect_all_connectors(el)
                    messages.append("Disconnected Fabrication Part - ID: {}".format(
                        el.Id.IntegerValue
                    ))
                except Exception as e:
                    messages.append("Error disconnecting element {}: {}".format(
                        el.Id.IntegerValue,
                        str(e)
                    ))
            
            transaction.Commit()
            result_message = "\n".join(messages)
            TaskDialog.Show(
                "Disconnect Complete",
                result_message if result_message else "All fabrication parts disconnected successfully"
            )
            
        except Exception as e:
            transaction.RollBack()
            TaskDialog.Show("Error", "Transaction failed: {}".format(str(e)))
    
    except InvalidOperationException:
        TaskDialog.Show("Disconnect Fabrication Parts", "Selection cancelled by user")


disconnect()