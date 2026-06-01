# coding: utf8
from Autodesk.Revit.DB import Transaction, Reference
from Autodesk.Revit.UI import UIDocument, TaskDialog
from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.Exceptions import InvalidOperationException
from System.Collections.Generic import List

# Get the active document and UI document
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

# List to collect messages
messages = []


def get_connector_manager(element):
    """Get connector manager for both Revit MEP elements and Fabrication parts"""
    # Try MEP ConnectorManager
    if hasattr(element, 'ConnectorManager') and element.ConnectorManager:
        return element.ConnectorManager
    
    # Try MEP ConnectorManager on MEPModel
    if hasattr(element, 'MEPModel') and element.MEPModel:
        if hasattr(element.MEPModel, 'ConnectorManager') and element.MEPModel.ConnectorManager:
            return element.MEPModel.ConnectorManager
    
    # Try Fabrication ConnectorManager
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
        
        # Collect all connected connectors first to avoid collection modification
        connected_connectors = []
        for other_connector in connector.AllRefs:
            if other_connector.Owner.Id != connector.Owner.Id:
                connected_connectors.append(other_connector)
        
        # Disconnect from all connected connectors
        for other_connector in connected_connectors:
            try:
                connector.DisconnectFrom(other_connector)
            except Exception as ex:
                messages.append("Failed to disconnect connector: {}".format(str(ex)))
        
        # Handle MEP system division
        try:
            system = connector.MEPSystem
            if system and hasattr(system, 'IsMultipleNetwork') and system.IsMultipleNetwork:
                system.DivideSystem(doc)
        except:
            pass


def disconnect():
    """Main function to disconnect selected elements"""
    global messages
    messages = []
    
    try:
        # Prompt user to select elements
        selected_refs = uidoc.Selection.PickObjects(ObjectType.Element, "Select elements to disconnect (press Finish when done)")
        
        if not selected_refs:
            TaskDialog.Show("Disconnect Elements", "No elements selected")
            return
        
        transaction = Transaction(doc, "Disconnect elements")
        transaction.Start()
        
        try:
            for ref in selected_refs:
                el = doc.GetElement(ref.ElementId)
                try:
                    disconnect_all_connectors(el)
                    messages.append("Disconnected: {} - ID: {}".format(
                        el.Category.Name if el.Category else "Unknown",
                        el.Id.IntegerValue
                    ))
                except AttributeError as e:
                    messages.append("No connector manager found for {}: {}".format(
                        el.Category.Name if el.Category else "Unknown",
                        el.Id.IntegerValue
                    ))
                except Exception as e:
                    messages.append("Error disconnecting element {}: {}".format(
                        el.Id.IntegerValue,
                        str(e)
                    ))
            
            transaction.Commit()
            
            # Show results in TaskDialog
            result_message = "\n".join(messages)
            TaskDialog.Show("Disconnect Complete", result_message if result_message else "All elements disconnected successfully")
            
        except Exception as e:
            transaction.RollBack()
            TaskDialog.Show("Error", "Transaction failed: {}".format(str(e)))
    
    except InvalidOperationException:
        TaskDialog.Show("Disconnect Elements", "Selection cancelled by user")


disconnect()