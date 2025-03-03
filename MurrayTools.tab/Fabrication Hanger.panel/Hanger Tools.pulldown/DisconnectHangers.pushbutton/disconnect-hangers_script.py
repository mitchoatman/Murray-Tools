from Autodesk.Revit.DB import Transaction, FabricationPart
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# Custom selection filter for MEP Fabrication Ductwork and Pipework
class CustomISelectionFilter(ISelectionFilter):
    def __init__(self, allowed_categories):
        self.allowed_categories = allowed_categories
    def AllowElement(self, e):
        if e.Category and e.Category.Name in self.allowed_categories:
            return True
        return False
    def AllowReference(self, ref, point):
        return True

# Define the categories we want to select
allowed_categories = ["MEP Fabrication Ductwork", "MEP Fabrication Pipework"]

# Prompt user to select ductwork or pipework
selection_filter = CustomISelectionFilter(allowed_categories)
selected_parts = uidoc.Selection.PickObjects(ObjectType.Element, selection_filter, "Select Fabrication Ductwork or Pipework to disconnect")
fab_parts = [doc.GetElement(elId) for elId in selected_parts]

# Start a transaction to modify elements
t = Transaction(doc, "Disconnect Fabrication Parts")
t.Start()

for part in fab_parts:
    try:
        # Get all connectors of the fabrication part
        connectors = part.ConnectorManager.Connectors
        if connectors.Size > 0:
            for connector in connectors:
                # Check if the connector is connected
                if connector.IsConnected:
                    # Get all connecting connectors
                    connected_connectors = connector.AllRefs
                    for connected in connected_connectors:
                        # Disconnect from each connected connector
                        connector.DisconnectFrom(connected)
                    print("Disconnected connectors for part {}".format(part.Id))
                else:
                    print("No connections found for part {}".format(part.Id))
        else:
            print("No connectors found for part {}".format(part.Id))
    except Exception as e:
        print("Failed to disconnect part {}: {}".format(part.Id, e))

t.Commit()