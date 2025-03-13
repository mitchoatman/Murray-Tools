from Autodesk.Revit.DB import Transaction, FabricationPart
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# Custom selection filter for MEP Fabrication Hangers
class CustomISelectionFilter(ISelectionFilter):
    def __init__(self, allowed_categories):
        self.allowed_categories = allowed_categories
    def AllowElement(self, e):
        if e.Category and e.Category.Name in self.allowed_categories:
            # Additional check to ensure it's a hanger
            if isinstance(e, FabricationPart) and e.IsAHanger():
                return True
        return False
    def AllowReference(self, ref, point):
        return True

# Define the category for hangers
allowed_categories = ["MEP Fabrication Hangers"]

# Try to get user selection with error handling for cancellation
try:
    selection_filter = CustomISelectionFilter(allowed_categories)
    selected_parts = uidoc.Selection.PickObjects(ObjectType.Element, selection_filter, "Select Fabrication Hangers to disconnect from hosts")
    fab_hangers = [doc.GetElement(elId) for elId in selected_parts]

    # Start a transaction to modify elements only if selection was successful
    if fab_hangers:  # Check if list is not empty
        t = Transaction(doc, "Disconnect Hangers from Hosts")
        t.Start()

        successful_disconnects = 0
        for hanger in fab_hangers:
            try:
                # Verify it's a hanger
                if hanger.IsAHanger():
                    # Get the hosted info for the hanger
                    hosted_info = hanger.GetHostedInfo()
                    if hosted_info:
                        # Disconnect using the hosted info
                        hosted_info.DisconnectFromHost()
                        successful_disconnects += 1
                        print("Disconnected hanger {} from its host".format(hanger.Id))
                    else:
                        print("No host information found for hanger {}".format(hanger.Id))
                else:
                    print("Element {} is not a hanger".format(hanger.Id))
                    
            except Exception as e:
                print("Failed to disconnect hanger {}: {}".format(hanger.Id, str(e)))

        t.Commit()
        print("Successfully disconnected {} hangers".format(successful_disconnects))
    else:
        print("No hangers selected")

except:
    # Silently exit if user cancels the selection
    pass