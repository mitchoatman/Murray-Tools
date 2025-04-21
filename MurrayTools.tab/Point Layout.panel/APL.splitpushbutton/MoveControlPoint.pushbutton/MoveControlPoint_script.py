import clr
clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import FabricationConfiguration


doc = __revit__.ActiveUIDocument.Document
config = FabricationConfiguration.GetFabricationConfiguration(doc)

if config:
    # Get all status IDs
    status_ids = config.GetAllPartStatuses()
    
    if status_ids and status_ids.Count > 0:
        # Convert IDs to names directly
        status_names = []
        for status_id in status_ids:
            name = config.GetPartStatusDescription(status_id)
            status_names.append(name)
        
        # Print the list with numbers
        print "Complete List of Fabrication Status Names:"
        for i, name in enumerate(status_names):
            print "%d. %s" % (i + 1, name)
    else:
        print "No statuses found in the fabrication configuration."
else:
    print "No fabrication configuration found in this project."