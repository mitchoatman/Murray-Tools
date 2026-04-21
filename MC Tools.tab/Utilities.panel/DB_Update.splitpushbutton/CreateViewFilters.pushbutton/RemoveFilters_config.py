import clr
import Autodesk
from Autodesk.Revit.DB import FilterElement, Transaction

clr.AddReference("RevitServices")
from RevitServices.Persistence import DocumentManager

# Get the current document
doc = __revit__.ActiveUIDocument.Document

# Start a transaction
t = Transaction(doc, "Delete All Filters")
t.Start()

try:
    # Get all filter elements in the document
    collector = Autodesk.Revit.DB.FilteredElementCollector(doc).OfClass(FilterElement)
    filters = collector.ToElements()

    # Delete each filter
    for filter_element in filters:
        doc.Delete(filter_element.Id)
    
    t.Commit()
except Exception as e:
    t.RollBack()
    print("Error: ", str(e))
