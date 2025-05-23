import Autodesk
from Autodesk.Revit import DB
from Autodesk.Revit.DB import Transaction, FilteredElementCollector, BuiltInCategory, BuiltInParameter
from pyrevit import revit, DB, script, forms, UI
from Parameters.Get_Set_Params import get_parameter_value_by_name_AsValueString

#define the active Revit application and document
DB = Autodesk.Revit.DB
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView
fec = FilteredElementCollector
app = doc.Application
RevitVersion = app.VersionNumber
RevitINT = float (RevitVersion)

selection = revit.get_selection()

# Creating collector instance and collecting all the fabrication hangers from the model
pipe_collector = FilteredElementCollector(doc, curview.Id).OfCategory(BuiltInCategory.OST_FabricationPipework) \
                   .WhereElementIsNotElementType() \
                   .ToElements()

found = False

elementlist = []
t = Transaction(doc, 'Select Pipes')
#Start Transaction
t.Start()
output = script.get_output()
# Print the instruction once at the top
print('Pick on ID to select or Magnifying Glass to zoom to:')

for pipe in pipe_collector:
    CID = pipe.ItemCustomId
    if CID == 2041:
        pipelen = pipe.Parameter[BuiltInParameter.FABRICATION_PART_LENGTH].AsDouble()
        pipemat = pipe.Parameter[BuiltInParameter.FABRICATION_PART_MATERIAL].AsValueString()
        if pipelen > 10.002 and pipemat == 'Cast Iron: Cast Iron':
            # Get the family name instead of element name
            family_name = get_parameter_value_by_name_AsValueString(pipe, 'Family')
            print('{}: {}'.format(family_name, output.linkify(pipe.Id)))
            found = True

if not found:
    print('Nothing Found')
#End Transaction
t.Commit()