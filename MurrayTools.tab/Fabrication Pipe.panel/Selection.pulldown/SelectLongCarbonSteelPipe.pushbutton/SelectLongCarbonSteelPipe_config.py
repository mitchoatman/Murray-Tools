import Autodesk
from Autodesk.Revit import DB
from Autodesk.Revit.DB import Transaction, FilteredElementCollector, BuiltInCategory, BuiltInParameter, ElementParameterFilter, FilterStringRule, ParameterValueProvider, FilterStringEquals, ElementId, FilterStringBeginsWith
from pyrevit import script
from Parameters.Get_Set_Params import get_parameter_value_by_name_AsValueString

#define the active Revit application and document
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView
app = doc.Application
RevitVersion = float(app.VersionNumber)

selected_ids = uidoc.Selection.GetElementIds()
selection = [doc.GetElement(eid) for eid in selected_ids]

import time
start = time.time()

# Create material filter
param_id = ElementId(BuiltInParameter.FABRICATION_PART_MATERIAL)
provider = ParameterValueProvider(param_id)
evaluator = FilterStringEquals()

# Create material filter
param_id = ElementId(BuiltInParameter.FABRICATION_PART_MATERIAL)
provider = ParameterValueProvider(param_id)
evaluator = FilterStringBeginsWith()
if RevitVersion <= 2021:
    rule = FilterStringRule(provider, evaluator, "Carbon Steel:", True)
else:
    rule = FilterStringRule(provider, evaluator, "Carbon Steel:")
material_filter = ElementParameterFilter(rule)

# Creating collector instance with material filter
pipe_collector = FilteredElementCollector(doc, curview.Id).OfCategory(BuiltInCategory.OST_FabricationPipework) \
                   .WhereElementIsNotElementType() \
                   .WherePasses(material_filter) \
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
        if pipelen > 20.0:
            # Get the family name instead of element name
            family_name = get_parameter_value_by_name_AsValueString(pipe, 'Family')
            print('{}: Length (ft): {}: {}'.format(family_name, pipelen, output.linkify(pipe.Id)))
            found = True

if not found:
    print('Nothing Found')
#End Transaction
t.Commit()

print("Time taken: {} seconds".format(time.time() - start))
