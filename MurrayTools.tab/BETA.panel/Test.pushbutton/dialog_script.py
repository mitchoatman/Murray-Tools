import Autodesk
from Autodesk.Revit.DB import FilteredElementCollector, ElementCategoryFilter, FamilyInstance, BuiltInCategory, ElementClassFilter, BuiltInParameter, \
                                ElementId, ElementParameterFilter, ParameterValueProvider, FilterStringRule, FilterStringEquals, LogicalAndFilter, Transaction
from Parameters.Get_Set_Params import set_parameter_by_name, get_parameter_value_by_name_AsString, \
                                        get_parameter_value_by_name_AsInteger, get_parameter_value_by_name_AsValueString, get_parameter_value_by_name_AsDouble
from SharedParam.Add_Parameters import Shared_Params
Shared_Params()

DB = Autodesk.Revit.DB
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView
fec = FilteredElementCollector
app = doc.Application
RevitVersion = app.VersionNumber
RevitINT = float (RevitVersion)

# Create a filter for pipe accessories
pipe_accessory_filter = ElementCategoryFilter(BuiltInCategory.OST_PipeAccessory)

# Create a filter for family instances (optional, as pipe accessories are already family instances)
family_instance_filter = ElementClassFilter(FamilyInstance)

# Create a parameter value provider for the family name
provider = ParameterValueProvider(ElementId(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM))

# Define the rule to filter by the family name "Pipe Label"
# Remove False for new Revit versions
if RevitINT > 2022:
    rule = FilterStringRule(provider, FilterStringEquals(), "Pipe Label")
else:
    rule = FilterStringRule(provider, FilterStringEquals(), "Pipe Label", False)

# Create a parameter filter using the rule
family_name_filter = ElementParameterFilter(rule)

# Combine the filters
filter = LogicalAndFilter(pipe_accessory_filter, family_name_filter)

# Collect all elements that match the filter
pipe_labels = FilteredElementCollector(doc).WherePasses(filter).ToElements()

t = Transaction(doc, "Set Pipe Label type")
t.Start()
# Iterate over elements and fetch parameter values
for pipe_label in pipe_labels:
    try:
        diameter = (get_parameter_value_by_name_AsDouble(pipe_label, 'Diameter') * 12)
        # print diameter
        # print diameter <= 0.5
        if diameter <= 0.50:
            set_parameter_by_name(pipe_label, 'FP_Product Entry', 'AA')
        if 0.50 < diameter < 1.00:
            set_parameter_by_name(pipe_label, 'FP_Product Entry', 'A')
        if 1.00 < diameter < 2.375:
            set_parameter_by_name(pipe_label, 'FP_Product Entry', 'B')
        if 2.375 < diameter < 3.25:
            set_parameter_by_name(pipe_label, 'FP_Product Entry', 'C')
        if 3.25 < diameter < 4.50:
            set_parameter_by_name(pipe_label, 'FP_Product Entry', 'D')
        if 4.50 < diameter < 5.875:
            set_parameter_by_name(pipe_label, 'FP_Product Entry', 'E')
        if 5.875 < diameter < 7.875:
            set_parameter_by_name(pipe_label, 'FP_Product Entry', 'F')
        if 7.875 < diameter < 9.875:
            set_parameter_by_name(pipe_label, 'FP_Product Entry', 'G')
        if diameter > 9.875:
            set_parameter_by_name(pipe_label, 'FP_Product Entry', 'H')
    except:
        pass
t.Commit()
