from Autodesk.Revit import DB
from Autodesk.Revit.DB import Transaction, ParameterFilterRuleFactory, ElementParameterFilter, ParameterFilterElement, Color, ElementId, BuiltInCategory, OverrideGraphicSettings, FilteredElementCollector, FillPatternElement
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from Parameters.Add_SharedParameters import Shared_Params
from System.Collections.Generic import List

Shared_Params()

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# Define functions
def set_parameter_by_name(element, parameterName, value):
    element.LookupParameter(parameterName).Set(value)

def get_parameter_value_by_name(element, parameterName):
    return element.LookupParameter(parameterName).AsString()

class CustomISelectionFilter(ISelectionFilter):
    def __init__(self, nom_categorie):
        self.nom_categorie = nom_categorie
    def AllowElement(self, e):
        if e.Category.Name == self.nom_categorie:
            return True
        else:
            return False
    def AllowReference(self, ref, point):
        return True

# Get selection
pipesel = uidoc.Selection.PickObjects(ObjectType.Element,
    CustomISelectionFilter("MEP Fabrication Hangers"), 
    "Select Fabrication Hangers")            
Fhangers = [doc.GetElement(elId) for elId in pipesel]

view = doc.ActiveView
filter_name = "BEAM HANGERS"
filter_color = DB.Color(0, 255, 255)  # Cyan

# Get solid fill pattern
fill_patterns = FilteredElementCollector(doc).OfClass(FillPatternElement).ToElements()
solid_fill_id = None
for pattern in fill_patterns:
    if pattern.Name == "<Solid fill>" and pattern.GetFillPattern().IsSolidFill:
        solid_fill_id = pattern.Id
        break

if solid_fill_id is None:
    print("Solid fill pattern not found in the document")

# Create category list for filter
categories = List[ElementId]()
categories.Add(ElementId(BuiltInCategory.OST_FabricationHangers))

# Start transaction
t = Transaction(doc, 'Toggle Beam Hanger and Apply Filter')
t.Start()

# Toggle parameter value
for hanger in Fhangers:
    BHangerStatus = get_parameter_value_by_name(hanger, 'FP_Beam Hanger')
    if BHangerStatus == None or BHangerStatus == 'No':
        set_parameter_by_name(hanger, 'FP_Beam Hanger', 'Yes')
    else:
        set_parameter_by_name(hanger, 'FP_Beam Hanger', 'No')

# Get view to modify (view or view template)
view_template_id = view.ViewTemplateId
if not view_template_id.Equals(ElementId.InvalidElementId):
    view_to_modify = doc.GetElement(view_template_id)
else:
    view_to_modify = view

# Check existing filters
existing_filters = FilteredElementCollector(doc).OfClass(ParameterFilterElement).ToElements()
existing_filter_names = {filter.Name for filter in existing_filters}
existing_filter_dict = {filter.Name: filter.Id for filter in existing_filters}

# Get applied filters
applied_filters = {doc.GetElement(id).Name: id for id in view_to_modify.GetFilters()}

# Create and apply filter if it doesn't exist
if filter_name not in existing_filter_names:
    # Get parameter ID
    param_id = None
    if Fhangers:  # Use first selected hanger to get parameter
        for p in Fhangers[0].Parameters:
            if p.Definition.Name == "FP_Beam Hanger":
                param_id = p.Id
                break
    
    if param_id:
        # Create filter rule
        rule = ParameterFilterRuleFactory.CreateEqualsRule(param_id, "Yes", False)
        filter_element = ElementParameterFilter(rule)
        filter_elem = ParameterFilterElement.Create(doc, filter_name, categories)
        filter_elem.SetElementFilter(filter_element)
        filter_id = filter_elem.Id
        
        # Apply filter if not already applied
        if filter_name not in applied_filters and solid_fill_id:
            overrides = OverrideGraphicSettings()
            overrides.SetSurfaceForegroundPatternColor(filter_color)
            overrides.SetSurfaceForegroundPatternId(solid_fill_id)
            view_to_modify.AddFilter(filter_id)
            view_to_modify.SetFilterVisibility(filter_id, True)
            view_to_modify.SetFilterOverrides(filter_id, overrides)
    else:
        print("Could not find FP_Beam Hanger parameter")
else:
    # Filter exists, apply it if not already applied
    filter_id = existing_filter_dict[filter_name]
    if filter_name not in applied_filters and solid_fill_id:
        overrides = OverrideGraphicSettings()
        overrides.SetSurfaceForegroundPatternColor(filter_color)
        overrides.SetSurfaceForegroundPatternId(solid_fill_id)
        view_to_modify.AddFilter(filter_id)
        view_to_modify.SetFilterVisibility(filter_id, True)
        view_to_modify.SetFilterOverrides(filter_id, overrides)

# End transaction
t.Commit()