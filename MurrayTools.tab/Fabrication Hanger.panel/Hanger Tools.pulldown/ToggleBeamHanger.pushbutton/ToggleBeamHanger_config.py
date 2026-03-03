import Autodesk
from Autodesk.Revit import DB
from Autodesk.Revit.DB import Transaction, ParameterFilterRuleFactory, ElementParameterFilter, ParameterFilterElement, Color, ElementId, BuiltInCategory, OverrideGraphicSettings, FilteredElementCollector, FillPatternElement
from Parameters.Add_SharedParameters import Shared_Params
from System.Collections.Generic import List

Shared_Params()
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
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

# Create category list for filter
categories = List[ElementId]()
categories.Add(ElementId(BuiltInCategory.OST_FabricationHangers))

# Transaction to apply view filter
t = Transaction(doc, 'Apply View Filter')
t.Start()

# Get view to modify (template if applied, otherwise current view)
view_template_id = view.ViewTemplateId
if not view_template_id.Equals(ElementId.InvalidElementId):
    view_to_modify = doc.GetElement(view_template_id)
else:
    view_to_modify = view

# Check existing filters
existing_filters = FilteredElementCollector(doc).OfClass(ParameterFilterElement).ToElements()
existing_filter_names = {f.Name for f in existing_filters}
existing_filter_dict = {f.Name: f.Id for f in existing_filters}
applied_filters = {doc.GetElement(fid).Name: fid for fid in view_to_modify.GetFilters()}

# Find parameter ID for "FP_Beam Hanger" from any existing fabrication hanger
param_id = None
hanger_collector = FilteredElementCollector(doc)\
    .OfCategory(BuiltInCategory.OST_FabricationHangers)\
    .WhereElementIsNotElementType()
hangers = hanger_collector.ToElements()
if hangers:
    for hanger in hangers:
        param = hanger.LookupParameter("FP_Beam Hanger")
        if param:
            param_id = param.Id
            break

# Create filter if it does not exist
if filter_name not in existing_filter_names:
    if param_id and solid_fill_id:
        rule = ParameterFilterRuleFactory.CreateEqualsRule(param_id, "Yes", False)
        filter_element = ElementParameterFilter(rule)
        filter_elem = ParameterFilterElement.Create(doc, filter_name, categories)
        filter_elem.SetElementFilter(filter_element)
        filter_id = filter_elem.Id
        
        overrides = OverrideGraphicSettings()
        overrides.SetSurfaceForegroundPatternColor(filter_color)
        overrides.SetSurfaceForegroundPatternId(solid_fill_id)
        
        view_to_modify.AddFilter(filter_id)
        view_to_modify.SetFilterVisibility(filter_id, True)
        view_to_modify.SetFilterOverrides(filter_id, overrides)

# If filter already exists, ensure it is applied with correct overrides
else:
    filter_id = existing_filter_dict.get(filter_name)
    if filter_id and solid_fill_id:
        overrides = OverrideGraphicSettings()
        overrides.SetSurfaceForegroundPatternColor(filter_color)
        overrides.SetSurfaceForegroundPatternId(solid_fill_id)
        
        if filter_name not in applied_filters:
            view_to_modify.AddFilter(filter_id)
            view_to_modify.SetFilterVisibility(filter_id, True)
        
        view_to_modify.SetFilterOverrides(filter_id, overrides)

t.Commit()