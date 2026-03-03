from Autodesk.Revit import DB
from Autodesk.Revit.DB import Transaction, ParameterFilterRuleFactory, ElementParameterFilter, ParameterFilterElement, Color, ElementId, BuiltInCategory, OverrideGraphicSettings, FilteredElementCollector, FillPatternElement
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from Parameters.Add_SharedParameters import Shared_Params
from System.Collections.Generic import List

Shared_Params()
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

def set_parameter_by_name(element, parameterName, value):
    element.LookupParameter(parameterName).Set(value)

def get_parameter_value_by_name(element, parameterName):
    return element.LookupParameter(parameterName).AsString()

class CustomISelectionFilter(ISelectionFilter):
    def __init__(self, categories):
        self.categories = categories
    def AllowElement(self, e):
        return e.Category.Name in self.categories
    def AllowReference(self, ref, point):
        return True

pipesel = uidoc.Selection.PickObjects(ObjectType.Element,
    CustomISelectionFilter(["MEP Fabrication Hangers", "Structural Stiffeners", "Pipe Accessories"]),
    "Select Fabrication Hangers or Structural Stiffeners or Pipe Accessories")

Fhangers = [doc.GetElement(elId) for elId in pipesel]
view = doc.ActiveView
filter_name = "BEAM HANGERS"
filter_color = DB.Color(0, 255, 255)

fill_patterns = FilteredElementCollector(doc).OfClass(FillPatternElement).ToElements()
solid_fill_id = None
for pattern in fill_patterns:
    if pattern.Name == "<Solid fill>" and pattern.GetFillPattern().IsSolidFill:
        solid_fill_id = pattern.Id
        break

if solid_fill_id is None:
    print("Solid fill pattern not found in the document")

categories = List[ElementId]()
categories.Add(ElementId(BuiltInCategory.OST_FabricationHangers))
categories.Add(ElementId(BuiltInCategory.OST_StructuralStiffener))
categories.Add(ElementId(BuiltInCategory.OST_PipeAccessory))

t = Transaction(doc, 'Toggle Beam Hanger and Apply Filter')
t.Start()

for hanger in Fhangers:
    BHangerStatus = get_parameter_value_by_name(hanger, 'FP_Beam Hanger')
    if hanger.LookupParameter('FP_Beam Hanger'):
        if BHangerStatus is None or BHangerStatus == 'No':
            set_parameter_by_name(hanger, 'FP_Beam Hanger', 'Yes')
        else:
            set_parameter_by_name(hanger, 'FP_Beam Hanger', 'No')

view_template_id = view.ViewTemplateId
if not view_template_id.Equals(ElementId.InvalidElementId):
    view_to_modify = doc.GetElement(view_template_id)
else:
    view_to_modify = view

existing_filters = FilteredElementCollector(doc).OfClass(ParameterFilterElement).ToElements()
existing_filter_names = {filter.Name for filter in existing_filters}
existing_filter_dict = {filter.Name: filter.Id for filter in existing_filters}
applied_filters = {doc.GetElement(id).Name: id for id in view_to_modify.GetFilters()}


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
if filter_name not in existing_filter_names:
    if param_id:
        rule = ParameterFilterRuleFactory.CreateEqualsRule(param_id, "Yes", False)
        filter_element = ElementParameterFilter(rule)
        filter_elem = ParameterFilterElement.Create(doc, filter_name, categories)
        filter_elem.SetElementFilter(filter_element)
        filter_id = filter_elem.Id

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
    filter_id = existing_filter_dict[filter_name]
    if filter_name not in applied_filters and solid_fill_id:
        overrides = OverrideGraphicSettings()
        overrides.SetSurfaceForegroundPatternColor(filter_color)
        overrides.SetSurfaceForegroundPatternId(solid_fill_id)
        view_to_modify.AddFilter(filter_id)
        view_to_modify.SetFilterVisibility(filter_id, True)
        view_to_modify.SetFilterOverrides(filter_id, overrides)

t.Commit()