import Autodesk
from Autodesk.Revit import DB
from Autodesk.Revit.DB import Transaction, TransactionGroup, ParameterFilterRuleFactory, ElementParameterFilter, ParameterFilterElement, Color, ElementId, BuiltInCategory, OverrideGraphicSettings, FilteredElementCollector, FillPatternElement, BuiltInParameter
from Autodesk.Revit.UI.Selection import *
from Parameters.Add_SharedParameters import Shared_Params
from System.Collections.Generic import List
import os

Shared_Params()

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView
app = doc.Application

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

def get_existing_rod_controls():
    rod_controls = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFoundation)\
                                               .WhereElementIsNotElementType()\
                                               .ToElements()
    existing_locations = {}
    for rc in rod_controls:
        if rc.Symbol.Family.Name == FamilyName and rc.Symbol.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString() == FamilyType:
            loc = rc.Location
            if isinstance(loc, DB.LocationPoint):
                existing_locations[rc.Id] = loc.Point
    return existing_locations

def is_location_occupied(target_point, existing_locations, tolerance=0.01):
    for loc in existing_locations.values():
        distance = target_point.DistanceTo(loc)
        if distance <= tolerance:
            return True
    return False

# Get selection
pipesel = uidoc.Selection.PickObjects(ObjectType.Element,
    CustomISelectionFilter("MEP Fabrication Hangers"), 
    "Select Fabrication Hangers")            
Fhangers = [doc.GetElement(elId) for elId in pipesel]

view = doc.ActiveView
filter_name = "BEAM HANGERS"
filter_color = DB.Color(0, 255, 255)  # Cyan
FamilyName = 'FP_Rod Control'
FamilyType = 'FP_Rod Control'
folder_name = "c:\\Temp"
filepath = os.path.join(folder_name, 'Ribbon_ExtentHangerRod.txt')

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

# Initialize transaction group
tg = TransactionGroup(doc, "Add Rod Control and Set Beam Hanger")
tg.Start()

# Load family if needed
path, filename = os.path.split(__file__)
NewFilename = '\FP_Rod Control.rfa'
family_pathCC = path + NewFilename
families = FilteredElementCollector(doc).OfClass(DB.Family)
Fam_is_in_project = any(f.Name == FamilyName for f in families)

class FamilyLoaderOptionsHandler(DB.IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues.Value = False
        return True
    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        source.Value = DB.FamilySource.Family
        overwriteParameterValues.Value = False
        return True

t = Transaction(doc, 'Load Rod Control')
t.Start()
if not Fam_is_in_project:
    fload_handler = FamilyLoaderOptionsHandler()
    doc.LoadFamily(family_pathCC, fload_handler)
t.Commit()

# Set parameter and place rod controls
t = Transaction(doc, 'Set Beam Hanger and Populate Rod Control')
t.Start()

# Set parameter value to Yes
for hanger in Fhangers:
    set_parameter_by_name(hanger, 'FP_Beam Hanger', 'Yes')

# Get existing rod control locations
existing_rod_locations = get_existing_rod_controls()

# Place rod controls
familyTypes = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFoundation)\
                                           .OfClass(DB.FamilySymbol)\
                                           .ToElements()
ItmList1 = []
ItmList2 = []

for famtype in familyTypes:
    typeName = famtype.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
    if famtype.Family.Name == FamilyName and typeName == FamilyType:
        famtype.Activate()
        doc.Regenerate()
        for e in Fhangers:
            rod_info = e.GetRodInfo()
            rod_count = rod_info.RodCount
            ItmList1.append(rod_count)
            for n in range(rod_count):
                rodloc = rod_info.GetRodEndPosition(n)
                rodloc_top = DB.XYZ(rodloc.X, rodloc.Y, rodloc.Z + 0.01)  # Slight offset to place at top
                ItmList2.append(rodloc_top)
        for hangerlocation in ItmList2:
            if not is_location_occupied(hangerlocation, existing_rod_locations):
                doc.Create.NewFamilyInstance(hangerlocation, famtype, DB.Structure.StructuralType.NonStructural)

t.Commit()

# Attach rods to structure
t = Transaction(doc, 'Attach Rod to Structure')
t.Start()
for e in Fhangers:
    rod_info = e.GetRodInfo()
    rod_count = rod_info.RodCount
    for n in range(rod_count):
        rod_info.AttachToStructure()
t.Commit()

# Apply view filter
t = Transaction(doc, 'Apply View Filter')
t.Start()

# Get view to modify
view_template_id = view.ViewTemplateId
if not view_template_id.Equals(ElementId.InvalidElementId):
    view_to_modify = doc.GetElement(view_template_id)
else:
    view_to_modify = view

# Check existing filters
existing_filters = FilteredElementCollector(doc).OfClass(ParameterFilterElement).ToElements()
existing_filter_names = {filter.Name for filter in existing_filters}
existing_filter_dict = {filter.Name: filter.Id for filter in existing_filters}
applied_filters = {doc.GetElement(id).Name: id for id in view_to_modify.GetFilters()}

# Create and apply filter if it doesn't exist
if filter_name not in existing_filter_names:
    param_id = None
    if Fhangers:
        for p in Fhangers[0].Parameters:
            if p.Definition.Name == "FP_Beam Hanger":
                param_id = p.Id
                break
    
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
tg.Assimilate()