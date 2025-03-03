import os
from collections import defaultdict
from Autodesk.Revit import DB
from Autodesk.Revit.DB import (
    Transaction, 
    TransactionGroup,
    FilteredElementCollector, 
    BuiltInCategory, 
    FabricationPart, 
    IndependentTag, 
    TagOrientation, 
    Family, 
    FamilySymbol
)
from pyrevit import revit, forms
from SharedParam.Add_Parameters import Shared_Params
from Parameters.Get_Set_Params import get_parameter_value_by_name_AsString

# Initialize Revit application objects
app = __revit__.Application
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView

# Get Revit version
RevitVersion = app.VersionNumber
RevitINT = float(RevitVersion)
Shared_Params()

# File path setup
path, filename = os.path.split(__file__)
NewFilename1 = r'\Fabrication Pipe - Stratus Assembly.rfa'

class FamilyLoaderOptionsHandler(DB.IFamilyLoadOptions):
    """Handler for family loading options."""
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues.Value = False
        return True

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        source.Value = DB.FamilySource.Family
        overwriteParameterValues.Value = False
        return True

def load_family_if_missing():
    """Load the specified family if it's not already in the project."""
    families = FilteredElementCollector(doc).OfClass(Family)
    FamilyName = 'Fabrication Pipe - Stratus Assembly'
    Fam_is_in_project = any(f.Name == FamilyName for f in families)
    
    if not Fam_is_in_project:
        family_path = path + NewFilename1
        with Transaction(doc, 'Load Stratus Family Tag') as t:
            t.Start()
            fload_handler = FamilyLoaderOptionsHandler()
            doc.LoadFamily(family_path, fload_handler)
            t.Commit()

def get_tag_type():
    """Retrieve the specified tag type or alert if not found."""
    tag_family_name = "Fabrication Pipe - Stratus Assembly"
    all_tags = FilteredElementCollector(doc).OfClass(DB.FamilySymbol).ToElements()
    available_tags = {tag.Family.Name for tag in all_tags}
    tag_type = next((tag for tag in all_tags if tag.Family.Name == tag_family_name), None)
    
    if not tag_type:
        available_tags_sorted = sorted(available_tags)
        error_message = "Tag family '{}' not found.\nAvailable tag families:\n{}".format(
            tag_family_name, "\n".join(available_tags_sorted))
        forms.alert(error_message, exitscript=True)
    return tag_type

def get_safe_stratus_value(element):
    """Safely retrieve STRATUS Assembly parameter value."""
    try:
        value = get_parameter_value_by_name_AsString(element, 'STRATUS Assembly')
        return value if value else "Unknown"
    except:
        return "Unknown"

def get_existing_tagged_spools():
    """Collect 'STRATUS Assembly' values of elements already tagged in the current view."""
    existing_tags = FilteredElementCollector(doc, curview.Id).OfClass(IndependentTag).ToElements()
    tagged_spools = set()
    
    try:
        if RevitINT > 2021:
            for tag in existing_tags:
                tagged_element_ids = tag.GetTaggedLocalElementIds()  # For Revit 2022+
                for element_id in tagged_element_ids:
                    tagged_elem = doc.GetElement(element_id)
                    if tagged_elem and isinstance(tagged_elem, FabricationPart):
                        spool_value = get_safe_stratus_value(tagged_elem)
                        if spool_value != "Unknown":
                            tagged_spools.add(spool_value)
        else:
            for tag in existing_tags:
                element_id = tag.TaggedLocalElementId  # For Revit 2021 and below
                tagged_elem = doc.GetElement(element_id)
                if tagged_elem and isinstance(tagged_elem, FabricationPart):
                    spool_value = get_safe_stratus_value(tagged_elem)
                    if spool_value != "Unknown":
                        tagged_spools.add(spool_value)
    except Exception as e:
        forms.alert("Error retrieving tagged spools: {}".format(str(e)))
    
    return tagged_spools

def place_tag(element, part_collector):
    """Place a tag on the element using its location or a fallback point."""
    try:
        location = element.Location
        tag_point = None

        if isinstance(location, DB.LocationPoint):
            tag_point = location.Point
        elif isinstance(location, DB.LocationCurve) and location.Curve:
            tag_point = location.Curve.Evaluate(0.5, True)

        if not tag_point:
            for next_elem in part_collector:
                if next_elem != element:
                    next_location = next_elem.Location
                    if isinstance(next_location, DB.LocationCurve) and next_location.Curve:
                        tag_point = next_location.Curve.Evaluate(0.9, True)
                        if tag_point:
                            break
        if tag_point:
            with Transaction(doc, "Tag Fabrication Parts") as t:
                t.Start()
                tag = IndependentTag.Create(
                    doc,
                    tag_type.Id,
                    curview.Id,
                    DB.Reference(element),
                    True,
                    TagOrientation.Horizontal,
                    tag_point
                )
                stratus_assembly = get_safe_stratus_value(element)
                t.Commit()
                return "Tagged: {}".format(stratus_assembly)
        else:
            return "Skipping element {}: No valid tag placement point".format(element.Id)

    except Exception as e:
        return "Error tagging element {}:\n{}".format(element.Id, e)

# Main execution
load_family_if_missing()
tag_type = get_tag_type()

# Collect fabrication parts
part_collector = FilteredElementCollector(doc, curview.Id)\
    .OfClass(FabricationPart)\
    .WhereElementIsNotElementType()\
    .ToElements()

if not part_collector:
    forms.alert("No Fabrication parts found in the active view.", exitscript=True)

# Get unique STRATUS Assembly values, excluding those with "-MAP"
STRATUSAssemlist = [get_safe_stratus_value(x) for x in part_collector 
                   if x.LookupParameter("Fabrication Service") and "-MAP" not in get_safe_stratus_value(x)]
STRATUSAssemlist_set = sorted(set(STRATUSAssemlist))

if not STRATUSAssemlist_set:
    forms.alert("No STRATUS Assembly values (excluding '-MAP') found in the active view.", exitscript=True)

# Use all assemblies (spool selection is commented out)
selected_assemblies = STRATUSAssemlist_set

# Get existing tagged spools
existing_tagged_spools = get_existing_tagged_spools()

# Tag elements
elementlist = []
tagged_assemblies = set()

for elem in part_collector:
    assembly_value = get_safe_stratus_value(elem)
    # Skip if the spool contains "-MAP", is already tagged, or not in selected assemblies
    if (assembly_value in selected_assemblies and 
        "-MAP" not in assembly_value and
        assembly_value not in tagged_assemblies and 
        assembly_value not in existing_tagged_spools):
        elementlist.append(elem)
        tagged_assemblies.add(assembly_value)

if not elementlist:
    forms.alert("No untagged elements (excluding '-MAP') found for tagging.", exitscript=True)

# Process tagging
tag_results = []
with TransactionGroup(doc, "Tag Spools") as tg:
    tg.Start()
    for elem in elementlist:
        result = place_tag(elem, part_collector)
        tag_results.append(result)
    tg.Assimilate()