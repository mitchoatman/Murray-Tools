import Autodesk
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
import os
import System
from Autodesk.Revit.DB import (
    IFamilyLoadOptions,
    FamilySource,
    Transaction,
    FilteredElementCollector,
    Family,
    TransactionGroup,
    BuiltInCategory,
    FamilySymbol,
    BuiltInParameter,
    Reference,
    IndependentTag,
    TagMode,
    TagOrientation
)

# Get the current script's directory path and define the family file name
path, filename = os.path.split(__file__)
NewFilename = '\Fabrication Hanger - Pointload Tag.rfa'

# Set up document references
DB = Autodesk.Revit.DB
doc = __revit__.ActiveUIDocument.Document  # Current Revit document
uidoc = __revit__.ActiveUIDocument  # Current UI document
curview = doc.ActiveView  # Current active view
app = doc.Application
RevitVersion = app.VersionNumber
RevitINT = float (RevitVersion)

file_path = doc.PathName  # Full path of the current document
file_name = System.IO.Path.GetFileNameWithoutExtension(file_path)  # Document name without extension

# Define family loading options handler to control how families are loaded
class FamilyLoaderOptionsHandler(DB.IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues.Value = False  # Don't overwrite parameter values
        return True

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        source.Value = DB.FamilySource.Family  # Use family source
        overwriteParameterValues.Value = False  # Don't overwrite parameter values
        return True

# Check if the target family is already in the project
families = FilteredElementCollector(doc).OfClass(Family)  # Collect all families in document
FamilyName = 'Fabrication Hanger - Pointload Tag'  # Target family name
FamilyType = 'POINTLOAD'  # Target family type name
Fam_is_in_project = any(f.Name == FamilyName for f in families)  # Boolean if family exists
family_pathCC = path + NewFilename  # Full path to family file

def get_existing_tags(element):
    """Get all existing tag IDs for a given element in the current view.
    Args:
        element: The Revit element to check
    Returns:
        set: Set of ElementIds of tags associated with the element"""
    tags = FilteredElementCollector(doc, curview.Id).OfCategory(BuiltInCategory.OST_FabricationHangerTags)\
                                                    .WhereElementIsNotElementType()\
                                                    .ToElements()
    tagged_elements = set()

    try:
        if RevitINT >= 2022:  # Revit 2022 and later use GetTaggedLocalElementIds
            for tag in tags:
                tagged_ids = tag.GetTaggedLocalElementIds()  # Call the method to get the collection
                if tagged_ids and element.Id in tagged_ids:  # Check if element.Id is in the collection
                    tagged_elements.add(tag.Id)
        else:  # Revit 2021 and earlier use TaggedLocalElementId
            for tag in tags:
                if tag.TaggedLocalElementId == element.Id:
                    tagged_elements.add(tag.Id)
    except:
        # If any error occurs, return an empty set (could log this for debugging)
        return tagged_elements

    return tagged_elements  # Return the set of tag IDs, empty if no tags found

def needs_tagging(element, rod_count, existing_tags):
    """Determine if the element needs additional tags based on rod count.
    Args:
        element: The Revit element to check
        rod_count: Number of rods in the element
        existing_tags: Set of existing tag IDs
    Returns:
        bool: True if more tags are needed, False otherwise"""
    return len(existing_tags) < rod_count


# Shift-click mode: Manual selection of hangers
if __shiftclick__:
    file_name = doc.Title  # Get document title
    
    # Custom filter for selecting only MEP Fabrication Hangers
    class CustomISelectionFilter(ISelectionFilter):
        def __init__(self, nom_categorie):
            self.nom_categorie = nom_categorie
        def AllowElement(self, e):
            return e.Category.Name == self.nom_categorie
        def AllowReference(self, ref, point):
            return True

    # Prompt user to select hangers
    fhangers = uidoc.Selection.PickObjects(ObjectType.Element,
        CustomISelectionFilter("MEP Fabrication Hangers"), "Select Fabrication Hangers")
    Hanger_collector = [doc.GetElement(elId) for elId in fhangers]  # Convert selections to elements

    # Collect all family symbols for hanger tags
    familyTypes = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_FabricationHangerTags)\
                                               .OfClass(FamilySymbol)\
                                               .ToElements()

    # Start transaction group for manual tagging
    tg = TransactionGroup(doc, "Selected Pointload Tags")
    tg.Start()

    # Load family if not present
    t = Transaction(doc, 'Load PointLoad Family')
    t.Start()
    if not Fam_is_in_project:
        fload_handler = FamilyLoaderOptionsHandler()
        family = doc.LoadFamily(family_pathCC, fload_handler)  # Load the tag family
    t.Commit()

    ItmList1 = list()  # List to store rod counts

    # Tag creation transaction
    t = Transaction(doc, 'Tag Pointloads')
    t.Start()
    # Activate the correct family symbol
    for famtype in familyTypes:
        typeName = famtype.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
        if famtype.Family.Name == FamilyName and typeName == FamilyType:
            if not famtype.IsActive:
                famtype.Activate()
                doc.Regenerate()

    # Tag each hanger's rods if not already fully tagged
    for e in Hanger_collector:
        R = Reference(e)
        STName = e.GetRodInfo().RodCount  # Number of rods in hanger
        ItmList1.append(STName)
        existing_tags = get_existing_tags(e)
        
        if needs_tagging(e, STName, existing_tags):
            STName1 = e.GetRodInfo()
            remaining_tags = STName - len(existing_tags)
            for n in range(remaining_tags):
                rodloc = STName1.GetRodEndPosition(n)  # Position of each rod
                IndependentTag.Create(doc, curview.Id, R, False, TagMode.TM_ADDBY_CATEGORY, 
                                    TagOrientation.Horizontal, rodloc)  # Create tag
    t.Commit()
    tg.Assimilate()  # Complete transaction group

# Automatic mode: Tag all hangers in current view
else:
    # Start transaction group for automatic tagging
    tg = TransactionGroup(doc, "Add Pointload Tags")
    tg.Start()

    # Load family if not present
    t = Transaction(doc, 'Load PointLoad Family')
    t.Start()
    if not Fam_is_in_project:
        fload_handler = FamilyLoaderOptionsHandler()
        family = doc.LoadFamily(family_pathCC, fload_handler)
    t.Commit()

    # Collect all hangers in current view
    Hanger_collector = FilteredElementCollector(doc, curview.Id).OfCategory(BuiltInCategory.OST_FabricationHangers)\
                                                               .WhereElementIsNotElementType()
    # Collect all family symbols for hanger tags
    familyTypes = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_FabricationHangerTags)\
                                               .OfClass(FamilySymbol)\
                                               .ToElements()

    ItmList1 = list()

    # Tag creation transaction
    t = Transaction(doc, 'Tag Pointloads')
    t.Start()
    # Activate the correct family symbol
    for famtype in familyTypes:
        typeName = famtype.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
        if famtype.Family.Name == FamilyName and typeName == FamilyType:
            if not famtype.IsActive:
                famtype.Activate()
                doc.Regenerate()

    # Tag each hanger's rods if not already fully tagged
    for e in Hanger_collector:
        R = Reference(e)
        STName = e.GetRodInfo().RodCount
        ItmList1.append(STName)
        existing_tags = get_existing_tags(e)
        
        if needs_tagging(e, STName, existing_tags):
            STName1 = e.GetRodInfo()
            remaining_tags = STName - len(existing_tags)
            for n in range(remaining_tags):
                rodloc = STName1.GetRodEndPosition(n)
                IndependentTag.Create(doc, curview.Id, R, False, TagMode.TM_ADDBY_CATEGORY, 
                                    TagOrientation.Horizontal, rodloc)
    t.Commit()
    tg.Assimilate()