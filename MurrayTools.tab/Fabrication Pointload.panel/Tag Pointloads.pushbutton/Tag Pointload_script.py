import Autodesk
from Autodesk.Revit.UI import TaskDialog
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
    TagOrientation,
    ViewType,
    View3D
)

# Get the current script's directory path and define the family file name
path, filename = os.path.split(__file__)
NewFilename = r'\Fabrication Hanger - Pointload Tag.rfa'

# Set up document references
DB = Autodesk.Revit.DB
doc = __revit__.ActiveUIDocument.Document  # Current Revit document
uidoc = __revit__.ActiveUIDocument  # Current UI document
curview = doc.ActiveView  # Current active view
app = doc.Application
RevitVersion = app.VersionNumber
RevitINT = float(RevitVersion)

file_path = doc.PathName  # Full path of the current document
file_name = System.IO.Path.GetFileNameWithoutExtension(file_path)  # Document name without extension

# Check if the current view is a 3D view and verify if it is locked
if curview.ViewType == ViewType.ThreeD:
    view3d = curview  # Cast to View3D is implicit as curview is already the active view
    if not view3d.IsLocked:
        TaskDialog.Show(
            "Error",
            "This script cannot be run in an unlocked 3D view. Please lock the 3D view or switch to a plan, section, or elevation view and try again."
        )
        raise SystemExit("Script terminated: Unlocked 3D view detected.")
    # If the 3D view is locked, proceed with the script
else:
    # Non-3D views (e.g., plan, section, elevation) are valid, so proceed
    pass

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
family_pathCC = os.path.join(path, NewFilename.strip('\\'))  # Full path to family file, normalized

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
    except Exception as e:
        # If any error occurs, return an empty set and log the error
        TaskDialog.Show("Warning", "Error retrieving existing tags: {}. Continuing with empty tag set.".format(str(e)))
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

# Function to validate family file existence
def validate_family_file(file_path):
    """Validate if the family file exists and is accessible.
    Args:
        file_path: Path to the family file
    Returns:
        bool: True if file exists and is accessible, False otherwise"""
    try:
        return os.path.isfile(file_path)
    except:
        return False

# Main script logic wrapped in try-except to catch any unforeseen errors
try:
    # Automatic mode: Tag all hangers in current view
    tg = TransactionGroup(doc, "Add Pointload Tags")
    tg.Start()

    # Load family if not present
    t = Transaction(doc, 'Load PointLoad Family')
    t.Start()
    if not Fam_is_in_project:
        # Validate family file before attempting to load
        if not validate_family_file(family_pathCC):
            TaskDialog.Show(
                "Error",
                "Family file not found at path: {}. Please ensure the file exists and try again.".format(family_pathCC)
            )
            t.RollBack()  # Roll back the transaction to prevent freezing
            tg.RollBack()  # Roll back the transaction group
            raise SystemExit("Family file not found. Script terminated to prevent issues.")
        
        try:
            fload_handler = FamilyLoaderOptionsHandler()
            family = doc.LoadFamily(family_pathCC, fload_handler)
            if family is None:
                TaskDialog.Show(
                    "Error",
                    "Failed to load family from path: {}. Please check the family file and try again.".format(family_pathCC)
                )
                t.RollBack()  # Roll back the transaction
                tg.RollBack()  # Roll back the transaction group
                raise SystemExit("Family loading failed. Script terminated to prevent issues.")
        except Exception as e:
            TaskDialog.Show(
                "Error",
                "Error loading family: {}. Please check the family file and try again.".format(str(e))
            )
            t.RollBack()  # Roll back the transaction to prevent freezing
            tg.RollBack()  # Roll back the transaction group
            raise SystemExit("Family loading error: {}. Script terminated to prevent issues.".format(str(e)))

    t.Commit()  # Commit the transaction if loading was successful

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
        try:
            R = Reference(e)
            STName = e.GetRodInfo().RodCount
            ItmList1.append(STName)
            existing_tags = get_existing_tags(e)
            
            if needs_tagging(e, STName, existing_tags):
                STName1 = e.GetRodInfo()
                remaining_tags = STName - len(existing_tags)
                for n in range(remaining_tags):
                    rodloc = STName1.GetRodEndPosition(n)
                    try:
                        IndependentTag.Create(doc, curview.Id, R, False, TagMode.TM_ADDBY_CATEGORY, 
                                             TagOrientation.Horizontal, rodloc)
                    except Exception as tag_error:
                        TaskDialog.Show(
                            "Error",
                            "Failed to create tag for element ID {}: {}. Skipping this tag.".format(e.Id, str(tag_error))
                        )
                        continue  # Skip to the next tag to avoid stopping the loop
        except Exception as e:
            TaskDialog.Show(
                "Error",
                "Error processing element ID {}: {}. Skipping this element.".format(e.Id, str(e))
            )
            continue  # Skip to the next element to avoid stopping the loop

    t.Commit()
    tg.Assimilate()

except Exception as e:
    # Catch any unforeseen errors in the main logic
    TaskDialog.Show(
        "Error",
        "An unexpected error occurred: {}. Please contact support or check the script.".format(str(e))
    )
    # Ensure any open transactions are rolled back
    if t.HasStarted() and not t.HasEnded():
        t.RollBack()
    if tg.HasStarted() and not tg.HasEnded():
        tg.RollBack()
    raise SystemExit("Script terminated due to unexpected error: {}".format(str(e)))