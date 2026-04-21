from Autodesk.Revit import DB
from Autodesk.Revit.DB import FilteredElementCollector, Transaction, BuiltInCategory, FamilySymbol, Family
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from Autodesk.Revit.UI import TaskDialog
import os
import math
import clr
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
from System.Drawing import Point

path, filename = os.path.split(__file__)
NewFilename = '\Fabrication Pipe - Size Tag.rfa'
ElevationTagFilename = '\Fabrication Pipe - Elevation Tag.rfa'

app = __revit__.Application
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

class FamilyLoaderOptionsHandler(DB.IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues.Value = False
        return True
    def OnSharedFamilyFound(self, familyInUse, source, overwriteParameterValues):
        source.Value = DB.FamilySource.Family
        overwriteParameterValues.Value = False
        return True

SizeFamilyName = 'Fabrication Pipe - Size Tag'
SizeFamilyType = 'Size Tag'
ElevationFamilyName = 'Fabrication Pipe - Elevation Tag'
ElevationFamilyType = 'BOP'

family_pathCC2 = path + NewFilename
family_pathCC1 = path + ElevationTagFilename

# Validate file paths
if not os.path.exists(family_pathCC2):
    TaskDialog.Show("Error", "Size tag family file not found: " + family_pathCC2)
    raise Exception("Size tag family file not found: " + family_pathCC2)
if not os.path.exists(family_pathCC1):
    TaskDialog.Show("Error", "Elevation tag family file not found: " + family_pathCC1)
    raise Exception("Elevation tag family file not found: " + family_pathCC1)

# Check if families need loading
families = FilteredElementCollector(doc).OfClass(Family)
needs_size_loading = not any(f.Name == SizeFamilyName for f in families)
needs_elevation_loading = not any(f.Name == ElevationFamilyName for f in families)

if needs_size_loading:
    with Transaction(doc, 'Load Pipe Size Family') as t:
        t.Start()
        fload_handler = FamilyLoaderOptionsHandler()
        doc.LoadFamily(family_pathCC2, fload_handler)
        t.Commit()

if needs_elevation_loading:
    with Transaction(doc, 'Load Elevation Tag Family') as t:
        t.Start()
        fload_handler = FamilyLoaderOptionsHandler()
        doc.LoadFamily(family_pathCC1, fload_handler)
        t.Commit()

# Collect family symbols
collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_FabricationPipeworkTags).OfClass(FamilySymbol)

size_symbol = None
elevation_symbol = None
for symbol in collector:
    if symbol.Family.Name == SizeFamilyName and symbol.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString() == SizeFamilyType:
        size_symbol = symbol
    if symbol.Family.Name == ElevationFamilyName and symbol.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString() == ElevationFamilyType:
        elevation_symbol = symbol

# Validate symbols
if not size_symbol or not size_symbol.IsValidObject:
    TaskDialog.Show("Error", "Required size tag family type not found or invalid")
    raise Exception("Required size tag family type not found or invalid")
if not elevation_symbol or not elevation_symbol.IsValidObject:
    TaskDialog.Show("Error", "Required elevation tag family type not found or invalid")
    raise Exception("Required elevation tag family type not found or invalid")

# Activate symbols if needed
if not size_symbol.IsActive:
    with Transaction(doc, 'Activate Size Family Symbol') as t:
        t.Start()
        size_symbol.Activate()
        t.Commit()

if not elevation_symbol.IsActive:
    with Transaction(doc, 'Activate Elevation Family Symbol') as t:
        t.Start()
        elevation_symbol.Activate()
        t.Commit()

try:
    class FabricationPipeFilter(ISelectionFilter):
        def AllowElement(self, element):
            if element.Category.Id.IntegerValue == int(BuiltInCategory.OST_FabricationPipework):
                param = element.ItemCustomId
                return param == 2041
            return False
        def AllowReference(self, reference, point):
            return False

    selected_refs = uidoc.Selection.PickObjects(ObjectType.Element, FabricationPipeFilter(), "Select fabrication pipes CID 2041")
    selected_pipes = [doc.GetElement(ref.ElementId) for ref in selected_refs]
    
    if not selected_pipes:
        TaskDialog.Show("Error", "No fabrication pipes selected")
        raise Exception("No fabrication pipes selected")

    is_x_direction = False
    if selected_pipes:
        connectors = selected_pipes[0].ConnectorManager.Connectors
        if connectors.Size >= 2:
            conn_list = list(connectors)
            direction = (conn_list[1].Origin - conn_list[0].Origin).Normalize()
            is_x_direction = abs(direction.X) > abs(direction.Y)

    view_scale = doc.ActiveView.Scale
    spacing_inches = 6.0 * (view_scale / 48.0)
    spacing = DB.UnitUtils.ConvertToInternalUnits(spacing_inches, DB.UnitTypeId.Inches)
    base_point = uidoc.Selection.PickPoint("Select base point for stacked tags")
    
    # Sort pipes
    if is_x_direction:
        sorted_pipes = sorted(selected_pipes, key=lambda p: list(p.ConnectorManager.Connectors)[0].Origin.Y if p.ConnectorManager.Connectors.Size > 0 else 0, reverse=True)
    else:
        sorted_pipes = sorted(selected_pipes, key=lambda p: list(p.ConnectorManager.Connectors)[0].Origin.X if p.ConnectorManager.Connectors.Size > 0 else 0)
    
    x_offset = 0.0
    y_offset = -spacing

    # Batch tag creation
    batch_size = 10
    for i in range(0, len(sorted_pipes), batch_size):
        with Transaction(doc, 'Place Stacked Tags Batch') as t:
            t.Start()
            for j, pipe in enumerate(sorted_pipes[i:i+batch_size]):
                try:
                    if pipe.ConnectorManager.Connectors.Size == 0:
                        TaskDialog.Show("Warning", "Skipping pipe with no connectors: " + str(pipe.Id))
                        continue
                    tag_position = DB.XYZ(base_point.X + ((i+j) * x_offset), base_point.Y + ((i+j) * y_offset), base_point.Z)
                    DB.IndependentTag.Create(doc, size_symbol.Id, doc.ActiveView.Id, DB.Reference(pipe), False, DB.TagOrientation.Horizontal, tag_position)
                except Exception, e:
                    TaskDialog.Show("Error", "Failed to create tag for pipe " + str(pipe.Id) + ": " + str(e))
                    continue
            t.Commit()

    # Place elevation tag
    with Transaction(doc, 'Place Elevation Tag') as t:
        t.Start()
        num_pipes = len(sorted_pipes)
        elevation_position = DB.XYZ(
            base_point.X + (num_pipes * x_offset),
            base_point.Y + (num_pipes * y_offset),
            base_point.Z
        )
        try:
            elevation_tag = DB.IndependentTag.Create(
                doc, 
                elevation_symbol.Id, 
                doc.ActiveView.Id, 
                DB.Reference(sorted_pipes[-1]), 
                False, 
                DB.TagOrientation.Horizontal, 
                elevation_position
            )
            elevation_tag.HasLeader = True
        except Exception, e:
            TaskDialog.Show("Error", "Failed to create elevation tag: " + str(e))
        t.Commit()

except Exception, e:
    TaskDialog.Show("Error", "Operation cancelled or failed: " + str(e))