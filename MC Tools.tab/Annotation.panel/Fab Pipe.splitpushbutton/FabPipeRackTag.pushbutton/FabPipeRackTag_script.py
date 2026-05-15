from Autodesk.Revit import DB
from Autodesk.Revit.DB import FilteredElementCollector, Transaction, BuiltInCategory, FamilySymbol, Family
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from Autodesk.Revit.UI import TaskDialog
import os
import clr
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

path, filename = os.path.split(__file__)
size_family_path = os.path.join(path, 'Fabrication Pipe - Size Tag.rfa')
elev_family_path = os.path.join(path, 'Fabrication Pipe - Elevation Tag.rfa')

app = __revit__.Application
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

class FamilyLoaderOptionsHandler(DB.IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues.Value = False
        return True

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        source.Value = DB.FamilySource.Family
        overwriteParameterValues.Value = False
        return True

SizeFamilyName = 'Fabrication Pipe - Size Tag'
SizeFamilyType = 'Size Tag'
ElevationFamilyName = 'Fabrication Pipe - Elevation Tag'
ElevationFamilyType = 'BOP'

def get_pipe_sort_point(pipe):
    try:
        connectors = pipe.ConnectorManager.Connectors
        if connectors.Size >= 2:
            conn_list = list(connectors)
            return (conn_list[0].Origin + conn_list[1].Origin) / 2.0
        elif connectors.Size == 1:
            return list(connectors)[0].Origin
    except:
        pass

    try:
        bbox = pipe.get_BoundingBox(doc.ActiveView)
        if bbox:
            return (bbox.Min + bbox.Max) / 2.0
    except:
        pass

    return DB.XYZ(0, 0, 0)

# Validate file paths
if not os.path.exists(size_family_path):
    TaskDialog.Show("Error", "Size tag family file not found: " + size_family_path)
    raise Exception("Size tag family file not found: " + size_family_path)

if not os.path.exists(elev_family_path):
    TaskDialog.Show("Error", "Elevation tag family file not found: " + elev_family_path)
    raise Exception("Elevation tag family file not found: " + elev_family_path)

# Check if families need loading
families = FilteredElementCollector(doc).OfClass(Family)
needs_size_loading = not any(f.Name == SizeFamilyName for f in families)
needs_elevation_loading = not any(f.Name == ElevationFamilyName for f in families)

if needs_size_loading:
    with Transaction(doc, 'Load Pipe Size Family') as t:
        t.Start()
        fload_handler = FamilyLoaderOptionsHandler()
        doc.LoadFamily(size_family_path, fload_handler)
        t.Commit()

if needs_elevation_loading:
    with Transaction(doc, 'Load Elevation Tag Family') as t:
        t.Start()
        fload_handler = FamilyLoaderOptionsHandler()
        doc.LoadFamily(elev_family_path, fload_handler)
        t.Commit()

# Collect family symbols
collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_FabricationPipeworkTags).OfClass(FamilySymbol)

size_symbol = None
elevation_symbol = None
for symbol in collector:
    type_param = symbol.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
    type_name = type_param.AsString() if type_param else ""

    if symbol.Family.Name == SizeFamilyName and type_name == SizeFamilyType:
        size_symbol = symbol
    if symbol.Family.Name == ElevationFamilyName and type_name == ElevationFamilyType:
        elevation_symbol = symbol

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
        doc.Regenerate()
        t.Commit()

if not elevation_symbol.IsActive:
    with Transaction(doc, 'Activate Elevation Family Symbol') as t:
        t.Start()
        elevation_symbol.Activate()
        doc.Regenerate()
        t.Commit()

class FabricationPipeFilter(ISelectionFilter):
    def AllowElement(self, element):
        if element.Category and element.Category.Id.IntegerValue == int(BuiltInCategory.OST_FabricationPipework):
            try:
                return element.ItemCustomId == 2041
            except:
                return False
        return False

    def AllowReference(self, reference, point):
        return False

try:
    selected_refs = uidoc.Selection.PickObjects(ObjectType.Element, FabricationPipeFilter(), "Select fabrication pipes CID 2041")
    selected_pipes = [doc.GetElement(ref.ElementId) for ref in selected_refs]

    if not selected_pipes:
        TaskDialog.Show("Error", "No fabrication pipes selected")
        raise Exception("No fabrication pipes selected")

    view_scale = doc.ActiveView.Scale
    spacing_inches = 6.0 * (view_scale / 48.0)
    spacing = DB.UnitUtils.ConvertToInternalUnits(spacing_inches, DB.UnitTypeId.Inches)
    base_point = uidoc.Selection.PickPoint("Select base point for stacked tags")

    is_x_direction = False
    if selected_pipes:
        connectors = selected_pipes[0].ConnectorManager.Connectors
        if connectors.Size >= 2:
            conn_list = list(connectors)
            direction = (conn_list[1].Origin - conn_list[0].Origin).Normalize()
            is_x_direction = abs(direction.X) > abs(direction.Y)

    # Better sorting using pipe midpoint
    if is_x_direction:
        # horizontal run -> top to bottom
        sorted_pipes = sorted(
            selected_pipes,
            key=lambda p: (-get_pipe_sort_point(p).Y, get_pipe_sort_point(p).X)
        )
    else:
        # vertical run -> left to right
        sorted_pipes = sorted(
            selected_pipes,
            key=lambda p: (get_pipe_sort_point(p).X, -get_pipe_sort_point(p).Y)
        )

    failed_pipes = []

    with Transaction(doc, 'Place Stacked Tags Batch') as t:
        t.Start()
        for i, pipe in enumerate(sorted_pipes):
            try:
                if pipe.ConnectorManager.Connectors.Size == 0:
                    failed_pipes.append(str(pipe.Id))
                    continue

                # tags always stack top to bottom
                tag_position = DB.XYZ(base_point.X, base_point.Y - (i * spacing), base_point.Z)

                DB.IndependentTag.Create(
                    doc,
                    size_symbol.Id,
                    doc.ActiveView.Id,
                    DB.Reference(pipe),
                    False,
                    DB.TagOrientation.Horizontal,
                    tag_position
                )
            except Exception, e:
                failed_pipes.append(str(pipe.Id))
                continue
        t.Commit()

    with Transaction(doc, 'Place Elevation Tag') as t:
        t.Start()
        elevation_position = DB.XYZ(
            base_point.X,
            base_point.Y - (len(sorted_pipes) * spacing),
            base_point.Z
        )
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
        t.Commit()

    if failed_pipes:
        TaskDialog.Show("Warning", "Some pipes were skipped: " + ", ".join(failed_pipes))

except Exception, e:
    TaskDialog.Show("Error", "Operation cancelled or failed: " + str(e))