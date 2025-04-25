from pyrevit import revit, DB
from Autodesk.Revit.DB import NavisworksExportOptions, FilteredElementCollector, Transaction, TransactionGroup, Material, FabricationPart, ElementId, BuiltInCategory
from System.Windows.Forms import SaveFileDialog, DialogResult
from SharedParam.Add_Parameters import Shared_Params
import os
import json

Shared_Params()

doc = revit.doc
active_view = doc.ActiveView

# Validate active view
if not active_view or active_view.ViewType not in [DB.ViewType.FloorPlan, DB.ViewType.CeilingPlan, DB.ViewType.ThreeD, DB.ViewType.Section, DB.ViewType.Elevation, DB.ViewType.Detail]:
    from Autodesk.Revit.UI import TaskDialog, TaskDialogCommonButtons, TaskDialogResult
    dialog = TaskDialog("Invalid Active View")
    dialog.MainInstruction = "The active view is not valid for export."
    dialog.MainContent = "Please click into a graphical view (e.g., 3D View, Floor Plan, Section, or Elevation) and try again."
    dialog.CommonButtons = TaskDialogCommonButtons.Close
    dialog.Show()
    import sys
    sys.exit()

# File save dialog for NWC export
folder_name = "C:\\Temp"
filepath = os.path.join(folder_name, "Ribbon_NWC-Export.txt")

if not os.path.exists(folder_name):
    os.makedirs(folder_name)

default_desktop_path = os.path.expandvars("%USERPROFILE%\\Desktop")

if os.path.exists(filepath):
    with open(filepath, 'r') as f:
        last_save_path = f.read().strip()
    if not os.path.exists(last_save_path):
        last_save_path = default_desktop_path
else:
    last_save_path = default_desktop_path

save_dialog = SaveFileDialog()
save_dialog.Title = "Save NWC File"
save_dialog.Filter = "Navisworks Files (*.nwc)|*.nwc"
save_dialog.DefaultExt = "nwc"
save_dialog.InitialDirectory = last_save_path
save_dialog.FileName = active_view.Name

if save_dialog.ShowDialog() != DialogResult.OK:
    print "Export cancelled by user."
    import sys
    sys.exit()

full_file_path = save_dialog.FileName
folder_path = os.path.dirname(full_file_path)
file_name = os.path.splitext(os.path.basename(full_file_path))[0]

SHARED_PARAM_NAME = "FP_Material"

def get_or_update_material(doc, material_name, color):
    materials = FilteredElementCollector(doc).OfClass(Material).ToElements()
    for mat in materials:
        if mat.Name == material_name:
            mat_color = mat.Color
            needs_update = mat_color.Red != color.Red or mat_color.Green != color.Green or mat_color.Blue != color.Blue
            is_insulation = material_name.lower() in ['fp_insulation', 'insulation', 'mp insulation']
            if needs_update or (is_insulation and mat.Transparency != 50):
                try:
                    with revit.Transaction("Update Material Color"):
                        mat.Color = color
                        if is_insulation:
                            mat.Transparency = 50
                except Exception as e:
                    print "Failed to update material %s: %s" % (material_name, str(e))
                    continue
            return mat.Id
    try:
        with revit.Transaction("Create Material"):
            new_mat_id = Material.Create(doc, material_name)
            new_mat = doc.GetElement(new_mat_id)
            if new_mat:
                new_mat.Color = color
                if material_name.lower() in ['fp_insulation', 'insulation', 'mp insulation']:
                    new_mat.Transparency = 50
                new_mat.SurfaceForegroundPatternId = DB.ElementId.InvalidElementId
                new_mat.SurfaceBackgroundPatternId = DB.ElementId.InvalidElementId
            return new_mat_id
    except Exception as e:
        print "Failed to create material %s: %s" % (material_name, str(e))
        return None

filter_material_mapping = {}
elements_to_update = {}
insulation_material_id = None

with revit.Transaction("Analyze View Filters"):
    filters = active_view.GetFilters()
    for filter_id in filters:
        filter_element = doc.GetElement(filter_id)
        filter_name = filter_element.Name
        
        override_settings = active_view.GetFilterOverrides(filter_id)
        pattern_color = override_settings.SurfaceForegroundPatternColor
        line_color = override_settings.ProjectionLineColor
        filter_color = pattern_color if pattern_color.IsValid else line_color

        if filter_color.IsValid:
            material_id = get_or_update_material(doc, filter_name, filter_color)
            if material_id is None:
                continue  # Skip if material creation/update failed
            filter_material_mapping[filter_name] = (filter_color, material_id)
            
            if filter_name.lower() in ['fp_insulation', 'insulation', 'mp insulation']:
                insulation_material_id = material_id
            else:
                filter_rules = filter_element.GetElementFilter()
                if filter_rules:
                    fab_parts = FilteredElementCollector(doc, active_view.Id).OfClass(FabricationPart).WherePasses(filter_rules).ToElements()
                    for element in fab_parts:
                        if element.Id not in elements_to_update:
                            param = element.LookupParameter(SHARED_PARAM_NAME)
                            if param and param.HasValue and param.AsElementId() == material_id:
                                continue  # Skip if element already has the correct material
                            elements_to_update[element.Id] = material_id

categories = doc.Settings.Categories
insulation_category = categories.get_Item(BuiltInCategory.OST_FabricationPipeworkInsulation)

tg = TransactionGroup(doc, "Assign Materials")
try:
    tg.Start()
    
    with revit.Transaction("Assign Materials to Parts"):
        assigned_count = 0
        for element_id, material_id in elements_to_update.items():
            element = doc.GetElement(element_id)
            if element:
                param = element.LookupParameter(SHARED_PARAM_NAME)
                if param is None:
                    continue
                try:
                    param.Set(material_id)
                    assigned_count += 1
                except Exception as e:
                    print "Failed to assign material to element %s: %s" % (element_id.IntegerValue, str(e))
        
        if insulation_material_id and insulation_category:
            try:
                if insulation_category.Material is None or insulation_category.Material.Id != insulation_material_id:
                    insulation_category.Material = doc.GetElement(insulation_material_id)
            except Exception as e:
                pass
    
    tg.Assimilate()
    
except Exception as e:
    print "Error during material assignment: %s" % str(e)
    if tg.HasStarted():
        tg.RollBack()
finally:
    if tg.HasStarted():
        tg.RollBack()

# Export the NWC
OPTIONS_FILE = r"C:\Temp\Ribbon_NavisworksExportOptions.txt"

# Default options (matching current hard-coded values)
default_options = {
    "ExportScope": "View",
    "Coordinates": "Shared",
    "FindMissingMaterials": True,
    "DivideFileIntoLevels": False,
    "ConvertElementProperties": True,
    "ExportUrls": False,
    "ConvertLinkedCADFormats": False,
    "ExportLinks": False
}

# Load saved options
def load_options():
    if os.path.exists(OPTIONS_FILE):
        with open(OPTIONS_FILE, 'r') as f:
            return json.load(f)
    return default_options

export_options = NavisworksExportOptions()
options = load_options()

# Map Coordinates option
coordinates_map = {"Shared": DB.NavisworksCoordinates.Shared, "Project Internal": DB.NavisworksCoordinates.Internal}
coordinates_value = coordinates_map.get(options["Coordinates"], DB.NavisworksCoordinates.Shared)

# Apply options
export_options.ExportScope = getattr(DB.NavisworksExportScope, options["ExportScope"])
export_options.ViewId = active_view.Id
export_options.Coordinates = coordinates_value
export_options.FindMissingMaterials = options["FindMissingMaterials"]
export_options.DivideFileIntoLevels = options["DivideFileIntoLevels"]
export_options.ConvertElementProperties = options["ConvertElementProperties"]
export_options.ExportUrls = options["ExportUrls"]
export_options.ConvertLinkedCADFormats = options["ConvertLinkedCADFormats"]
export_options.ExportLinks = options["ExportLinks"]

try:
    doc.Export(folder_path, file_name, export_options)
    with open(filepath, 'w') as f:
        f.write(folder_path)
except Exception as e:
    print "Export failed: %s" % str(e)