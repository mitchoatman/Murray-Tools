from pyrevit import revit, DB
from Autodesk.Revit.DB import NavisworksExportOptions, FilteredElementCollector, Transaction, TransactionGroup, Material, FabricationPart, ElementId, BuiltInCategory, WorksharingUtils, WorksetId
from Autodesk.Revit.UI import TaskDialog
from System.Windows.Forms import SaveFileDialog, DialogResult, FolderBrowserDialog
from Parameters.Add_SharedParameters import Shared_Params
import os
import json
import sys

Shared_Params()

doc = revit.doc

# Get selected views
uidoc = __revit__.ActiveUIDocument
selected_ids = uidoc.Selection.GetElementIds()
selected_views = []
for element_id in selected_ids:
    element = doc.GetElement(element_id)
    if isinstance(element, DB.View) and element.ViewType in [
        DB.ViewType.FloorPlan, DB.ViewType.CeilingPlan, DB.ViewType.ThreeD,
        DB.ViewType.Section, DB.ViewType.Elevation, DB.ViewType.Detail
    ] and not element.IsTemplate:
        selected_views.append(element)

# If no valid views are selected, use active view
if not selected_views:
    active_view = doc.ActiveView
    if not active_view or active_view.ViewType not in [
        DB.ViewType.FloorPlan, DB.ViewType.CeilingPlan, DB.ViewType.ThreeD,
        DB.ViewType.Section, DB.ViewType.Elevation, DB.ViewType.Detail
    ] or active_view.IsTemplate:
        TaskDialog.Show("Error", "No valid views selected, and the active view is not valid for export. Please select graphical views or activate one.")
        sys.exit()
    selected_views = [active_view]

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

folder_path = None
if len(selected_views) == 1:
    save_dialog = SaveFileDialog()
    save_dialog.Title = "Save NWC File"
    save_dialog.Filter = "Navisworks Files (*.nwc)|*.nwc"
    save_dialog.DefaultExt = "nwc"
    save_dialog.InitialDirectory = last_save_path
    save_dialog.FileName = selected_views[0].Name

    if save_dialog.ShowDialog() != DialogResult.OK:
        TaskDialog.Show("Error", "Export cancelled by user. Please select a valid file path.")
        sys.exit()

    full_file_path = save_dialog.FileName
    folder_path = os.path.dirname(full_file_path)
    view_file_mapping = [(selected_views[0], os.path.basename(full_file_path))]
else:
    folder_dialog = FolderBrowserDialog()
    folder_dialog.Description = "Select Folder to Save NWC Files"
    folder_dialog.SelectedPath = last_save_path
    folder_dialog.ShowNewFolderButton = True

    if folder_dialog.ShowDialog() != DialogResult.OK:
        TaskDialog.Show("Error", "Export cancelled by user. Please select a valid folder path.")
        sys.exit()

    folder_path = folder_dialog.SelectedPath
    view_file_mapping = [(view, view.Name + ".nwc") for view in selected_views]

# Save the folder path
with open(filepath, 'w') as f:
    f.write(folder_path)

# Check if model is workshared and handle workset checkout
if doc.IsWorkshared:
    view_workset_ids = {}
    for view in selected_views:
        element_ids = set()
        collectors = [
            FilteredElementCollector(doc, view.Id).OfClass(FabricationPart).WhereElementIsNotElementType(),
        ]
        for collector in collectors:
            for element in collector.ToElements():
                element_ids.add(element.Id)
        workset_ids = set()
        for element_id in element_ids:
            try:
                workset_id = WorksharingUtils.GetWorksetId(doc, element_id)
                if workset_id != WorksetId.InvalidWorksetId:
                    workset_ids.add(workset_id)
            except:
                pass
        view_workset_ids[view.Id] = workset_ids

    checkout_failed = False
    all_workset_ids = set()
    for workset_ids in view_workset_ids.values():
        all_workset_ids.update(workset_ids)

    for workset_id in all_workset_ids:
        try:
            if WorksharingUtils.GetCheckoutStatus(doc, workset_id) != DB.CheckoutStatus.OwnedByCurrentUser:
                WorksharingUtils.CheckoutWorksets(doc, [workset_id])
        except Exception as e:
            checkout_failed = True
            TaskDialog.Show("Error", "Failed to check out workset {}: {}".format(workset_id.IntegerValue, str(e)))
            break

    if checkout_failed:
        TaskDialog.Show("Error", "Unable to check out one or more worksets. They may be owned by another user.")
        sys.exit()

SHARED_PARAM_NAME = "FP_Material"

def get_or_update_material(doc, material_name, color):
    materials = FilteredElementCollector(doc).OfClass(Material).ToElements()
    for mat in materials:
        if mat.Name == material_name:
            mat_color = mat.Color
            needs_update = mat_color.Red != color.Red or mat_color.Green != color.Green or mat_color.Blue != color.Blue
            is_insulation = material_name.lower() in ['insulation_material', 'fp_insulation', 'insulation', 'mp insulation']
            if needs_update or (is_insulation and mat.Transparency != 50):
                try:
                    with revit.Transaction("Update Material Color"):
                        mat.Color = color
                        if is_insulation:
                            mat.Transparency = 50
                    return mat.Id
                except Exception as e:
                    TaskDialog.Show("Error", "Failed to update material {}: {}".format(material_name, str(e)))
                    continue
            return mat.Id
    try:
        with revit.Transaction("Create Material"):
            new_mat_id = Material.Create(doc, material_name)
            new_mat = doc.GetElement(new_mat_id)
            if new_mat:
                new_mat.Color = color
                if material_name.lower() in ['insulation_material', 'fp_insulation', 'insulation', 'mp insulation']:
                    new_mat.Transparency = 50
                new_mat.SurfaceForegroundPatternId = DB.ElementId.InvalidElementId
                new_mat.SurfaceBackgroundPatternId = DB.ElementId.InvalidElementId
            return new_mat_id
    except Exception as e:
        TaskDialog.Show("Error", "Failed to create material {}: {}".format(material_name, str(e)))
        return None

# Process each view
for view, file_name in view_file_mapping:
    filter_material_mapping = {}
    elements_to_update = {}

    with revit.Transaction("Analyze View Filters"):
        filters = view.GetFilters()
        for filter_id in filters:
            filter_element = doc.GetElement(filter_id)
            filter_name = filter_element.Name
            # Skip filters with insulation-related names
            if filter_name.lower() in ['insulation_material', 'fp_insulation', 'insulation', 'mp insulation']:
                continue
            override_settings = view.GetFilterOverrides(filter_id)
            pattern_color = override_settings.SurfaceForegroundPatternColor
            line_color = override_settings.ProjectionLineColor
            filter_color = pattern_color if pattern_color.IsValid else line_color

            if filter_color.IsValid:
                material_id = get_or_update_material(doc, filter_name, filter_color)
                if material_id is None:
                    TaskDialog.Show("Error", "Failed to create material for filter {}.".format(filter_name))
                    continue
                filter_material_mapping[filter_name] = (filter_color, material_id)
                
                filter_rules = filter_element.GetElementFilter()
                if filter_rules:
                    fab_parts = FilteredElementCollector(doc, view.Id).OfClass(FabricationPart).WherePasses(filter_rules).ToElements()
                    for element in fab_parts:
                        elements_to_update[element.Id] = material_id

    categories = doc.Settings.Categories
    pipe_insulation_category = categories.get_Item(BuiltInCategory.OST_FabricationPipeworkInsulation)
    duct_insulation_category = categories.get_Item(BuiltInCategory.OST_FabricationDuctworkInsulation)

    tg = TransactionGroup(doc, "Assign Materials for View: " + view.Name)
    try:
        tg.Start()
        
        with revit.Transaction("Assign Materials to Parts"):
            assigned_count = 0
            element_assignment_errors = []
            for element_id, material_id in elements_to_update.items():
                element = doc.GetElement(element_id)
                if element:
                    param = element.LookupParameter(SHARED_PARAM_NAME)
                    if param is None:
                        element_assignment_errors.append("Element {} in view {} does not have parameter {}.".format(element_id.IntegerValue, view.Name, SHARED_PARAM_NAME))
                        continue
                    try:
                        current_material_id = param.AsElementId()
                        current_material_name = doc.GetElement(current_material_id).Name if current_material_id and current_material_id != ElementId.InvalidElementId else "None"
                        new_material_name = doc.GetElement(material_id).Name if material_id else "None"
                        if current_material_id != material_id:
                            param.Set(material_id)
                            assigned_count += 1
                            # Log assignment for debugging
                            # print("Assigned material {} to element {} in view {}".format(new_material_name, element_id.IntegerValue, view.Name))
                    except Exception as e:
                        element_assignment_errors.append("Failed to assign material {} to element {} in view {}: {}".format(new_material_name, element_id.IntegerValue, view.Name, str(e)))
            
            if element_assignment_errors:
                TaskDialog.Show("Error", "Element material assignment errors in view {}:\n{}".format(view.Name, "\n".join(element_assignment_errors)))
            
            default_color = DB.Color(128, 128, 128)
            insulation_material_id = get_or_update_material(doc, "Insulation_Material", default_color)
            if insulation_material_id:
                if pipe_insulation_category:
                    current_material = pipe_insulation_category.Material
                    if current_material is None or current_material.Id != insulation_material_id:
                        try:
                            pipe_insulation_category.Material = doc.GetElement(insulation_material_id)
                        except Exception as e:
                            TaskDialog.Show("Error", "Failed to set material for OST_FabricationPipeworkInsulation: {}".format(str(e)))
                
                if duct_insulation_category:
                    current_material = duct_insulation_category.Material
                    if current_material is None or current_material.Id != insulation_material_id:
                        try:
                            duct_insulation_category.Material = doc.GetElement(insulation_material_id)
                        except Exception as e:
                            TaskDialog.Show("Error", "Failed to set material for OST_FabricationDuctworkInsulation: {}".format(str(e)))
            else:
                TaskDialog.Show("Error", "Failed to create or retrieve material 'Insulation_Material'.")
        
        tg.Assimilate()
        
    except Exception as e:
        TaskDialog.Show("Error", "Error during material assignment for view {}: {}".format(view.Name, str(e)))
        if tg.HasStarted():
            tg.RollBack()
    finally:
        if tg.HasStarted() and not tg.HasEnded():
            tg.RollBack()

    OPTIONS_FILE = r"C:\Temp\Ribbon_NavisworksExportOptions.txt"
    default_options = {
        "ExportScope": "View",
        "Coordinates": "Shared",
        "FindMissingMaterials": True,
        "DivideFileIntoLevels": False,
        "ConvertElementProperties": True,
        "ExportUrls": False,
        "ConvertLinkedCADFormats": False,
        "ExportLinks": False,
        "SuccessMessage": True
    }

    def load_options():
        if os.path.exists(OPTIONS_FILE):
            with open(OPTIONS_FILE, 'r') as f:
                return json.load(f)
        return default_options

    export_options = NavisworksExportOptions()
    options = load_options()
    coordinates_map = {"Shared": DB.NavisworksCoordinates.Shared, "Project Internal": DB.NavisworksCoordinates.Internal}
    coordinates_value = coordinates_map.get(options["Coordinates"], DB.NavisworksCoordinates.Shared)

    export_options.ExportScope = getattr(DB.NavisworksExportScope, options["ExportScope"])
    export_options.ViewId = view.Id
    export_options.Coordinates = coordinates_value
    export_options.FindMissingMaterials = options["FindMissingMaterials"]
    export_options.DivideFileIntoLevels = options["DivideFileIntoLevels"]
    export_options.ConvertElementProperties = options["ConvertElementProperties"]
    export_options.ExportUrls = options["ExportUrls"]
    export_options.ConvertLinkedCADFormats = options["ConvertLinkedCADFormats"]
    export_options.ExportLinks = options["ExportLinks"]

    try:
        file_name_no_ext = os.path.splitext(file_name)[0]
        doc.Export(folder_path, file_name_no_ext, export_options)
        if options.get("SuccessMessage", True):
            TaskDialog.Show("Success", "Successfully exported view {} to {}".format(view.Name, os.path.join(folder_path, file_name)))
    except Exception as e:
        TaskDialog.Show("Error", "Export failed for view {}: {}".format(view.Name, str(e)))