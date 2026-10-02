# -*- coding: utf-8 -*-
import Autodesk
import clr
import os

# Reference .NET and Revit API
clr.AddReference('System')
clr.AddReference('System.Windows.Forms')
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

# Import .NET classes properly
from System.Windows.Forms import SaveFileDialog, DialogResult
from System.Diagnostics import Process, ProcessStartInfo
import System

from Autodesk.Revit.UI import (
    TaskDialog,
    TaskDialogCommandLinkId,
    TaskDialogCommonButtons,
    TaskDialogResult
)

# Get the current Revit document
doc = __revit__.ActiveUIDocument.Document

# Settings storage configuration (shared with your points exporter)
SETTINGS_FOLDER = r"C:\Temp"
SETTINGS_FILE = os.path.join(SETTINGS_FOLDER, "Ribbon_ExportPoints.txt")


def load_dwg_output_path(default_fallback_path):
    """Loads the previously saved output path from settings if available."""
    try:
        if os.path.exists(SETTINGS_FILE):
            with open(SETTINGS_FILE, "r") as f:
                for line in f.readlines():
                    line = line.strip()
                    if line.startswith("output_path="):
                        path_val = line.split("=", 1)[1].strip()
                        if path_val:
                            # Swap extension to .dwg if the setting pointed to a .csv
                            base_dir = os.path.dirname(path_val)
                            file_name = os.path.splitext(os.path.basename(default_fallback_path))[0] + ".dwg"
                            if base_dir and os.path.exists(base_dir):
                                return os.path.join(base_dir, file_name)
    except:
        pass
    return default_fallback_path


def show_export_success_dialog(filepath):
    try:
        td = TaskDialog("DWG Export")
        td.MainInstruction = "DWG export completed successfully."
        td.MainContent = filepath
        td.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Open File")
        td.CommonButtons = TaskDialogCommonButtons.Close

        result = td.Show()

        if result == TaskDialogResult.CommandLink1:
            psi = ProcessStartInfo(filepath)
            psi.UseShellExecute = True
            Process.Start(psi)

    except Exception as ex:
        TaskDialog.Show(
            "Error",
            "Export succeeded, but could not open dialog/file:\n{}".format(str(ex))
        )


def export_to_dwg(view, output_path):
    try:
        # Create DWG export options
        dwg_options = Autodesk.Revit.DB.DWGExportOptions()

        # Set export to use shared coordinates
        dwg_options.SharedCoords = True

        # Set MergedViews to true to prevent Xrefs
        dwg_options.MergedViews = True

        # Set HideUnreferenceViewTags
        dwg_options.HideUnreferenceViewTags = True

        # Additional DWG export settings
        dwg_options.ExportingAreas = False
        dwg_options.LineScaling = Autodesk.Revit.DB.LineScaling.PaperSpace
        dwg_options.TargetUnit = Autodesk.Revit.DB.ExportUnit.Inch

        # Define the view set to export
        view_set = System.Collections.Generic.List[Autodesk.Revit.DB.ElementId]()
        view_set.Add(view.Id)

        # Get the directory and file name from the output path
        output_folder = os.path.dirname(output_path)
        file_name = os.path.splitext(os.path.basename(output_path))[0]

        # Perform the export
        result = doc.Export(output_folder, file_name, view_set, dwg_options)
        return result

    except Exception as ex:
        TaskDialog.Show("Error", "Error during DWG export: {}".format(str(ex)))
        return False


def main():
    """
    Main function to export the active floor plan view to DWG using saved paths.
    """
    active_view = doc.ActiveView
    if not active_view:
        TaskDialog.Show("Error", "No active view found")
        return

    # Check if the active view is a floor plan
    if active_view.ViewType not in [
        Autodesk.Revit.DB.ViewType.FloorPlan,
        Autodesk.Revit.DB.ViewType.CeilingPlan
    ]:
        TaskDialog.Show("Error", "The active view is not a floor plan")
        return

    # Sanitize the view name for the default file name
    default_file_name = active_view.Name.replace(":", "_").replace("/", "_").replace("\\", "_") + ".dwg"
    
    # Fallback to Desktop if settings folder/file doesn't exist yet
    desktop_folder = os.path.join(os.path.expanduser("~"), "Desktop")
    fallback_path = os.path.join(desktop_folder, default_file_name)

    # Retrieve path from settings if available, keeping the target folder preference
    target_path = load_dwg_output_path(fallback_path)
    default_folder = os.path.dirname(target_path)
    if not os.path.exists(default_folder):
        default_folder = desktop_folder

    # File save dialog for DWG export using .NET
    save_dialog = SaveFileDialog()
    save_dialog.Title = "Save DWG File"
    save_dialog.Filter = "AutoCAD DWG Files (*.dwg)|*.dwg"
    save_dialog.DefaultExt = "dwg"
    save_dialog.InitialDirectory = default_folder
    save_dialog.FileName = default_file_name

    # Show the save dialog and check if the user clicked OK
    if save_dialog.ShowDialog() != DialogResult.OK:
        TaskDialog.Show("Error", "No save location selected")
        return

    # Get the selected file path
    file_path = save_dialog.FileName

    # Export the active floor plan view
    success = export_to_dwg(active_view, file_path)

    if success and os.path.exists(file_path):
        show_export_success_dialog(file_path)
    elif success:
        TaskDialog.Show("Export Complete", "DWG export completed.")


if __name__ == "__main__":
    main()