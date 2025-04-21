# -*- coding: utf-8 -*-
import Autodesk
import clr
import os

# Reference .NET and Revit API
clr.AddReference('System.Windows.Forms')
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

# Import .NET SaveFileDialog and System for generic collections
from System.Windows.Forms import SaveFileDialog, DialogResult
import System

# Get the current Revit document
doc = __revit__.ActiveUIDocument.Document

def export_to_dwg(view, output_path):
    try:
        # Create DWG export options
        dwg_options = Autodesk.Revit.DB.DWGExportOptions()

        # Debug: List all properties of DWGExportOptions
        # print("DWGExportOptions properties: {}".format([attr for attr in dir(dwg_options) if not attr.startswith('_')]))

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
        doc.Export(output_folder, file_name, view_set, dwg_options)
        # print("DWG exported successfully to: {}".format(output_path))
    
    except Exception as ex:
        print("Error during DWG export: {}".format(str(ex)))

def main():
    """
    Main function to export the active floor plan view to DWG.
    """
    # Get the active view
    active_view = doc.ActiveView
    if not active_view:
        print("No active view found")
        return

    # Check if the active view is a floor plan
    if active_view.ViewType != Autodesk.Revit.DB.ViewType.FloorPlan:
        print("The active view is not a floor plan")
        return

    # Sanitize the view name for the default file name
    default_file_name = active_view.Name.replace(":", "_").replace("/", "_").replace("\\", "_") + ".dwg"
    default_folder = os.path.join(os.environ['USERPROFILE'], 'Desktop')

    # File save dialog for DWG export using .NET
    save_dialog = SaveFileDialog()
    save_dialog.Title = "Save DWG File"
    save_dialog.Filter = "AutoCAD DWG Files (*.dwg)|*.dwg"
    save_dialog.DefaultExt = "dwg"
    save_dialog.InitialDirectory = default_folder
    save_dialog.FileName = default_file_name

    # Show the save dialog and check if the user clicked OK
    if save_dialog.ShowDialog() != DialogResult.OK:
        print("No save location selected")
        return

    # Get the selected file path
    file_path = save_dialog.FileName

    # Export the active floor plan view
    export_to_dwg(active_view, file_path)

if __name__ == "__main__":
    main()