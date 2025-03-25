from Autodesk.Revit.UI import TaskDialog, TaskDialogCommonButtons, TaskDialogResult
from pyrevit import revit, EXEC_PARAMS

args = __eventargs__
doc = revit.doc

if not doc.IsFamilyDocument:
    # Create the dialog
    dialog = TaskDialog("Warning")
    dialog.MainInstruction = "Do NOT use IMPORT CAD! \nUse LINK CAD instead."
    
    # Set the buttons to OK and Cancel
    dialog.CommonButtons = TaskDialogCommonButtons.Ok | TaskDialogCommonButtons.Cancel
    
    # Set Cancel as the default button (optional)
    dialog.DefaultButton = TaskDialogResult.Cancel
    
    # Show the dialog and get the user's response
    result = dialog.Show()
    
    # Check if Cancel was clicked
    if result == TaskDialogResult.Cancel:
        # Cancel the operation
        EXEC_PARAMS.event_args.Cancel = True