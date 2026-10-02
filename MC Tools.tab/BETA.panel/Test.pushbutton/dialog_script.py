# -*- coding: utf-8 -*-
from pyrevit import revit, script
import Autodesk.Revit.DB as DB

doc = revit.doc
output = script.get_output()

# Get current pre-selection
selection = revit.get_selection()

# If nothing is pre-selected, prompt the user to pick elements interactively
if not selection:
    selection = revit.pick_elements("Select elements to clear STRATUS parameters")
    
    # If the user cancels the picker (e.g., presses ESC), exit safely
    if not selection:
        script.exit("No elements selected. Operation cancelled.")

# Open a Revit transaction to modify parameters
with revit.Transaction("Clear STRATUS Parameters"):
    for element in selection:
        elem_id = element.Id
        parameters = element.Parameters

        for param in parameters:
            param_name = param.Definition.Name
            
            # Check if parameter matches your target names
            is_target_param = param_name.startswith("STRATUS") or param_name == "FP_Spool Map"
            
            if is_target_param:
                # Check if it's read-only and log it
                if param.IsReadOnly:
                    msg = "⚠️ **Read-Only Skipped:** Element ID `{0}` | Parameter: **`{1}`**".format(elem_id, param_name)
                    output.print_md(msg)
                    continue
                
                storage_type = param.StorageType
                
                # Attempt to clear value safely
                try:
                    if storage_type == DB.StorageType.String:
                        param.Set("")
                    elif storage_type == DB.StorageType.ElementId:
                        param.Set(DB.ElementId.InvalidElementId)
                    elif storage_type == DB.StorageType.Integer:
                        param.Set(0)
                    elif storage_type == DB.StorageType.Double:
                        param.Set(0.0)
                except Exception as e:
                    err_msg = "❌ **Error:** Element ID `{0}` | Parameter: **`{1}`** | Error: `{2}`".format(elem_id, param_name, e)
                    output.print_md(err_msg)