# -*- coding: utf-8 -*-
from pyrevit import revit, script
import Autodesk.Revit.DB as DB

doc = revit.doc


# Get current selection
selection = revit.get_selection()

# Open a Revit transaction to modify parameters
with revit.Transaction("Clear STRATUS Parameters"):
    updated_elements_count = 0
    total_params_cleared = 0

    for element in selection:
        element_modified = False
        parameters = element.Parameters

        for param in parameters:
            param_name = param.Definition.Name
            
            # Check if parameter name starts with "STRATUS" and is editable
            is_target_param = param_name.startswith("STRATUS") or param_name == "FP_Spool Map"
                storage_type = param.StorageType
                
                # Clear value based on parameter storage type
                if storage_type == DB.StorageType.String:
                    param.Set("")
                    element_modified = True
                    total_params_cleared += 1
                elif storage_type == DB.StorageType.Double:
                    param.Set(0.0)
                    element_modified = True
                    total_params_cleared += 1
                elif storage_type == DB.StorageType.Integer:
                    param.Set(0)
                    element_modified = True
                    total_params_cleared += 1
                elif storage_type == DB.StorageType.ElementId:
                    param.Set(DB.ElementId.InvalidElementId)
                    element_modified = True
                    total_params_cleared += 1

        if element_modified:
            updated_elements_count += 1