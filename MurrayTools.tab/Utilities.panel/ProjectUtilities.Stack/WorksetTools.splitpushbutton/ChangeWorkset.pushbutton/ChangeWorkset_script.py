# Imports
import Autodesk
from pyrevit import revit
from Autodesk.Revit.DB import Transaction, Workset, WorksetKind, FilteredWorksetCollector, BuiltInParameter
from Autodesk.Revit.UI import Selection
from pyrevit import forms

# Get the active document
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# Initialize a list to hold workset names
workset_names = []

# Use FilteredWorksetCollector to get all user worksets
all_worksets = FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset).ToWorksets()

# Iterate through all worksets to retrieve names
for workset in all_worksets:
    workset_names.append(workset.Name)

# Check if any worksets were found and display them
if not workset_names:
    forms.alert("No worksets found in this model.")
else:
    try:
        # Start selecting elements
        sel = uidoc.Selection.PickObjects(Selection.ObjectType.Element, 'Select Elements')
        selected_elements = [doc.GetElement(elId) for elId in sel]

        # Prompt user to select a workset
        selected_workset_name = forms.ask_for_one_item(
            workset_names,
            default=workset_names[0],
            prompt='Select a workset:',
            title='Workset Selection'
        )

        # Find the selected workset by name
        selected_workset = next(ws for ws in all_worksets if ws.Name == selected_workset_name)
        workset_int = selected_workset.Id.IntegerValue

        # Variables to track skipped elements
        skipped_elements = []

        # START TRANSACTION
        t = Transaction(doc, 'Set Workset for Elements')
        t.Start()
        for el in selected_elements:
            workset_param = el.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM)

            # Check if the parameter is writable
            if workset_param and not workset_param.IsReadOnly:
                workset_param.Set(workset_int)
            else:
                skipped_elements.append(el)

        t.Commit()

    except Exception as e:
        forms.alert("Error: {0}".format(e))
