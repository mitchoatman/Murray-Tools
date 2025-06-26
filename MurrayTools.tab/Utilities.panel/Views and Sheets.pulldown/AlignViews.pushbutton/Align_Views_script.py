# -*- coding: utf-8 -*-
from pyrevit import forms
from pyrevit import revit
from Autodesk.Revit.DB import ViewSheet, Viewport, FilteredElementCollector, XYZ
import sys

doc = __revit__.ActiveUIDocument.Document

# Get all sheets
all_sheets = FilteredElementCollector(doc).OfClass(ViewSheet).ToElements()

# Step 1 - User selects template sheet
template_sheet = forms.SelectFromList.show(
    all_sheets,
    name_attr='Name',
    title='Select Template Sheet (To Get View Position)',
    button_name='Select'
)
if not template_sheet:
    sys.exit()

# Step 2 - Get first viewport's position from template sheet
template_vp = None
for vp_id in template_sheet.GetAllViewports():
    template_vp = doc.GetElement(vp_id)
    break

if not template_vp:
    forms.alert("No views found on the selected sheet.", title="Error")
    sys.exit()

template_position = template_vp.GetBoxCenter()

# Step 3 - Select target sheets to align views
target_sheets = forms.SelectFromList.show(
    all_sheets,
    name_attr='Name',
    multiselect=True,
    title='Select Sheets to Align Views',
    button_name='Align Views'
)
if not target_sheets:
    sys.exit()

# Step 4 - Align views on selected sheets
aligned = 0
with revit.Transaction("Align Views on Sheets"):
    for sheet in target_sheets:
        for vp_id in sheet.GetAllViewports():
            vp = doc.GetElement(vp_id)
            vp.SetBoxCenter(template_position)
            aligned += 1

forms.show_balloon("Alignment Complete", "Aligned %d view(s) to match template position." % aligned)