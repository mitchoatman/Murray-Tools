# -*- coding: UTF-8 -*-
from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory, BuiltInParameter,
    FabricationPart
)
from pyrevit import script
from collections import Counter

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
output = script.get_output()

# -------------------------------------------------------------
# CONFIG: Fabrication ItemCustomId values (CIDs).
# These depend on your loaded Fabrication config. If results are empty,
# the diagnostic table below lists the CIDs actually present.
# -------------------------------------------------------------
PIPE_CID = 2041
HANGER_CID = 838


# -------------------------------------------------------------
# Helpers (defined locally, no external lib import needed)
# -------------------------------------------------------------
def get_parameter_value_by_name_AsString(element, parameterName):
    param = element.LookupParameter(parameterName)
    if param and param.HasValue:
        return param.AsString() or ""
    return ""


def get_parameter_value_by_name(element, parameterName):
    param = element.LookupParameter(parameterName)
    if param and param.HasValue:
        return param.AsValueString() or ""
    return ""


def is_straight(part):
    # IsAStraight is a METHOD on FabricationPart, it must be called
    try:
        return bool(part.IsAStraight())
    except Exception:
        return False


def get_level_name(element):
    # Invalid or missing LevelId returns None from GetElement
    level_elem = doc.GetElement(element.LevelId)
    return level_elem.Name if level_elem else "Unassigned Level"


# -------------------------------------------------------------
# Project title
# -------------------------------------------------------------
project_info_collector = FilteredElementCollector(doc).OfCategory(
    BuiltInCategory.OST_ProjectInformation).ToElements()
project_name = None
if project_info_collector:
    project_name_param = project_info_collector[0].LookupParameter("Project Name")
    if project_name_param and project_name_param.HasValue:
        project_name = project_name_param.AsString()
project_name = project_name or "Untitled Project"
file_name = doc.Title or "Untitled File"

output.print_md("# {} - {}".format(project_name, file_name))

# -------------------------------------------------------------
# Scope: selection if any, otherwise whole active document
# -------------------------------------------------------------
selected_ids = uidoc.Selection.GetElementIds()

if selected_ids:
    elements_to_process = [doc.GetElement(i) for i in selected_ids
                           if doc.GetElement(i) is not None]
    mode_label = "Selected "
    scope_label = "current selection"
else:
    pipe_collector = FilteredElementCollector(doc).OfCategory(
        BuiltInCategory.OST_FabricationPipework).WhereElementIsNotElementType()
    hanger_collector = FilteredElementCollector(doc).OfCategory(
        BuiltInCategory.OST_FabricationHangers).WhereElementIsNotElementType()
    elements_to_process = list(pipe_collector) + list(hanger_collector)
    mode_label = ""
    scope_label = "entire project"

fab_parts = [e for e in elements_to_process if isinstance(e, FabricationPart)]

# -------------------------------------------------------------
# 1. STRAIGHT PIPES
# -------------------------------------------------------------
pipes = [p for p in fab_parts if p.ItemCustomId == PIPE_CID and is_straight(p)]

material_totals = {}
level_material_totals = {}
level_system_material_totals = {}

for pipe in pipes:
    len_param = pipe.get_Parameter(BuiltInParameter.FABRICATION_PART_LENGTH)
    if not (len_param and len_param.HasValue):
        continue
    length = len_param.AsDouble()  # internal units: feet

    level_name = get_level_name(pipe)
    material_name = get_parameter_value_by_name(pipe, 'Part Material') or "No Material"
    system_name = get_parameter_value_by_name_AsString(pipe, 'Fabrication Service Name') or "No Service"

    material_totals[material_name] = material_totals.get(material_name, 0.0) + length

    key_lm = (level_name, material_name)
    level_material_totals[key_lm] = level_material_totals.get(key_lm, 0.0) + length

    key_lsm = (level_name, system_name, material_name)
    level_system_material_totals[key_lsm] = level_system_material_totals.get(key_lsm, 0.0) + length

material_table_data = sorted(
    [[m, "{:.2f}".format(l)] for m, l in material_totals.items()],
    key=lambda x: x[0])
level_material_table_data = sorted(
    [[k[0], k[1], "{:.2f}".format(l)] for k, l in level_material_totals.items()],
    key=lambda x: (x[0], x[1]))
level_system_material_table_data = sorted(
    [[k[0], k[1], k[2], "{:.2f}".format(l)] for k, l in level_system_material_totals.items()],
    key=lambda x: (x[0], x[1], x[2]))

if material_table_data:
    output.print_table(
        table_data=material_table_data,
        title="Total Lengths of {}Fabrication Pipes by Material".format(mode_label),
        columns=["Material", "Length (Linear Feet)"],
        formats=['', '{}'])

    output.print_table(
        table_data=level_material_table_data,
        title="Total Lengths of {}Fabrication Pipes by Level and Material".format(mode_label),
        columns=["Level", "Material", "Length (Linear Feet)"],
        formats=['', '', '{}'])

    output.print_table(
        table_data=level_system_material_table_data,
        title="Total Lengths of {}Fabrication Pipes by Level, System, and Material".format(mode_label),
        columns=["Level", "System", "Material", "Length (Linear Feet)"],
        formats=['', '', '', '{}'])
else:
    output.print_md("**No straight fabrication pipes with CID {} found in the {}.**".format(
        PIPE_CID, scope_label))

# -------------------------------------------------------------
# 2. HANGERS
# -------------------------------------------------------------
hanger_parts = [h for h in fab_parts if h.ItemCustomId == HANGER_CID]

hanger_counts = {}
for hanger in hanger_parts:
    level_name = get_level_name(hanger)  # no level is allowed, shows as Unassigned
    system_name = get_parameter_value_by_name_AsString(hanger, 'Fabrication Service Name') or "No Service"
    hanger_name = get_parameter_value_by_name(hanger, 'Family') or "Unnamed Hanger"
    key = (level_name, system_name, hanger_name)
    hanger_counts[key] = hanger_counts.get(key, 0) + 1

hanger_table_data = sorted(
    [[k[0], k[1], k[2], str(c)] for k, c in hanger_counts.items()],
    key=lambda x: (x[0], x[1], x[2]))

if hanger_table_data:
    output.print_table(
        table_data=hanger_table_data,
        title="Total Counts of {}Fabrication Hangers by Level, System, and Hanger Type".format(mode_label),
        columns=["Level", "System", "Hanger Type", "Count"],
        formats=['', '', '', '{}'])
else:
    output.print_md("**No fabrication hangers with CID {} found in the {}.**".format(
        HANGER_CID, scope_label))

# -------------------------------------------------------------
# 3. DIAGNOSTIC: shown only if either section came up empty,
# lists the CIDs actually present so PIPE_CID / HANGER_CID can be corrected
# -------------------------------------------------------------
if not material_table_data or not hanger_table_data:
    counts = Counter((p.ItemCustomId, p.Name) for p in fab_parts)
    diag_rows = [[str(cid), name, str(n)] for (cid, name), n in counts.most_common()]
    if diag_rows:
        output.print_table(
            table_data=diag_rows,
            title="Diagnostic: Fabrication parts found in {} (CID, Name, Count)".format(scope_label),
            columns=["CID", "Part Name", "Count"],
            formats=['', '', '{}'])
    else:
        output.print_md("**Diagnostic: no FabricationPart elements found in the {}.**".format(scope_label))