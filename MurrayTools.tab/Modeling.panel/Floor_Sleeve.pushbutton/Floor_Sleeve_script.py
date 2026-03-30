# -*- coding: utf-8 -*-
from Autodesk.Revit import DB
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    Family,
    BuiltInCategory,
    FamilySymbol,
    Transaction,
    XYZ,
    ViewType,
    Transform,
    ProjectLocation,
)
from Autodesk.Revit.UI import TaskDialog
from fractions import Fraction
import re
import os

from Parameters.Add_SharedParameters import Shared_Params
Shared_Params()
from Parameters.Get_Set_Params import (
    set_parameter_by_name,
    get_parameter_value_by_name_AsString
)

class PointConverter:
    """Convert coordinates between internal / project / survey systems."""
    def __init__(self, x, y, z, coord_sys='internal', doc=None):
        if doc is None:
            doc = __revit__.ActiveUIDocument.Document
        self.doc = doc
        pt = XYZ(x, y, z)

        srv_trans = self._get_survey_transform()
        proj_trans = self._get_project_transform()

        if coord_sys.lower() == 'internal':
            self.internal = pt
            self.survey   = srv_trans.Inverse.OfPoint(pt)
            self.project  = proj_trans.Inverse.OfPoint(pt)
        elif coord_sys.lower() == 'project':
            self.project  = pt
            self.internal = proj_trans.OfPoint(pt)
            self.survey   = srv_trans.Inverse.OfPoint(self.internal)
        elif coord_sys.lower() == 'survey':
            self.survey   = pt
            self.internal = srv_trans.OfPoint(pt)
            self.project  = proj_trans.Inverse.OfPoint(self.internal)
        else:
            raise ValueError("coord_sys must be 'internal', 'project' or 'survey'")

    def _get_survey_transform(self):
        return self.doc.ActiveProjectLocation.GetTotalTransform()

    def _get_project_transform(self):
        collector = FilteredElementCollector(self.doc).OfClass(ProjectLocation).WhereElementIsNotElementType()
        for loc in collector:
            if loc.Name == "Project":
                return loc.GetTotalTransform()
        return Transform.Identity

# Document & setup
doc   = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView

path, _ = os.path.split(__file__)
family_path = os.path.join(path, 'Round Floor Sleeve.rfa')

# Sleeve length from file
temp_folder = r"c:\Temp"
sleeve_length_file = os.path.join(temp_folder, 'Ribbon_Sleeve.txt')
if not os.path.exists(temp_folder):
    os.makedirs(temp_folder)
if not os.path.exists(sleeve_length_file):
    with open(sleeve_length_file, 'w') as f:
        f.write('6')
with open(sleeve_length_file, 'r') as f:
    sleeve_length = float(f.read().strip())

DIAMETER_MAP = {
    (0.0, 1.0): 2.0, (1.0, 1.25): 2.5, (1.25, 1.5): 3.0,
    (1.5, 2.5): 4.0, (2.5, 3.5): 5.0, (3.5, 4.5): 6.0,
    (4.5, 7.5): 8.0, (7.5, 8.5): 10.0, (8.5, 10.5): 12.0,
    (10.5, 14.5): 16.0, (14.5, 16.5): 18.0, (16.5, 18.5): 20.0,
    (18.5, 20.5): 22.0, (20.5, 22.5): 24.0, (22.5, 24.5): 26.0,
    (24.5, 26.5): 28.0, (26.5, 28.5): 30.0, (28.5, 30.5): 32.0,
    (30.5, 32.5): 34.0, (32.5, 34.5): 36.0
}

def load_family():
    families = FilteredElementCollector(doc).OfClass(Family)
    family_name = 'Round Floor Sleeve'
    if not any(f.Name == family_name for f in families):
        t = Transaction(doc, 'Load Family')
        t.Start()
        doc.LoadFamily(family_path)
        t.Commit()

    collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeAccessory).OfClass(FamilySymbol)
    for fs in collector:
        name_param = fs.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
        if fs.Family.Name == family_name and name_param and name_param.AsString() == family_name:
            return fs
    return None

def get_diameter_from_size(pipe_diameter_feet):
    inches = pipe_diameter_feet * 12
    for (lo, hi), sleeve_in in DIAMETER_MAP.items():
        if lo < inches <= hi:
            return sleeve_in / 12.0
    return 2.0 / 12.0

def is_vertical_pipe(pipe):
    if pipe.ItemCustomId != 2041:
        return False
    conns = list(pipe.ConnectorManager.Connectors)
    if len(conns) < 2:
        return False
    direction = (conns[1].Origin - conns[0].Origin).Normalize()
    return abs(direction.Z) > 0.99

def get_pipe_intersections(pipe, level):
    if not is_vertical_pipe(pipe):
        return []

    bbox = pipe.get_BoundingBox(None)
    if bbox is None:
        return []

    # Use ProjectElevation → always in internal coordinates, independent of Elevation Base
    plane_z_internal = level.ProjectElevation

    if not (bbox.Min.Z < plane_z_internal < bbox.Max.Z):
        return []

    conns = list(pipe.ConnectorManager.Connectors)
    if len(conns) < 2:
        return []

    cx = (conns[0].Origin.X + conns[1].Origin.X) / 2.0
    cy = (conns[0].Origin.Y + conns[1].Origin.Y) / 2.0

    return [XYZ(cx, cy, plane_z_internal)]

def is_duplicate_sleeve(point, existing, tol=0.02):
    for el in existing:
        loc = getattr(el.Location, 'Point', None)
        if loc and all(abs(a - b) < tol for a, b in zip((loc.X, loc.Y, loc.Z), (point.X, point.Y, point.Z))):
            return True
    return False

def clean_size_string(size_str):
    return re.sub(r'["\']|ø', '', size_str.strip())

def place_sleeve_at_intersection(pipe, pt, symbol, level, existing):
    if is_duplicate_sleeve(pt, existing):
        return None

    inst = doc.Create.NewFamilyInstance(
        pt, symbol, level, DB.Structure.StructuralType.NonStructural)

    size_str = pipe.get_Parameter(DB.BuiltInParameter.RBS_REFERENCE_OVERALLSIZE).AsString() or ""
    cleaned = clean_size_string(size_str)

    try:
        dia_in = float(cleaned)
    except ValueError:
        m = re.match(r'(?:(\d+)[-\s])?(\d+/\d+)', cleaned)
        if m:
            int_part, frac_part = m.groups()
            dia_in = float(Fraction(frac_part))
            if int_part:
                dia_in += float(int_part)
        else:
            dia_in = 0.5

    dia_ft = dia_in / 12.0
    inst.LookupParameter('Diameter').Set(get_diameter_from_size(dia_ft))

    inst.LookupParameter('Length').Set(sleeve_length)
    inst.LookupParameter('Schedule Level').Set(level.Id)

    mapping = {
        'FP_Product Entry': 'Overall Size',
        'FP_Service Name': 'Fabrication Service Name',
        'FP_Service Abbreviation': 'Fabrication Service Abbreviation'
    }
    for fam_p, pipe_p in mapping.items():
        try:
            val = get_parameter_value_by_name_AsString(pipe, pipe_p)
            cleaned_val = clean_size_string(val)
            set_parameter_by_name(inst, fam_p, cleaned_val if cleaned_val is not None else "")
        except:
            set_parameter_by_name(inst, fam_p, "")

    return inst

def get_upper_level(view, all_levels):
    gen_level = view.GenLevel
    if not gen_level:
        return None
    elev = gen_level.Elevation
    candidates = [lvl for lvl in all_levels if lvl.Elevation > elev]
    if not candidates:
        return None
    return min(candidates, key=lambda l: l.Elevation)

def main():
    symbol = load_family()
    if not symbol:
        TaskDialog.Show("Error", "Cannot load family symbol.\nPlease load manually:\n" + family_path)
        return

    selected_ids = uidoc.Selection.GetElementIds()
    all_levels = FilteredElementCollector(doc).OfClass(DB.Level).ToElements()
    existing_sleeves = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeAccessory)\
                                                     .WhereElementIsNotElementType()\
                                                     .ToElements()

    with Transaction(doc, 'Place Sleeves at Intersections') as t:
        t.Start()
        if not symbol.IsActive:
            symbol.Activate()
            doc.Regenerate()

        placed = 0

        is_3d = curview.ViewType == ViewType.ThreeD
        is_plan = curview.ViewType == ViewType.FloorPlan

        section_box = curview.GetSectionBox() if is_3d and curview.IsSectionBoxActive else None
        if section_box:
            sb_min = section_box.Transform.OfPoint(section_box.Min)
            sb_max = section_box.Transform.OfPoint(section_box.Max)

        # Mode selection
        if selected_ids.Count > 0 and is_plan:
            upper = get_upper_level(curview, all_levels)
            if not upper:
                TaskDialog.Show("Error", "No level above current floor plan view.")
                t.Commit()
                return
            levels_to_check = [upper]

        elif selected_ids.Count > 0 and is_3d:
            visible_pipes = FilteredElementCollector(doc, curview.Id)\
                            .OfCategory(BuiltInCategory.OST_FabricationPipework)\
                            .WhereElementIsNotElementType().ToElements()
            visible_ids = set([p.Id for p in visible_pipes])
            levels_to_check = all_levels

        elif is_3d:
            visible_pipes = FilteredElementCollector(doc, curview.Id)\
                            .OfCategory(BuiltInCategory.OST_FabricationPipework)\
                            .WhereElementIsNotElementType().ToElements()
            levels_to_check = all_levels

        elif is_plan:
            upper = get_upper_level(curview, all_levels)
            if not upper:
                TaskDialog.Show("Error", "No level above current view.")
                t.Commit()
                return
            levels_to_check = [upper]

        else:
            TaskDialog.Show("Error", "This script supports 3D and Floor Plan views only.")
            t.Commit()
            return

        # Actual processing
        pipes = []
        if selected_ids.Count > 0:
            for eid in selected_ids:
                el = doc.GetElement(eid)
                if el and el.Category and el.Category.Id.IntegerValue == int(BuiltInCategory.OST_FabricationPipework):
                    pipes.append(el)
        else:
            pipes = visible_pipes if 'visible_pipes' in locals() else \
                    FilteredElementCollector(doc, curview.Id)\
                    .OfCategory(BuiltInCategory.OST_FabricationPipework)\
                    .WhereElementIsNotElementType().ToElements()

        for pipe in pipes:
            if not is_vertical_pipe(pipe):
                continue
            for lvl in levels_to_check:
                pts = get_pipe_intersections(pipe, lvl)
                for pt in pts:
                    if section_box and not (sb_min.X <= pt.X <= sb_max.X and
                                            sb_min.Y <= pt.Y <= sb_max.Y and
                                            sb_min.Z <= pt.Z <= sb_max.Z):
                        continue
                    sleeve = place_sleeve_at_intersection(pipe, pt, symbol, lvl, existing_sleeves)
                    if sleeve:
                        placed += 1
                        existing_sleeves.Add(sleeve)

        t.Commit()
        TaskDialog.Show("Result", "Placed {} sleeve instances.".format(placed))

if __name__ == '__main__':
    main()