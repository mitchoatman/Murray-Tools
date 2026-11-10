# -*- coding: utf-8 -*-
import os
import json
import clr
import System
import math
from collections import defaultdict
from pyrevit import script, revit
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    BuiltInCategory,
    BuiltInParameter,
    ElementId,
    FabricationPart,
    LocationPoint,
    XYZ
)

clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

from System.Windows import (
    Window, Thickness, WindowStyle, ResizeMode,
    WindowStartupLocation, HorizontalAlignment, VerticalAlignment, GridLength
)
from System.Windows.Controls import (
    Label, TextBox, Button, Grid, RowDefinition,
    ColumnDefinition, GroupBox, StackPanel, Orientation, CheckBox, ScrollViewer
)

# ==============================================================================
# CONFIG FILE STORAGE & DEFAULTS
# ==============================================================================
CONFIG_DIR = r"C:\Temp"
CONFIG_FILE = os.path.join(CONFIG_DIR, "Ribbon_HangerQC.json")

DEFAULT_SPECS = {
    "show_passing_details": False,
    "rod_specs": [
        {"pipe_size": 1.00,  "label": "1\" Pipe:",       "rod_size": 0.375},
        {"pipe_size": 1.25,  "label": "1-1/4\" Pipe:",   "rod_size": 0.375},
        {"pipe_size": 1.50,  "label": "1-1/2\" Pipe:",   "rod_size": 0.375},
        {"pipe_size": 2.00,  "label": "2\" Pipe:",       "rod_size": 0.375},
        {"pipe_size": 2.50,  "label": "2-1/2\" Pipe:",   "rod_size": 0.375},
        {"pipe_size": 3.00,  "label": "3\" Pipe:",       "rod_size": 0.375},
        {"pipe_size": 4.00,  "label": "4\" Pipe:",       "rod_size": 0.375},
        {"pipe_size": 5.00,  "label": "5\" Pipe:",       "rod_size": 0.500},
        {"pipe_size": 6.00,  "label": "6\" Pipe:",       "rod_size": 0.500},
        {"pipe_size": 8.00,  "label": "8\" Pipe:",       "rod_size": 0.500},
        {"pipe_size": 10.00, "label": "10\" Pipe:",      "rod_size": 0.625},
        {"pipe_size": 12.00, "label": "12\" Pipe:",      "rod_size": 0.625}
    ],
    "spacing_specs": {
        "cast_iron": 10.0,       # Cast Iron Default: 10.0'
        "copper_small": 6.0,     # Copper ≤ 1.5": 6.0'
        "copper_large": 10.0,    # Copper ≥ 2": 10.0'
        "steel_small": 10.0,     # Steel Water/DWV ≤ 3/4": 10.0'
        "steel_large": 12.0,     # Steel Water/DWV ≥ 1": 12.0'
        "pvc_abs": 4.0,          # Sch 40 PVC & ABS: 4.0'
        "cpvc_small": 3.0,       # CPVC ≤ 1": 3.0'
        "cpvc_large": 4.0        # CPVC ≥ 1.25": 4.0'
    }
}

def load_specs():
    if not os.path.exists(CONFIG_DIR):
        try:
            os.makedirs(CONFIG_DIR)
        except:
            pass
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r') as f:
                data = json.load(f)
                for k, v in DEFAULT_SPECS["spacing_specs"].items():
                    if k not in data.get("spacing_specs", {}):
                        data["spacing_specs"][k] = v
                if "show_passing_details" not in data:
                    data["show_passing_details"] = False
                if "rod_specs" in data and len(data["rod_specs"]) != len(DEFAULT_SPECS["rod_specs"]):
                    data["rod_specs"] = DEFAULT_SPECS["rod_specs"]
                return data
        except:
            pass
    return DEFAULT_SPECS

def save_specs(specs):
    try:
        with open(CONFIG_FILE, 'w') as f:
            json.dump(specs, f, indent=4)
    except Exception as e:
        print("Could not save config file: {}".format(e))

# ==============================================================================
# WPF SPECIFICATION CONFIGURATION FORM
# ==============================================================================
class HangerSpecForm(Window):
    def __init__(self, current_specs):
        self.Title = "Hanger QC Specification Setup"
        self.Width = 510
        self.Height = 640
        self.WindowStyle = WindowStyle.SingleBorderWindow
        self.ResizeMode = ResizeMode.NoResize
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.specs = current_specs

        root_grid = Grid()
        root_grid.Margin = Thickness(12)
        root_grid.RowDefinitions.Add(RowDefinition(Height=GridLength(1, System.Windows.GridUnitType.Star)))
        root_grid.RowDefinitions.Add(RowDefinition(Height=GridLength.Auto))
        root_grid.RowDefinitions.Add(RowDefinition(Height=GridLength.Auto))

        scroll = ScrollViewer()
        scroll.VerticalScrollBarVisibility = System.Windows.Controls.ScrollBarVisibility.Auto
        scroll_content = StackPanel()

        rod_box = GroupBox()
        rod_box.Header = " Min Rod Size (Inches) per Individual Pipe Size "
        
        rod_grid = Grid()
        rod_grid.Margin = Thickness(8)
        col1 = ColumnDefinition()
        col1.Width = GridLength(1, System.Windows.GridUnitType.Star)
        col2 = ColumnDefinition()
        col2.Width = GridLength.Auto
        rod_grid.ColumnDefinitions.Add(col1)
        rod_grid.ColumnDefinitions.Add(col2)

        self.rod_inputs = []
        for i, rod_item in enumerate(self.specs["rod_specs"]):
            rod_grid.RowDefinitions.Add(RowDefinition(Height=GridLength.Auto))

            lbl = Label()
            lbl.Content = rod_item.get("label", "{}\" Pipe:".format(rod_item["pipe_size"]))
            lbl.Margin = Thickness(0, 1, 0, 1)
            lbl.VerticalAlignment = VerticalAlignment.Center
            Grid.SetRow(lbl, i)
            Grid.SetColumn(lbl, 0)
            rod_grid.Children.Add(lbl)

            tb = TextBox()
            tb.Text = str(rod_item["rod_size"])
            tb.Width = 80
            tb.Height = 22
            tb.Margin = Thickness(0, 1, 0, 1)
            tb.HorizontalAlignment = HorizontalAlignment.Right
            tb.VerticalContentAlignment = VerticalAlignment.Center
            Grid.SetRow(tb, i)
            Grid.SetColumn(tb, 1)
            rod_grid.Children.Add(tb)
            self.rod_inputs.append((rod_item["pipe_size"], rod_item.get("label", ""), tb))

        rod_box.Content = rod_grid
        scroll_content.Children.Add(rod_box)

        space_box = GroupBox()
        space_box.Header = " Max Horizontal Hanger Spacing (Feet) "
        space_box.Margin = Thickness(0, 10, 0, 0)

        space_grid = Grid()
        space_grid.Margin = Thickness(8)
        
        col1_s = ColumnDefinition()
        col1_s.Width = GridLength(1, System.Windows.GridUnitType.Star)
        col2_s = ColumnDefinition()
        col2_s.Width = GridLength.Auto
        space_grid.ColumnDefinitions.Add(col1_s)
        space_grid.ColumnDefinitions.Add(col2_s)

        spacing_fields = [
            ("Cast Iron / Hubless (All Sizes):", "cast_iron"),
            ("Copper & Copper Alloys (≤ 1.5\"): ", "copper_small"),
            ("Copper & Copper Alloys (≥ 2\"): ", "copper_large"),
            ("Steel Pipe Water/DWV (≤ 3/4\"): ", "steel_small"),
            ("Steel Pipe Water/DWV (≥ 1\"): ", "steel_large"),
            ("Sch 40 PVC & ABS DWV (All Sizes): ", "pvc_abs"),
            ("CPVC (≤ 1\"): ", "cpvc_small"),
            ("CPVC (≥ 1.25\"): ", "cpvc_large"),
        ]
        self.space_inputs = {}
        for i, (label_text, key) in enumerate(spacing_fields):
            space_grid.RowDefinitions.Add(RowDefinition(Height=GridLength.Auto))

            lbl = Label()
            lbl.Content = label_text
            lbl.Margin = Thickness(0, 1, 0, 1)
            lbl.VerticalAlignment = VerticalAlignment.Center
            Grid.SetRow(lbl, i)
            Grid.SetColumn(lbl, 0)
            space_grid.Children.Add(lbl)

            tb = TextBox()
            tb.Text = str(self.specs["spacing_specs"].get(key, 10.0))
            tb.Width = 80
            tb.Height = 22
            tb.Margin = Thickness(0, 1, 0, 1)
            tb.HorizontalAlignment = HorizontalAlignment.Right
            tb.VerticalContentAlignment = VerticalAlignment.Center
            Grid.SetRow(tb, i)
            Grid.SetColumn(tb, 1)
            space_grid.Children.Add(tb)
            self.space_inputs[key] = tb

        space_box.Content = space_grid
        scroll_content.Children.Add(space_box)

        scroll.Content = scroll_content
        Grid.SetRow(scroll, 0)
        root_grid.Children.Add(scroll)

        self.chk_details = CheckBox()
        self.chk_details.Content = "List individual passing hanger line items below summary"
        self.chk_details.IsChecked = self.specs.get("show_passing_details", False)
        self.chk_details.Margin = Thickness(5, 10, 0, 0)
        Grid.SetRow(self.chk_details, 1)
        root_grid.Children.Add(self.chk_details)

        btn_panel = StackPanel()
        btn_panel.Orientation = Orientation.Horizontal
        btn_panel.HorizontalAlignment = HorizontalAlignment.Right
        btn_panel.Margin = Thickness(0, 12, 0, 0)

        ok_btn = Button()
        ok_btn.Content = "Run Audit"
        ok_btn.Width = 90
        ok_btn.Height = 26
        ok_btn.Margin = Thickness(0, 0, 10, 0)
        ok_btn.Click += self.on_ok

        cancel_btn = Button()
        cancel_btn.Content = "Cancel"
        cancel_btn.Width = 75
        cancel_btn.Height = 26
        cancel_btn.Click += self.on_cancel

        btn_panel.Children.Add(ok_btn)
        btn_panel.Children.Add(cancel_btn)
        Grid.SetRow(btn_panel, 2)
        root_grid.Children.Add(btn_panel)

        self.Content = root_grid

    def on_ok(self, sender, args):
        try:
            new_rod_specs = []
            for pipe_s, lbl, tb in self.rod_inputs:
                new_rod_specs.append({
                    "pipe_size": pipe_s,
                    "label": lbl,
                    "rod_size": float(tb.Text)
                })
            self.specs["rod_specs"] = new_rod_specs

            for key, tb in self.space_inputs.items():
                self.specs["spacing_specs"][key] = float(tb.Text)

            self.specs["show_passing_details"] = bool(self.chk_details.IsChecked)

            save_specs(self.specs)
            self.DialogResult = True
            self.Close()
        except Exception as ex:
            print("Invalid input value: {}".format(ex))

    def on_cancel(self, sender, args):
        self.DialogResult = False
        self.Close()

# ==============================================================================
# AUDIT CORE & HELPER FUNCTIONS
# ==============================================================================
doc = revit.doc
output = script.get_output()

def get_required_rod_size(pipe_size_in, rod_specs):
    sorted_specs = sorted(rod_specs, key=lambda x: x["pipe_size"])
    for item in sorted_specs:
        if pipe_size_in <= item["pipe_size"]:
            return item["rod_size"]
    return sorted_specs[-1]["rod_size"] if sorted_specs else 0.625

def get_max_allowed_spacing(material_name, pipe_size_in, spacing_specs):
    name_lower = material_name.lower() if material_name else ""
    
    if "cast iron" in name_lower or "cast-iron" in name_lower:
        return spacing_specs.get("cast_iron", 10.0)
    elif "copper" in name_lower:
        if pipe_size_in <= 1.5:
            return spacing_specs.get("copper_small", 6.0)
        else:
            return spacing_specs.get("copper_large", 10.0)
    elif "pvc" in name_lower or "abs" in name_lower:
        if "cpvc" in name_lower:
            if pipe_size_in <= 1.0:
                return spacing_specs.get("cpvc_small", 3.0)
            else:
                return spacing_specs.get("cpvc_large", 4.0)
        else:
            return spacing_specs.get("pvc_abs", 4.0)
    else:
        if pipe_size_in <= 0.75:
            return spacing_specs.get("steel_small", 10.0)
        else:
            return spacing_specs.get("steel_large", 12.0)

def parse_pipe_size(size_str):
    if not size_str:
        return 0.0
    cleaned = str(size_str).replace('"', '').strip()
    try:
        if '/' in cleaned:
            parts = cleaned.split()
            if len(parts) == 2:
                num, denom = parts[1].split('/')
                return float(parts[0]) + (float(num) / float(denom))
            else:
                num, denom = cleaned.split('/')
                return float(num) / float(denom)
        return float(cleaned)
    except:
        return 0.0

def get_host_pipe_material(host_elem):
    if not host_elem:
        return "Steel"

    p_mat = host_elem.get_Parameter(BuiltInParameter.FABRICATION_PART_MATERIAL)
    if p_mat and p_mat.HasValue:
        val_str = p_mat.AsValueString() or p_mat.AsString()
        if val_str:
            return val_str.split(':')[0].strip()

    p_pm = host_elem.LookupParameter("Part Material")
    if p_pm and p_pm.HasValue:
        val_str = p_pm.AsValueString() or p_pm.AsString()
        if val_str:
            return val_str.split(':')[0].strip()

    if hasattr(host_elem, "MaterialName") and host_elem.MaterialName:
        return host_elem.MaterialName.split(':')[0].strip()

    return "Steel"

def get_actual_rod_diameter(hanger):
    try:
        for anc in hanger.GetPartAncillaryUsage():
            if anc.AncillaryWidthOrDiameter > 0:
                return anc.AncillaryWidthOrDiameter * 12.0
    except:
        pass

    p_rod = hanger.LookupParameter('FP_Rod Size')
    if p_rod and p_rod.HasValue:
        val = p_rod.AsDouble()
        if val > 0:
            return val * 12.0
        val_str = p_rod.AsString() or p_rod.AsValueString()
        if val_str:
            return parse_pipe_size(val_str)

    return 0.0

def is_trapeze(hanger):
    p = hanger.LookupParameter("Family")
    if p and p.HasValue:
        fam = (p.AsValueString() or p.AsString() or "").lower()
        return "trapeze" in fam
    return False

def get_hanger_location(hanger):
    loc = hanger.Location
    if isinstance(loc, LocationPoint):
        return loc.Point
    elif hasattr(loc, "Point"):
        return loc.Point

    conn_mgr = hanger.ConnectorManager
    if conn_mgr and conn_mgr.Connectors.Size > 0:
        for c in conn_mgr.Connectors:
            return c.Origin

    return None

def planar_distance(pt1, pt2):
    return math.sqrt((pt2.X - pt1.X)**2 + (pt2.Y - pt1.Y)**2)

def is_pipe(element):
    try:
        p = element.LookupParameter('Part Pattern Number')
        if p and p.HasValue:
            return p.AsInteger() == 2041
    except:
        pass
    return False

def vertical_fab(element):
    """Calculates if pipe vector angle from horizontal exceeds 22.5 degrees (Riser Check)."""
    try:
        pts = [c.Origin for c in element.ConnectorManager.Connectors]
        if len(pts) >= 2:
            v = pts[1].Subtract(pts[0])
            if v.GetLength() < 0.0001:
                return False
            v = v.Normalize()
            angle_from_horizontal = math.asin(abs(v.Z))
            threshold = math.radians(22.5)
            return angle_from_horizontal > threshold
    except:
        pass
    return False

def get_pipe_centerline_elevation(pipe_elem):
    try:
        conn_mgr = pipe_elem.ConnectorManager
        if conn_mgr and conn_mgr.Connectors.Size >= 2:
            conns = [c for c in conn_mgr.Connectors]
            return (conns[0].Origin.Z + conns[1].Origin.Z) / 2.0
    except:
        pass
    loc = pipe_elem.Location
    if isinstance(loc, LocationPoint):
        return loc.Point.Z
    elif hasattr(loc, "Point"):
        return loc.Point.Z
    return 0.0

def is_pipe_supported_by_trapeze(pipe_elem, trapeze_boxes):
    """Checks if the pipe's centerline intersects any trapeze bounding box."""
    try:
        conn_mgr = pipe_elem.ConnectorManager
        if not conn_mgr or conn_mgr.Connectors.Size < 2:
            return False
        conns = [c.Origin for c in conn_mgr.Connectors]
        p1, p2 = conns[0], conns[1]

        for bbox in trapeze_boxes:
            # Expand bounding box height slightly (± 0.5 ft) to account for pipe centerline elevation
            min_z = bbox.Min.Z - 0.5
            max_z = bbox.Max.Z + 0.5

            # Quick Z-check
            if (p1.Z < min_z and p2.Z < min_z) or (p1.Z > max_z and p2.Z > max_z):
                continue

            # Check 2D horizontal box intersection
            min_x, max_x = min(bbox.Min.X, bbox.Max.X), max(bbox.Min.X, bbox.Max.X)
            min_y, max_y = min(bbox.Min.Y, bbox.Max.Y), max(bbox.Min.Y, bbox.Max.Y)

            # Check if pipe centerline endpoints or segment pass through box XY limits
            pipe_min_x, pipe_max_x = min(p1.X, p2.X), max(p1.X, p2.X)
            pipe_min_y, pipe_max_y = min(p1.Y, p2.Y), max(p1.Y, p2.Y)

            if (pipe_max_x >= min_x and pipe_min_x <= max_x) and \
               (pipe_max_y >= min_y and pipe_min_y <= max_y):
                return True
    except:
        pass
    return False

# ==============================================================================
# MAIN AUDIT LOGIC
# ==============================================================================
def audit_hangers_and_spacing(user_specs):
    # Collector 1: Fabrication Hangers
    hanger_collector = FilteredElementCollector(doc, doc.ActiveView.Id) \
                        .OfCategory(BuiltInCategory.OST_FabricationHangers) \
                        .WhereElementIsNotElementType()

    hangers = [h for h in hanger_collector if isinstance(h, FabricationPart)]

    # Collector 2: Fabrication Pipe Parts (Horizontal only)
    pipe_collector = FilteredElementCollector(doc, doc.ActiveView.Id) \
                      .OfCategory(BuiltInCategory.OST_FabricationPipework) \
                      .WhereElementIsNotElementType()

    all_pipes = [p for p in pipe_collector if isinstance(p, FabricationPart) and is_pipe(p)]

    output.print_md("# 🛠️ Hanger QC Report: Exceptions-Only Audit")
    output.print_md("---")

    unhosted_hangers = []
    failed_rows = []
    passed_by_group = defaultdict(list)

    rod_fail_count = 0
    space_fail_count = 0
    passed_count = 0

    rod_specs = user_specs["rod_specs"]
    spacing_specs = user_specs["spacing_specs"]
    show_details = user_specs.get("show_passing_details", False)

    # 1. Separate Trapeze Hangers and Collect Bounding Boxes
    trapeze_boxes = []
    host_to_hangers = defaultdict(list)
    hosted_pipe_ids = set()

    for hanger in hangers:
        if is_trapeze(hanger):
            bbox = hanger.get_BoundingBox(doc.ActiveView)
            if bbox:
                trapeze_boxes.append(bbox)
            continue

        hosted_info = hanger.GetHostedInfo()
        host_id = hosted_info.HostId if hosted_info else ElementId.InvalidElementId

        if not host_id or host_id == ElementId.InvalidElementId:
            unhosted_hangers.append(hanger)
            continue

        host_elem = doc.GetElement(host_id)
        if not host_elem or not isinstance(host_elem, FabricationPart):
            unhosted_hangers.append(hanger)
            continue

        h_loc = get_hanger_location(hanger)
        if not h_loc:
            unhosted_hangers.append(hanger)
            continue

        hosted_pipe_ids.add(host_id)

        size_param = host_elem.LookupParameter("Size")
        size_str = size_param.AsString() if size_param else ""
        pipe_size_in = parse_pipe_size(size_str)
        material_str = get_host_pipe_material(host_elem)

        actual_rod_in = round(get_actual_rod_diameter(hanger), 3)
        req_rod_in = get_required_rod_size(pipe_size_in, rod_specs)
        
        is_rod_ok = (actual_rod_in + 0.001) >= req_rod_in
        rod_status = "OK ✅" if is_rod_ok else "FAIL ❌ (Undersized)"

        if not is_rod_ok:
            rod_fail_count += 1

        host_to_hangers[host_id].append({
            "hanger": hanger,
            "host_id": host_id,
            "host_elem": host_elem,
            "loc": h_loc,
            "size": pipe_size_in,
            "material": material_str,
            "actual_rod": actual_rod_in,
            "req_rod": req_rod_in,
            "rod_status": rod_status,
            "is_rod_ok": is_rod_ok
        })

    # 2. Audit Missing Hangers (>10ft long AND >5ft elevation AND 0 hangers AND HORIZONTAL ONLY)
    missing_hanger_pipes = []
    for pipe_elem in all_pipes:
        # Check horizontal orientation using vertical_fab()
        if vertical_fab(pipe_elem):
            continue

        p_length = pipe_elem.CenterlineLength
        p_elev = get_pipe_centerline_elevation(pipe_elem)

        if p_length > 10.0 and p_elev > 5.0:
            if pipe_elem.Id not in hosted_pipe_ids:
                # If pipe isn't explicitly hosted by a single hanger, check trapeze support geometry
                if is_pipe_supported_by_trapeze(pipe_elem, trapeze_boxes):
                    continue

                size_param = pipe_elem.LookupParameter("Size")
                size_str = size_param.AsString() if size_param else ""
                pipe_size_in = parse_pipe_size(size_str)
                mat_str = get_host_pipe_material(pipe_elem)
                
                missing_hanger_pipes.append([
                    output.linkify(pipe_elem.Id),
                    "{:.2f}\"".format(pipe_size_in),
                    mat_str,
                    "{:.2f}'".format(p_length),
                    "{:.2f}'".format(p_elev),
                    "FAIL ❌ (Missing Hanger)"
                ])

    # 3. Process Spacing per Host Pipe Segment
    for host_id, hanger_list in host_to_hangers.items():
        host_pipe = doc.GetElement(host_id)
        
        conn_mgr = host_pipe.ConnectorManager
        if conn_mgr and conn_mgr.Connectors.Size >= 2:
            conns = [c for c in conn_mgr.Connectors]
            p1 = conns[0].Origin
            p2 = conns[1].Origin
            dir_vec = (p2 - p1).Normalize()
            hanger_list.sort(key=lambda item: item["loc"].DotProduct(dir_vec))

        for i in range(len(hanger_list)):
            item = hanger_list[i]
            hanger_id = output.linkify(item["hanger"].Id)
            host_pipe_id = output.linkify(host_id)

            pipe_size_in = item["size"]
            material_str = item["material"]
            max_allowed_space = get_max_allowed_spacing(material_str, pipe_size_in, spacing_specs)

            if len(hanger_list) > 1 and i > 0:
                prev_item = hanger_list[i - 1]
                spacing_ft = planar_distance(item["loc"], prev_item["loc"])
                is_spacing_ok = spacing_ft <= (max_allowed_space + 0.15)
                spacing_str = "{:.2f}'".format(spacing_ft)
            else:
                spacing_str = "N/A (Single Hanger)"
                is_spacing_ok = True

            if not is_spacing_ok:
                space_fail_count += 1
                spacing_status = "FAIL ❌ (Over Max)"
            else:
                spacing_status = "PASS ✅"

            row_data = [
                hanger_id,
                host_pipe_id,
                "{:.2f}\"".format(pipe_size_in),
                material_str,
                "{:.3f}\" / {:.3f}\"".format(item["actual_rod"], item["req_rod"]),
                item["rod_status"],
                spacing_str,
                "{:.1f}'".format(max_allowed_space),
                spacing_status
            ]

            if not item["is_rod_ok"] or not is_spacing_ok:
                failed_rows.append(row_data)
            else:
                passed_count += 1
                group_key = "{:.2f}\" {}".format(pipe_size_in, material_str)
                passed_by_group[group_key].append(row_data)

    # OUTPUT SECTION 1: Missing Hangers Flagged
    if missing_hanger_pipes:
        output.print_md("### ⚠️ Horizontal Pipes Missing Hangers (>10' Length & >5' Elevation) ({})".format(len(missing_hanger_pipes)))
        output.print_table(
            table_data=missing_hanger_pipes,
            columns=[
                "Pipe Element ID",
                "Pipe Size",
                "Part Material",
                "Pipe Length",
                "Centerline Elev",
                "QC Flag"
            ]
        )
        output.print_md("---")

    # OUTPUT SECTION 2: Unhosted Hangers
    if unhosted_hangers:
        output.print_md("### ❌ Unhosted / Invalid Hangers ({})".format(len(unhosted_hangers)))
        for h in unhosted_hangers:
            output.print_md("* Unhosted Hanger ID: {}".format(output.linkify(h.Id)))
        output.print_md("---")

    # OUTPUT SECTION 3: Failed Hangers Table
    if failed_rows:
        output.print_md("### ❌ QC Violations Found ({})".format(len(failed_rows)))
        output.print_table(
            table_data=failed_rows,
            columns=[
                "Hanger ID", 
                "Host Pipe ID", 
                "Pipe Size", 
                "Part Material", 
                "Rod Size (Act / Req)", 
                "Rod Status", 
                "Hanger Spacing", 
                "Max Allowed", 
                "Spacing Status"
            ]
        )
        output.print_md("---")
    elif not missing_hanger_pipes:
        output.print_md("### 🎉 Zero QC Violations Found!")
        output.print_md("All hosted hangers meet rod size and max spacing requirements.")
        output.print_md("---")

    # OUTPUT SECTION 4: Rolled-Up Summary of Passing Items
    if passed_by_group:
        output.print_md("### 📦 Rolled-Up Summary of Passing Runs")
        summary_table = []
        for grp_name, rows in sorted(passed_by_group.items()):
            group_size = rows[0][2]
            group_mat = rows[0][3]
            group_rod = rows[0][4]
            group_max_space = rows[0][7]
            summary_table.append([
                grp_name,
                group_size,
                group_mat,
                group_rod,
                str(len(rows)),
                group_max_space,
                "PASS ✅"
            ])

        output.print_table(
            table_data=summary_table,
            columns=[
                "Pipe Group", 
                "Pipe Size", 
                "Part Material", 
                "Rod Size (Act / Req)", 
                "Hanger Count", 
                "Max Allowed", 
                "QC Status"
            ]
        )

        if show_details:
            output.print_md("---")
            output.print_md("### 📋 Detailed List of Passing Hangers ({})".format(passed_count))
            all_passed_rows = []
            for rows in passed_by_group.values():
                all_passed_rows.extend(rows)

            output.print_table(
                table_data=all_passed_rows,
                columns=[
                    "Hanger ID", 
                    "Host Pipe ID", 
                    "Pipe Size", 
                    "Part Material", 
                    "Rod Size (Act / Req)", 
                    "Rod Status", 
                    "Hanger Spacing", 
                    "Max Allowed", 
                    "Spacing Status"
                ]
            )

    output.print_md("---")
    output.print_md("**Summary:** {} Passed | {} Missing Hanger Pipes | {} Undersized Rods | {} Over-Spaced Runs | Total Checked: {}".format(
        passed_count, len(missing_hanger_pipes), rod_fail_count, space_fail_count, len(hangers)
    ))

if __name__ == "__main__":
    current_specs = load_specs()
    form = HangerSpecForm(current_specs)
    if form.ShowDialog() and form.DialogResult:
        audit_hangers_and_spacing(form.specs)