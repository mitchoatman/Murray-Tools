# -*- coding: utf-8 -*-
import os
import json
import clr
import System
from collections import defaultdict
from pyrevit import script, revit
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    BuiltInCategory,
    BuiltInParameter,
    FabricationPart
)

clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

from System.Windows import (
    Window, Thickness, WindowStyle, ResizeMode,
    WindowStartupLocation, HorizontalAlignment, VerticalAlignment, GridLength
)
from System.Windows.Controls import (
    Label, Button, Grid, RowDefinition, ColumnDefinition, GroupBox, StackPanel, Orientation, CheckBox
)

# ==============================================================================
# CONFIG FILE STORAGE & DEFAULTS
# ==============================================================================
CONFIG_DIR = r"C:\Temp"
CONFIG_FILE = os.path.join(CONFIG_DIR, "Ribbon_PTrapQC.json")

DEFAULT_SPECS = {
    "show_passing_details": False
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
                if "show_passing_details" not in data:
                    data["show_passing_details"] = False
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
class PTrapSpecForm(Window):
    def __init__(self, current_specs):
        self.Title = "P-Trap QC Setup"
        self.Width = 420
        self.Height = 180
        self.WindowStyle = WindowStyle.SingleBorderWindow
        self.ResizeMode = ResizeMode.NoResize
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.specs = current_specs

        root_grid = Grid()
        root_grid.Margin = Thickness(12)
        root_grid.RowDefinitions.Add(RowDefinition(Height=GridLength(1, System.Windows.GridUnitType.Star)))
        root_grid.RowDefinitions.Add(RowDefinition(Height=GridLength.Auto))

        opt_box = GroupBox()
        opt_box.Header = " Reporting Options "
        
        opt_panel = StackPanel()
        opt_panel.Margin = Thickness(10)

        self.chk_details = CheckBox()
        self.chk_details.Content = "List individual passing P-Trap line items below summary"
        self.chk_details.IsChecked = self.specs.get("show_passing_details", False)
        opt_panel.Children.Add(self.chk_details)

        opt_box.Content = opt_panel
        Grid.SetRow(opt_box, 0)
        root_grid.Children.Add(opt_box)

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
        Grid.SetRow(btn_panel, 1)
        root_grid.Children.Add(btn_panel)

        self.Content = root_grid

    def on_ok(self, sender, args):
        self.specs["show_passing_details"] = bool(self.chk_details.IsChecked)
        save_specs(self.specs)
        self.DialogResult = True
        self.Close()

    def on_cancel(self, sender, args):
        self.DialogResult = False
        self.Close()

# ==============================================================================
# CODE COMPLIANCE RULES & TRAVERSAL
# ==============================================================================
TRAP_ARM_LIMITS = {
    1.25: {"min": 2.5,  "max": 30.0},   # 2.5" min | 2'-6" max
    1.5:  {"min": 3.0,  "max": 42.0},   # 3.0" min | 3'-6" max
    2.0:  {"min": 4.0,  "max": 60.0},   # 4.0" min | 5'-0" max
    3.0:  {"min": 6.0,  "max": 72.0},   # 6.0" min | 6'-0" max
    4.0:  {"min": 8.0,  "max": 120.0}   # 8.0" min | 10'-0" max
}

MAX_TAILPIECE_LIMIT_INCHES = 24.0

doc = revit.doc
output = script.get_output()

def get_trap_arm_limits(size_in_inches):
    if size_in_inches <= 1.25:
        return TRAP_ARM_LIMITS[1.25]["min"], TRAP_ARM_LIMITS[1.25]["max"]
    elif size_in_inches <= 1.5:
        return TRAP_ARM_LIMITS[1.5]["min"], TRAP_ARM_LIMITS[1.5]["max"]
    elif size_in_inches <= 2.0:
        return TRAP_ARM_LIMITS[2.0]["min"], TRAP_ARM_LIMITS[2.0]["max"]
    elif size_in_inches <= 3.0:
        return TRAP_ARM_LIMITS[3.0]["min"], TRAP_ARM_LIMITS[3.0]["max"]
    else:
        min_in = 2.0 * size_in_inches
        return min_in, TRAP_ARM_LIMITS[4.0]["max"]

def is_strictly_straight_pipe(part):
    if part.ItemCustomId != 2041:
        return False
    param = part.get_Parameter(BuiltInParameter.FABRICATION_PART_LENGTH)
    if param and param.HasValue and param.AsDouble() > 0:
        return True
    return False

def get_fab_part_length_in_inches(pipe_part):
    param = pipe_part.get_Parameter(BuiltInParameter.FABRICATION_PART_LENGTH)
    if param and param.HasValue:
        return param.AsDouble() * 12.0

    p_length = pipe_part.LookupParameter("Length")
    if p_length and p_length.HasValue:
        return p_length.AsDouble() * 12.0

    conn_mgr = pipe_part.ConnectorManager
    if conn_mgr:
        conns = [c for c in conn_mgr.Connectors]
        if len(conns) >= 2:
            return conns[0].Origin.DistanceTo(conns[1].Origin) * 12.0

    return 0.0

def get_dynamic_weir_offset_in_inches(trap_part):
    dims = trap_part.GetDimensions()
    for dim_def in dims:
        name = dim_def.Name.lower()
        if "right radius" in name or "bottom radius" in name:
            val_ft = trap_part.GetDimensionValue(dim_def)
            if val_ft > 0:
                return val_ft * 12.0

    p_rr = trap_part.LookupParameter("Right Radius")
    if p_rr and p_rr.HasValue:
        return p_rr.AsDouble() * 12.0

    p_br = trap_part.LookupParameter("Bottom Radius")
    if p_br and p_br.HasValue:
        return p_br.AsDouble() * 12.0

    return 0.0

def get_connector_by_index(part, index):
    conn_mgr = part.ConnectorManager
    if not conn_mgr:
        return None

    for conn in conn_mgr.Connectors:
        if conn.Id == index:
            return conn

    conns = [c for c in conn_mgr.Connectors]
    if len(conns) >= index:
        return conns[index - 1]

    return None

def traverse_trap_arm_from_c1(trap):
    straight_pipes = []
    visited_ids = set([trap.Id])
    total_length_in = 0.0
    
    c1_conn = get_connector_by_index(trap, 1)
    if not c1_conn:
        return straight_pipes, total_length_in, 0.0

    pipe_size_in = c1_conn.Radius * 2.0 * 12.0

    if not c1_conn.IsConnected:
        return straight_pipes, total_length_in, pipe_size_in

    queue = []
    for ref_conn in c1_conn.AllRefs:
        owner = ref_conn.Owner
        if owner and owner.Id not in visited_ids and isinstance(owner, FabricationPart):
            queue.append((owner, ref_conn))

    while queue:
        current_part, entry_conn = queue.pop(0)
        if current_part.Id in visited_ids:
            continue
        visited_ids.add(current_part.Id)

        curr_conn_mgr = current_part.ConnectorManager
        if not curr_conn_mgr:
            continue

        if is_strictly_straight_pipe(current_part):
            len_in = get_fab_part_length_in_inches(current_part)
            if len_in > 0:
                straight_pipes.append((current_part, len_in))
                total_length_in += len_in
            break

        if curr_conn_mgr.Connectors.Size >= 2:
            for c in curr_conn_mgr.Connectors:
                if c.Id != entry_conn.Id and c.IsConnected:
                    for ref_conn in c.AllRefs:
                        next_owner = ref_conn.Owner
                        if next_owner and next_owner.Id not in visited_ids and isinstance(next_owner, FabricationPart):
                            queue.append((next_owner, ref_conn))

    return straight_pipes, total_length_in, pipe_size_in

def traverse_tailpiece_from_c2(trap):
    tailpiece_parts = []
    visited_ids = set([trap.Id])
    total_vertical_len_in = 0.0

    c2_conn = get_connector_by_index(trap, 2)
    if not c2_conn or not c2_conn.IsConnected:
        return tailpiece_parts, total_vertical_len_in

    queue = []
    for ref_conn in c2_conn.AllRefs:
        owner = ref_conn.Owner
        if owner and owner.Id not in visited_ids and isinstance(owner, FabricationPart):
            queue.append((owner, ref_conn))

    while queue:
        current_part, entry_conn = queue.pop(0)
        if current_part.Id in visited_ids:
            continue
        visited_ids.add(current_part.Id)

        curr_conn_mgr = current_part.ConnectorManager
        if not curr_conn_mgr:
            continue

        if is_strictly_straight_pipe(current_part):
            len_in = get_fab_part_length_in_inches(current_part)
            tailpiece_parts.append((current_part, len_in))
            total_vertical_len_in += len_in

        if curr_conn_mgr.Connectors.Size >= 2:
            for c in curr_conn_mgr.Connectors:
                if c.Id != entry_conn.Id and c.IsConnected:
                    for ref_conn in c.AllRefs:
                        next_owner = ref_conn.Owner
                        if next_owner and next_owner.Id not in visited_ids and isinstance(next_owner, FabricationPart):
                            queue.append((next_owner, ref_conn))

    return tailpiece_parts, total_vertical_len_in

# ==============================================================================
# AUDIT CORE
# ==============================================================================
def check_trap_arm_and_tailpiece(user_specs):
    output.set_width(1200)
    output.set_height(600)

    collector = FilteredElementCollector(doc, doc.ActiveView.Id) \
                .OfCategory(BuiltInCategory.OST_FabricationPipework) \
                .WhereElementIsNotElementType()

    traps = [elem for elem in collector if isinstance(elem, FabricationPart) and elem.ItemCustomId == 916]

    if not traps:
        output.print_md("### ⚠️ No CID 916 P-Traps found in the active view.")
        return

    output.print_md("# 🛠️ Plumbing QC Report: P-Trap Exceptions Audit (CPC / UPC)")
    output.print_md("---")

    failed_rows = []
    warning_rows = []
    passed_by_group = defaultdict(list)

    passed_count = 0
    failed_count = 0
    warning_count = 0

    show_details = user_specs.get("show_passing_details", False)

    for trap in traps:
        trap_id = output.linkify(trap.Id)
        straight_pipes_info, pipe_len_in, pipe_size_in = traverse_trap_arm_from_c1(trap)
        tailpiece_info, tailpiece_len_in = traverse_tailpiece_from_c2(trap)

        c1_pipe_ids = ", ".join([output.linkify(p[0].Id) for p in straight_pipes_info]) if straight_pipes_info else "None"
        c2_pipe_ids = ", ".join([output.linkify(p[0].Id) for p in tailpiece_info]) if tailpiece_info else "None"

        if straight_pipes_info:
            weir_offset_in = get_dynamic_weir_offset_in_inches(trap)
            true_trap_arm_in = pipe_len_in + weir_offset_in
            true_trap_arm_ft = true_trap_arm_in / 12.0
            
            min_allowed_arm_in, max_allowed_arm_in = get_trap_arm_limits(pipe_size_in)
            
            is_arm_max_ok = true_trap_arm_in <= max_allowed_arm_in
            is_arm_min_ok = true_trap_arm_in >= min_allowed_arm_in
            is_tailpiece_ok = tailpiece_len_in <= MAX_TAILPIECE_LIMIT_INCHES

            tailpiece_str = "{:.2f}\"".format(tailpiece_len_in) if tailpiece_len_in > 0 else "0.00\""

            row_data = [
                trap_id,
                c1_pipe_ids,
                c2_pipe_ids,
                "{:.2f}\"".format(pipe_size_in),
                "{:.2f}\" ({:.2f}')".format(true_trap_arm_in, true_trap_arm_ft),
                "{:.1f}\" to {:.1f}\"".format(min_allowed_arm_in, max_allowed_arm_in),
                tailpiece_str
            ]

            if is_arm_max_ok and is_arm_min_ok and is_tailpiece_ok:
                row_data.append("PASS ✅")
                passed_count += 1
                group_key = "{:.2f}\" P-Trap".format(pipe_size_in)
                passed_by_group[group_key].append(row_data)
            else:
                fail_reasons = []
                if not is_arm_min_ok:
                    fail_reasons.append("Arm < Min ({:.1f}\")".format(min_allowed_arm_in))
                if not is_arm_max_ok:
                    fail_reasons.append("Arm > Max")
                if not is_tailpiece_ok:
                    fail_reasons.append("Tailpiece > 24\"")
                
                row_data.append("FAIL ❌ ({})".format(", ".join(fail_reasons)))
                failed_rows.append(row_data)
                failed_count += 1
        else:
            warning_count += 1
            warning_rows.append([
                trap_id,
                "No C1 pipe",
                c2_pipe_ids,
                "{:.2f}\"".format(pipe_size_in) if pipe_size_in > 0 else "-",
                "-",
                "-",
                "-",
                "WARNING ⚠️ (Unconnected C1 Arm)"
            ])

    if failed_rows:
        output.print_md("### ❌ P-Trap QC Violations Found ({})".format(len(failed_rows)))
        output.print_table(
            table_data=failed_rows,
            columns=[
                "Trap Element ID", 
                "Pipe (C1 Arm)",
                "Pipe (C2 Tailpiece)",
                "Pipe Size", 
                "Trap Arm Length", 
                "Arm (Min - Max)", 
                "Tailpiece (C2)",
                "QC Status"
            ]
        )
        output.print_md("---")
    elif not warning_rows:
        output.print_md("### 🎉 Zero QC Violations Found!")
        output.print_md("All P-Traps meet trap arm min/max limits and tailpiece length requirements.")
        output.print_md("---")

    if warning_rows:
        output.print_md("### ⚠️ Warnings / Unconnected Traps ({})".format(len(warning_rows)))
        output.print_table(
            table_data=warning_rows,
            columns=[
                "Trap Element ID", 
                "Pipe (C1 Arm)",
                "Pipe (C2 Tailpiece)",
                "Pipe Size", 
                "Trap Arm Length", 
                "Arm (Min - Max)", 
                "Tailpiece (C2)",
                "QC Status"
            ]
        )
        output.print_md("---")

    if passed_by_group:
        output.print_md("### 📦 Rolled-Up Summary of Passing P-Traps")
        summary_table = []
        for grp_name, rows in sorted(passed_by_group.items()):
            group_size = rows[0][3]
            group_arm_limit = rows[0][5]
            summary_table.append([
                grp_name,
                group_size,
                group_arm_limit,
                "≤ 24.0\"",
                str(len(rows)),
                "PASS ✅"
            ])

        output.print_table(
            table_data=summary_table,
            columns=[
                "P-Trap Group", 
                "Pipe Size", 
                "Allowed Arm (Min - Max)", 
                "Max Tailpiece", 
                "Total Quantity", 
                "QC Status"
            ]
        )

        if show_details:
            output.print_md("---")
            output.print_md("### 📋 Detailed List of Passing P-Traps ({})".format(passed_count))
            all_passed_rows = []
            for rows in passed_by_group.values():
                all_passed_rows.extend(rows)

            output.print_table(
                table_data=all_passed_rows,
                columns=[
                    "Trap Element ID", 
                    "Pipe (C1 Arm)",
                    "Pipe (C2 Tailpiece)",
                    "Pipe Size", 
                    "Trap Arm Length", 
                    "Arm (Min - Max)", 
                    "Tailpiece (C2)",
                    "QC Status"
                ]
            )

    output.print_md("---")
    output.print_md("**Summary:** {} Passed | {} Failed | {} Warnings | Total Traps Checked: {}".format(
        passed_count, failed_count, warning_count, len(traps)
    ))

if __name__ == "__main__":
    current_specs = load_specs()
    form = PTrapSpecForm(current_specs)
    if form.ShowDialog() and form.DialogResult:
        check_trap_arm_and_tailpiece(form.specs)