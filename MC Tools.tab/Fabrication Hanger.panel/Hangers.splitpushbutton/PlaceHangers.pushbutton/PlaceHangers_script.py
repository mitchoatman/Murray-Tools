import Autodesk
from Autodesk.Revit.DB import Transaction, FabricationConfiguration, FabricationPart, XYZ
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
import math
import os
import clr
clr.AddReference("PresentationCore")
clr.AddReference("PresentationFramework")
clr.AddReference("WindowsBase")
from System.Windows import Window, Thickness, WindowStartupLocation, ResizeMode, HorizontalAlignment
from System.Windows.Controls import StackPanel, Label, ComboBox, TextBox, CheckBox, Button, Orientation
from System.Windows.Media import FontFamily
from System import Array

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = doc.Application
RevitVersion = app.VersionNumber
RevitINT = float(RevitVersion)

DIRECTION_DOT_THRESHOLD = 0.999
MARGIN = 0.01

class FabricationPartSelectionFilter(ISelectionFilter):
    def AllowElement(self, element):
        return isinstance(element, FabricationPart)
    def AllowReference(self, reference, point):
        return False

class CustomISelectionFilter(ISelectionFilter):
    def __init__(self, nom_categorie):
        self.nom_categorie = nom_categorie
    def AllowElement(self, e):
        return e.LookupParameter('Fabrication Service').AsValueString() == self.nom_categorie
    def AllowReference(self, ref, point):
        return True

def is_cid_2875(element):
    try:
        return element.ItemCustomId in (2875, 875)
    except:
        return False

def find_nearest_connector(element, pick_point):
    nearest = None
    min_dist = float('inf')
    for c in element.ConnectorManager.Connectors:
        d = pick_point.DistanceTo(c.Origin)
        if d < min_dist:
            min_dist = d
            nearest = c
    return nearest

def vertical_fab(element):
    pts = [c.Origin for c in element.ConnectorManager.Connectors]
    if len(pts) >= 2:
        v = pts[1].Subtract(pts[0])
        if v.GetLength() < 0.0001:
            return False

        v = v.Normalize()

        # angle from horizontal = arcsin(|Z|)
        angle_from_horizontal = math.asin(abs(v.Z))

        # 30 degrees in radians
        threshold = math.radians(22.5)

        return angle_from_horizontal > threshold

    return False

def is_pipe(element):
    return element.LookupParameter('Part Pattern Number').AsInteger() in (2041, 866, 40)

def get_pipe_direction(entry_xyz, exit_xyz):
    v = exit_xyz.Subtract(entry_xyz)
    if v.GetLength() < 0.0001:
        return None
    return v.Normalize()

# ---------------------------------------------------------------------------
# Walk selected elements starting from start_element/start_connector.
# At each step advance through the exit connector of the current element to
# find the next adjacent selected element. Returns an ordered list of all
# selected elements reachable from the start, plus a dict of entry connectors.
# Elements NOT reachable are returned as leftovers (branches).
# ---------------------------------------------------------------------------
def walk_chain(selected_elements, start_element, start_connector):
    selected_ids = {e.Id: e for e in selected_elements}
    ordered = [start_element]
    entry_conns = {start_element.Id: start_connector}
    visited = {start_element.Id}

    all_start_conns = list(start_element.ConnectorManager.Connectors)
    exit_conns_of_start = [c for c in all_start_conns if c.Id != start_connector.Id]
    if not exit_conns_of_start:
        leftovers = [e for e in selected_elements if e.Id not in visited]
        return ordered, entry_conns, leftovers

    current_exit_conn = exit_conns_of_start[0]

    last_pipe_dir = None
    if is_pipe(start_element) and not vertical_fab(start_element) and not is_cid_2875(start_element):
        last_pipe_dir = get_pipe_direction(start_connector.Origin, current_exit_conn.Origin)

    while True:
        candidates = []
        for eid, e in selected_ids.items():
            if eid in visited:
                continue
            for c in e.ConnectorManager.Connectors:
                if current_exit_conn.Origin.DistanceTo(c.Origin) < 0.1:
                    candidates.append((e, c))
                    break

        if not candidates:
            break

        found_elem = candidates[0][0]
        found_entry = candidates[0][1]

        if len(candidates) > 1 and last_pipe_dir is not None:
            best_dot = -2.0
            for cand_elem, cand_entry in candidates:
                other_conns = [c for c in cand_elem.ConnectorManager.Connectors
                               if c.Id != cand_entry.Id]
                if other_conns:
                    d = get_pipe_direction(cand_entry.Origin, other_conns[0].Origin)
                    if d is not None:
                        dot = last_pipe_dir.DotProduct(d)
                        if dot > best_dot:
                            best_dot = dot
                            found_elem = cand_elem
                            found_entry = cand_entry

        ordered.append(found_elem)
        entry_conns[found_elem.Id] = found_entry
        visited.add(found_elem.Id)

        all_conns = list(found_elem.ConnectorManager.Connectors)
        exits = [c for c in all_conns if c.Id != found_entry.Id]
        if not exits:
            break

        if is_pipe(found_elem) and not vertical_fab(found_elem) and not is_cid_2875(found_elem):
            last_pipe_dir = get_pipe_direction(found_entry.Origin, exits[0].Origin)

        if is_cid_2875(found_elem):
            current_exit_conn = exits[0]

        elif len(exits) == 1:
            current_exit_conn = exits[0]

        else:
            best_exit = exits[0]
            if last_pipe_dir is not None:
                best_dot = -2.0
                for ex in exits:
                    d = get_pipe_direction(current_exit_conn.Origin, ex.Origin)
                    if d is not None:
                        dot = last_pipe_dir.DotProduct(d)
                        if dot > best_dot:
                            best_dot = dot
                            best_exit = ex
            current_exit_conn = best_exit

    leftovers = [e for e in selected_elements if e.Id not in visited]
    return ordered, entry_conns, leftovers

# ---------------------------------------------------------------------------
# From an ordered chain of elements + entry connectors, extract only the
# horizontal pipes and split into direction-based segments.
# Returns list of segments, each segment is a list of pipe dicts.
# ---------------------------------------------------------------------------
def chain_to_segments(ordered_chain, entry_conns):
    pipe_dicts = []
    for e in ordered_chain:
        if not is_pipe(e) or vertical_fab(e):
            continue
        entry_conn = entry_conns.get(e.Id)
        if entry_conn is None:
            entry_conn = next(iter(e.ConnectorManager.Connectors), None)
        if entry_conn is None:
            continue
        exit_conn = None
        for c in e.ConnectorManager.Connectors:
            if c.Id != entry_conn.Id:
                exit_conn = c
                break
        if exit_conn is None:
            continue
        direction = get_pipe_direction(entry_conn.Origin, exit_conn.Origin)
        pipe_dicts.append({
            'element':    e,
            'length':     e.CenterlineLength,
            'entry_xyz':  entry_conn.Origin,
            'exit_xyz':   exit_conn.Origin,
            'entry_conn': entry_conn,
            'direction':  direction,
        })

    if not pipe_dicts:
        return []

    segments = []
    current_seg = [pipe_dicts[0]]
    for i in range(1, len(pipe_dicts)):
        prev_dir = pipe_dicts[i-1]['direction']
        curr_dir = pipe_dicts[i]['direction']
        if prev_dir is not None and curr_dir is not None:
            dot = prev_dir.DotProduct(curr_dir)
        else:
            dot = 1.0
        if dot < DIRECTION_DOT_THRESHOLD:
            segments.append(current_seg)
            current_seg = [pipe_dicts[i]]
        else:
            current_seg.append(pipe_dicts[i])
    segments.append(current_seg)
    return segments

# ---------------------------------------------------------------------------
# Place hangers on one segment (list of pipe dicts, all same direction).
# force_end_hanger: True only for the last segment of a run.
# ---------------------------------------------------------------------------
def place_segment(pipe_list, fab_btn, distancefromend, spacing, atos,
                  force_end_hanger, label, debug_lines):

    def walk(start_idx, start_xyz, distance):
        idx = start_idx
        remaining = distance
        cur_xyz = start_xyz
        while idx < len(pipe_list):
            pd = pipe_list[idx]
            dist_to_exit = cur_xyz.DistanceTo(pd['exit_xyz'])
            if remaining <= dist_to_exit + MARGIN:
                direction = pd['exit_xyz'].Subtract(cur_xyz).Normalize()
                landing = cur_xyz.Add(direction.Multiply(remaining))
                local = pd['entry_xyz'].DistanceTo(landing)
                return (idx, landing, local)
            remaining -= dist_to_exit
            next_idx = idx + 1
            if next_idx >= len(pipe_list):
                return None
            gap = pd['exit_xyz'].DistanceTo(pipe_list[next_idx]['entry_xyz'])
            remaining -= gap
            if remaining < 0:
                remaining = 0.0
            cur_xyz = pipe_list[next_idx]['entry_xyz']
            idx = next_idx
        return None

    first = pipe_list[0]
    first_local = first['length'] / 2.0 if first['length'] < 2 * distancefromend else distancefromend
    unit_dir = first['exit_xyz'].Subtract(first['entry_xyz']).Normalize()
    cur_hanger_xyz = first['entry_xyz'].Add(unit_dir.Multiply(first_local))
    cur_pipe_idx = 0

    debug_lines.append("  [{}] first hanger pipe[0] local={:.6f} xyz=({:.4f},{:.4f},{:.4f})".format(
        label, first_local, cur_hanger_xyz.X, cur_hanger_xyz.Y, cur_hanger_xyz.Z))
    try:
        FabricationPart.CreateHanger(doc, fab_btn, first['element'].Id,
                                     first['entry_conn'], first_local, atos)
    except Exception as ex:
        debug_lines.append("  FAILED first: {}".format(str(ex)))

    last_pipe = pipe_list[-1]
    last_exit_xyz = last_pipe['exit_xyz']
    end_local = last_pipe['length'] / 2.0 if last_pipe['length'] < 2 * distancefromend else last_pipe['length'] - distancefromend

    while True:
        result = walk(cur_pipe_idx, cur_hanger_xyz, spacing)
        if result is None:
            debug_lines.append("  [{}] end of run".format(label))
            break
        next_idx, next_xyz, local_offset = result
        # Stop before end hanger zone only if we are forcing an end hanger
        if force_end_hanger:
            dist_to_end = next_xyz.DistanceTo(last_exit_xyz)
            if dist_to_end < distancefromend - MARGIN:
                debug_lines.append("  [{}] stopping for end zone dist={:.6f}".format(label, dist_to_end))
                break
        pd = pipe_list[next_idx]
        local_offset = max(MARGIN, min(local_offset, pd['length'] - MARGIN))
        actual_dist = cur_hanger_xyz.DistanceTo(next_xyz)
        debug_lines.append("  [{}] hanger pipe[{}] local={:.6f} xyz=({:.4f},{:.4f},{:.4f}) dist_from_prev={:.6f}".format(
            label, next_idx, local_offset,
            next_xyz.X, next_xyz.Y, next_xyz.Z, actual_dist))
        try:
            FabricationPart.CreateHanger(doc, fab_btn, pd['element'].Id,
                                         pd['entry_conn'], local_offset, atos)
        except Exception as ex:
            debug_lines.append("  FAILED: {}".format(str(ex)))
        cur_hanger_xyz = next_xyz
        cur_pipe_idx = next_idx

    if force_end_hanger:
        debug_lines.append("  [{}] forced end hanger pipe[{}] local={:.6f}".format(
            label, len(pipe_list)-1, end_local))
        try:
            FabricationPart.CreateHanger(doc, fab_btn, last_pipe['element'].Id,
                                         last_pipe['entry_conn'], end_local, atos)
        except Exception as ex:
            debug_lines.append("  FAILED end: {}".format(str(ex)))

# ---------------------------------------------------------------------------
# Process a full run: walk chain, split into segments, place hangers.
# Only the last segment gets a forced end hanger.
# ---------------------------------------------------------------------------
def get_element_id_value(eid):
    try:
        return eid.IntegerValue
    except:
        return eid.Value


def process_run(ordered_chain, entry_conns, fab_btn, distancefromend,
                spacing, atos, run_label, debug_lines):
    
    segments = chain_to_segments(ordered_chain, entry_conns)
    debug_lines.append("[{}] {} direction segment(s)".format(run_label, len(segments)))

    for i, seg in enumerate(segments):
        is_last = (i == len(segments) - 1)
        label = "{}-seg{}".format(run_label, i)
        debug_lines.append("[{}] {} pipe(s)".format(label, len(seg)))

        for j, pd in enumerate(seg):
            eid_val = get_element_id_value(pd['element'].Id)

            debug_lines.append(
                "  pipe[{}] id={} len={:.6f} entry=({:.4f},{:.4f},{:.4f}) exit=({:.4f},{:.4f},{:.4f})".format(
                    j,
                    eid_val,
                    pd['length'],
                    pd['entry_xyz'].X, pd['entry_xyz'].Y, pd['entry_xyz'].Z,
                    pd['exit_xyz'].X, pd['exit_xyz'].Y, pd['exit_xyz'].Z
                )
            )

        place_segment(seg, fab_btn, distancefromend, spacing, atos,
                      is_last, label, debug_lines)

# ---------------------------------------------------------------------------
# Group leftover elements into connected clusters for branch processing
# ---------------------------------------------------------------------------
def group_leftovers(leftovers):
    if not leftovers:
        return []
    remaining = list(leftovers)
    groups = []
    while remaining:
        group = [remaining.pop(0)]
        changed = True
        while changed:
            changed = False
            still_out = []
            for e in remaining:
                connected = False
                for ge in group:
                    for gc in ge.ConnectorManager.Connectors:
                        for ec in e.ConnectorManager.Connectors:
                            if gc.Origin.DistanceTo(ec.Origin) < 0.1:
                                connected = True
                                break
                        if connected:
                            break
                    if connected:
                        break
                if connected:
                    group.append(e)
                    changed = True
                else:
                    still_out.append(e)
            remaining = still_out
        groups.append(group)
    return groups

# ---------------------------------------------------------------------------
# Dialog
# ---------------------------------------------------------------------------
class HangerSpacingDialog(Window):
    def __init__(self, button_names, lines):
        super(HangerSpacingDialog, self).__init__()
        self.Title = "Hanger and Spacing"
        self.Width = 390
        self.Height = 300
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.ResizeMode = ResizeMode.NoResize
        stack = StackPanel()
        stack.Orientation = Orientation.Vertical
        stack.Margin = Thickness(10)
        # rebuild properly
        self.Content = None
        stack = StackPanel()
        stack.Orientation = Orientation.Vertical
        stack.Margin = Thickness(10)
        def lbl(txt):
            l = Label(); l.Content = txt; l.FontSize = 12; l.FontFamily = FontFamily("Arial")
            return l
        stack.Children.Add(lbl("Choose Hanger:"))
        self.combobox_hanger = ComboBox()
        self.combobox_hanger.Width = 350; self.combobox_hanger.Height = 20
        self.combobox_hanger.FontSize = 12; self.combobox_hanger.FontFamily = FontFamily("Arial")
        self.combobox_hanger.ItemsSource = Array[object](button_names)
        self.combobox_hanger.SelectedItem = lines[0] if lines[0] in button_names else button_names[0]
        self.combobox_hanger.Margin = Thickness(0, 0, 0, 10)
        self.combobox_hanger.HorizontalAlignment = HorizontalAlignment.Left
        stack.Children.Add(self.combobox_hanger)
        stack.Children.Add(lbl("Distance from End (In):"))
        self.textbox_end_dist = TextBox()
        self.textbox_end_dist.Width = 200; self.textbox_end_dist.Height = 20
        self.textbox_end_dist.FontSize = 12; self.textbox_end_dist.FontFamily = FontFamily("Arial")
        self.textbox_end_dist.Text = str(round(float(lines[1]) * 12.0, 4)); self.textbox_end_dist.Margin = Thickness(0, 0, 0, 10)
        self.textbox_end_dist.HorizontalAlignment = HorizontalAlignment.Left
        stack.Children.Add(self.textbox_end_dist)
        stack.Children.Add(lbl("Hanger Spacing (Ft):"))
        self.textbox_spacing = TextBox()
        self.textbox_spacing.Width = 200; self.textbox_spacing.Height = 20
        self.textbox_spacing.FontSize = 12; self.textbox_spacing.FontFamily = FontFamily("Arial")
        self.textbox_spacing.Text = lines[2]; self.textbox_spacing.Margin = Thickness(0, 0, 0, 10)
        self.textbox_spacing.HorizontalAlignment = HorizontalAlignment.Left
        stack.Children.Add(self.textbox_spacing)
        self.checkbox_atos = CheckBox()
        self.checkbox_atos.Content = "Attach to Structure"; self.checkbox_atos.FontSize = 12
        self.checkbox_atos.FontFamily = FontFamily("Arial"); self.checkbox_atos.IsChecked = True
        self.checkbox_atos.Margin = Thickness(0, 0, 0, 5)
        stack.Children.Add(self.checkbox_atos)
        self.checkbox_support_joints = CheckBox()
        self.checkbox_support_joints.Content = "Support Joints"; self.checkbox_support_joints.FontSize = 12
        self.checkbox_support_joints.FontFamily = FontFamily("Arial")
        self.checkbox_support_joints.IsChecked = lines[3].lower() == 'true'
        self.checkbox_support_joints.Margin = Thickness(0, 0, 0, 10)
        stack.Children.Add(self.checkbox_support_joints)
        btn = Button(); btn.Content = "OK"; btn.FontSize = 12; btn.FontFamily = FontFamily("Arial")
        btn.Width = 74; btn.Height = 25; btn.HorizontalAlignment = HorizontalAlignment.Center
        btn.Click += self.ok_button_clicked
        stack.Children.Add(btn)
        self.Content = stack

    def ok_button_clicked(self, sender, event):
        self.DialogResult = True
        self.Close()

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
try:
    selected_ref = uidoc.Selection.PickObject(
        ObjectType.Element, FabricationPartSelectionFilter(),
        'Select the starting Fabrication Part')
    element = doc.GetElement(selected_ref.ElementId)
    if not isinstance(element, FabricationPart):
        raise Exception("Not a FabricationPart.")
    pick_point = selected_ref.GlobalPoint
    start_connector = find_nearest_connector(element, pick_point)
    if not start_connector:
        raise Exception("No connector found.")

    parameters = element.LookupParameter('Fabrication Service').AsValueString()
    from Autodesk.Revit.DB import FabricationConfiguration
    Config = FabricationConfiguration.GetFabricationConfiguration(doc)
    LoadedServices = Config.GetAllLoadedServices()
    servicenamelist = []
    for s in LoadedServices:
        try: servicenamelist.append(s.Name)
        except: servicenamelist.append('')
    Servicenum = servicenamelist.index(parameters)
    FabService = LoadedServices[Servicenum]
    buttonnames = []
    button_data = []
    grp_count = FabService.PaletteCount if RevitINT > 2022 else FabService.GroupCount
    for gi in range(grp_count):
        for bi in range(FabService.GetButtonCount(gi)):
            bt = FabService.GetButton(gi, bi)
            if bt.IsAHanger:
                buttonnames.append(bt.Name)
                button_data.append((gi, bi, bt.Name))

    folder_name = "c:\\Temp"
    filepath = os.path.join(folder_name, 'Ribbon_PlaceHangers.txt')
    if not os.path.exists(folder_name):
        os.makedirs(folder_name)
    if not os.path.exists(filepath):
        with open(filepath, 'w') as f:
            f.writelines([str(buttonnames[0]) + '\n', '1\n', '4\n', 'True'])
    with open(filepath, 'r') as f:
        lines = [l.rstrip() for l in f.readlines()]
    if len(lines) < 4:
        with open(filepath, 'w') as f:
            f.writelines([str(buttonnames[0]) + '\n', '1\n', '4\n', 'True'])
        with open(filepath, 'r') as f:
            lines = [l.rstrip() for l in f.readlines()]

    form = HangerSpacingDialog(buttonnames, lines)
    if form.ShowDialog():
        Selectedbutton = form.combobox_hanger.SelectedItem
        distancefromend = float(form.textbox_end_dist.Text) / 12.0
        Spacing = float(form.textbox_spacing.Text)
        AtoS = form.checkbox_atos.IsChecked
        SupportJoint = form.checkbox_support_joints.IsChecked

        with open(filepath, 'w') as f:
            f.writelines([Selectedbutton + '\n', str(distancefromend) + '\n',
                          str(Spacing) + '\n', str(SupportJoint)])

        for gi, bi, bn in button_data:
            if bn == Selectedbutton:
                Servicegroupnum = gi; Buttonnum = bi; break
        FabServiceButton = FabService.GetButton(Servicegroupnum, Buttonnum)

        selected_refs = uidoc.Selection.PickObjects(
            ObjectType.Element, CustomISelectionFilter(parameters),
            "Select Fabrication Parts")
        selected_elements = [doc.GetElement(r) for r in selected_refs]
        if element.Id not in [e.Id for e in selected_elements]:
            selected_elements.insert(0, element)

        t = Transaction(doc, 'Place Hangers')
        t.Start()

        debug_lines = ["=== HANGER DEBUG ===",
                       "Spacing={} distancefromend={}".format(Spacing, distancefromend)]

        if SupportJoint:
            for e in selected_elements:
                if not is_pipe(e) or vertical_fab(e):
                    continue
                pipelen = e.CenterlineLength
                pipe_connectors = list(e.ConnectorManager.Connectors)
                if not pipe_connectors:
                    continue
                if pipelen < 2 * distancefromend:
                    try:
                        FabricationPart.CreateHanger(doc, FabServiceButton, e.Id,
                                                     pipe_connectors[0], pipelen / 2.0, AtoS)
                    except: pass
                else:
                    try:
                        for c in pipe_connectors:
                            FabricationPart.CreateHanger(doc, FabServiceButton, e.Id,
                                                         c, distancefromend, AtoS)
                        if pipelen > Spacing + 2 * distancefromend:
                            pos = distancefromend
                            for _ in range(int((math.floor(pipelen) - 2 * distancefromend) / Spacing)):
                                pos += Spacing
                                FabricationPart.CreateHanger(doc, FabServiceButton, e.Id,
                                                             pipe_connectors[0], pos, AtoS)
                    except: pass

        else:
            # Walk main chain from start element/connector
            main_chain, main_entry_conns, leftovers = walk_chain(
                selected_elements, element, start_connector)

            debug_lines.append("Main chain: {} elements, {} leftovers".format(
                len(main_chain), len(leftovers)))

            process_run(main_chain, main_entry_conns, FabServiceButton,
                        distancefromend, Spacing, AtoS, "main", debug_lines)

            # Process branches
            branch_groups = group_leftovers(leftovers)
            debug_lines.append("Branch groups: {}".format(len(branch_groups)))
            for bi, branch_elems in enumerate(branch_groups):
                # Find the branch element whose connector touches the main chain
                branch_start = branch_elems[0]
                branch_start_conn = next(iter(branch_start.ConnectorManager.Connectors), None)
                # Try to find a better start: elem in branch connected to main chain
                main_chain_ids = {e.Id for e in main_chain}
                for be in branch_elems:
                    for bc in be.ConnectorManager.Connectors:
                        for me in main_chain:
                            for mc in me.ConnectorManager.Connectors:
                                if bc.Origin.DistanceTo(mc.Origin) < 0.1:
                                    branch_start = be
                                    branch_start_conn = bc
                                    break
                            if branch_start_conn == bc:
                                break
                        if branch_start_conn == bc:
                            break

                branch_chain, branch_entry_conns, _ = walk_chain(
                    branch_elems, branch_start, branch_start_conn)
                debug_lines.append("Branch {}: {} elements".format(bi, len(branch_chain)))
                process_run(branch_chain, branch_entry_conns, FabServiceButton,
                            distancefromend, Spacing, AtoS, "branch{}".format(bi), debug_lines)

        t.Commit()

        # debug_path = os.path.join("c:\\Temp", "PlaceHangers_debug.txt")
        # with open(debug_path, 'w') as f:
            # f.write('\n'.join(debug_lines))

except Exception as ex:
    pass
