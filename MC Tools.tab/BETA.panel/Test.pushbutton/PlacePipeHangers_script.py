# -*- coding: UTF-8 -*-
import Autodesk
from Autodesk.Revit.DB import Transaction, FabricationConfiguration, FabricationPart, XYZ, ConnectorProfileType
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from Autodesk.Revit.UI import TaskDialog
import math
import os

# Import .NET namespaces for native Windows Balloon Notification and path handling
import clr
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")
import System
from System.Windows.Forms import NotifyIcon, ToolTipIcon
from System.Drawing import SystemIcons, Bitmap

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = doc.Application
RevitVersion = app.VersionNumber
RevitINT = float(RevitVersion)

DIRECTION_DOT_THRESHOLD = 0.999
MARGIN = 0.01
DIST_FROM_END = 1.0  # Default 1ft (12") from end since new UI omits this parameter
ATOS = True

# Setup project-specific dynamic file path resolution to match the config creator script
file_path = doc.PathName
file_name = System.IO.Path.GetFileNameWithoutExtension(file_path)
if not file_name:
    file_name = doc.Title

FOLDER_NAME = r"c:\Temp"
project_name = file_name.replace(" ", "_")
FILEPATH = os.path.join(FOLDER_NAME, "Ribbon_Pipe-Hanger-Config_{}.txt".format(project_name))


def show_balloon_notification(title, message, timeout=5000):
    """Displays a native Windows balloon notification in the system tray area."""
    notify_icon = NotifyIcon()
    try:
        notify_icon.Icon = SystemIcons.Information
        notify_icon.Visible = True
        notify_icon.ShowBalloonTip(timeout, title, message, ToolTipIcon.Info)
    except Exception:
        pass


class FabricationPartSelectionFilter(ISelectionFilter):
    def AllowElement(self, element):
        return isinstance(element, FabricationPart)
    def AllowReference(self, reference, point):
        return False


def load_service_settings(path):
    settings = {}
    if not os.path.exists(path):
        return settings

    with open(path, 'r') as f:
        for line in f:
            line = line.strip()
            if "=" not in line: continue
            parts = line.split("=", 1)
            if len(parts) != 2: continue

            key = parts[0].strip()
            val_string = parts[1].strip()
            rules = []
            
            if ":" not in val_string:
                vals = val_string.split("|")
                if len(vals) == 3:
                    try:
                        rules.append({
                            "size": 999.0,
                            "hanger": vals[0].strip(),
                            "spacing": float(vals[1].strip()),
                            "joints": vals[2].strip().lower() == "true"
                        })
                    except: pass
            else:
                rule_strings = val_string.split("|")
                for rs in rule_strings:
                    r_parts = rs.split(":")
                    if len(r_parts) == 4:
                        try:
                            rules.append({
                                "size": float(r_parts[0].strip()),
                                "hanger": r_parts[1].strip(),
                                "spacing": float(r_parts[2].strip()),
                                "joints": r_parts[3].strip().lower() == "true"
                            })
                        except: pass
                        
            if rules:
                rules.sort(key=lambda x: x["size"])
                settings[key] = rules
    return settings


def safe_param_string(elem, param_name):
    try:
        p = elem.LookupParameter(param_name)
        if p:
            val = p.AsString()
            if val and val.strip(): return val.strip()
            val = p.AsValueString()
            if val and val.strip(): return val.strip()
    except: pass
    return None


def get_service_name(elem):
    for p_name in ["Fabrication Service Name", "Service Name", "Fabrication Service"]:
        val = safe_param_string(elem, p_name)
        if val: return val
    try:
        return elem.LookupParameter('Fabrication Service').AsValueString()
    except:
        pass
    return "UNASSIGNED"


def get_pipe_size(elem):
    try:
        conns = list(elem.ConnectorManager.Connectors)
        for c in conns:
            if c.Shape == ConnectorProfileType.Round:
                return c.Radius * 2.0 * 12.0
            elif c.Shape == ConnectorProfileType.Rectangular:
                return max(c.Width, c.Height) * 12.0
    except: pass
    return 999.0


def is_disabled_hanger(hanger_name):
    if not hanger_name: return True
    h = hanger_name.strip().upper()
    return h == "" or h == "--- NONE ---" or "NONE" in h


def get_rule_for_element(elem, settings):
    svc_name = get_service_name(elem).strip().lower()
    size_inches = get_pipe_size(elem)
    
    rules = []
    # 1. Exact match
    for k, v in settings.items():
        if k.strip().lower() == svc_name:
            rules = v
            break
            
    # 2. Substring/fuzzy match if exact match fails
    if not rules:
        for k, v in settings.items():
            k_clean = k.strip().lower()
            if k_clean in svc_name or svc_name in k_clean:
                rules = v
                break

    if not rules:
        return "--- NONE ---", 10.0, True
    
    for rule in rules:
        if size_inches <= rule["size"]:
            return rule["hanger"], rule["spacing"], rule["joints"]
            
    last_rule = rules[-1]
    return last_rule["hanger"], last_rule["spacing"], last_rule["joints"]


def get_hanger_button(doc, elem, hanger_name):
    if is_disabled_hanger(hanger_name): return None
    try:
        config = FabricationConfiguration.GetFabricationConfiguration(doc)
        svc_name = get_service_name(elem).strip().lower()
        target_hanger = hanger_name.strip().lower()
        
        for svc in config.GetAllLoadedServices():
            if svc.Name:
                s_name = svc.Name.strip().lower()
                if s_name == svc_name or svc_name in s_name or s_name in svc_name:
                    grp_count = svc.PaletteCount if RevitINT > 2022 else svc.GroupCount
                    for gi in range(grp_count):
                        for bi in range(svc.GetButtonCount(gi)):
                            bt = svc.GetButton(gi, bi)
                            if bt and bt.IsAHanger and bt.Name:
                                if bt.Name.strip().lower() == target_hanger:
                                    return bt
                                    
        for svc in config.GetAllLoadedServices():
            grp_count = svc.PaletteCount if RevitINT > 2022 else svc.GroupCount
            for gi in range(grp_count):
                for bi in range(svc.GetButtonCount(gi)):
                    bt = svc.GetButton(gi, bi)
                    if bt and bt.IsAHanger and bt.Name:
                        if bt.Name.strip().lower() == target_hanger:
                            return bt
    except: pass
    return None


def is_cid_2875(element):
    try:
        return element.ItemCustomId in (2875, 875)
    except:
        return False


def vertical_fab(element):
    pts = [c.Origin for c in element.ConnectorManager.Connectors]
    if len(pts) >= 2:
        v = pts[1].Subtract(pts[0])
        if v.GetLength() < 0.0001:
            return False
        v = v.Normalize()
        angle_from_horizontal = math.asin(abs(v.Z))
        threshold = math.radians(22.5)
        return angle_from_horizontal > threshold
    return False


def is_pipe(element):
    try:
        return element.LookupParameter('Part Pattern Number').AsInteger() == 2041
    except:
        return False


def get_pipe_direction(entry_xyz, exit_xyz):
    v = exit_xyz.Subtract(entry_xyz)
    if v.GetLength() < 0.0001:
        return None
    return v.Normalize()


# ==============================================================================
# 1-to-1 ORIGINAL LOGIC FOR WALKING CHAINS & PLACING HANGERS
# ==============================================================================

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
                other_conns = [c for c in cand_elem.ConnectorManager.Connectors if c.Id != cand_entry.Id]
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


def place_segment(pipe_list, fab_btn, distancefromend, spacing, atos, force_end_hanger, doc):
    placed_count = 0
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

    try:
        FabricationPart.CreateHanger(doc, fab_btn, first['element'].Id, first['entry_conn'], first_local, atos)
        placed_count += 1
    except: pass

    last_pipe = pipe_list[-1]
    last_exit_xyz = last_pipe['exit_xyz']
    end_local = last_pipe['length'] / 2.0 if last_pipe['length'] < 2 * distancefromend else last_pipe['length'] - distancefromend

    while True:
        result = walk(cur_pipe_idx, cur_hanger_xyz, spacing)
        if result is None:
            break
        next_idx, next_xyz, local_offset = result
        if force_end_hanger:
            dist_to_end = next_xyz.DistanceTo(last_exit_xyz)
            if dist_to_end < distancefromend - MARGIN:
                break
        pd = pipe_list[next_idx]
        local_offset = max(MARGIN, min(local_offset, pd['length'] - MARGIN))
        
        try:
            FabricationPart.CreateHanger(doc, fab_btn, pd['element'].Id, pd['entry_conn'], local_offset, atos)
            placed_count += 1
        except: pass
        
        cur_hanger_xyz = next_xyz
        cur_pipe_idx = next_idx

    if force_end_hanger:
        try:
            FabricationPart.CreateHanger(doc, fab_btn, last_pipe['element'].Id, last_pipe['entry_conn'], end_local, atos)
            placed_count += 1
        except: pass
        
    return placed_count


def process_run(ordered_chain, entry_conns, settings, doc, atos):
    run_placed = 0
    segments = chain_to_segments(ordered_chain, entry_conns)
    for i, seg in enumerate(segments):
        is_last = (i == len(segments) - 1)
        
        # Read rules for the first pipe of this specific segment
        first_pipe = seg[0]['element']
        hanger_name, spacing, _ = get_rule_for_element(first_pipe, settings)
        
        if is_disabled_hanger(hanger_name): 
            continue
            
        fab_btn = get_hanger_button(doc, first_pipe, hanger_name)
        if not fab_btn: 
            continue

        run_placed += place_segment(seg, fab_btn, DIST_FROM_END, spacing, atos, is_last, doc)
        
    return run_placed


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


# Finds a viable starting element and its open connector from a contiguous network
def find_best_start(network):
    if not network: return None, None
    if len(network) == 1:
        conns = list(network[0].ConnectorManager.Connectors)
        return network[0], conns[0] if conns else None
        
    for e in network:
        connected_count = 0
        open_conn = None
        for c in e.ConnectorManager.Connectors:
            is_connected = False
            for other in network:
                if other.Id == e.Id: continue
                for oc in other.ConnectorManager.Connectors:
                    if c.Origin.DistanceTo(oc.Origin) < 0.1:
                        is_connected = True
                        break
                if is_connected: break
            
            if is_connected:
                connected_count += 1
            else:
                open_conn = c
        
        if open_conn and connected_count >= 0:
            return e, open_conn
            
    # Fallback
    conns = list(network[0].ConnectorManager.Connectors)
    return network[0], conns[0] if conns else None


# ---------------------------------------------------------------------------
# MAIN EXECUTION
# ---------------------------------------------------------------------------
try:
    settings = load_service_settings(FILEPATH)
    if not settings:
        TaskDialog.Show("Place Hangers", "Config file not found or empty for project '{}'!\nRun the configuration tool first.".format(file_name))
        import sys
        sys.exit()

    # Multi-selection workflow: box select elements freely without a single-element anchor prompt
    selected_refs = uidoc.Selection.PickObjects(
        ObjectType.Element, FabricationPartSelectionFilter(),
        "Select Fabrication Parts for Hanger Placement")
        
    selected_elements = [doc.GetElement(r) for r in selected_refs]
    if not selected_elements:
        import sys
        sys.exit()

    t = Transaction(doc, 'Place Hangers')
    t.Start()

    placed_count = 0

    # Compartmentalize elements strictly by their individual Revit Service Name
    elements_by_service = {}
    for e in selected_elements:
        svc_name = get_service_name(e)
        if svc_name not in elements_by_service:
            elements_by_service[svc_name] = []
        elements_by_service[svc_name].append(e)

    # Process each service container entirely isolated from other services
    for svc_name, svc_elements in elements_by_service.items():
        joints_elements = []
        chain_elements = []
        
        for e in svc_elements:
            hanger_name, _, joints = get_rule_for_element(e, settings)
            if is_disabled_hanger(hanger_name): 
                continue
            
            if joints:
                joints_elements.append(e)
            else:
                chain_elements.append(e)

        # 1. Process Joint-Supported independently per service
        for e in joints_elements:
            if not is_pipe(e) or vertical_fab(e): continue
            
            h_name, sp, _ = get_rule_for_element(e, settings)
            if is_disabled_hanger(h_name): continue
            
            btn = get_hanger_button(doc, e, h_name)
            if not btn: continue
            
            pipelen = e.CenterlineLength
            pipe_connectors = list(e.ConnectorManager.Connectors)
            if not pipe_connectors: continue
            
            if pipelen < 2 * DIST_FROM_END:
                try: 
                    FabricationPart.CreateHanger(doc, btn, e.Id, pipe_connectors[0], pipelen / 2.0, ATOS)
                    placed_count += 1
                except: pass
            else:
                try:
                    for c in pipe_connectors:
                        FabricationPart.CreateHanger(doc, btn, e.Id, c, DIST_FROM_END, ATOS)
                        placed_count += 1
                    if pipelen > sp + 2 * DIST_FROM_END:
                        pos = DIST_FROM_END
                        for _ in range(int((math.floor(pipelen) - 2 * DIST_FROM_END) / sp)):
                            pos += sp
                            FabricationPart.CreateHanger(doc, btn, e.Id, pipe_connectors[0], pos, ATOS)
                            placed_count += 1
                except: pass

        # 2. Process Continuous Chain runs strictly isolated within this service's element group
        if chain_elements:
            networks = group_leftovers(chain_elements)
            
            for network in networks:
                start_element, start_connector = find_best_start(network)
                if not start_element or not start_connector: continue
                
                main_chain, main_entry_conns, leftovers = walk_chain(network, start_element, start_connector)
                placed_count += process_run(main_chain, main_entry_conns, settings, doc, ATOS)

                branch_groups = group_leftovers(leftovers)
                for branch_elems in branch_groups:
                    branch_start = branch_elems[0]
                    branch_start_conn = next(iter(branch_start.ConnectorManager.Connectors), None)
                    
                    # Verify attachment point to the main chain
                    for be in branch_elems:
                        for bc in be.ConnectorManager.Connectors:
                            for me in main_chain:
                                for mc in me.ConnectorManager.Connectors:
                                    if bc.Origin.DistanceTo(mc.Origin) < 0.1:
                                        branch_start = be
                                        branch_start_conn = bc
                                        break
                                if branch_start_conn == bc: break
                            if branch_start_conn == bc: break
                    
                    branch_chain, branch_entry_conns, _ = walk_chain(branch_elems, branch_start, branch_start_conn)
                    placed_count += process_run(branch_chain, branch_entry_conns, settings, doc, ATOS)

    t.Commit()
    
    show_balloon_notification(
        "Place Hangers", 
        "Hanger placement complete.\nTotal successfully placed: {}".format(placed_count)
    )

except Exception as ex:
    msg = str(ex).lower()
    if "cancel" not in msg:
        TaskDialog.Show("Place Hangers", str(ex))