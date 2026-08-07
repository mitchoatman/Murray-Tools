# -*- coding: UTF-8 -*-
import Autodesk
from Autodesk.Revit.DB import Transaction, FabricationConfiguration, FabricationPart, ConnectorProfileType
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
import math
import os

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = doc.Application
RevitINT = float(app.VersionNumber)

CONFIG_FOLDER = r"C:\Temp"
CONFIG_PATH = os.path.join(CONFIG_FOLDER, "Ribbon_Duct-Hanger-Config.txt")
CONNECT_TOL = 0.1

DEFAULTS = {
    "ROUND_HANGER": "",
    "RECT_HANGER": "",
    "END_DIST_IN": "12",
    "ROUND_MAX_SPACING_FT": "8",
    "RECT_MAX_SPACING_FT": "8",
    "ATTACH_TO_STRUCTURE": "True",
}


class FabricationPartSelectionFilter(ISelectionFilter):
    def AllowElement(self, element):
        return isinstance(element, FabricationPart)

    def AllowReference(self, reference, point):
        return False


def get_connectors(element):
    try: return list(element.ConnectorManager.Connectors)
    except: return []


def load_config(path):
    cfg = dict(DEFAULTS)
    if not os.path.exists(path):
        raise Exception("Config file not found:\n{}\n\nRun the config tool first.".format(path))

    with open(path, "r") as f:
        for line in f:
            line = line.strip()
            if "=" not in line: continue
            k, v = line.split("=", 1)
            cfg[k.strip()] = v.strip()

    if not cfg.get("ROUND_HANGER") or not cfg.get("RECT_HANGER"):
        raise Exception("Hanger names are blank in config. Please re-run the configuration tool.")
    return cfg


def get_param_double(elem, names):
    for name in names:
        try:
            p = elem.LookupParameter(name)
            if p and p.HasValue: return p.AsDouble()
        except: pass
    return None


def get_param_text(elem, names):
    for name in names:
        try:
            p = elem.LookupParameter(name)
            if not p: continue
            s = p.AsString()
            if s and s.strip(): return s.strip()
            s = p.AsValueString()
            if s and s.strip(): return s.strip()
        except: pass
    return None


def get_service_name(elem):
    return get_param_text(elem, ["Fabrication Service", "Fabrication Service Name", "Service Name"])


def vertical_fab(element):
    conns = get_connectors(element)
    if len(conns) < 2: return False
    v = conns[1].Origin.Subtract(conns[0].Origin)
    length = v.GetLength()
    if length < 0.0001: return False
    angle = math.asin(abs(v.Z / length))
    return angle > math.radians(85)


def get_shape_info(elem):
    conns = get_connectors(elem)
    for c in conns:
        try:
            if c.Shape == ConnectorProfileType.Round:
                dia = c.Radius * 2.0
                if dia > 0: return {"shape": "ROUND", "diameter": dia}
            elif c.Shape == ConnectorProfileType.Rectangular:
                w, h = c.Width, c.Height
                if w > 0 and h > 0: return {"shape": "RECT", "width": w, "height": h}
        except: pass

    diameter = get_param_double(elem, ["Main Diameter", "Diameter", "Item Diameter", "Overall Diameter"])
    if diameter and diameter > 0: return {"shape": "ROUND", "diameter": diameter}
    
    width = get_param_double(elem, ["Main Primary Width", "Width", "Item Width", "Overall Width"])
    height = get_param_double(elem, ["Depth", "Main Secondary Depth", "Height", "Item Height", "Overall Height"])
    if width and height and width > 0 and height > 0:
        return {"shape": "RECT", "width": width, "height": height}

    return None


def get_main_connectors(elem):
    conns = get_connectors(elem)
    if len(conns) <= 2: return conns
    
    try:
        L = elem.CenterlineLength
    except:
        L = 0.0
        
    if L <= 0:
        return conns[:2]
        
    best_pair = (conns[0], conns[1])
    min_diff = float('inf')
    
    for i in range(len(conns)):
        for j in range(i+1, len(conns)):
            d = conns[i].Origin.DistanceTo(conns[j].Origin)
            diff = abs(d - L)
            if diff < min_diff:
                min_diff = diff
                best_pair = (conns[i], conns[j])
    return list(best_pair)


def is_supported_hanger_host(elem):
    if not isinstance(elem, FabricationPart): return False
    if vertical_fab(elem): return False
    
    try:
        if hasattr(elem, "IsStraight") and not elem.IsStraight: return False
    except: pass
    try:
        if hasattr(elem, "IsStraightSegment") and not elem.IsStraightSegment(): return False
    except: pass

    try:
        if elem.CenterlineLength <= 0: return False
    except: return False
    
    m_conns = get_main_connectors(elem)
    if len(m_conns) < 2: return False

    # GEOMETRIC VECTOR CHECK: Ignores elbows, 45s, and offsets based on connector vectors
    try:
        z1 = m_conns[0].CoordinateSystem.BasisZ
        z2 = m_conns[1].CoordinateSystem.BasisZ
        if z1.DotProduct(z2) > -0.98: 
            return False
    except: pass

    name = get_param_text(elem, ["Item Name", "Item Description", "Description", "Family", "Type Name"]) or ""
    name = name.lower()
    
    exclusions = [
        "tee", "cross", "wye", "elbow", "bend", "transition", "taper", 
        "shoe", "gored", "stamped", "radius", "offset", "takeoff", 
        "cap", "end", "plug", "spud", "reducer", "branch", "collar", "tap", "45", "drop", "rise", "miter"
    ]
    if any(x in name for x in exclusions):
        return False

    return get_shape_info(elem) is not None


def get_spacing_for_element(elem, cfg):
    si = get_shape_info(elem)
    if not si: return float(cfg["ROUND_MAX_SPACING_FT"])
    if si["shape"] == "ROUND": 
        return float(cfg["ROUND_MAX_SPACING_FT"])
    else: 
        return float(cfg["RECT_MAX_SPACING_FT"])


def build_loaded_service_map(doc):
    out = {}
    try:
        config = FabricationConfiguration.GetFabricationConfiguration(doc)
        for svc in config.GetAllLoadedServices():
            out[svc.Name] = svc
    except: pass
    return out


def get_hanger_button(elem, cfg, service_map):
    si = get_shape_info(elem)
    if not si: return None
    
    button_name = cfg["ROUND_HANGER"] if si["shape"] == "ROUND" else cfg["RECT_HANGER"]
    svc_name = get_service_name(elem)
    
    if not svc_name or svc_name not in service_map: 
        return None
        
    svc = service_map[svc_name]
    grp_count = svc.PaletteCount if RevitINT > 2022 else svc.GroupCount
    
    for gi in range(grp_count):
        for bi in range(svc.GetButtonCount(gi)):
            try:
                bt = svc.GetButton(gi, bi)
                if bt.IsAHanger and bt.Name and bt.Name.strip() == button_name.strip(): 
                    return bt
            except: pass
    return None


def group_into_runs(valid_hosts):
    adj = {e.Id: [] for e in valid_hosts}
    host_dict = {e.Id: e for e in valid_hosts}
    
    for idx, e1 in enumerate(valid_hosts):
        m1 = get_main_connectors(e1)
        for e2 in valid_hosts[idx+1:]:
            m2 = get_main_connectors(e2)
            
            for c1 in m1:
                for c2 in m2:
                    if c1.Origin.DistanceTo(c2.Origin) < CONNECT_TOL:
                        adj[e1.Id].append((e2.Id, c1, c2))
                        adj[e2.Id].append((e1.Id, c2, c1))
                        break

    runs = []
    unvisited = set(host_dict.keys())
    
    while unvisited:
        start_id = unvisited.pop()
        comp = [start_id]
        queue = [start_id]
        
        while queue:
            curr = queue.pop(0)
            for nxt, _, _ in adj[curr]:
                if nxt in unvisited:
                    unvisited.remove(nxt)
                    queue.append(nxt)
                    comp.append(nxt)
                    
        sub_adj = {nid: [x for x in adj[nid] if x[0] in comp] for nid in comp}
        endpoints = [nid for nid in comp if len(sub_adj[nid]) <= 1]
        
        start_node = endpoints[0] if endpoints else comp[0]
        
        ordered = []
        curr = start_node
        prev = None
        
        entry_conn = None
        m_conns = get_main_connectors(host_dict[curr])
        if len(sub_adj[curr]) == 1:
            shared_c = sub_adj[curr][0][1]
            entry_conn = m_conns[0] if m_conns[1].Id == shared_c.Id else m_conns[1]
        else:
            entry_conn = m_conns[0]

        safety_counter = 0
        while curr and safety_counter < 1000:
            safety_counter += 1
            ordered.append((host_dict[curr], entry_conn))
            next_step = None
            for nxt, my_conn, their_conn in sub_adj[curr]:
                if nxt != prev:
                    next_step = (nxt, their_conn) 
                    break
            if next_step:
                prev = curr
                curr = next_step[0]
                entry_conn = next_step[1]
            else:
                break
                
        runs.append(ordered)
        
    return runs


try:
    cfg = load_config(CONFIG_PATH)
    service_map = build_loaded_service_map(doc)
    dist_from_end = float(cfg["END_DIST_IN"]) / 12.0
    atos = str(cfg["ATTACH_TO_STRUCTURE"]).lower() == "true"

    selected_refs = uidoc.Selection.PickObjects(
        ObjectType.Element,
        FabricationPartSelectionFilter(),
        "Select fabrication ductwork to hang"
    )

    valid_hosts = []
    for r in selected_refs:
        try:
            e = doc.GetElement(r.ElementId)
            if is_supported_hanger_host(e):
                valid_hosts.append(e)
        except: pass

    if not valid_hosts:
        raise Exception("No valid straight duct parts found in selection.")

    runs = group_into_runs(valid_hosts)

    t = Transaction(doc, "Place Duct Hangers")
    t.Start()

    placed_count = 0

    for run in runs:
        spacing = get_spacing_for_element(run[0][0], cfg)
        if spacing <= 0: spacing = 8.0 
        
        run_length = 0.0
        part_spans = []
        
        for i in range(len(run)):
            part, entry_conn = run[i]
            m_conns = get_main_connectors(part)
            exit_conn = m_conns[1] if m_conns[0].Id == entry_conn.Id else m_conns[0]
            
            L = entry_conn.Origin.DistanceTo(exit_conn.Origin) 
            if L <= 0: L = 0.1
            
            start_dist = run_length
            end_dist = run_length + L
            part_spans.append({
                "part": part, 
                "entry_conn": entry_conn, 
                "exit_conn": exit_conn, 
                "start": start_dist, 
                "end": end_dist,
                "L": L
            })
            
            run_length += L
            if i < len(run) - 1:
                next_part, next_entry = run[i+1]
                gap = exit_conn.Origin.DistanceTo(next_entry.Origin)
                run_length += gap

        # Map all tap/branch connectors in the run to avoid them
        tap_zones = []
        for span in part_spans:
            conns = get_connectors(span["part"])
            main_ids = {span["entry_conn"].Id, span["exit_conn"].Id}
            if span["exit_conn"].Origin.IsAlmostEqualTo(span["entry_conn"].Origin): continue
            direction = span["exit_conn"].Origin.Subtract(span["entry_conn"].Origin).Normalize()
            
            for c in conns:
                if c.Id not in main_ids:
                    v = c.Origin.Subtract(span["entry_conn"].Origin)
                    tap_loc = span["start"] + v.DotProduct(direction)
                    
                    tap_width = 1.0
                    try:
                        if c.Shape == ConnectorProfileType.Rectangular: tap_width = max(c.Width, c.Height)
                        elif c.Shape == ConnectorProfileType.Round: tap_width = c.Radius * 2.0
                    except: pass
                    
                    half_w = tap_width / 2.0
                    # Create an exclusion zone representing the tap + a 6 inch buffer on both sides
                    tap_zones.append((tap_loc - half_w - 0.5, tap_loc + half_w + 0.5))

        # Dynamic Strict Target Generation
        targets = []
        start_target = dist_from_end
        end_target = run_length - dist_from_end
        gap_length = end_target - start_target

        if run_length <= dist_from_end * 2.0 or gap_length < 1.0:
            targets.append(run_length / 2.0)
        else:
            targets.append(start_target)
            current_pos = start_target + spacing
            
            loop_guard = 0
            while current_pos < end_target - 0.5 and loop_guard < 500:
                loop_guard += 1
                
                # Check if current step lands in a tap zone
                adjusted = False
                for z_start, z_end in tap_zones:
                    if z_start <= current_pos <= z_end:
                        # Attempt to pull it back slightly outside the tap zone
                        shifted_pos = z_start - 0.25
                        # If pulling it back puts it too close to the previous hanger, push it forward instead
                        if shifted_pos < targets[-1] + 1.0:
                            shifted_pos = z_end + 0.25
                        
                        current_pos = shifted_pos
                        adjusted = True
                        break
                
                # If adjustment pushes it beyond the end target, break the loop
                if current_pos >= end_target - 0.5:
                    break
                    
                targets.append(current_pos)
                
                # Next hanger is strictly 'spacing' distance from the newly adjusted position
                current_pos += spacing
                
            targets.append(end_target)

        # Place hangers
        for t_abs in targets:
            for span in part_spans:
                if span["start"] - CONNECT_TOL <= t_abs <= span["end"] + CONNECT_TOL:
                    local_pos = t_abs - span["start"]
                    # Provide tiny safe clamping to prevent placing exactly on the connector face
                    local_pos = max(0.1, min(local_pos, span["L"] - 0.1))
                    
                    try:
                        bt = get_hanger_button(span["part"], cfg, service_map)
                        if bt: 
                            FabricationPart.CreateHanger(doc, bt, span["part"].Id, span["entry_conn"], local_pos, atos)
                            placed_count += 1
                    except: pass
                    break

    t.Commit()
    TaskDialog.Show("Place Duct Hangers", "Hanger placement complete. Total placed: {}".format(placed_count))

except Exception as ex:
    msg = str(ex).lower()
    if "cancel" not in msg:
        TaskDialog.Show("Place Duct Hangers", str(ex))