# -*- coding: utf-8 -*-

from pyrevit import forms, script
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from Autodesk.Revit.DB import (
    FabricationPart,
    FabricationConfiguration,
    FabricationPartType,
    ElementId,
    LocationCurve,
    LocationPoint,
    Transaction,
    TransactionGroup,
    XYZ
)
from System.Collections.Generic import List

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = doc.Application
output = script.get_output()


# -----------------------------------------------------------------------------
# helpers
# -----------------------------------------------------------------------------

# def log(msg):
    # try:
        # output.print_md(msg)
    # except:
        # print(msg)


def norm(s):
    if s is None:
        return ""
    return str(s).strip().lower()


def get_service_param_value(elem):
    p = elem.LookupParameter("Fabrication Service")
    if not p:
        return None
    try:
        v = p.AsValueString()
        if v:
            return v
    except:
        pass
    try:
        v = p.AsString()
        if v:
            return v
    except:
        pass
    return None


def safe_service_name(service):
    try:
        if service.Name:
            return service.Name
    except:
        pass
    try:
        if service.FabricationSystemName:
            return service.FabricationSystemName
    except:
        pass
    return "<Unknown Service>"


def get_all_services(config):
    services = []
    seen = set()

    try:
        for s in config.GetAllLoadedServices():
            sid = s.ServiceId
            if sid not in seen:
                services.append(s)
                seen.add(sid)
    except:
        pass

    try:
        for s in config.GetAllUsedServices():
            sid = s.ServiceId
            if sid not in seen:
                services.append(s)
                seen.add(sid)
    except:
        pass

    return services


def get_group_count(service):
    try:
        return service.PaletteCount
    except:
        pass
    try:
        return service.GroupCount
    except:
        pass
    return 0


def get_group_name(service, idx):
    try:
        return service.GetPaletteName(idx)
    except:
        pass
    try:
        return service.GetGroupName(idx)
    except:
        pass
    return "Group {}".format(idx)


def safe_button_name(button):
    try:
        if button.Name:
            return button.Name
    except:
        pass
    return "<Unnamed Button>"


def button_is_hanger(button):
    try:
        return bool(button.IsAHanger)
    except:
        return False


def get_hanger_buttons(service):
    items = []
    gcount = get_group_count(service)

    for gi in range(gcount):
        group_name = get_group_name(service, gi)
        try:
            bcount = service.GetButtonCount(gi)
        except:
            bcount = 0

        for bi in range(bcount):
            try:
                btn = service.GetButton(gi, bi)
            except:
                btn = None

            if not btn:
                continue

            if not button_is_hanger(btn):
                continue

            btn_name = safe_button_name(btn)

            items.append({
                "group_index": gi,
                "group_name": group_name,
                "button_index": bi,
                "button_name": btn_name,
                "button": btn,
                "display": u"{} | {}".format(group_name, btn_name)
            })

    return items


def get_element_center(elem):
    bb = elem.get_BoundingBox(None)
    if not bb:
        return None
    return XYZ(
        (bb.Min.X + bb.Max.X) * 0.5,
        (bb.Min.Y + bb.Max.Y) * 0.5,
        (bb.Min.Z + bb.Max.Z) * 0.5
    )


def get_point_for_element(elem):
    loc = elem.Location
    if isinstance(loc, LocationPoint):
        return loc.Point
    return get_element_center(elem)


def get_connectors(elem):
    conns = []
    try:
        for c in elem.ConnectorManager.Connectors:
            conns.append(c)
    except:
        pass
    return conns


def make_net_id_list(ids):
    net_ids = List[ElementId]()
    for i in ids:
        net_ids.Add(i)
    return net_ids


def is_fab_hanger(elem):
    try:
        return isinstance(elem, FabricationPart) and elem.IsAHanger()
    except:
        return False


# -----------------------------------------------------------------------------
# selection filters
# -----------------------------------------------------------------------------

class SeedHangerFilter(ISelectionFilter):
    def AllowElement(self, elem):
        return is_fab_hanger(elem)

    def AllowReference(self, reference, point):
        return False


class SameFabServiceHangerFilter(ISelectionFilter):
    def __init__(self, service_value):
        self.service_value = norm(service_value)

    def AllowElement(self, elem):
        if not is_fab_hanger(elem):
            return False
        val = get_service_param_value(elem)
        return norm(val) == self.service_value

    def AllowReference(self, reference, point):
        return False


# -----------------------------------------------------------------------------
# UI wrapper
# -----------------------------------------------------------------------------

class HangerChoice(forms.TemplateListItem):
    @property
    def name(self):
        return self.item["display"]


# -----------------------------------------------------------------------------
# service resolution
# -----------------------------------------------------------------------------

def resolve_seed_service_by_parameter(seed):
    service_value = get_service_param_value(seed)
    if not service_value:
        return None

    config = FabricationConfiguration.GetFabricationConfiguration(doc)
    if not config:
        return None

    services = get_all_services(config)
    exact_matches = []
    loose_matches = []

    for s in services:
        s_name = None
        s_sys = None

        try:
            s_name = s.Name
        except:
            pass

        try:
            s_sys = s.FabricationSystemName
        except:
            pass

        if norm(service_value) == norm(s_name) or norm(service_value) == norm(s_sys):
            exact_matches.append(s)
        elif norm(service_value) in norm(s_name) or norm(s_name) in norm(service_value):
            loose_matches.append(s)
        elif norm(service_value) in norm(s_sys) or norm(s_sys) in norm(service_value):
            loose_matches.append(s)

    candidates = exact_matches if exact_matches else loose_matches

    # log("## Seed Service Resolution")
    # log("- Seed parameter value: `{}`".format(service_value))
    # log("- Candidate matches: `{}`".format(len(candidates)))

    # for s in candidates:
        # log("- Candidate service: `{}`".format(safe_service_name(s)))

    # for s in candidates:
        # buttons = get_hanger_buttons(s)
        # if buttons:
            # return {
                # "service_value": service_value,
                # "service": s,
                # "service_name": safe_service_name(s),
                # "buttons": buttons,
                # "source": "parameter"
            # }

    # if candidates:
        # return {
            # "service_value": service_value,
            # "service": candidates[0],
            # "service_name": safe_service_name(candidates[0]),
            # "buttons": [],
            # "source": "parameter-no-buttons"
        # }

    # return None


def resolve_seed_service_by_type(seed):
    seed_type_id = seed.GetTypeId()
    config = FabricationConfiguration.GetFabricationConfiguration(doc)
    if not config:
        return None

    services = get_all_services(config)

    for s in services:
        buttons = get_hanger_buttons(s)
        matched = False

        for item in buttons:
            btn = item["button"]
            try:
                cond_count = btn.ConditionCount
            except:
                cond_count = 1

            for ci in range(cond_count):
                try:
                    tid = FabricationPartType.Lookup(doc, btn, ci)
                except:
                    tid = ElementId.InvalidElementId

                if tid and tid != ElementId.InvalidElementId and tid.IntegerValue == seed_type_id.IntegerValue:
                    matched = True
                    break
            if matched:
                break

        if matched:
            return {
                "service_value": get_service_param_value(seed),
                "service": s,
                "service_name": safe_service_name(s),
                "buttons": buttons,
                "source": "type-fallback"
            }

    return None


def resolve_seed_service(seed):
    result = resolve_seed_service_by_parameter(seed)
    if result and result["buttons"]:
        return result

    fallback = resolve_seed_service_by_type(seed)
    if fallback:
        return fallback

    return result


# -----------------------------------------------------------------------------
# hosted placement
# -----------------------------------------------------------------------------

def try_get_host_placement(hanger):
    try:
        hosted_info = hanger.GetHostedInfo()
    except:
        hosted_info = None

    if not hosted_info:
        return (False, None, None, None)

    try:
        host_id = hosted_info.HostId
    except:
        host_id = ElementId.InvalidElementId

    if not host_id or host_id == ElementId.InvalidElementId:
        return (False, None, None, None)

    host = doc.GetElement(host_id)
    if not host or not isinstance(host, FabricationPart):
        return (False, None, None, None)

    host_loc = host.Location
    if not isinstance(host_loc, LocationCurve):
        return (False, None, None, None)

    hanger_pt = get_point_for_element(hanger)
    if not hanger_pt:
        return (False, None, None, None)

    try:
        ir = host_loc.Curve.Project(hanger_pt)
    except:
        ir = None

    if not ir:
        return (False, None, None, None)

    projected = ir.XYZPoint
    conns = get_connectors(host)
    if not conns:
        return (False, None, None, None)

    conns = sorted(conns, key=lambda c: c.Origin.DistanceTo(projected))
    host_conn = conns[0]

    try:
        distance = host_conn.Origin.DistanceTo(projected)
    except:
        return (False, None, None, None)

    return (True, host_id, host_conn, distance)


# -----------------------------------------------------------------------------
# main
# -----------------------------------------------------------------------------

def main():
    try:
        seed_ref = uidoc.Selection.PickObject(
            ObjectType.Element,
            SeedHangerFilter(),
            "Select one fabrication hanger to define the service"
        )
    except:
        return

    seed = doc.GetElement(seed_ref.ElementId)
    if not is_fab_hanger(seed):
        forms.alert("Selected element is not a fabrication hanger.", exitscript=True)

    seed_ctx = resolve_seed_service(seed)
    if not seed_ctx:
        forms.alert(
            "Could not resolve the seed hanger service.\n\n"
            "Check pyRevit output for debug info.",
            exitscript=True
        )

    service_value = seed_ctx["service_value"]
    service = seed_ctx["service"]
    service_name = seed_ctx["service_name"]
    buttons = seed_ctx["buttons"]
    source = seed_ctx["source"]

    # log("- Resolved service: `{}`".format(service_name))
    # log("- Resolution source: `{}`".format(source))
    # log("- Button count: `{}`".format(len(buttons)))

    if not service_value:
        forms.alert("Seed hanger does not have a readable 'Fabrication Service' value.", exitscript=True)

    if not buttons:
        forms.alert(
            "Service resolved but no hanger buttons were found in that service.\n\n"
            "Service: {}".format(service_name),
            exitscript=True
        )

    wrapped = [HangerChoice(x) for x in buttons]
    selected_button_item = forms.SelectFromList.show(
        wrapped,
        title="Select Replacement Hanger - {}".format(service_name),
        multiselect=False,
        width=700,
        button_name="Use Selected Hanger"
    )

    if not selected_button_item:
        return

    # pyRevit may return either TemplateListItem or raw dict
    button_data = getattr(selected_button_item, "item", selected_button_item)

    fab_button = button_data["button"]
    button_name = button_data["button_name"]

    attach_to_structure = forms.alert(
        "Attach new hangers to structure?",
        yes=True,
        no=True,
        ok=False
    )

    try:
        target_refs = uidoc.Selection.PickObjects(
            ObjectType.Element,
            SameFabServiceHangerFilter(service_value),
            "Select fabrication hangers on the same Fabrication Service"
        )
    except:
        return

    if not target_refs:
        return

    targets = []
    seen = set()
    for r in target_refs:
        e = doc.GetElement(r.ElementId)
        if e and is_fab_hanger(e) and e.Id.IntegerValue not in seen:
            targets.append(e)
            seen.add(e.Id.IntegerValue)

    if seed.Id.IntegerValue not in seen:
        targets.insert(0, seed)

    swapped = 0
    skipped = []
    delete_ids = []

    tg = TransactionGroup(doc, "Swap Fabrication Hangers")
    tg.Start()

    t = Transaction(doc, "Swap Fabrication Hangers")
    t.Start()

    try:
        for old_hanger in targets:
            old_id = old_hanger.Id.IntegerValue

            ok, host_id, host_conn, distance = try_get_host_placement(old_hanger)
            if not ok:
                skipped.append("Id {}: could not determine host placement".format(old_id))
                continue

            try:
                new_hanger = FabricationPart.CreateHanger(
                    doc,
                    fab_button,
                    host_id,
                    host_conn,
                    distance,
                    attach_to_structure
                )
            except Exception as ex:
                skipped.append("Id {}: create failed - {}".format(old_id, ex))
                continue

            if new_hanger:
                delete_ids.append(old_hanger.Id)
                swapped += 1
            else:
                skipped.append("Id {}: create returned no hanger".format(old_id))

        if delete_ids:
            doc.Delete(make_net_id_list(delete_ids))

        t.Commit()
        tg.Assimilate()

    except Exception as ex:
        try:
            t.RollBack()
        except:
            pass
        try:
            tg.RollBack()
        except:
            pass

        forms.alert("Swap failed: {}".format(ex), exitscript=True)

    if skipped:
        log("## Skipped")
        for s in skipped:
            log("- {}".format(s))

    # forms.alert(
        # "Fabrication hanger swap complete.\n\n"
        # "Service: {}\n"
        # "Replacement: {}\n"
        # "Selected: {}\n"
        # "Swapped: {}\n"
        # "Skipped: {}".format(
            # service_name,
            # button_name,
            # len(targets),
            # swapped,
            # len(skipped)
        # )
    # )


if __name__ == "__main__":
    main()