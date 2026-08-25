# -*- coding: utf-8 -*-
import os
import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("PresentationCore")
clr.AddReference("PresentationFramework")
clr.AddReference("WindowsBase")
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")

from Autodesk.Revit.DB import (
    FabricationPart, FabricationConfiguration, ElementId,
    LocationCurve, LocationPoint, Transaction, TransactionGroup, XYZ
)
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from Autodesk.Revit.UI import TaskDialog

from System.Windows import Window, Thickness, WindowStartupLocation, ResizeMode, GridLength, GridUnitType, HorizontalAlignment
from System.Windows.Controls import Grid, RowDefinition, Label, ListBox, Button, StackPanel, Orientation, SelectionMode, CheckBox
from System.Windows.Media import FontFamily
from System.Collections.Generic import List
from System.Windows.Forms import NotifyIcon, ToolTipIcon
from System.Drawing import SystemIcons

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = doc.Application

# -----------------------------------------------------------------------------
# Helpers
# -----------------------------------------------------------------------------
def norm(s):
    if s is None: return ""
    return str(s).strip().lower()

def get_service_param_value(elem):
    p = elem.LookupParameter("Fabrication Service")
    if not p: return None
    try: 
        v = p.AsValueString()
        if v: return v
    except: pass
    try: 
        v = p.AsString()
        if v: return v
    except: pass
    return None

def safe_service_name(service):
    try:
        if service.Name: return service.Name
    except: pass
    try:
        if service.FabricationSystemName: return service.FabricationSystemName
    except: pass
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
    except: pass
    try:
        for s in config.GetAllUsedServices():
            sid = s.ServiceId
            if sid not in seen:
                services.append(s)
                seen.add(sid)
    except: pass
    return services

def get_hanger_buttons(service):
    items = []
    gcount = getattr(service, 'PaletteCount', getattr(service, 'GroupCount', 0))
    for gi in range(gcount):
        group_name = "Group {}".format(gi)
        try:
            group_name = service.GetPaletteName(gi)
        except:
            try: group_name = service.GetGroupName(gi)
            except: pass
        try:
            bcount = service.GetButtonCount(gi)
        except:
            bcount = 0
        for bi in range(bcount):
            try:
                btn = service.GetButton(gi, bi)
            except:
                btn = None
            if not btn or not getattr(btn, 'IsAHanger', False):
                continue
            btn_name = getattr(btn, 'Name', "<Unnamed Button>")
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
    if not bb: return None
    return XYZ((bb.Min.X + bb.Max.X) * 0.5, (bb.Min.Y + bb.Max.Y) * 0.5, (bb.Min.Z + bb.Max.Z) * 0.5)

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
    except: pass
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

# Global list to keep NotifyIcon alive and prevent premature garbage collection
_active_notifications = []

def show_balloon_notification(title, message, timeout=5000):
    notify_icon = NotifyIcon()
    _active_notifications.append(notify_icon)
    try:
        notify_icon.Icon = SystemIcons.Information
        notify_icon.Visible = True
        notify_icon.ShowBalloonTip(timeout, title, message, ToolTipIcon.Info)
    except Exception:
        pass

# -----------------------------------------------------------------------------
# Selection Filters
# -----------------------------------------------------------------------------
class SeedHangerFilter(ISelectionFilter):
    def AllowElement(self, elem):
        return is_fab_hanger(elem)
    def AllowReference(self, reference, point):
        return False

# -----------------------------------------------------------------------------
# WPF Dialog with Checkbox
# -----------------------------------------------------------------------------
class HangerSelectionDialog(Window):
    def __init__(self, button_items, service_name):
        super(HangerSelectionDialog, self).__init__()
        self.Title = "Select Replacement Hanger - {}".format(service_name)
        self.Width = 560
        self.Height = 560
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.ResizeMode = ResizeMode.CanResizeWithGrip

        grid = Grid()
        grid.Margin = Thickness(12)

        row0 = RowDefinition()  # Label
        row0.Height = GridLength.Auto
        row1 = RowDefinition()  # ListBox
        row1.Height = GridLength(1, GridUnitType.Star)
        row2 = RowDefinition()  # Checkbox
        row2.Height = GridLength.Auto
        row3 = RowDefinition()  # Buttons
        row3.Height = GridLength.Auto

        grid.RowDefinitions.Add(row0)
        grid.RowDefinitions.Add(row1)
        grid.RowDefinitions.Add(row2)
        grid.RowDefinitions.Add(row3)

        # Label
        lbl = Label()
        lbl.Content = "Choose replacement hanger:"
        lbl.FontSize = 14
        lbl.Margin = Thickness(0, 0, 0, 8)
        Grid.SetRow(lbl, 0)
        grid.Children.Add(lbl)

        # ListBox
        self.listbox = ListBox()
        self.listbox.SelectionMode = SelectionMode.Single
        self.listbox.FontSize = 13
        self.listbox.Margin = Thickness(0, 0, 0, 12)
        for item in button_items:
            self.listbox.Items.Add(item["display"])
        self.listbox.MouseDoubleClick += self.on_listbox_double_click
        Grid.SetRow(self.listbox, 1)
        grid.Children.Add(self.listbox)

        # Checkbox
        self.chk_attach = CheckBox()
        self.chk_attach.Content = "Attach new hangers to structure"
        self.chk_attach.IsChecked = True
        self.chk_attach.FontSize = 13
        self.chk_attach.Margin = Thickness(4, 8, 0, 12)
        Grid.SetRow(self.chk_attach, 2)
        grid.Children.Add(self.chk_attach)

        # Button Panel
        btn_panel = StackPanel()
        btn_panel.Orientation = Orientation.Horizontal
        btn_panel.HorizontalAlignment = HorizontalAlignment.Right
        btn_panel.Margin = Thickness(0, 0, 0, 4)

        btn_ok = Button()
        btn_ok.Content = "Use Selected Hanger"
        btn_ok.Width = 170
        btn_ok.Height = 34
        btn_ok.Margin = Thickness(8, 0, 8, 0)
        btn_ok.Click += self.on_ok
        btn_panel.Children.Add(btn_ok)

        btn_cancel = Button()
        btn_cancel.Content = "Cancel"
        btn_cancel.Width = 100
        btn_cancel.Height = 34
        btn_cancel.Click += self.on_cancel
        btn_panel.Children.Add(btn_cancel)

        Grid.SetRow(btn_panel, 3)
        grid.Children.Add(btn_panel)

        self.Content = grid
        self.selected_item = None

    def on_ok(self, sender, e):
        if self.listbox.SelectedIndex >= 0:
            self.selected_item = self.listbox.SelectedItem
            self.DialogResult = True
        self.Close()

    def on_listbox_double_click(self, sender, e):
        if self.listbox.SelectedIndex >= 0:
            self.selected_item = self.listbox.SelectedItem
            self.DialogResult = True
            self.Close()

    def on_cancel(self, sender, e):
        self.DialogResult = False
        self.Close()

# -----------------------------------------------------------------------------
# Main
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
        TaskDialog.Show("Error", "Selected element is not a fabrication hanger.")
        return

    config = FabricationConfiguration.GetFabricationConfiguration(doc)
    services = get_all_services(config)
    seed_service_value = get_service_param_value(seed)
    service = None
    buttons = []

    for s in services:
        if norm(seed_service_value) == norm(safe_service_name(s)):
            buttons = get_hanger_buttons(s)
            service = s
            break

    if not service or not buttons:
        TaskDialog.Show("Error", "Could not resolve the seed hanger service or no hanger buttons found.")
        return

    service_name = safe_service_name(service)

    try:
        dlg = HangerSelectionDialog(buttons, service_name)
        if not dlg.ShowDialog():
            return
    except Exception as ex:
        TaskDialog.Show("Dialog Error", "Failed to show dialog:\n{}".format(str(ex)))
        return

    selected_display = dlg.selected_item
    if not selected_display:
        return

    fab_button = None
    button_name = ""
    for b in buttons:
        if b["display"] == selected_display:
            fab_button = b["button"]
            button_name = b["button_name"]
            break

    if not fab_button:
        TaskDialog.Show("Error", "Failed to retrieve selected hanger button.")
        return

    attach_to_structure = bool(dlg.chk_attach.IsChecked)

    # Continue with selection and swap
    try:
        target_refs = uidoc.Selection.PickObjects(
            ObjectType.Element,
            SeedHangerFilter(),
            "Select fabrication hangers to swap"
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
            ok = False
            host_id = ElementId.InvalidElementId
            host_conn = None
            distance = 0.0

            try:
                hosted_info = old_hanger.GetHostedInfo()
                if hosted_info and hosted_info.HostId != ElementId.InvalidElementId:
                    host_id = hosted_info.HostId
                    host = doc.GetElement(host_id)
                    if host and isinstance(host, FabricationPart):
                        host_loc = host.Location
                        if isinstance(host_loc, LocationCurve):
                            hanger_pt = get_point_for_element(old_hanger)
                            if hanger_pt:
                                ir = host_loc.Curve.Project(hanger_pt)
                                if ir:
                                    projected = ir.XYZPoint
                                    conns = sorted(get_connectors(host), key=lambda c: c.Origin.DistanceTo(projected))
                                    if conns:
                                        host_conn = conns[0]
                                        distance = host_conn.Origin.DistanceTo(projected)
                                        ok = True
            except:
                pass

            if not ok:
                skipped.append("Id {}: could not determine host placement".format(old_id))
                continue

            try:
                new_hanger = FabricationPart.CreateHanger(
                    doc, fab_button, host_id, host_conn, distance, attach_to_structure
                )
                if new_hanger:
                    delete_ids.append(old_hanger.Id)
                    swapped += 1
                else:
                    skipped.append("Id {}: create returned no hanger".format(old_id))
            except:
                skipped.append("Id {}: create failed".format(old_id))

        if delete_ids:
            doc.Delete(make_net_id_list(delete_ids))

        t.Commit()
        tg.Assimilate()

        show_balloon_notification(
            "Fabrication Hanger Swap Complete",
            "Swapped: {}\nSkipped: {}".format(swapped, len(skipped))
        )

    except Exception as ex:
        try: t.RollBack()
        except: pass
        try: tg.RollBack()
        except: pass
        TaskDialog.Show("Error", "Swap failed: {}".format(str(ex)))

if __name__ == "__main__":
    main()