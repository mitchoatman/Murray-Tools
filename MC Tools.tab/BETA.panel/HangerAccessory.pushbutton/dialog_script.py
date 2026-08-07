# -*- coding: utf-8 -*-
import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.Exceptions import OperationCanceledException
from System.Windows import Window, Thickness, WindowStartupLocation, ResizeMode, HorizontalAlignment
from System.Windows.Controls import StackPanel, TextBox, ListBox, Label, ComboBox, Button, DockPanel, Dock, Orientation
from System.Windows.Input import Keyboard
from System.Windows.Forms import NotifyIcon, ToolTipIcon
from System.Drawing import SystemIcons

# Revit
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

def get_fab_group_count(fab_service):
    try:
        return fab_service.PaletteCount   # Revit 2022+
    except:
        return fab_service.GroupCount     # Revit 2021 and older


def get_fab_group_name(fab_service, index):
    try:
        return fab_service.GetPaletteName(index)   # Revit 2022+
    except:
        return fab_service.GetGroupName(index)     # Revit 2021 and older


# -------------------------------------------------------
# SELECTION FILTERS
# -------------------------------------------------------
class FabricationPartOrHangerSelectionFilter(ISelectionFilter):
    def AllowElement(self, elem):
        try:
            return isinstance(elem, FabricationPart)
        except:
            return False

    def AllowReference(self, reference, point):
        return False


class SameServiceHangerSelectionFilter(ISelectionFilter):
    def __init__(self, service_id):
        self.service_id = service_id

    def AllowElement(self, elem):
        try:
            return (
                isinstance(elem, FabricationPart)
                and elem.IsAHanger()
                and elem.ServiceId == self.service_id
            )
        except:
            return False

    def AllowReference(self, reference, point):
        return False


# -------------------------------------------------------
# HELPERS
# -------------------------------------------------------
def get_hanger_rod_points(hanger):
    rod_points = []
    try:
        rod_info = hanger.GetRodInfo()
        if not rod_info or rod_info.RodCount == 0:
            return rod_points

        for i in range(rod_info.RodCount):
            pt = rod_info.GetRodEndPosition(i)
            if pt:
                rod_points.append(pt)
    except:
        pass

    return rod_points


def get_service_from_part(part):
    config = FabricationConfiguration.GetFabricationConfiguration(doc)
    services = config.GetAllLoadedServices()

    for s in services:
        try:
            if s.ServiceId == part.ServiceId:
                return s
        except:
            pass
    return None


def show_balloon_notification(title, message, timeout=5000):
    """Displays a native Windows balloon notification."""
    notify_icon = NotifyIcon()
    try:
        notify_icon.Icon = SystemIcons.Information
        notify_icon.Visible = True
        notify_icon.ShowBalloonTip(timeout, title, message, ToolTipIcon.Info)
    except Exception:
        pass


# -------------------------------------------------------
# PICK FIRST FAB PART (PIPE OR HANGER) TO DEFINE SERVICE
# -------------------------------------------------------
first_pick_filter = FabricationPartOrHangerSelectionFilter()

try:
    first_ref = uidoc.Selection.PickObject(
        ObjectType.Element,
        first_pick_filter,
        "Select fabrication pipe or hanger to define service"
    )
except OperationCanceledException:
    import sys
    sys.exit()

first_part = doc.GetElement(first_ref.ElementId)

if not first_part:
    TaskDialog.Show("Error", "No fabrication part selected.")
    raise Exception("No fabrication part selected")

target_service = get_service_from_part(first_part)
if not target_service:
    TaskDialog.Show("Error", "Could not determine service from selected fabrication part.")
    raise Exception("Service not found")

service_filter = SameServiceHangerSelectionFilter(first_part.ServiceId)

# -------------------------------------------------------
# PICK MULTIPLE HANGERS ON SAME SERVICE
# -------------------------------------------------------
try:
    picked_refs = uidoc.Selection.PickObjects(
        ObjectType.Element,
        service_filter,
        "Select fabrication hangers on the same service"
    )
except OperationCanceledException:
    picked_refs = []

selected_hangers = []

# If first picked element was already a hanger, include it automatically
try:
    if first_part.IsAHanger():
        selected_hangers.append(first_part)
except:
    pass

for r in picked_refs:
    e = doc.GetElement(r.ElementId)
    if e and e.Id not in [h.Id for h in selected_hangers]:
        selected_hangers.append(e)

if not selected_hangers:
    TaskDialog.Show("Error", "No hangers selected.")
    raise Exception("No hangers selected")


# -------------------------------------------------------
# GET SERVICE + BUTTONS
# -------------------------------------------------------
palette_names = []
button_records = []

grp_count = get_fab_group_count(target_service)

for gi in range(grp_count):
    palette_name = get_fab_group_name(target_service, gi)
    palette_names.append(palette_name)

    btn_count = target_service.GetButtonCount(gi)

    for bi in range(btn_count):
        btn = target_service.GetButton(gi, bi)

        if btn.IsAHanger:
            continue

        if btn.ConditionCount > 1:
            for ci in range(btn.ConditionCount):
                cond_name = btn.GetConditionName(ci)
                display = u"{0} - {1}".format(btn.Name, cond_name)

                button_records.append({
                    "palette_index": gi,
                    "palette_name": palette_name,
                    "display": display,
                    "button": btn,
                    "condition_index": ci
                })
        else:
            button_records.append({
                "palette_index": gi,
                "palette_name": palette_name,
                "display": btn.Name,
                "button": btn,
                "condition_index": 0
            })


# -------------------------------------------------------
# WPF DIALOG
# -------------------------------------------------------
class PartPicker(Window):
    def __init__(self, records, palettes):
        self.all_records = list(records)
        self.filtered_records = list(records)
        self.selected_record = None

        self.Title = "Select Fabrication Part"
        self.Width = 450
        self.Height = 635
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.ResizeMode = ResizeMode.CanResize

        root_panel = DockPanel()
        root_panel.Margin = Thickness(10)

        # OK Button docked to bottom so it stays pinned at the bottom right
        self.ok_button = Button()
        self.ok_button.Content = "OK"
        self.ok_button.Height = 30
        self.ok_button.Width = 100
        self.ok_button.HorizontalAlignment = HorizontalAlignment.Right
        self.ok_button.Margin = Thickness(0, 10, 0, 0)
        self.ok_button.Click += self.on_ok_click
        DockPanel.SetDock(self.ok_button, Dock.Bottom)
        root_panel.Children.Add(self.ok_button)

        # Top controls container
        top_stack = StackPanel()
        DockPanel.SetDock(top_stack, Dock.Top)

        palette_label = Label()
        palette_label.Content = "Palette:"
        top_stack.Children.Add(palette_label)

        self.palette_combo = ComboBox()
        self.palette_combo.Margin = Thickness(0, 0, 0, 10)
        self.palette_combo.Items.Add("All Palettes")
        for p in palettes:
            self.palette_combo.Items.Add(p)
        self.palette_combo.SelectedIndex = 0
        self.palette_combo.SelectionChanged += self.apply_filters
        top_stack.Children.Add(self.palette_combo)

        label = Label()
        label.Content = "Search Part:"
        top_stack.Children.Add(label)

        self.search_box = TextBox()
        self.search_box.Margin = Thickness(0, 0, 0, 5)
        self.search_box.TextChanged += self.apply_filters
        top_stack.Children.Add(self.search_box)

        # Quick Search Label
        quick_label = Label()
        quick_label.Content = "Quick Search:"
        quick_label.Margin = Thickness(0, 0, 0, 2)
        top_stack.Children.Add(quick_label)

        # Quick Search Buttons Row
        quick_panel = StackPanel()
        quick_panel.Orientation = Orientation.Horizontal
        quick_panel.Margin = Thickness(0, 0, 0, 10)

        for tag in ["BEAM", "CLAMP", "CLEAR"]:
            quick_btn = Button()
            quick_btn.Content = tag
            quick_btn.Height = 22
            quick_btn.Margin = Thickness(0, 0, 5, 0)
            quick_btn.Padding = Thickness(8, 0, 8, 0)
            quick_btn.Click += self.create_quick_click_handler(tag)
            quick_panel.Children.Add(quick_btn)

        top_stack.Children.Add(quick_panel)

        instr_label = Label()
        instr_label.Content = "Double click to insert or click OK button"
        instr_label.Margin = Thickness(0, 0, 0, 5)
        top_stack.Children.Add(instr_label)

        root_panel.Children.Add(top_stack)

        # ListBox fills the remaining space dynamically
        self.list_box = ListBox()
        self.list_box.Margin = Thickness(0, 0, 0, 0)
        self.list_box.MouseDoubleClick += self.on_double_click
        root_panel.Children.Add(self.list_box)

        self.Content = root_panel

        self.refresh_list()

        self.search_box.Focus()
        Keyboard.Focus(self.search_box)

    def create_quick_click_handler(self, tag):
        def handler(sender, args):
            if tag == "CLEAR":
                self.search_box.Text = ""
            else:
                self.search_box.Text = tag
        return handler

    def refresh_list(self):
        self.list_box.ItemsSource = [r["display"] for r in self.filtered_records]

    def apply_filters(self, sender, args):
        selected_palette = self.palette_combo.SelectedItem
        search_text = self.search_box.Text.lower().strip()

        records = self.all_records

        if selected_palette and selected_palette != "All Palettes":
            records = [r for r in records if r["palette_name"] == selected_palette]

        if search_text:
            records = [r for r in records if search_text in r["display"].lower()]

        self.filtered_records = records
        self.refresh_list()

    def confirm_selection(self):
        idx = self.list_box.SelectedIndex
        if idx < 0 or idx >= len(self.filtered_records):
            TaskDialog.Show("Error", "Please select a part.")
            return False
        self.selected_record = self.filtered_records[idx]
        return True

    def on_double_click(self, sender, args):
        if self.confirm_selection():
            self.DialogResult = True
            self.Close()

    def on_ok_click(self, sender, args):
        if self.confirm_selection():
            self.DialogResult = True
            self.Close()


# -------------------------------------------------------
# SHOW DIALOG
# -------------------------------------------------------
dlg = PartPicker(button_records, palette_names)

if not dlg.ShowDialog():
    import sys
    sys.exit()

selected_record = dlg.selected_record
fab_btn = selected_record["button"]
condition_index = selected_record["condition_index"]

if fab_btn.IsAHanger:
    TaskDialog.Show("Invalid Selection", "Selected button is a hanger. Only non-hanger parts can be placed at rod points.")
    import sys
    sys.exit()


# -------------------------------------------------------
# CREATE PARTS AT ALL ROD POINTS OF ALL SELECTED HANGERS
# -------------------------------------------------------
placed_count = 0
skipped_count = 0

t = Transaction(doc, "Place Fabrication Part at Multiple Hanger Rods")
t.Start()

try:
    for hanger in selected_hangers:
        rod_points = get_hanger_rod_points(hanger)

        if not rod_points:
            skipped_count += 1
            continue

        for pt in rod_points:
            try:
                new_part = FabricationPart.Create(doc, fab_btn, condition_index, hanger.LevelId)

                loc = new_part.Origin
                target_pos = XYZ(pt.X, pt.Y, pt.Z)
                translation = target_pos - loc

                ElementTransformUtils.MoveElement(doc, new_part.Id, translation)
                placed_count += 1
            except:
                pass

    t.Commit()
except:
    t.RollBack()
    raise

completion_message = "Processed hangers: {0}\nPlaced parts: {1}\nSkipped hangers with no rods: {2}".format(
    len(selected_hangers),
    placed_count,
    skipped_count
)

show_balloon_notification("Fabrication Parts Placed", completion_message)