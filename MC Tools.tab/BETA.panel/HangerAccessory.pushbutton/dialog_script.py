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
from System.Windows.Controls import StackPanel, TextBox, ListBox, Label, ComboBox, Button, DockPanel, Dock, Orientation, RadioButton
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


class AnyElementSelectionFilter(ISelectionFilter):
    def AllowElement(self, elem):
        try:
            return elem is not None
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


def get_element_bbox_elevation(elem, use_top=True):
    try:
        box = elem.get_BoundingBox(None)
        if not box:
            return None
        return box.Max.Z if use_top else box.Min.Z
    except:
        return None


def get_part_alignment_z(part, align_mode):
    try:
        if align_mode == "origin":
            return part.Origin.Z

        box = part.get_BoundingBox(None)
        if not box:
            return part.Origin.Z

        if align_mode == "top":
            return box.Max.Z
        elif align_mode == "bottom":
            return box.Min.Z
        else:
            return part.Origin.Z
    except:
        try:
            return part.Origin.Z
        except:
            return None


def show_balloon_notification(title, message, timeout=5000):
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
        self.selected_mode = "rod_end"
        self.selected_align_mode = "origin"

        self.Title = "Select Fabrication Part"
        self.Width = 470
        self.Height = 790
        self.Topmost = True
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.ResizeMode = ResizeMode.CanResize

        root_panel = DockPanel()
        root_panel.Margin = Thickness(10)

        # OK Button
        self.ok_button = Button()
        self.ok_button.Content = "OK"
        self.ok_button.Height = 30
        self.ok_button.Width = 100
        self.ok_button.HorizontalAlignment = HorizontalAlignment.Right
        self.ok_button.Margin = Thickness(0, 10, 0, 0)
        self.ok_button.IsEnabled = False
        self.ok_button.Click += self.on_ok_click
        DockPanel.SetDock(self.ok_button, Dock.Bottom)
        root_panel.Children.Add(self.ok_button)

        # Bottom options panel
        options_panel = StackPanel()
        options_panel.Margin = Thickness(0, 10, 0, 0)
        DockPanel.SetDock(options_panel, Dock.Bottom)

        options_row = StackPanel()
        options_row.Orientation = Orientation.Horizontal

        # Left group - Placement Elevation
        placement_panel = StackPanel()
        placement_panel.Margin = Thickness(0, 0, 30, 0)

        placement_label = Label()
        placement_label.Content = "Placement Elevation:"
        placement_panel.Children.Add(placement_label)

        self.rb_rod_end = RadioButton()
        self.rb_rod_end.Content = "End of Rod"
        self.rb_rod_end.GroupName = "PlacementModeGroup"
        self.rb_rod_end.IsChecked = True
        placement_panel.Children.Add(self.rb_rod_end)

        self.rb_bbox_top = RadioButton()
        self.rb_bbox_top.Content = "Reference Element - BBox Top"
        self.rb_bbox_top.GroupName = "PlacementModeGroup"
        placement_panel.Children.Add(self.rb_bbox_top)

        self.rb_bbox_bottom = RadioButton()
        self.rb_bbox_bottom.Content = "Reference Element - BBox Bottom"
        self.rb_bbox_bottom.GroupName = "PlacementModeGroup"
        placement_panel.Children.Add(self.rb_bbox_bottom)

        # Right group - Placed Part Alignment
        alignment_panel = StackPanel()

        align_label = Label()
        align_label.Content = "Placed Part Alignment:"
        alignment_panel.Children.Add(align_label)

        self.rb_align_origin = RadioButton()
        self.rb_align_origin.Content = "Origin"
        self.rb_align_origin.GroupName = "AlignmentModeGroup"
        self.rb_align_origin.IsChecked = True
        alignment_panel.Children.Add(self.rb_align_origin)

        self.rb_align_top = RadioButton()
        self.rb_align_top.Content = "Top"
        self.rb_align_top.GroupName = "AlignmentModeGroup"
        alignment_panel.Children.Add(self.rb_align_top)

        self.rb_align_bottom = RadioButton()
        self.rb_align_bottom.Content = "Bottom"
        self.rb_align_bottom.GroupName = "AlignmentModeGroup"
        alignment_panel.Children.Add(self.rb_align_bottom)

        options_row.Children.Add(placement_panel)
        options_row.Children.Add(alignment_panel)
        options_panel.Children.Add(options_row)

        root_panel.Children.Add(options_panel)

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

        quick_label = Label()
        quick_label.Content = "Quick Search:"
        quick_label.Margin = Thickness(0, 0, 0, 2)
        top_stack.Children.Add(quick_label)

        quick_panel = StackPanel()
        quick_panel.Orientation = Orientation.Horizontal
        quick_panel.Margin = Thickness(0, 0, 0, 10)

        for tag in ["BEAM", "CLAMP", "HEX", "CLEAR"]:
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

        self.list_box = ListBox()
        self.list_box.Margin = Thickness(0, 0, 0, 0)
        self.list_box.MouseDoubleClick += self.on_double_click
        self.list_box.SelectionChanged += self.on_list_selection_changed
        root_panel.Children.Add(self.list_box)

        self.Content = root_panel

        self.refresh_list()
        self.update_ok_state()

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
        self.update_ok_state()

    def update_ok_state(self):
        idx = self.list_box.SelectedIndex
        self.ok_button.IsEnabled = (idx >= 0 and idx < len(self.filtered_records))

    def on_list_selection_changed(self, sender, args):
        self.update_ok_state()

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
            self.update_ok_state()
            return False

        self.selected_record = self.filtered_records[idx]

        if self.rb_bbox_top.IsChecked:
            self.selected_mode = "bbox_top"
        elif self.rb_bbox_bottom.IsChecked:
            self.selected_mode = "bbox_bottom"
        else:
            self.selected_mode = "rod_end"

        if self.rb_align_top.IsChecked:
            self.selected_align_mode = "top"
        elif self.rb_align_bottom.IsChecked:
            self.selected_align_mode = "bottom"
        else:
            self.selected_align_mode = "origin"

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
placement_mode = dlg.selected_mode
align_mode = dlg.selected_align_mode

if fab_btn.IsAHanger:
    TaskDialog.Show("Invalid Selection", "Selected button is a hanger. Only non-hanger parts can be placed.")
    import sys
    sys.exit()

reference_z = None

if placement_mode in ["bbox_top", "bbox_bottom"]:
    any_filter = AnyElementSelectionFilter()

    try:
        ref = uidoc.Selection.PickObject(
            ObjectType.Element,
            any_filter,
            "Pick reference element for bounding box elevation"
        )
    except OperationCanceledException:
        import sys
        sys.exit()

    ref_elem = doc.GetElement(ref.ElementId)

    if not ref_elem:
        TaskDialog.Show("Error", "No reference element selected.")
        raise Exception("No reference element selected")

    reference_z = get_element_bbox_elevation(
        ref_elem,
        use_top=(placement_mode == "bbox_top")
    )

    if reference_z is None:
        TaskDialog.Show("Error", "Could not get bounding box elevation from selected element.")
        raise Exception("Reference element has no bounding box")


# -------------------------------------------------------
# CREATE PARTS
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
                target_z = pt.Z if placement_mode == "rod_end" else reference_z

                dx = pt.X - loc.X
                dy = pt.Y - loc.Y

                current_align_z = get_part_alignment_z(new_part, align_mode)
                if current_align_z is None:
                    current_align_z = loc.Z

                dz = target_z - current_align_z

                translation = XYZ(dx, dy, dz)
                ElementTransformUtils.MoveElement(doc, new_part.Id, translation)
                placed_count += 1
            except:
                pass

    t.Commit()
except:
    t.RollBack()
    raise

mode_text = {
    "rod_end": "End of Rod",
    "bbox_top": "Reference Element - BBox Top",
    "bbox_bottom": "Reference Element - BBox Bottom"
}.get(placement_mode, "Unknown")

align_text = {
    "origin": "Origin",
    "top": "Top",
    "bottom": "Bottom"
}.get(align_mode, "Unknown")

completion_message = "Mode: {0}\nAlignment: {1}\nProcessed hangers: {2}\nPlaced parts: {3}\nSkipped hangers with no rods: {4}".format(
    mode_text,
    align_text,
    len(selected_hangers),
    placed_count,
    skipped_count
)

show_balloon_notification("Fabrication Parts Placed", completion_message)