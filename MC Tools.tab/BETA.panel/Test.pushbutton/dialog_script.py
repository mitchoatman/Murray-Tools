# -*- coding: utf-8 -*-
import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.Exceptions import OperationCanceledException
from System.Windows import Window, Thickness, WindowStartupLocation, ResizeMode
from System.Windows.Controls import StackPanel, TextBox, ListBox, Label, ComboBox
from System.Windows.Input import Keyboard

# Revit
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument


# -------------------------------------------------------
# SELECTION FILTERS
# -------------------------------------------------------
class FabricationHangerSelectionFilter(ISelectionFilter):
    def AllowElement(self, elem):
        try:
            return isinstance(elem, FabricationPart) and elem.IsAHanger()
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


def get_service_from_hanger(hanger):
    config = FabricationConfiguration.GetFabricationConfiguration(doc)
    services = config.GetAllLoadedServices()

    for s in services:
        try:
            if s.ServiceId == hanger.ServiceId:
                return s
        except:
            pass
    return None


# -------------------------------------------------------
# PICK FIRST HANGER
# -------------------------------------------------------
hanger_filter = FabricationHangerSelectionFilter()

try:
    first_ref = uidoc.Selection.PickObject(
        ObjectType.Element,
        hanger_filter,
        "Select first Fabrication Hanger to define service"
    )
except OperationCanceledException:
    import sys
    sys.exit()

first_hanger = doc.GetElement(first_ref.ElementId)

if not first_hanger:
    TaskDialog.Show("Error", "No hanger selected.")
    raise Exception("No hanger selected")

target_service = get_service_from_hanger(first_hanger)
if not target_service:
    TaskDialog.Show("Error", "Could not determine service from selected hanger.")
    raise Exception("Service not found")

service_filter = SameServiceHangerSelectionFilter(first_hanger.ServiceId)

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

selected_hangers = [first_hanger]

for r in picked_refs:
    e = doc.GetElement(r.ElementId)
    if e and e.Id != first_hanger.Id:
        selected_hangers.append(e)

if not selected_hangers:
    TaskDialog.Show("Error", "No hangers selected.")
    raise Exception("No hangers selected")


# -------------------------------------------------------
# GET SERVICE + BUTTONS
# -------------------------------------------------------
palette_names = []
button_records = []

palette_count = target_service.PaletteCount

for p in range(palette_count):
    palette_name = target_service.GetPaletteName(p)
    palette_names.append(palette_name)

    count = target_service.GetButtonCount(p)

    for i in range(count):
        btn = target_service.GetButton(p, i)

        # skip hanger buttons
        if btn.IsAHanger:
            continue

        if btn.ConditionCount > 1:
            for c in range(btn.ConditionCount):
                cond_name = btn.GetConditionName(c)
                display = u"{0} - {1}".format(btn.Name, cond_name)

                button_records.append({
                    "palette_index": p,
                    "palette_name": palette_name,
                    "display": display,
                    "button": btn,
                    "condition_index": c
                })
        else:
            display = u"{0}".format(btn.Name)

            button_records.append({
                "palette_index": p,
                "palette_name": palette_name,
                "display": display,
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
        self.Width = 400
        self.Height = 620
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.ResizeMode = ResizeMode.CanResize

        stack = StackPanel()
        stack.Margin = Thickness(10)

        palette_label = Label()
        palette_label.Content = "Palette:"
        stack.Children.Add(palette_label)

        self.palette_combo = ComboBox()
        self.palette_combo.Margin = Thickness(0, 0, 0, 10)
        self.palette_combo.Items.Add("All Palettes")
        for p in palettes:
            self.palette_combo.Items.Add(p)
        self.palette_combo.SelectedIndex = 0
        self.palette_combo.SelectionChanged += self.apply_filters
        stack.Children.Add(self.palette_combo)

        label = Label()
        label.Content = "Search Part:"
        stack.Children.Add(label)

        self.search_box = TextBox()
        self.search_box.Margin = Thickness(0, 0, 0, 10)
        self.search_box.TextChanged += self.apply_filters
        stack.Children.Add(self.search_box)

        instr_label = Label()
        instr_label.Content = "Double Click Item to Insert"
        instr_label.Margin = Thickness(0, 0, 0, 5)
        stack.Children.Add(instr_label)

        self.list_box = ListBox()
        self.list_box.Height = 430
        self.list_box.Margin = Thickness(0, 0, 0, 10)
        self.list_box.MouseDoubleClick += self.on_double_click
        stack.Children.Add(self.list_box)

        self.Content = stack

        self.refresh_list()

        self.search_box.Focus()
        Keyboard.Focus(self.search_box)

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

    def on_double_click(self, sender, args):
        idx = self.list_box.SelectedIndex

        if idx < 0 or idx >= len(self.filtered_records):
            TaskDialog.Show("Error", "Please select a part.")
            return

        self.selected_record = self.filtered_records[idx]
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

TaskDialog.Show(
    "Complete",
    "Processed hangers: {0}\nPlaced parts: {1}\nSkipped hangers with no rods: {2}".format(
        len(selected_hangers),
        placed_count,
        skipped_count
    )
)