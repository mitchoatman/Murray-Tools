# -*- coding: utf-8 -*-
import Autodesk
import os
import clr
import sys

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('System')
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference("System.Core")

from System import Action
import System.Windows.Threading
from System.Windows import Window, Thickness, HorizontalAlignment, VerticalAlignment, WindowStartupLocation
from System.Windows.Controls import Grid, RowDefinition, ColumnDefinition, Label, TextBox, Button, ListBox, StackPanel, CheckBox, SelectionMode
from System.Windows.Interop import WindowInteropHelper
from System.Collections.Generic import List

from System.Windows.Forms import NotifyIcon, ToolTipIcon
from System.Drawing import Icon, SystemIcons

from Autodesk.Revit.UI import TaskDialog, UIApplication
from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.DB import (
    IFamilyLoadOptions,
    FamilySource,
    Transaction,
    TransactionGroup,
    FilteredElementCollector,
    Family,
    FamilySymbol,
    BuiltInCategory,
    BuiltInParameter,
    IndependentTag,
    TagOrientation,
    Reference,
    ElementMulticategoryFilter,
    LocationPoint,
    LocationCurve,
    XYZ,
    View3D,
    FamilyInstance,
    ElementId
)

DB = Autodesk.Revit.DB
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView

PARAM_NAME = "TS_Point_Number"
FAMILY_NAME = 'Multi Category Tag -TS_Point_Number'
FAMILY_TYPE = 'Multi Category Tag -TS_Point_Number'

SCRIPT_DIR = os.path.dirname(__file__)
FAMILY_FILE = 'Multi Category Tag -TS_Point_Number.rfa'
FAMILY_PATH = os.path.join(SCRIPT_DIR, FAMILY_FILE)

TARGET_CATEGORIES = [
    BuiltInCategory.OST_PipeAccessory,
    BuiltInCategory.OST_PlumbingFixtures,
    BuiltInCategory.OST_StructuralStiffener,
    BuiltInCategory.OST_GenericModel,
    BuiltInCategory.OST_DuctAccessory,
]

CANDIDATE_OFFSETS = [
    XYZ(0.0, 0.0, 0.0),
    XYZ(0.0, 0.75, 0.0),
    XYZ(0.0, -0.75, 0.0),
    XYZ(0.0, 1.50, 0.0),
    XYZ(0.0, -1.50, 0.0),
    XYZ(0.0, 2.25, 0.0),
    XYZ(0.0, -2.25, 0.0),
    XYZ(0.0, 3.00, 0.0),
    XYZ(0.0, -3.00, 0.0),
    XYZ(0.0, 3.75, 0.0),
    XYZ(0.0, -3.75, 0.0),
    XYZ(0.75, 0.75, 0.0),
    XYZ(-0.75, 0.75, 0.0),
    XYZ(0.75, -0.75, 0.0),
    XYZ(-0.75, -0.75, 0.0),
    XYZ(1.50, 0.0, 0.0),
    XYZ(-1.50, 0.0, 0.0),
    XYZ(1.50, 1.50, 0.0),
    XYZ(-1.50, 1.50, 0.0),
    XYZ(1.50, -1.50, 0.0),
    XYZ(-1.50, -1.50, 0.0),
]

BASE_LABEL_WIDTH = 0.55
PER_CHAR_WIDTH = 0.18
LABEL_HEIGHT = 0.45
LABEL_PADDING_X = 0.20
LABEL_PADDING_Y = 0.12


# --------------------------------------------------
# FAMILY LOAD OPTIONS
# --------------------------------------------------
class FamilyLoaderOptionsHandler(IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues.Value = False
        return True

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        source.Value = FamilySource.Family
        overwriteParameterValues.Value = False
        return True


# --------------------------------------------------
# HELPERS
# --------------------------------------------------
def natural_key(s):
    import re
    return [int(text) if text.isdigit() else text.lower() for text in re.split(r'([0-9]+)', s)]

def get_revit_window_handle():
    try:
        return uidoc.Application.MainWindowHandle
    except:
        return System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle

def show_balloon_notification(title, message, icon_path=None, timeout=5000):
    """Displays a native Windows balloon notification."""
    notify_icon = NotifyIcon()
    try:
        if icon_path and os.path.exists(icon_path):
            notify_icon.Icon = Icon(icon_path)
        else:
            notify_icon.Icon = SystemIcons.Information
            
        notify_icon.Visible = True
        notify_icon.ShowBalloonTip(timeout, title, message, ToolTipIcon.Info)
    except Exception:
        pass

def get_tag_symbol(doc, family_name, type_name):
    symbols = FilteredElementCollector(doc).OfClass(FamilySymbol).ToElements()
    for sym in symbols:
        try:
            sym_name = sym.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
            if sym.Family.Name == family_name and sym_name == type_name:
                return sym
        except:
            pass
    return None

def is_parent_family_instance(element):
    if not isinstance(element, FamilyInstance):
        return True
    return element.SuperComponent is None

def get_tag_point(element, view):
    loc = element.Location
    if isinstance(loc, LocationPoint):
        return loc.Point
    if isinstance(loc, LocationCurve):
        try:
            return loc.Curve.Evaluate(0.5, True)
        except:
            pass
    try:
        bbox = element.get_BoundingBox(view)
        if bbox:
            return XYZ(
                (bbox.Min.X + bbox.Max.X) / 2.0,
                (bbox.Min.Y + bbox.Max.Y) / 2.0,
                (bbox.Min.Z + bbox.Max.Z) / 2.0
            )
    except:
        pass
    return None

def get_all_tagged_element_ids(tags):
    tagged_ids = set()
    for tag in tags:
        try:
            for eid in tag.GetTaggedLocalElementIds():
                if eid and eid.IntegerValue != -1:
                    tagged_ids.add(eid.IntegerValue)
            continue
        except:
            pass
        try:
            eid = tag.TaggedLocalElementId
            if eid and eid.IntegerValue != -1:
                tagged_ids.add(eid.IntegerValue)
        except:
            pass
    return tagged_ids

def build_multicategory_filter(categories):
    cat_list = List[BuiltInCategory]()
    for cat in categories:
        cat_list.Add(cat)
    return ElementMulticategoryFilter(cat_list)

def get_tag_text(tag):
    try:
        txt = tag.TagText
        if txt:
            return txt.strip()
    except:
        pass
    return "TAG"

def estimate_label_box_from_head(head_point, tag_text):
    char_count = max(len(tag_text), 1)
    width = BASE_LABEL_WIDTH + (char_count * PER_CHAR_WIDTH) + LABEL_PADDING_X
    height = LABEL_HEIGHT + LABEL_PADDING_Y
    half_w = width / 2.0
    half_h = height / 2.0
    return (
        head_point.X - half_w,
        head_point.Y - half_h,
        head_point.X + half_w,
        head_point.Y + half_h
    )

def boxes_overlap_2d(b1, b2):
    return not (
        b1[2] < b2[0] or
        b1[0] > b2[2] or
        b1[3] < b2[1] or
        b1[1] > b2[3]
    )

def collect_existing_tag_boxes(tags):
    boxes = []
    for tag in tags:
        try:
            head = tag.TagHeadPosition
            txt = get_tag_text(tag)
            box = estimate_label_box_from_head(head, txt)
            boxes.append(box)
        except:
            pass
    return boxes

def choose_tag_head_position(tag, occupied_boxes):
    base_head = tag.TagHeadPosition
    tag_text = get_tag_text(tag)
    fallback_head = base_head
    fallback_box = estimate_label_box_from_head(base_head, tag_text)

    for offset in CANDIDATE_OFFSETS:
        candidate = XYZ(
            base_head.X + offset.X,
            base_head.Y + offset.Y,
            base_head.Z + offset.Z
        )
        candidate_box = estimate_label_box_from_head(candidate, tag_text)
        overlap = any(boxes_overlap_2d(candidate_box, ex) for ex in occupied_boxes)
        if not overlap:
            return candidate, candidate_box

    return fallback_head, fallback_box

def get_selected_or_picked_elements():
    """Use preselection if available, otherwise prompt user to pick."""
    selected_ids = list(uidoc.Selection.GetElementIds())
    if selected_ids:
        return [doc.GetElement(el_id) for el_id in selected_ids]

    try:
        picked_refs = uidoc.Selection.PickObjects(
            ObjectType.Element,
            'Select Elements or Finish Button'
        )
        return [doc.GetElement(picked_ref.ElementId) for picked_ref in picked_refs]
    except:
        sys.exit()


# --------------------------------------------------
# WPF SELECTION WINDOW
# --------------------------------------------------
class PointNumberSelectWindow(Window):
    def __init__(self, point_data_dict, revit_window_handle):
        Window.__init__(self)
        self.Title = "Select TS Point Numbers to Tag"
        self.Width = 360
        self.Height = 450
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.Topmost = True

        self.point_data_dict = point_data_dict # Dict mapping point_number -> [elements]
        self.all_points = sorted(point_data_dict.keys(), key=natural_key)
        self.selected_points = []
        self.action_mode = None  # "list" or "fence"
        self.include_leader = True

        self.initialize_components()
        try:
            WindowInteropHelper(self).Owner = revit_window_handle
        except:
            pass

    def initialize_components(self):
        grid = Grid()
        grid.Margin = Thickness(10)
        self.Content = grid

        # Define Rows:
        # Row 0: Search Label (Auto)
        # Row 1: Search TextBox & Select All Button (Auto)
        # Row 2: List Label (Auto)
        # Row 3: ListBox (Star - stretches dynamically when resized)
        # Row 4: Checkbox Include Leader (Auto)
        # Row 5: Button Panel (Auto)
        for h in [
            System.Windows.GridLength.Auto,
            System.Windows.GridLength.Auto,
            System.Windows.GridLength.Auto,
            System.Windows.GridLength(1, System.Windows.GridUnitType.Star),
            System.Windows.GridLength.Auto,
            System.Windows.GridLength.Auto
        ]:
            grid.RowDefinitions.Add(RowDefinition(Height=h))

        # Define Columns:
        # Col 0: Main content (Star)
        # Col 1: Action buttons like Select All (Auto)
        grid.ColumnDefinitions.Add(ColumnDefinition(Width=System.Windows.GridLength(1, System.Windows.GridUnitType.Star)))
        grid.ColumnDefinitions.Add(ColumnDefinition(Width=System.Windows.GridLength.Auto))

        row_index = 0

        # Row 0: Label
        lbl = Label()
        lbl.Content = "Search Point Number:"
        lbl.Margin = Thickness(0, 0, 0, 2)
        Grid.SetRow(lbl, row_index)
        Grid.SetColumnSpan(lbl, 2)
        grid.Children.Add(lbl)
        row_index += 1

        # Row 1: TextBox (Col 0) and Select All Button (Col 1)
        self.textbox = TextBox()
        self.textbox.Height = 24
        self.textbox.VerticalContentAlignment = VerticalAlignment.Center
        self.textbox.Margin = Thickness(0, 0, 5, 5)
        self.textbox.TextChanged += self.on_text_changed
        Grid.SetRow(self.textbox, row_index)
        Grid.SetColumn(self.textbox, 0)
        grid.Children.Add(self.textbox)

        self.select_all_btn = Button()
        self.select_all_btn.Content = "Select All"
        self.select_all_btn.Width = 75
        self.select_all_btn.Height = 24
        self.select_all_btn.Margin = Thickness(0, 0, 0, 5)
        self.select_all_btn.Click += self.on_select_all_click
        Grid.SetRow(self.select_all_btn, row_index)
        Grid.SetColumn(self.select_all_btn, 1)
        grid.Children.Add(self.select_all_btn)
        row_index += 1

        # Row 2: List Label
        lbl2 = Label()
        lbl2.Content = "Select Points to Tag (Ctrl / Shift supported):"
        lbl2.Margin = Thickness(0, 5, 0, 2)
        Grid.SetRow(lbl2, row_index)
        Grid.SetColumnSpan(lbl2, 2)
        grid.Children.Add(lbl2)
        row_index += 1

        # Row 3: ListBox (Star sizing expands/contracts cleanly)
        self.listbox = ListBox()
        self.listbox.SelectionMode = SelectionMode.Extended
        self.listbox.Margin = Thickness(0, 0, 0, 8)
        Grid.SetRow(self.listbox, row_index)
        Grid.SetColumnSpan(self.listbox, 2)
        grid.Children.Add(self.listbox)
        row_index += 1

        # Row 4: Checkbox Include Leader
        self.checkbox_leader = CheckBox()
        self.checkbox_leader.Content = "Include Leader"
        self.checkbox_leader.IsChecked = True
        self.checkbox_leader.Margin = Thickness(0, 0, 0, 8)
        Grid.SetRow(self.checkbox_leader, row_index)
        Grid.SetColumnSpan(self.checkbox_leader, 2)
        grid.Children.Add(self.checkbox_leader)
        row_index += 1

        # Row 5: Action Button Panel
        btn_panel = StackPanel()
        btn_panel.Orientation = System.Windows.Controls.Orientation.Horizontal
        btn_panel.HorizontalAlignment = HorizontalAlignment.Center
        btn_panel.Margin = Thickness(0, 0, 0, 0)
        Grid.SetRow(btn_panel, row_index)
        Grid.SetColumnSpan(btn_panel, 2)
        grid.Children.Add(btn_panel)

        tag_btn = Button()
        tag_btn.Content = "Tag Selected"
        tag_btn.Width = 90
        tag_btn.Height = 26
        tag_btn.Margin = Thickness(4, 0, 4, 0)
        tag_btn.Click += self.on_tag_click
        btn_panel.Children.Add(tag_btn)

        fence_btn = Button()
        fence_btn.Content = "Fence"
        fence_btn.Width = 70
        fence_btn.Height = 26
        fence_btn.Margin = Thickness(4, 0, 4, 0)
        fence_btn.Click += self.on_fence_click
        btn_panel.Children.Add(fence_btn)

        cancel_btn = Button()
        cancel_btn.Content = "Cancel"
        cancel_btn.Width = 70
        cancel_btn.Height = 26
        cancel_btn.Margin = Thickness(4, 0, 4, 0)
        cancel_btn.Click += self.on_cancel_click
        btn_panel.Children.Add(cancel_btn)

        self.refresh_listbox("")

    def refresh_listbox(self, filter_text):
        self.listbox.Items.Clear()
        filter_text = (filter_text or "").strip().lower()
        for pt in self.all_points:
            if not filter_text or filter_text in pt.lower():
                self.listbox.Items.Add(pt)

    def on_text_changed(self, sender, event):
        self.refresh_listbox(self.textbox.Text)

    def on_select_all_click(self, sender, event):
        self.listbox.SelectAll()

    def on_fence_click(self, sender, event):
        self.action_mode = "fence"
        self.include_leader = bool(self.checkbox_leader.IsChecked)
        self.Close()

    def on_tag_click(self, sender, event):
        self.action_mode = "list"
        self.include_leader = bool(self.checkbox_leader.IsChecked)
        for item in self.listbox.SelectedItems:
            self.selected_points.append(str(item))
        self.Close()

    def on_cancel_click(self, sender, event):
        self.action_mode = "cancel"
        self.selected_points = []
        self.Close()


# --------------------------------------------------
# MAIN EXECUTION
# --------------------------------------------------
try:
    if isinstance(curview, View3D):
        TaskDialog.Show("Unsupported View", "Please run this tool from a 2D view (plan, section, or elevation).")
        sys.exit()

    multi_cat_filter = build_multicategory_filter(TARGET_CATEGORIES)
    elements_in_view = list(
        FilteredElementCollector(doc, curview.Id)
        .WherePasses(multi_cat_filter)
        .WhereElementIsNotElementType()
        .ToElements()
    )

    point_data_dict = {}
    valid_elements_map = {elem.Id.IntegerValue: elem for elem in elements_in_view}

    for elem in elements_in_view:
        if not is_parent_family_instance(elem):
            continue
        param = elem.LookupParameter(PARAM_NAME)
        if param and param.HasValue:
            val = param.AsString()
            if val:
                if val not in point_data_dict:
                    point_data_dict[val] = []
                point_data_dict[val].append(elem)

    if not point_data_dict:
        TaskDialog.Show("No Elements Found", "No elements with a valid '{}' parameter were found in the active view.".format(PARAM_NAME))
        sys.exit()

    handle = get_revit_window_handle()
    window = PointNumberSelectWindow(point_data_dict, handle)
    window.ShowDialog()

    elements_to_tag = []
    include_leader = window.include_leader

    if window.action_mode == "list":
        for pt_str in window.selected_points:
            elements_to_tag.extend(point_data_dict.get(pt_str, []))
    elif window.action_mode == "fence":
        picked_elems = get_selected_or_picked_elements()
        for elem in picked_elems:
            if elem:
                if not is_parent_family_instance(elem):
                    continue
                if elem.Id.IntegerValue not in valid_elements_map:
                    continue
                param = elem.LookupParameter(PARAM_NAME)
                if param and param.HasValue and param.AsString():
                    if elem not in elements_to_tag:
                        elements_to_tag.append(elem)

    if not elements_to_tag:
        sys.exit()

    families = FilteredElementCollector(doc).OfClass(Family)
    family_in_project = any(f.Name == FAMILY_NAME for f in families)

    tg = TransactionGroup(doc, "Tag Selected Elements with TS Point Number")
    tg.Start()

    t1 =Transaction(doc, "Load Tag Family")
    try:
        t1.Start()
        if not family_in_project:
            if not os.path.exists(FAMILY_PATH):
                raise Exception("Family file not found:\n{}".format(FAMILY_PATH))
            load_options = FamilyLoaderOptionsHandler()
            if not doc.LoadFamily(FAMILY_PATH, load_options):
                raise Exception("Revit could not load the family.")
        t1.Commit()
    except Exception as e:
        if t1.HasStarted():
            t1.RollBack()
        raise Exception("Failed to load family: {}".format(str(e)))

    tag_symbol = get_tag_symbol(doc, FAMILY_NAME, FAMILY_TYPE)
    if not tag_symbol:
        tg.RollBack()
        raise Exception("Could not find tag type '{}' in family '{}'.".format(FAMILY_TYPE, FAMILY_NAME))

    existing_tags = FilteredElementCollector(doc, curview.Id).OfClass(IndependentTag).ToElements()
    already_tagged_ids = get_all_tagged_element_ids(existing_tags)
    occupied_tag_boxes = collect_existing_tag_boxes(existing_tags)

    t2 = Transaction(doc, "Tag Selected Elements")
    try:
        t2.Start()
        if not tag_symbol.IsActive:
            tag_symbol.Activate()
            doc.Regenerate()

        tagged_count = 0

        for element in elements_to_tag:
            if element.Id.IntegerValue in already_tagged_ids:
                continue

            host_point = get_tag_point(element, curview)
            if not host_point:
                continue

            ref = Reference(element)
            new_tag = IndependentTag.Create(
                doc,
                tag_symbol.Id,
                curview.Id,
                ref,
                include_leader,
                TagOrientation.Horizontal,
                host_point
            )

            if not new_tag:
                continue

            doc.Regenerate()
            final_head, final_box = choose_tag_head_position(new_tag, occupied_tag_boxes)

            try:
                new_tag.TagHeadPosition = final_head
                doc.Regenerate()
            except:
                pass

            occupied_tag_boxes.append(final_box)
            already_tagged_ids.add(element.Id.IntegerValue)
            tagged_count += 1

        t2.Commit()
        tg.Assimilate()
        
        # Show balloon notification on completion
        show_balloon_notification("Tagging Complete", "Successfully tagged: {} elements.".format(tagged_count))

    except Exception as e:
        if t2.HasStarted():
            t2.RollBack()
        tg.RollBack()
        raise Exception("Failed while tagging elements: {}".format(str(e)))

except Exception as e:
    TaskDialog.Show("Error", str(e))
    raise