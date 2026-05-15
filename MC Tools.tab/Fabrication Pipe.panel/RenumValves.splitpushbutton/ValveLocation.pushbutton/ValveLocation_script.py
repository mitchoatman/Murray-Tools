# -*- coding: utf-8 -*-
import sys
import clr
from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory, Transaction,
    BuiltInParameter, RevitLinkInstance, FabricationPart
)
from Autodesk.Revit.DB.Architecture import Room
from Autodesk.Revit.UI import TaskDialog
from Parameters.Add_SharedParameters import Shared_Params
from Parameters.Get_Set_Params import set_parameter_by_name

clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference('WindowsBase')
clr.AddReference('System')

from System.Windows import (
    Window, Thickness, HorizontalAlignment, VerticalAlignment,
    ResizeMode, WindowStartupLocation, GridLength, GridUnitType
)
from System.Windows.Controls import (
    Button, CheckBox, Grid, RowDefinition, ColumnDefinition,
    Label, StackPanel, ScrollViewer, TextBox
)
from System.Windows.Media import Brushes, FontFamily
from System.Windows.Controls.Primitives import UniformGrid
from System.Windows.Input import Keyboard, ModifierKeys

# ── init ──────────────────────────────────────────────────────────────────────
Shared_Params()

doc = __revit__.ActiveUIDocument.Document
curview = doc.ActiveView

# ── collect fabrication valves in active view only ────────────────────────────
FAB_VALVE_SERVICE_TYPE = 53

all_pipework = (
    FilteredElementCollector(doc, curview.Id)
    .OfCategory(BuiltInCategory.OST_FabricationPipework)
    .WhereElementIsNotElementType()
    .ToElements()
)

elements_in_view = [
    p for p in all_pipework
    if isinstance(p, FabricationPart) and p.ServiceType == FAB_VALVE_SERVICE_TYPE
]

if not elements_in_view:
    TaskDialog.Show("Error", "No fabrication valves found in the current view.")
    sys.exit()

# ── collect all link instances ────────────────────────────────────────────────
all_link_instances = list(
    FilteredElementCollector(doc)
    .OfClass(RevitLinkInstance)
    .WhereElementIsNotElementType()
    .ToElements()
)

valid_links = [
    (lnk, lnk.GetLinkDocument())
    for lnk in all_link_instances
    if lnk.GetLinkDocument() is not None
]

if not valid_links:
    TaskDialog.Show("Error", "No loaded Revit links found in the model.")
    sys.exit()

# ── Link Selection Dialog ─────────────────────────────────────────────────────
class LinkSelectionForm(object):
    def __init__(self, link_pairs):
        self.link_pairs = link_pairs
        self.selected = []
        self.checkboxes = []
        self.check_all_state = False
        self.last_checked_index = None
        self._confirmed = False
        self._build()

    def _build(self):
        win = Window()
        win.Title = "Select Linked Models"
        win.Width = 440
        win.Height = 420
        win.MinWidth = win.Width
        win.MinHeight = win.Height
        win.ResizeMode = ResizeMode.NoResize
        win.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self._window = win

        root = Grid()
        root.Margin = Thickness(10)

        for h in [
            GridLength.Auto,
            GridLength.Auto,
            GridLength(1, GridUnitType.Star),
            GridLength.Auto
        ]:
            rd = RowDefinition()
            rd.Height = h
            root.RowDefinitions.Add(rd)

        root.ColumnDefinitions.Add(ColumnDefinition())

        lbl = Label()
        lbl.Content = "Choose linked model(s) to read rooms from:"
        lbl.FontFamily = FontFamily("Segoe UI")
        lbl.FontSize = 14
        lbl.Margin = Thickness(0, 0, 0, 4)
        Grid.SetRow(lbl, 0)
        root.Children.Add(lbl)

        self._search = TextBox()
        self._search.Height = 22
        self._search.FontFamily = FontFamily("Segoe UI")
        self._search.FontSize = 12
        self._search.Margin = Thickness(0, 0, 0, 4)
        self._search.TextChanged += self._on_search
        Grid.SetRow(self._search, 1)
        root.Children.Add(self._search)

        self._panel = StackPanel()

        sv = ScrollViewer()
        sv.Content = self._panel
        sv.VerticalAlignment = VerticalAlignment.Stretch
        Grid.SetRow(sv, 2)
        root.Children.Add(sv)

        self._populate(self.link_pairs)

        btn_grid = UniformGrid()
        btn_grid.Columns = 3
        btn_grid.HorizontalAlignment = HorizontalAlignment.Center
        btn_grid.Margin = Thickness(0, 4, 0, 0)

        for label_text, handler, color in [
            ("All / None", self._toggle_all, None),
            ("Cancel", self._cancel, Brushes.IndianRed),
            ("OK", self._ok, Brushes.SteelBlue),
        ]:
            btn = Button()
            btn.Content = label_text
            btn.FontFamily = FontFamily("Segoe UI")
            btn.FontSize = 12
            btn.Height = 28
            btn.Margin = Thickness(5, 0, 5, 0)
            if color:
                btn.Background = color
                btn.Foreground = Brushes.White
            btn.Click += handler
            btn_grid.Children.Add(btn)

        Grid.SetRow(btn_grid, 3)
        root.Children.Add(btn_grid)

        win.Content = root

    def _get_display_name(self, link, link_doc):
        try:
            return link_doc.Title
        except:
            return str(link.Id)

    def _populate(self, pairs):
        self._panel.Children.Clear()
        self.checkboxes = []

        already_selected_names = set(
            self._get_display_name(lnk, ld) for lnk, ld in self.selected
        )

        for lnk, ld in pairs:
            name = self._get_display_name(lnk, ld)
            cb = CheckBox()
            cb.Content = name
            cb.Tag = (lnk, ld)
            cb.FontFamily = FontFamily("Segoe UI")
            cb.FontSize = 12
            cb.Margin = Thickness(4, 2, 4, 2)
            cb.IsChecked = name in already_selected_names
            cb.Checked += self._cb_changed
            cb.Unchecked += self._cb_changed
            self._panel.Children.Add(cb)
            self.checkboxes.append(cb)

    def _on_search(self, sender, args):
        q = self._search.Text.lower()
        filtered = [
            (lnk, ld) for lnk, ld in self.link_pairs
            if q in self._get_display_name(lnk, ld).lower()
        ]
        self._populate(filtered)

    def _cb_changed(self, sender, args):
        try:
            idx = self.checkboxes.index(sender)
            if Keyboard.Modifiers == ModifierKeys.Shift and self.last_checked_index is not None:
                start = min(self.last_checked_index, idx)
                end = max(self.last_checked_index, idx)
                state = sender.IsChecked
                for i in range(start, end + 1):
                    self.checkboxes[i].IsChecked = state
            self.last_checked_index = idx
            self.selected = [cb.Tag for cb in self.checkboxes if cb.IsChecked]
        except Exception as ex:
            print("Checkbox error: {}".format(ex))

    def _toggle_all(self, sender, args):
        self.check_all_state = not self.check_all_state
        for cb in self.checkboxes:
            cb.IsChecked = self.check_all_state
        self.selected = [cb.Tag for cb in self.checkboxes if cb.IsChecked]

    def _ok(self, sender, args):
        self.selected = [cb.Tag for cb in self.checkboxes if cb.IsChecked]
        if not self.selected:
            TaskDialog.Show("Warning", "Please select at least one linked model.")
            return
        self._confirmed = True
        self._window.Close()

    def _cancel(self, sender, args):
        self._confirmed = False
        self._window.Close()

    def ShowDialog(self):
        self._window.ShowDialog()
        return self._confirmed

# ── show dialog ───────────────────────────────────────────────────────────────
form = LinkSelectionForm(valid_links)
confirmed = form.ShowDialog()

if not confirmed or not form.selected:
    sys.exit()

chosen_links = form.selected

# ── collect rooms only from chosen links ──────────────────────────────────────
linked_rooms = []
for lnk, link_doc in chosen_links:
    xform = lnk.GetTotalTransform()
    rooms = (
        FilteredElementCollector(link_doc)
        .OfCategory(BuiltInCategory.OST_Rooms)
        .WhereElementIsNotElementType()
        .ToElements()
    )

    for room in rooms:
        if isinstance(room, Room):
            linked_rooms.append((room, xform))

if not linked_rooms:
    TaskDialog.Show("Error", "No Rooms found in the selected linked model(s).")
    sys.exit()

# ── helpers ───────────────────────────────────────────────────────────────────
def get_element_point(e):
    try:
        return e.Origin
    except:
        return None

# ── assign FP_Location from linked rooms ──────────────────────────────────────
t = Transaction(doc, "FP_Location <- Linked Room")
t.Start()

for e in elements_in_view:
    pt_host = get_element_point(e)
    if not pt_host:
        continue

    fp = e.LookupParameter("FP_Location")
    if fp is None:
        continue

    for room, xform in linked_rooms:
        pt_link = xform.Inverse.OfPoint(pt_host)

        try:
            if room.IsPointInRoom(pt_link):
                name_param = room.get_Parameter(BuiltInParameter.ROOM_NAME)
                room_name = name_param.AsString() if name_param else ""

                if not room_name:
                    lp = room.LookupParameter("Name")
                    room_name = lp.AsString() if lp else ""

                if room_name:
                    set_parameter_by_name(e, "FP_Location", room_name)
                break
        except:
            pass

t.Commit()