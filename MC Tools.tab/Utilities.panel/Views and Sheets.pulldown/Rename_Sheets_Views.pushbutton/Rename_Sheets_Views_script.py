# coding: utf8
from Autodesk.Revit.DB import (
    ViewSheet,
    View,
    ViewType,
    FilteredElementCollector,
    Transaction
)
from Autodesk.Revit.UI import TaskDialog
import os
import clr
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference('System.Xaml')
import System
from System.Windows import (
    Window, Thickness, ResizeMode, HorizontalAlignment, VerticalAlignment,
    GridLength, GridUnitType, FontWeights, WindowStartupLocation,
    TextTrimming, SystemParameters
)
from System.Windows.Controls import (
    Grid, StackPanel, ScrollViewer, TextBox, TextBlock, Button, CheckBox,
    Border, ColumnDefinition, RowDefinition, Orientation, WrapPanel,
    ScrollBarVisibility
)
from System.Windows.Media import Brushes, FontFamily
from System.Windows.Forms import NotifyIcon, ToolTipIcon
from System.Drawing import Icon, SystemIcons
from System.Windows.Input import Keyboard, Key

uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document

# Chip order and the view types that are included
GROUP_ORDER = ["SheetName", "SheetNumber", "Schedule", "FloorPlan", "CeilingPlan",
               "ThreeD", "Section", "Elevation", "DraftingView"]
# Display labels for chips whose key differs from the label
CHIP_LABELS = {"SheetName": "Sheet Name", "SheetNumber": "Sheet Number"}
VIEW_TYPES = [ViewType.FloorPlan, ViewType.CeilingPlan, ViewType.Elevation,
              ViewType.Section, ViewType.ThreeD, ViewType.DraftingView,
              ViewType.Schedule]


# Set to your icon path if desired, e.g. r"C:\path\to\Murray.ico"
icon_file = None


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


class Item(object):
    """One renamable value: a view name, a sheet name, or a sheet number."""
    def __init__(self, element, kind, group, display):
        self.element = element
        self.kind = kind            # "name" or "number"
        self.group = group          # chip key
        self.display = display      # text for the TYPE column
        self.checked = False
        self.new_name = ""
        self.status = "UNCHANGED"
        self.cb = None
        self.new_tb = None
        self.status_tb = None
        self.status_border = None

    @property
    def current(self):
        if self.kind == "number":
            return self.element.SheetNumber
        return self.element.Name


def collect_items():
    items = []

    sheets = sorted(FilteredElementCollector(doc).OfClass(ViewSheet).ToElements(),
                    key=lambda s: s.SheetNumber)
    for s in sheets:
        items.append(Item(s, "name", "SheetName", "DrawingSheet (Name)"))
        items.append(Item(s, "number", "SheetNumber", "DrawingSheet (Number)"))

    views = []
    for v in FilteredElementCollector(doc).OfClass(View).ToElements():
        if isinstance(v, ViewSheet) or v.IsTemplate or v.ViewType not in VIEW_TYPES:
            continue
        # skip system schedules that cannot be renamed
        if getattr(v, "IsTitleblockRevisionSchedule", False) or getattr(v, "IsInternalKeynoteSchedule", False):
            continue
        views.append(v)

    def sort_key(v):
        g = str(v.ViewType)
        idx = GROUP_ORDER.index(g) if g in GROUP_ORDER else len(GROUP_ORDER)
        return (idx, v.Name.lower())

    for v in sorted(views, key=sort_key):
        g = str(v.ViewType)
        items.append(Item(v, "name", g, g))
    return items


def bordered_box(title, *controls):
    """Small titled input group (top bar accent in black)."""
    border = Border()
    border.BorderBrush = Brushes.Black
    border.BorderThickness = Thickness(1, 3, 1, 1)
    border.Padding = Thickness(8, 6, 8, 8)
    border.Margin = Thickness(3, 0, 3, 0)
    stack = StackPanel(Orientation=Orientation.Vertical)
    head = TextBlock()
    head.Text = title
    head.FontWeight = FontWeights.Bold
    head.FontSize = 10
    head.Margin = Thickness(0, 0, 0, 5)
    stack.Children.Add(head)
    for c in controls:
        stack.Children.Add(c)
    border.Child = stack
    return border


def make_textbox(hint):
    tb = TextBox()
    tb.Height = 24
    tb.Padding = Thickness(4, 2, 4, 2)
    tb.VerticalContentAlignment = VerticalAlignment.Center
    tb.BorderBrush = Brushes.Black
    tb.Margin = Thickness(0, 0, 0, 4)
    tb.ToolTip = hint
    return tb


def table_columns(grid):
    for width, star in [(32, False), (170, False), (1, True), (1, True), (110, False)]:
        cd = ColumnDefinition()
        cd.Width = GridLength(width, GridUnitType.Star) if star else GridLength(width)
        grid.ColumnDefinitions.Add(cd)


def cell_text(text, col, bold=False, size=12):
    tb = TextBlock()
    tb.Text = text
    tb.FontSize = size
    tb.VerticalAlignment = VerticalAlignment.Center
    tb.TextTrimming = TextTrimming.CharacterEllipsis
    tb.Margin = Thickness(4, 0, 8, 0)
    if bold:
        tb.FontWeight = FontWeights.Bold
    tb.ToolTip = text
    Grid.SetColumn(tb, col)
    return tb


class RenamerForm(Window):
    def __init__(self, items):
        self.items = items
        self.visible = []
        self.active_group = "ALL"
        self.chips = {}
        self.last_clicked_index = None
        self.ready = False
        self.init_components()
        self.ready = True
        self.rebuild_rows()

    # ------------------------------------------------------------------ UI
    def init_components(self):
        self.Title = "Renamer"
        self.Width = 780
        self.Height = 650
        self.MinWidth = 700
        self.MinHeight = 520
        self.ResizeMode = ResizeMode.CanResize
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.FontFamily = FontFamily("Arial")
        self.FontSize = 12
        self.Background = Brushes.White

        root = Grid()
        root.Margin = Thickness(12)
        for i in range(5):
            rd = RowDefinition()
            rd.Height = GridLength(1, GridUnitType.Star) if i == 3 else GridLength.Auto
            root.RowDefinitions.Add(rd)

        # Row 0: Find/Replace, Prefix, Suffix, Search
        top = Grid()
        for i in range(4):
            top.ColumnDefinitions.Add(ColumnDefinition())
        self.find_input = make_textbox("Find (case sensitive)")
        self.replace_input = make_textbox("Replace")
        self.prefix_input = make_textbox("e.g. L1-")
        self.suffix_input = make_textbox("e.g. -DRAFT")
        self.search_box = make_textbox("Filter by current name")
        boxes = [
            bordered_box("FIND / REPLACE", self.find_input, self.replace_input),
            bordered_box("PREFIX", self.prefix_input),
            bordered_box("SUFFIX", self.suffix_input),
            bordered_box("SEARCH", self.search_box),
        ]
        for i, b in enumerate(boxes):
            Grid.SetColumn(b, i)
            top.Children.Add(b)
        Grid.SetRow(top, 0)
        root.Children.Add(top)

        for tb in (self.find_input, self.replace_input, self.prefix_input, self.suffix_input):
            tb.TextChanged += self.inputs_changed
        self.search_box.TextChanged += self.search_changed

        # Row 1: type chips
        chip_panel = WrapPanel()
        chip_panel.Margin = Thickness(3, 12, 3, 8)
        present = []
        for it in self.items:
            if it.group not in present:
                present.append(it.group)
        for key in ["ALL"] + present:
            btn = Button()
            btn.Content = CHIP_LABELS.get(key, key)
            btn.Tag = key
            btn.Padding = Thickness(10, 3, 10, 3)
            btn.Margin = Thickness(0, 0, 6, 4)
            btn.BorderBrush = Brushes.Black
            btn.FontSize = 11
            btn.Click += self.chip_clicked
            chip_panel.Children.Add(btn)
            self.chips[key] = btn
        Grid.SetRow(chip_panel, 1)
        root.Children.Add(chip_panel)
        self.style_chips()

        # Row 2: table header
        header_border = Border()
        header_border.BorderBrush = Brushes.Black
        header_border.BorderThickness = Thickness(1, 1, 1, 2)
        header_border.Padding = Thickness(0, 6, 0, 6)
        header_border.Margin = Thickness(3, 0, 3, 0)
        header = Grid()
        header.Margin = Thickness(0, 0, SystemParameters.VerticalScrollBarWidth, 0)
        table_columns(header)
        self.check_all = CheckBox()
        self.check_all.VerticalAlignment = VerticalAlignment.Center
        self.check_all.HorizontalAlignment = HorizontalAlignment.Center
        self.check_all.Click += self.check_all_clicked
        Grid.SetColumn(self.check_all, 0)
        header.Children.Add(self.check_all)
        for col, text in [(1, "TYPE"), (2, "CURRENT NAME"), (3, "NEW NAME PREVIEW"), (4, "STATUS")]:
            header.Children.Add(cell_text(text, col, bold=True, size=10))
        header_border.Child = header
        Grid.SetRow(header_border, 2)
        root.Children.Add(header_border)

        # Row 3: scrollable rows
        self.row_panel = StackPanel(Orientation=Orientation.Vertical)
        scroll = ScrollViewer()
        scroll.Content = self.row_panel
        scroll.VerticalScrollBarVisibility = ScrollBarVisibility.Visible
        scroll.HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled
        scroll_border = Border()
        scroll_border.BorderBrush = Brushes.Black
        scroll_border.BorderThickness = Thickness(1, 0, 1, 1)
        scroll_border.Margin = Thickness(3, 0, 3, 0)
        scroll_border.Child = scroll
        Grid.SetRow(scroll_border, 3)
        root.Children.Add(scroll_border)

        # Row 4: rename button
        self.rename_button = Button()
        self.rename_button.Height = 36
        self.rename_button.Margin = Thickness(3, 12, 3, 0)
        self.rename_button.FontWeight = FontWeights.Bold
        self.rename_button.FontSize = 13
        self.rename_button.Background = Brushes.White
        self.rename_button.Foreground = Brushes.Black
        self.rename_button.BorderBrush = Brushes.Black
        self.rename_button.BorderThickness = Thickness(2)
        self.rename_button.Click += self.rename_clicked
        Grid.SetRow(self.rename_button, 4)
        root.Children.Add(self.rename_button)

        self.Content = root

    def style_chips(self):
        for key, btn in self.chips.items():
            if key == self.active_group:
                btn.Background = Brushes.Black
                btn.Foreground = Brushes.White
            else:
                btn.Background = Brushes.White
                btn.Foreground = Brushes.Black

    # ------------------------------------------------------------ rows
    def rebuild_rows(self):
        search = self.search_box.Text.lower().strip()
        self.row_panel.Children.Clear()
        self.visible = []
        for it in self.items:
            it.cb = None
            it.new_tb = None
            it.status_tb = None
            it.status_border = None

        for it in self.items:
            if self.active_group != "ALL" and it.group != self.active_group:
                continue
            if search and search not in it.current.lower():
                continue
            self.visible.append(it)
            self.row_panel.Children.Add(self.build_row(it))

        self.last_clicked_index = None
        self.refresh()

    def build_row(self, it):
        border = Border()
        border.BorderBrush = Brushes.Black
        border.BorderThickness = Thickness(0, 0, 0, 1)
        border.Padding = Thickness(0, 8, 0, 8)
        row = Grid()
        table_columns(row)

        cb = CheckBox()
        cb.IsChecked = it.checked
        cb.Tag = it
        cb.VerticalAlignment = VerticalAlignment.Center
        cb.HorizontalAlignment = HorizontalAlignment.Center
        cb.Click += self.checkbox_clicked
        Grid.SetColumn(cb, 0)
        row.Children.Add(cb)
        it.cb = cb

        row.Children.Add(cell_text(it.display, 1))
        row.Children.Add(cell_text(it.current, 2))

        new_tb = cell_text(it.current, 3, bold=True)
        row.Children.Add(new_tb)
        it.new_tb = new_tb

        badge = Border()
        badge.BorderBrush = Brushes.Black
        badge.BorderThickness = Thickness(1)
        badge.Padding = Thickness(6, 1, 6, 1)
        badge.HorizontalAlignment = HorizontalAlignment.Left
        badge.VerticalAlignment = VerticalAlignment.Center
        badge_tb = TextBlock()
        badge_tb.FontSize = 10
        badge_tb.FontWeight = FontWeights.Bold
        badge.Child = badge_tb
        Grid.SetColumn(badge, 4)
        row.Children.Add(badge)
        it.status_tb = badge_tb
        it.status_border = badge

        border.Child = row
        return border

    # ------------------------------------------------------- preview logic
    def compute_new(self, it):
        s = it.current
        find = self.find_input.Text
        if find:
            s = s.replace(find, self.replace_input.Text)
        return self.prefix_input.Text + s + self.suffix_input.Text

    def refresh(self):
        if not self.ready:
            return

        # new names for every item
        for it in self.items:
            it.new_name = self.compute_new(it)

        # items that would actually change
        changing = [it for it in self.items if it.checked and it.new_name != it.current]

        # names that stay as they are, per uniqueness scope
        taken = {}
        for it in self.items:
            if it in changing:
                continue
            taken.setdefault((it.group, it.kind), set()).add(it.current)

        new_counts = {}
        for it in changing:
            k = (it.group, it.kind, it.new_name)
            new_counts[k] = new_counts.get(k, 0) + 1

        for it in self.items:
            if it.new_name == it.current:
                it.status = "UNCHANGED"
            elif not it.checked:
                it.status = "NOT SELECTED"
            else:
                conflict = not it.new_name.strip()
                # sheet names may repeat in Revit; views and sheet numbers may not
                if it.group != "SheetName":
                    if it.new_name in taken.get((it.group, it.kind), set()):
                        conflict = True
                    if new_counts.get((it.group, it.kind, it.new_name), 0) > 1:
                        conflict = True
                it.status = "CONFLICT" if conflict else "WILL RENAME"

        # update widgets that are on screen
        for it in self.visible:
            if it.new_tb is None:
                continue
            it.new_tb.Text = it.new_name
            it.new_tb.ToolTip = it.new_name
            it.status_tb.Text = it.status
            if it.status == "WILL RENAME":
                it.status_border.Background = Brushes.Black
                it.status_tb.Foreground = Brushes.White
            else:
                it.status_border.Background = Brushes.White
                it.status_tb.Foreground = Brushes.Black
            it.status_border.BorderThickness = Thickness(2) if it.status == "CONFLICT" else Thickness(1)

        count = len([it for it in self.items if it.status == "WILL RENAME"])
        self.rename_button.Content = "RENAME {} ITEM(S)".format(count)
        self.rename_button.IsEnabled = count > 0

        if self.visible:
            self.check_all.IsChecked = all(it.checked for it in self.visible)
        else:
            self.check_all.IsChecked = False

    # ------------------------------------------------------------ events
    def inputs_changed(self, sender, args):
        self.refresh()

    def search_changed(self, sender, args):
        if self.ready:
            self.rebuild_rows()

    def chip_clicked(self, sender, args):
        self.active_group = sender.Tag
        self.style_chips()
        self.rebuild_rows()

    def check_all_clicked(self, sender, args):
        state = bool(self.check_all.IsChecked)
        for it in self.visible:
            it.checked = state
            if it.cb is not None:
                it.cb.IsChecked = state
        self.refresh()

    def checkbox_clicked(self, sender, args):
        it = sender.Tag
        try:
            current_index = self.visible.index(it)
        except ValueError:
            return

        shift = Keyboard.IsKeyDown(Key.LeftShift) or Keyboard.IsKeyDown(Key.RightShift)
        state = bool(sender.IsChecked)
        if shift and self.last_clicked_index is not None:
            start = min(self.last_clicked_index, current_index)
            end = max(self.last_clicked_index, current_index)
            for i in range(start, end + 1):
                v = self.visible[i]
                v.checked = state
                if v.cb is not None:
                    v.cb.IsChecked = state
        else:
            it.checked = state

        self.last_clicked_index = current_index
        self.refresh()

    def rename_clicked(self, sender, args):
        targets = [it for it in self.items if it.status == "WILL RENAME"]
        if not targets:
            TaskDialog.Show("Rename Operation", "Nothing to rename.")
            return

        t = Transaction(doc, 'Rename Sheets and Views')
        failed = []
        renamed = 0
        try:
            t.Start()
            for it in targets:
                try:
                    if it.kind == "number":
                        it.element.SheetNumber = it.new_name
                    else:
                        it.element.Name = it.new_name
                    renamed += 1
                except Exception as ex:
                    failed.append("{}: {}".format(it.current, str(ex)))
            t.Commit()
        except Exception as e:
            t.RollBack()
            TaskDialog.Show("Rename Error", "Operation aborted due to error: {}".format(str(e)))
            return

        if failed:
            # failures need to be read in full, so keep the dialog for those
            TaskDialog.Show("Rename Operation",
                            "{} item(s) renamed.\n\n{} failed:\n".format(renamed, len(failed)) + "\n".join(failed[:15]))
        else:
            show_balloon_notification("Renamer", "{} item(s) renamed.".format(renamed), icon_path=icon_file)
        self.Close()


items = collect_items()
if items:
    form = RenamerForm(items)
    form.ShowDialog()
else:
    TaskDialog.Show("No Elements Found", "No valid sheets or views found in the document.")