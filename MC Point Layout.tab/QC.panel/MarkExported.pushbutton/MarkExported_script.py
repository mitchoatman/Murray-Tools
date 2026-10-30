# -*- coding: UTF-8 -*-
"""
Mark Exported Points
Reads a points CSV (POINT NUMBER, Y, X, Z, DESCRIPTION), matches each row to
fabrication hanger rod ends / GTP generic models in the ACTIVE VIEW by shared
coordinates, then applies a color override in the view and/or writes a tag
into Comments on the PARENT element (e.g. the sleeve, not the nested GTP).

Coordinates are computed exactly like the export script (shared coordinates,
feet), so a CSV produced from that export will match.
"""
import os
import sys
import csv
import math
import clr
from collections import defaultdict

clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

from System.Windows import (
    Window, WindowStartupLocation, ResizeMode, SizeToContent, GridLength,
    GridUnitType, HorizontalAlignment, VerticalAlignment, TextWrapping,
    Thickness, FontWeights
)
from System.Windows.Controls import (
    Grid, ColumnDefinition, StackPanel, Orientation, Button, TextBox,
    TextBlock, CheckBox, RadioButton, GroupBox, ComboBox, Border
)
from System.Windows.Media import Color as WpfColor, SolidColorBrush, Brushes
from System.Windows.Interop import WindowInteropHelper
from Microsoft.Win32 import OpenFileDialog
from System.Windows.Forms import ColorDialog, DialogResult as WinDialogResult
from System.Drawing import Color as DrawColor

from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory, BuiltInParameter,
    Transaction, OverrideGraphicSettings, FillPatternElement,
    FillPatternTarget
)
from Autodesk.Revit.DB import Color as RvtColor
from Autodesk.Revit.UI import TaskDialog

doc = __revit__.ActiveUIDocument.Document
curview = doc.ActiveView
RevitINT = float(doc.Application.VersionNumber)

# Set from the dialog before matching starts
TOL_FT = 0.125 / 12.0
CELL = TOL_FT * 2.0


# ---- helpers (same logic as the export script) -------------------------------

def get_id_value(id_obj):
    if id_obj is None:
        return None
    try:
        return id_obj.Value if RevitINT > 2025 else id_obj.IntegerValue
    except:
        try:
            return id_obj.Value
        except:
            return id_obj.IntegerValue


def is_fab_hanger(element):
    try:
        return element and element.Category and \
            get_id_value(element.Category.Id) == int(BuiltInCategory.OST_FabricationHangers)
    except:
        return False


def is_gtp_element(element):
    if not element or not element.Category:
        return False
    if get_id_value(element.Category.Id) != int(BuiltInCategory.OST_GenericModel):
        return False
    try:
        symbol = element.Symbol
        if symbol:
            if getattr(symbol, 'FamilyName', '') == "GTP":
                return True
            if hasattr(symbol, 'Family') and symbol.Family and symbol.Family.Name == "GTP":
                return True
    except:
        pass
    return False


def clean_z(val, tol=1e-9):
    try:
        return 0.0 if abs(val) < tol else val
    except:
        return val


def get_shared_coords(point):
    if point is None:
        return None, None, None
    try:
        pp = doc.ActiveProjectLocation.GetProjectPosition(point)
        return pp.NorthSouth, pp.EastWest, clean_z(pp.Elevation)
    except:
        return point.Y, point.X, clean_z(point.Z)


def element_points(element):
    """Return a list of (east, north, elev) in feet for the element's export points."""
    pts = []
    try:
        if is_fab_hanger(element):
            info = element.GetRodInfo()
            if info is None:
                return pts
            for i in range(info.RodCount):
                pos = info.GetRodEndPosition(i)
                n, e, z = get_shared_coords(pos)
                if n is not None:
                    pts.append((e, n, z))
        else:
            loc = element.Location
            if hasattr(loc, 'Point') and loc.Point is not None:
                pos = loc.Point
            else:
                pos = element.GetTransform().Origin
            n, e, z = get_shared_coords(pos)
            if n is not None:
                pts.append((e, n, z))
    except:
        pass
    return pts


def get_owner(element):
    """Top-level parent of a nested family instance (the sleeve, not the GTP)."""
    current = element
    visited = set()
    while current:
        cid = get_id_value(current.Id)
        if cid in visited:
            break
        visited.add(cid)
        try:
            sup = current.SuperComponent
            if sup:
                current = sup
                continue
        except:
            pass
        break
    return current


def collect_elements():
    hangers = list(
        FilteredElementCollector(doc, curview.Id)
        .OfCategory(BuiltInCategory.OST_FabricationHangers)
        .WhereElementIsNotElementType().ToElements())
    generics = list(
        FilteredElementCollector(doc, curview.Id)
        .OfCategory(BuiltInCategory.OST_GenericModel)
        .WhereElementIsNotElementType().ToElements())
    return hangers + [el for el in generics if is_gtp_element(el)]


# ---- CSV and matching ---------------------------------------------------------

def load_csv(path, to_feet):
    if sys.version_info[0] == 2:
        f = open(path, 'rb')
    else:
        f = open(path, 'r', newline='', encoding='utf-8-sig')

    def norm(s):
        return s.replace('\xef\xbb\xbf', '').replace(u'\ufeff', '').strip().upper()

    points, bad = [], 0
    with f:
        reader = csv.reader(f)
        header = [norm(h) for h in next(reader)]
        for name in ("POINT NUMBER", "X", "Y", "Z"):
            if name not in header:
                return None, "Column '{}' not found. Header read: {}".format(name, header)
        ix = dict((name, header.index(name)) for name in ("POINT NUMBER", "X", "Y", "Z"))
        for row in reader:
            if not row or not any(c.strip() for c in row):
                continue
            try:
                points.append({
                    "num": row[ix["POINT NUMBER"]].strip(),
                    "x": float(row[ix["X"]]) * to_feet,
                    "y": float(row[ix["Y"]]) * to_feet,
                    "z": float(row[ix["Z"]]) * to_feet,
                })
            except:
                bad += 1
    return points, bad


def cell_of(x, y, z):
    return (int(math.floor(x / CELL)), int(math.floor(y / CELL)), int(math.floor(z / CELL)))


def build_index(points):
    index = defaultdict(list)
    for i, p in enumerate(points):
        index[cell_of(p["x"], p["y"], p["z"])].append(i)
    return index


def find_match(points, index, x, y, z):
    cx, cy, cz = cell_of(x, y, z)
    best, best_d = None, None
    for dx in (-1, 0, 1):
        for dy in (-1, 0, 1):
            for dz in (-1, 0, 1):
                for i in index.get((cx + dx, cy + dy, cz + dz), ()):
                    p = points[i]
                    if abs(p["x"] - x) <= TOL_FT and abs(p["y"] - y) <= TOL_FT and abs(p["z"] - z) <= TOL_FT:
                        d = (p["x"] - x) ** 2 + (p["y"] - y) ** 2 + (p["z"] - z) ** 2
                        if best is None or d < best_d:
                            best, best_d = i, d
    return best


# ---- override and comment helpers --------------------------------------------

def make_ogs(rgb, halftone, solid_fill):
    c = RvtColor(rgb[0], rgb[1], rgb[2])
    ogs = OverrideGraphicSettings()
    ogs.SetProjectionLineColor(c)
    ogs.SetCutLineColor(c)
    if halftone:
        ogs.SetHalftone(True)
    if solid_fill:
        try:
            for fp in FilteredElementCollector(doc).OfClass(FillPatternElement):
                pat = fp.GetFillPattern()
                if pat.IsSolidFill and pat.Target == FillPatternTarget.Drafting:
                    ogs.SetSurfaceForegroundPatternId(fp.Id)
                    ogs.SetSurfaceForegroundPatternColor(c)
                    break
        except:
            pass
    return ogs


def get_comments_param(el):
    p = el.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
    if p is None:
        p = el.LookupParameter("Comments")
    return p


def strip_tag(existing, tag):
    parts = [s.strip() for s in existing.split("|")]
    return " | ".join(p for p in parts if p and not p.upper().startswith(tag.upper()))


def add_tag(existing, tag, tag_text):
    existing = existing or ""
    base = strip_tag(existing, tag) if tag.upper() in existing.upper() else existing.strip()
    return (base + " | " + tag_text) if base else tag_text


# ---- WPF dialog (built in code, no XAML) ------------------------------------

PRESETS = [
    ("Grey", (150, 150, 150)),
    ("Red", (220, 40, 40)),
    ("Orange", (240, 140, 0)),
    ("Green", (40, 160, 70)),
    ("Blue", (40, 100, 220)),
    ("Magenta", (200, 40, 180)),
]


def brush_of(rgb):
    return SolidColorBrush(WpfColor.FromRgb(rgb[0], rgb[1], rgb[2]))


def make_group(title):
    g = GroupBox()
    g.Header = title
    g.Margin = Thickness(0, 0, 0, 8)
    g.Padding = Thickness(6)
    sp = StackPanel()
    g.Content = sp
    return g, sp


def make_row():
    sp = StackPanel()
    sp.Orientation = Orientation.Horizontal
    sp.Margin = Thickness(0, 3, 0, 3)
    return sp


def make_text(text, width=None):
    tb = TextBlock()
    tb.Text = text
    tb.Margin = Thickness(0, 0, 8, 0)
    tb.VerticalAlignment = VerticalAlignment.Center
    if width:
        tb.Width = width
    return tb


class MarkExportedWindow(Window):
    def __init__(self):
        self.Title = "Mark Exported Points"
        self.Width = 500
        self.SizeToContent = SizeToContent.Height
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.ResizeMode = ResizeMode.NoResize
        self.settings = None
        self.color = (150, 150, 150)

        try:
            WindowInteropHelper(self).Owner = __revit__.MainWindowHandle
        except:
            pass

        root = StackPanel()
        root.Margin = Thickness(12)
        self.Content = root

        # CSV file
        g_csv, sp_csv = make_group("Points CSV")
        grid = Grid()
        c0 = ColumnDefinition()
        c0.Width = GridLength(1, GridUnitType.Star)
        c1 = ColumnDefinition()
        c1.Width = GridLength.Auto
        grid.ColumnDefinitions.Add(c0)
        grid.ColumnDefinitions.Add(c1)
        self.txt_csv = TextBox()
        self.txt_csv.IsReadOnly = True
        self.txt_csv.Height = 24
        self.txt_csv.VerticalContentAlignment = VerticalAlignment.Center
        Grid.SetColumn(self.txt_csv, 0)
        grid.Children.Add(self.txt_csv)
        btn_browse = Button()
        btn_browse.Content = "Browse..."
        btn_browse.Width = 80
        btn_browse.Height = 24
        btn_browse.Margin = Thickness(6, 0, 0, 0)
        btn_browse.Click += self.on_browse
        Grid.SetColumn(btn_browse, 1)
        grid.Children.Add(btn_browse)
        sp_csv.Children.Add(grid)
        row_units = make_row()
        row_units.Children.Add(make_text("CSV units:"))
        self.cmb_units = ComboBox()
        self.cmb_units.Width = 80
        self.cmb_units.Items.Add("Feet")
        self.cmb_units.Items.Add("Inches")
        self.cmb_units.SelectedIndex = 0
        row_units.Children.Add(self.cmb_units)
        row_units.Children.Add(make_text("   Match tolerance (in):"))
        self.txt_tol = TextBox()
        self.txt_tol.Text = "0.125"
        self.txt_tol.Width = 60
        row_units.Children.Add(self.txt_tol)
        sp_csv.Children.Add(row_units)
        root.Children.Add(g_csv)

        # Action
        g_act, sp_act = make_group("Action")
        row_mode = make_row()
        self.rb_mark = RadioButton()
        self.rb_mark.Content = "Mark exported"
        self.rb_mark.GroupName = "mode"
        self.rb_mark.IsChecked = True
        self.rb_mark.Margin = Thickness(0, 0, 20, 0)
        self.rb_reset = RadioButton()
        self.rb_reset.Content = "Reset (clear checked items below)"
        self.rb_reset.GroupName = "mode"
        row_mode.Children.Add(self.rb_mark)
        row_mode.Children.Add(self.rb_reset)
        sp_act.Children.Add(row_mode)
        self.cb_override = CheckBox()
        self.cb_override.Content = "Override graphics in active view"
        self.cb_override.IsChecked = True
        self.cb_override.Margin = Thickness(0, 3, 0, 3)
        sp_act.Children.Add(self.cb_override)
        self.cb_comment = CheckBox()
        self.cb_comment.Content = "Write tag to Comments parameter"
        self.cb_comment.IsChecked = True
        self.cb_comment.Margin = Thickness(0, 3, 0, 3)
        sp_act.Children.Add(self.cb_comment)
        self.cb_partial = CheckBox()
        self.cb_partial.Content = "Also include partial matches (trapeze or sleeve with only some points found)"
        self.cb_partial.IsChecked = False
        self.cb_partial.Margin = Thickness(0, 3, 0, 3)
        sp_act.Children.Add(self.cb_partial)
        root.Children.Add(g_act)

        # Color
        self.grp_color, sp_col = make_group("Override color")
        row_col = make_row()
        self.swatch = Border()
        self.swatch.Width = 44
        self.swatch.Height = 24
        self.swatch.BorderBrush = Brushes.Black
        self.swatch.BorderThickness = Thickness(1)
        self.swatch.Margin = Thickness(0, 0, 8, 0)
        self.swatch.Background = brush_of(self.color)
        row_col.Children.Add(self.swatch)
        btn_pick = Button()
        btn_pick.Content = "Pick color..."
        btn_pick.Width = 90
        btn_pick.Height = 24
        btn_pick.Margin = Thickness(0, 0, 12, 0)
        btn_pick.Click += self.on_pick_color
        row_col.Children.Add(btn_pick)
        for name, rgb in PRESETS:
            b = Button()
            b.Width = 24
            b.Height = 24
            b.Margin = Thickness(0, 0, 4, 0)
            b.ToolTip = name
            b.Background = brush_of(rgb)
            b.Click += (lambda s, a, rgb=rgb: self.set_color(rgb))
            row_col.Children.Add(b)
        sp_col.Children.Add(row_col)
        row_opts = make_row()
        self.cb_solid = CheckBox()
        self.cb_solid.Content = "Solid surface fill"
        self.cb_solid.IsChecked = True
        self.cb_solid.Margin = Thickness(0, 0, 20, 0)
        self.cb_halftone = CheckBox()
        self.cb_halftone.Content = "Halftone"
        self.cb_halftone.IsChecked = False
        row_opts.Children.Add(self.cb_solid)
        row_opts.Children.Add(self.cb_halftone)
        sp_col.Children.Add(row_opts)
        root.Children.Add(self.grp_color)

        # Comment text
        g_com, sp_com = make_group("Comments")
        row_tag = make_row()
        row_tag.Children.Add(make_text("Tag text:"))
        self.txt_tag = TextBox()
        self.txt_tag.Text = "EXPORTED"
        self.txt_tag.Width = 200
        row_tag.Children.Add(self.txt_tag)
        sp_com.Children.Add(row_tag)
        self.cb_pointnums = CheckBox()
        self.cb_pointnums.Content = "Append CSV point numbers (e.g. EXPORTED: RCK01A,RCK01B)"
        self.cb_pointnums.IsChecked = False
        self.cb_pointnums.Margin = Thickness(0, 3, 0, 3)
        sp_com.Children.Add(self.cb_pointnums)
        root.Children.Add(g_com)

        # Status and buttons
        self.lbl_status = TextBlock()
        self.lbl_status.Foreground = Brushes.Firebrick
        self.lbl_status.TextWrapping = TextWrapping.Wrap
        self.lbl_status.Margin = Thickness(0, 0, 0, 6)
        root.Children.Add(self.lbl_status)

        btn_panel = StackPanel()
        btn_panel.Orientation = Orientation.Horizontal
        btn_panel.HorizontalAlignment = HorizontalAlignment.Right
        self.btn_run = Button()
        self.btn_run.Content = "Mark"
        self.btn_run.Width = 90
        self.btn_run.Height = 28
        self.btn_run.Margin = Thickness(0, 0, 8, 0)
        self.btn_run.Click += self.on_run
        btn_cancel = Button()
        btn_cancel.Content = "Cancel"
        btn_cancel.Width = 90
        btn_cancel.Height = 28
        btn_cancel.Click += self.on_cancel
        btn_panel.Children.Add(self.btn_run)
        btn_panel.Children.Add(btn_cancel)
        root.Children.Add(btn_panel)

        self.rb_mark.Checked += self.on_mode
        self.rb_reset.Checked += self.on_mode

    def set_color(self, rgb):
        self.color = rgb
        self.swatch.Background = brush_of(rgb)

    def on_mode(self, sender, args):
        reset = bool(self.rb_reset.IsChecked)
        self.grp_color.IsEnabled = not reset
        self.btn_run.Content = "Reset" if reset else "Mark"

    def on_browse(self, sender, args):
        dlg = OpenFileDialog()
        dlg.Title = "Select points CSV"
        dlg.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
        if dlg.ShowDialog(self):
            self.txt_csv.Text = dlg.FileName

    def on_pick_color(self, sender, args):
        dlg = ColorDialog()
        dlg.FullOpen = True
        dlg.AnyColor = True
        dlg.Color = DrawColor.FromArgb(self.color[0], self.color[1], self.color[2])
        if dlg.ShowDialog() == WinDialogResult.OK:
            c = dlg.Color
            self.set_color((int(c.R), int(c.G), int(c.B)))

    def on_cancel(self, sender, args):
        self.settings = None
        self.Close()

    def on_run(self, sender, args):
        path = (self.txt_csv.Text or "").strip()
        if not path or not os.path.isfile(path):
            self.lbl_status.Text = "Select a valid CSV file."
            return
        override = bool(self.cb_override.IsChecked)
        comment = bool(self.cb_comment.IsChecked)
        if not override and not comment:
            self.lbl_status.Text = "Check at least one of: override graphics, write Comments."
            return
        tag = (self.txt_tag.Text or "").strip()
        if comment and not tag:
            self.lbl_status.Text = "Enter the tag text for Comments."
            return
        try:
            tol_in = float(self.txt_tol.Text)
            if tol_in <= 0:
                raise ValueError()
        except:
            self.lbl_status.Text = "Tolerance must be a positive number of inches."
            return

        self.settings = {
            "csv": path,
            "to_feet": 1.0 if self.cmb_units.SelectedIndex == 0 else 1.0 / 12.0,
            "tol_ft": tol_in / 12.0,
            "reset": bool(self.rb_reset.IsChecked),
            "override": override,
            "comment": comment,
            "partial": bool(self.cb_partial.IsChecked),
            "color": self.color,
            "solid": bool(self.cb_solid.IsChecked),
            "halftone": bool(self.cb_halftone.IsChecked),
            "tag": tag,
            "pointnums": bool(self.cb_pointnums.IsChecked),
        }
        self.Close()


# ---- main ---------------------------------------------------------------------

win = MarkExportedWindow()
win.ShowDialog()
s = win.settings
if not s:
    sys.exit()

TOL_FT = s["tol_ft"]
CELL = TOL_FT * 2.0
do_reset = s["reset"]
tag = s["tag"]

points, info = load_csv(s["csv"], s["to_feet"])
if points is None:
    TaskDialog.Show("Mark Exported Points", info)
    sys.exit()
if not points:
    TaskDialog.Show("Mark Exported Points", "No valid rows found in the CSV.")
    sys.exit()

index = build_index(points)
elements = collect_elements()
if not elements:
    TaskDialog.Show("Mark Exported Points",
                    "No fabrication hangers or GTP elements found in the active view.")
    sys.exit()

# Group every point by its top-level owner (sleeve family, hanger, or standalone GTP)
owners = {}
for el in elements:
    pts = element_points(el)
    if not pts:
        continue
    owner = get_owner(el)
    oid = get_id_value(owner.Id)
    rec = owners.setdefault(oid, {"owner": owner, "total": 0, "hits": []})
    rec["total"] += len(pts)
    for (x, y, z) in pts:
        m = find_match(points, index, x, y, z)
        if m is not None:
            rec["hits"].append(m)

full, partial = [], []          # (owner element, [csv indices])
matched_csv = set()
for rec in owners.values():
    if not rec["hits"]:
        continue
    matched_csv.update(rec["hits"])
    if len(rec["hits"]) == rec["total"]:
        full.append((rec["owner"], rec["hits"]))
    else:
        partial.append((rec["owner"], rec["hits"]))

targets = list(full)
if s["partial"] or do_reset:
    targets += partial

mark_ogs = None
if s["override"] and not do_reset:
    mark_ogs = make_ogs(s["color"], s["halftone"], s["solid"])
blank_ogs = OverrideGraphicSettings()
comment_fail = 0
override_fail = 0

t = Transaction(doc, "Mark exported points")
t.Start()
try:
    for el, hits in targets:
        if s["override"]:
            try:
                curview.SetElementOverrides(el.Id, blank_ogs if do_reset else mark_ogs)
            except:
                override_fail += 1
        if s["comment"]:
            p = get_comments_param(el)
            if p is None or p.IsReadOnly:
                comment_fail += 1
                continue
            current = p.AsString() or ""
            try:
                if do_reset:
                    if tag.upper() in current.upper():
                        p.Set(strip_tag(current, tag))
                else:
                    tag_text = tag
                    if s["pointnums"]:
                        tag_text += ": " + ",".join(points[i]["num"] for i in hits)
                    p.Set(add_tag(current, tag, tag_text))
            except:
                comment_fail += 1
    t.Commit()
except Exception as ex:
    t.RollBack()
    TaskDialog.Show("Mark Exported Points", "Failed, nothing changed.\n{}".format(ex))
    sys.exit()

# ---- report -------------------------------------------------------------------

unmatched = [p["num"] for i, p in enumerate(points) if i not in matched_csv]

summary = [
    "Action: {}".format("Reset" if do_reset else "Mark"),
    "CSV points: {}".format(len(points)),
    "Elements scanned in view (incl. nested GTPs): {}".format(len(elements)),
    "Parent elements fully matched: {}".format(len(full)),
    "Parent elements partially matched (some points only): {}".format(len(partial)),
    "CSV points with no element found: {}".format(len(unmatched)),
]
if s["override"] and override_fail:
    summary.append("Overrides not applied: {}".format(override_fail))
if s["comment"] and comment_fail:
    summary.append("Comments not written (read-only, missing, or owned by another user): {}".format(comment_fail))
if unmatched:
    summary.append("First unmatched: " + ", ".join(unmatched[:15]))
if partial and not (s["partial"] or do_reset):
    summary.append("Partial matches were left unmarked. Element IDs: " +
                   ", ".join(str(get_id_value(el.Id)) for el, _ in partial[:15]))

# TaskDialog.Show("Mark Exported Points", "\n".join(summary))