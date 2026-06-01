# coding: utf8
import clr
import math
import sys

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
curview = doc.ActiveView

def get_fab_palette_count(fab_service):
    try:
        return fab_service.PaletteCount
    except:
        return fab_service.GroupCount

def get_fab_palette_name(fab_service, index):
    try:
        return fab_service.GetPaletteName(index)
    except:
        return fab_service.GetGroupName(index)

def get_param_text(elem, bip):
    try:
        p = elem.get_Parameter(bip)
        if p:
            s = p.AsString()
            if s:
                return s.strip()
            s = p.AsValueString()
            if s:
                return s.strip()
    except:
        pass
    return None


def normalize_size_text(s):
    if not s:
        return ""
    s = s.lower().strip()
    for ch in [" ", "-", "_", '"']:
        s = s.replace(ch, "")
    return s


def get_host_size_candidates(host_part):
    vals = []

    # strongest match: host's actual current product entry name
    try:
        if host_part.ProductListEntry >= 0:
            vals.append(host_part.GetProductListEntryName(host_part.ProductListEntry))
    except:
        pass

    # useful fabrication parameters
    for bip in [
        BuiltInParameter.FABRICATION_PART_DIAMETER_IN_OPTION,
        BuiltInParameter.FABRICATION_PART_DIAMETER_OUT_OPTION,
        BuiltInParameter.FABRICATION_PRODUCT_ENTRY,
    ]:
        v = get_param_text(host_part, bip)
        if v:
            vals.append(v)

    # optional
    try:
        if host_part.Size:
            vals.append(host_part.Size)
    except:
        pass

    # de-dupe while preserving order
    cleaned = []
    seen = set()
    for v in vals:
        key = normalize_size_text(v)
        if key and key not in seen:
            seen.add(key)
            cleaned.append(v)
    return cleaned


def try_match_product_entry_by_size(new_part, host_part):
    try:
        if not new_part.IsProductList():
            return False, "new part is not a product list"
    except:
        return False, "could not evaluate product list"

    desired_values = get_host_size_candidates(host_part)
    if not desired_values:
        return False, "no host size/product-entry text found"

    entry_count = new_part.GetProductListEntryCount()
    if entry_count < 1:
        return False, "no product entries on new part"

    # pass 1: exact normalized match
    for desired in desired_values:
        desired_norm = normalize_size_text(desired)

        for i in range(entry_count):
            try:
                entry_name = new_part.GetProductListEntryName(i)
                if normalize_size_text(entry_name) != desired_norm:
                    continue

                if new_part.IsProductListEntryCompatibleSize(i):
                    new_part.ProductListEntry = i
                    return True, entry_name
            except:
                pass

    # pass 2: contains match
    for desired in desired_values:
        desired_norm = normalize_size_text(desired)

        for i in range(entry_count):
            try:
                entry_name = new_part.GetProductListEntryName(i)
                entry_norm = normalize_size_text(entry_name)

                if desired_norm in entry_norm or entry_norm in desired_norm:
                    if new_part.IsProductListEntryCompatibleSize(i):
                        new_part.ProductListEntry = i
                        return True, entry_name
            except:
                pass

    return False, "no compatible matching product entry found"

# -----------------------------
# SELECTION FILTER
# -----------------------------
class FabricationStraightSelectionFilter(ISelectionFilter):
    def AllowElement(self, elem):
        try:
            return isinstance(elem, FabricationPart) and isinstance(elem.Location, LocationCurve)
        except:
            return False
    def AllowReference(self, reference, point):
        return False

# -----------------------------
# PIPE POINTS & DIRECTION
# -----------------------------
def get_projected_insert_point(host_part):
    picked_point = uidoc.Selection.PickPoint("Pick a point along the pipe centerline")
    curve = host_part.Location.Curve
    result = curve.Project(picked_point)
    if not result:
        raise Exception("Could not project picked point onto pipe centerline.")
    return result.XYZPoint

def get_pipe_midpoint(host_part):
    curve = host_part.Location.Curve
    p0 = curve.GetEndPoint(0)
    p1 = curve.GetEndPoint(1)
    return (p0 + p1) * 0.5

def get_pipe_direction(host_part):
    curve = host_part.Location.Curve
    p0 = curve.GetEndPoint(0)
    p1 = curve.GetEndPoint(1)
    return (p1 - p0).Normalize()

def is_3d_view(view):
    return view.ViewType == ViewType.ThreeD

def is_section_view(view):
    return view.ViewType == ViewType.Section

def set_section_workplane(view, origin):
    x_vec = view.RightDirection.Normalize()
    y_vec = view.UpDirection.Normalize()
    plane = Plane.CreateByOriginAndBasis(origin, x_vec, y_vec)

    t = Transaction(doc, "Set Section Work Plane")
    t.Start()
    sketch_plane = SketchPlane.Create(doc, plane)
    view.SketchPlane = sketch_plane
    doc.Regenerate()
    t.Commit()

# -----------------------------
# ROTATION UTILITIES
# -----------------------------
def rotate_to_vector(doc, element, origin, from_vec, to_vec):
    from_vec = from_vec.Normalize()
    to_vec = to_vec.Normalize()
    axis = from_vec.CrossProduct(to_vec)

    if axis.GetLength() < 1e-8:
        # Parallel or anti-parallel
        dot = from_vec.DotProduct(to_vec)
        if dot < 0:
            # 180 deg flip
            axis = XYZ.BasisZ
            angle = math.pi
        else:
            return
    else:
        axis = axis.Normalize()
        angle = math.acos(max(min(from_vec.DotProduct(to_vec), 1.0), -1.0))

    rot_line = Line.CreateBound(origin, origin + axis)
    ElementTransformUtils.RotateElement(doc, element.Id, rot_line, angle)

def ensure_facing_user(doc, element, origin):
    view = doc.ActiveView
    # Make part "face the user" in section/plan views
    view_dir = view.ViewDirection.Normalize()
    # Part's forward axis (assume after previous alignment it's along X)
    part_dir = XYZ.BasisX
    # Rotate about view_dir to align part's projected X to screen right
    # Project part_dir onto view plane
    part_proj = part_dir - view_dir.Multiply(part_dir.DotProduct(view_dir))
    part_proj_length = part_proj.GetLength()
    if part_proj_length < 1e-6:
        return
    part_proj = part_proj.Normalize()
    # Screen right in view
    screen_right = view.RightDirection.Normalize()
    axis = view_dir
    angle = math.acos(max(min(part_proj.DotProduct(screen_right), 1.0), -1.0))
    # Determine correct rotation direction
    if part_proj.CrossProduct(screen_right).DotProduct(axis) < 0:
        angle = -angle
    rot_line = Line.CreateBound(origin, origin + axis)
    ElementTransformUtils.RotateElement(doc, element.Id, rot_line, angle)

# -----------------------------
# SELECTION
# -----------------------------
straight_filter = FabricationStraightSelectionFilter()
try:
    ref = uidoc.Selection.PickObject(
        ObjectType.Element,
        straight_filter,
        "Select fabrication pipe/straight"
    )
except OperationCanceledException:
    sys.exit()

host_part = doc.GetElement(ref.ElementId)

try:
    if is_3d_view(doc.ActiveView):
        insert_point = get_pipe_midpoint(host_part)
    else:
        # Section view needs a work plane BEFORE PickPoint
        if is_section_view(doc.ActiveView):
            midpoint = get_pipe_midpoint(host_part)
            set_section_workplane(doc.ActiveView, midpoint)

        insert_point = get_projected_insert_point(host_part)

except OperationCanceledException:
    sys.exit()
except Exception as ex:
    TaskDialog.Show("Error", str(ex))
    sys.exit()

# -----------------------------
# GET SERVICE & BUTTONS
# -----------------------------
config = FabricationConfiguration.GetFabricationConfiguration(doc)
services = config.GetAllLoadedServices()
target_service = None
host_service_id = host_part.ServiceId

for s in services:
    if s.ServiceId == host_service_id:
        target_service = s
        break

if not target_service:
    TaskDialog.Show("Error", "Could not find loaded fabrication service for selected host part.")
    sys.exit()

palette_names = []
button_records = []

grp_count = get_fab_palette_count(target_service)

for p in range(grp_count):
    palette_name = get_fab_palette_name(target_service, p)
    palette_names.append(palette_name)

    for i in range(target_service.GetButtonCount(p)):
        btn = target_service.GetButton(p, i)

        if btn.ConditionCount > 1:
            for c in range(btn.ConditionCount):
                display = u"{0} - {1}".format(btn.Name, btn.GetConditionName(c))
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

if not button_records:
    TaskDialog.Show("Error", "No fabrication buttons found for the selected service.")
    sys.exit()

# -----------------------------
# WPF PART PICKER
# -----------------------------
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

        lbl_palette = Label(); lbl_palette.Content = "Palette:"; stack.Children.Add(lbl_palette)
        self.palette_combo = ComboBox(); self.palette_combo.Margin = Thickness(0,0,0,10)
        self.palette_combo.Items.Add("All Palettes")
        for p in palettes: self.palette_combo.Items.Add(p)
        self.palette_combo.SelectedIndex = 0
        self.palette_combo.SelectionChanged += self.apply_filters
        stack.Children.Add(self.palette_combo)

        lbl_search = Label(); lbl_search.Content = "Search Part:"; stack.Children.Add(lbl_search)
        self.search_box = TextBox(); self.search_box.Margin = Thickness(0,0,0,10)
        self.search_box.TextChanged += self.apply_filters; stack.Children.Add(self.search_box)

        lbl_instr = Label(); lbl_instr.Content = "Double Click Item to Insert"; lbl_instr.Margin = Thickness(0,0,0,5); stack.Children.Add(lbl_instr)
        self.list_box = ListBox(); self.list_box.Height = 430; self.list_box.Margin = Thickness(0,0,0,10)
        self.list_box.MouseDoubleClick += self.on_double_click
        stack.Children.Add(self.list_box)

        self.Content = stack
        self.refresh_list()
        self.search_box.Focus(); Keyboard.Focus(self.search_box)

    def refresh_list(self):
        self.list_box.ItemsSource = [r["display"] for r in self.filtered_records]
    def apply_filters(self, sender, args):
        sel_palette = self.palette_combo.SelectedItem
        search_text = self.search_box.Text.lower().strip()
        records = self.all_records
        if sel_palette and sel_palette != "All Palettes": records = [r for r in records if r["palette_name"] == sel_palette]
        if search_text: records = [r for r in records if search_text in r["display"].lower()]
        self.filtered_records = records
        self.refresh_list()
    def on_double_click(self, sender, args):
        idx = self.list_box.SelectedIndex
        if idx < 0 or idx >= len(self.filtered_records):
            TaskDialog.Show("Error","Please select a part."); return
        self.selected_record = self.filtered_records[idx]; self.DialogResult = True; self.Close()

# -----------------------------
# SHOW DIALOG
# -----------------------------
dlg = PartPicker(button_records, palette_names)
if not dlg.ShowDialog(): sys.exit()
selected_record = dlg.selected_record
fab_btn = selected_record["button"]
condition_index = selected_record["condition_index"]

# -----------------------------
# CREATE PART + MOVE + ROTATE
# -----------------------------
t = None
try:
    t = Transaction(doc, "Place Fabrication Part on Pipe Centerline")
    t.Start()

    new_part = FabricationPart.Create(doc, fab_btn, condition_index, host_part.LevelId)
    doc.Regenerate()

    matched, match_info = try_match_product_entry_by_size(new_part, host_part)
    doc.Regenerate()

    translation = insert_point - new_part.Origin
    ElementTransformUtils.MoveElement(doc, new_part.Id, translation)
    doc.Regenerate()

    pipe_dir = get_pipe_direction(host_part)
    rotate_to_vector(doc, new_part, insert_point, XYZ.BasisX, pipe_dir)

    if not is_3d_view(doc.ActiveView):
        ensure_facing_user(doc, new_part, insert_point)

    t.Commit()

except Exception as ex:
    if t and t.HasStarted() and not t.HasEnded():
        t.RollBack()
    TaskDialog.Show("Error", str(ex))