# -*- coding: utf-8 -*-
"""Create Floor Plan, Elevation, and 3D Isometric views around the current selection,
name them from a prompt, and optionally place them on a sheet.

Select fabrication parts and/or family instances in the active plan view, then run.

Requires pyRevit (uses pyrevit.forms). The window is modeless: Revit stays usable
(Properties, selection, view navigation) while it is open.
DRAFT: untested. Verify in a test project before use on a live model.
"""
__persistentengine__ = True

# ---------------------------------------------------------------------------
# Self-relaunch. pyRevit clears a script's own namespace when the script
# finishes, which breaks a modeless window (its buttons run later and would find
# the functions gone). So the script runs itself once inside a private namespace
# that pyRevit does not know about and cannot clear. Single file, no helper module.
# ---------------------------------------------------------------------------
if "__own_ns__" not in globals():
    import sys
    _ns = {
        "__own_ns__": True,
        "__name__": "views_from_selection",
        "__file__": __file__,
        "__revit__": __revit__,
    }
    exec(compile(open(__file__, "rb").read(), __file__, "exec"), _ns)
    sys.exit()

import os
import sys
import time
import shutil
import tempfile
import clr
from System.Collections.Generic import List
from Autodesk.Revit import DB
from Autodesk.Revit.DB import FabricationPart, FamilyInstance
from Autodesk.Revit.UI import TaskDialog, IExternalEventHandler, ExternalEvent
from Autodesk.Revit.Exceptions import OperationCanceledException
import System
from pyrevit import forms

clr.AddReference("PresentationCore")
clr.AddReference("PresentationFramework")
clr.AddReference("WindowsBase")
from System import Uri
from System.Windows import Visibility
from System.Windows.Media.Imaging import BitmapImage, BitmapCacheOption

_uiapp = None
doc = None
uidoc = None
active_view = None


def refresh_context():
    """Re-read the active document and view. The window stays open while the user
    works in Revit, so these can change between runs."""
    global doc, uidoc, active_view
    uidoc = _uiapp.ActiveUIDocument
    doc = uidoc.Document
    active_view = doc.ActiveView

SCRIPT_VERSION = "v4.2 (2026-10-20)"

# ---------------- Settings (feet unless noted) ----------------
PAD_X = 1.0
PAD_Y = 1.0
PAD_Z = 0.5
EYE_DISTANCE = 50.0

# True  = temporarily isolate the selected elements in each new view
# False = show selection in context, cropped around it
ISOLATE_SELECTION = False

# Optional view template names (exact match). None = no template.
VIEW_TEMPLATES = {"top": None, "iso": None, "front": None}

# Names shown for each view type. The suffix is added to the typed base name.
VIEW_LABELS = {"top": "Floor Plan", "front": "Elevation View", "iso": "3D Isometric View"}
VIEW_SUFFIX = {"top": " - Floor Plan", "front": " - Elevation View", "iso": " - 3D Isometric View"}

# Text written to each view's "Title on Sheet" (what shows under the viewport).
# {name} = typed name, {label} = Floor Plan / Elevation View / 3D Isometric View.
SHEET_TITLE_FORMAT = "{name} - {label}"

# Other view parameters that get the same text (skipped if the view does not have them).
EXTRA_TITLE_PARAMS = ["View Description"]

# False = the window shows one short result line. True = also show the full
# details box (what was kept/hidden, direction checks, version) for troubleshooting.
SHOW_DETAILS = False

# View scale denominators. 24 = 1/2" = 1'-0", 8 = 1-1/2" = 1'-0". None = Revit default.
VIEW_SCALES = {"top": 24, "front": 24, "iso": 8}

# True = after Create, show a preview of each side of the selection and let the user
# choose the face for the Elevation View. False = use the default (front) side.
CHOOSE_FACE_PREVIEW = True
PREVIEW_PIXELS = 600

# False = views are placed and fitted on the sheet automatically (Floor Plan over
# Elevation on the left, 3D Isometric on the right).
# True  = after Create, go to the sheet and let the user click where each view goes
# (Revit's own placement with a live preview).
PLACE_INTERACTIVE = False

# Hide fabrication parts whose assembly does not match the selected fabrication parts.
# Selected elements are always kept.
FILTER_BY_ASSEMBLY = True
# Also hide model family instances that are not selected and not in a selected assembly
FILTER_FAMILIES = True
STRATUS_PARAM = "STRATUS Assembly"

# Categories hidden in every new view (scope boxes, section lines, section heads,
# Revit links). Add more BuiltInCategory names if needed, for example "OST_Grids"
# or "OST_Levels".
HIDE_CATEGORIES = ["OST_VolumeOfInterest", "OST_Sections", "OST_Viewers", "OST_RvtLinks"]

# Sheet layout
SHEET_MARGIN = 0.3        # inset from title block extents
VIEWPORT_GAP = 0.15       # gap between viewports
FALLBACK_SHEET_W = 3.0    # used only if no title block is found (36 in)
FALLBACK_SHEET_H = 2.0    # (24 in)
OPEN_SHEET_AFTER = True    # True = open the sheet when done. False = stay in the view you started from

# Straight fabrication CIDs used as the alignment reference (same list as the original section tool)
ALLOWED_CIDS = set([2041, 866, 40])

ILLEGAL_CHARS = set('\\:{}[]|;<>?`~')


class UserCancel(Exception):
    pass


# ---------------- General helpers ----------------
def elem_name(e):
    try:
        return DB.Element.Name.GetValue(e)
    except:
        return e.Name


def clean_text(value, label):
    value = (value or "").strip()
    if not value:
        raise Exception("{0} cannot be blank.".format(label))
    bad = sorted(set(c for c in value if c in ILLEGAL_CHARS))
    if bad:
        raise Exception("{0} contains characters Revit does not allow: {1}".format(label, " ".join(bad)))
    return value


def get_selected_elements():
    elems = []
    for eid in uidoc.Selection.GetElementIds():
        e = doc.GetElement(eid)
        if isinstance(e, (FabricationPart, FamilyInstance)):
            elems.append(e)
    return elems


def get_view_vectors(view):
    """Right and up directions used to define 'front' (bottom of plan)."""
    if isinstance(view, DB.ViewPlan):
        right = view.RightDirection.Normalize()
        up = view.UpDirection.Normalize()
    else:
        right = DB.XYZ.BasisX
        up = DB.XYZ.BasisY
    return right, up


def longest_straight_pipe(elems):
    """Longest straight fabrication pipe (by horizontal length) in elems.
    Returns (pipe, horizontal_direction) or (None, None)."""
    best = None
    best_dir = None
    best_len = 1e-6
    for e in elems:
        try:
            if not isinstance(e, FabricationPart):
                continue
            if e.ItemCustomId not in ALLOWED_CIDS:
                continue
            loc = e.Location
            if not isinstance(loc, DB.LocationCurve):
                continue
            crv = loc.Curve
            if not isinstance(crv, DB.Line):
                continue
            d = crv.GetEndPoint(1).Subtract(crv.GetEndPoint(0))
            d = DB.XYZ(d.X, d.Y, 0)
            length = d.GetLength()
            if length > best_len:
                best_len = length
                best_dir = d
                best = e
        except:
            continue
    return best, best_dir


def get_pipe_frame(elems, view, assemblies=None):
    """Return (right, back, pipe). Frame is aligned to the longest straight
    fabrication pipe in the selection. If the selection has none (for example only
    the carrier is selected), the longest straight pipe in the same STRATUS assembly
    is used. 'Front' is the side of the pipe facing the bottom of the active plan.
    Returns the view orientation and pipe None if no straight pipe is found."""
    view_right, view_up = get_view_vectors(view)
    world_z = DB.XYZ.BasisZ

    best, best_dir = longest_straight_pipe(elems)

    if best is None and assemblies:
        cands = []
        col = DB.FilteredElementCollector(doc).OfClass(FabricationPart).WhereElementIsNotElementType()
        for e in col:
            try:
                if param_text(e, STRATUS_PARAM) in assemblies:
                    cands.append(e)
            except:
                continue
        best, best_dir = longest_straight_pipe(cands)

    if best is None:
        return view_right, view_up, None

    d = best_dir.Normalize()
    n = world_z.CrossProduct(d).Normalize()   # horizontal, perpendicular to pipe

    # Front = side facing the bottom of the plan (front . view_up < 0)
    if n.DotProduct(view_up) > 0:
        front = n.Negate()
    else:
        front = n
    # Pipe parallel to plan up: no bottom side exists, so take the right side
    if abs(front.DotProduct(view_up)) < 1e-3:
        front = n if n.DotProduct(view_right) > 0 else n.Negate()

    back = front.Negate()
    right = world_z.CrossProduct(front).Normalize()
    return right, back, best


def bbox_corners(bbox):
    mn = bbox.Min
    mx = bbox.Max
    pts = []
    for x in [mn.X, mx.X]:
        for y in [mn.Y, mx.Y]:
            for z in [mn.Z, mx.Z]:
                pts.append(DB.XYZ(x, y, z))
    return pts


def get_selection_points(elems):
    pts = []
    skipped = 0
    for e in elems:
        bb = e.get_BoundingBox(None)
        if bb is None:
            skipped += 1
            continue
        pts.extend(bbox_corners(bb))
    return pts, skipped


def get_local_bounds(points, transform):
    inv = transform.Inverse
    xs, ys, zs = [], [], []
    for p in points:
        lp = inv.OfPoint(p)
        xs.append(lp.X)
        ys.append(lp.Y)
        zs.append(lp.Z)
    return DB.XYZ(min(xs), min(ys), min(zs)), DB.XYZ(max(xs), max(ys), max(zs))


def get_view_family_type(family):
    for vft in DB.FilteredElementCollector(doc).OfClass(DB.ViewFamilyType):
        if vft.ViewFamily == family:
            return vft
    return None


def existing_view_names():
    return set(elem_name(v) for v in DB.FilteredElementCollector(doc).OfClass(DB.View))


def unique_name(base, taken):
    name = base
    i = 2
    while name in taken:
        name = "{0} ({1})".format(base, i)
        i += 1
    taken.add(name)
    return name


def apply_template(view, key):
    tname = VIEW_TEMPLATES.get(key)
    if not tname:
        return
    for v in DB.FilteredElementCollector(doc).OfClass(DB.View):
        if v.IsTemplate and elem_name(v) == tname:
            try:
                view.ViewTemplateId = v.Id
            except:
                pass
            return


def set_title_on_sheet(view, text):
    """Write the viewport title: the built-in Title on Sheet parameter, plus any
    extra parameters listed in EXTRA_TITLE_PARAMS (for example View Description)."""
    try:
        param = view.get_Parameter(DB.BuiltInParameter.VIEW_DESCRIPTION)
        if param and not param.IsReadOnly:
            param.Set(text)
    except:
        pass
    for name in EXTRA_TITLE_PARAMS:
        try:
            p = find_param(view, name)
            if p is not None and not p.IsReadOnly and p.StorageType == DB.StorageType.String:
                p.Set(text)
        except:
            pass


def set_fine_and_scale(view):
    try:
        view.DetailLevel = DB.ViewDetailLevel.Fine
    except:
        pass


def read_param_text(p):
    if p is None or not p.HasValue:
        return None
    if p.StorageType == DB.StorageType.String:
        v = p.AsString()
    else:
        v = p.AsValueString()
    v = (v or "").strip()
    return v or None


def find_param(elem, name):
    """Exact lookup first, then a case-insensitive name match."""
    p = elem.LookupParameter(name)
    if p is not None:
        return p
    target = name.strip().lower()
    for q in elem.Parameters:
        try:
            if q.Definition.Name.strip().lower() == target:
                return q
        except:
            continue
    return None


def param_text(elem, name):
    """Parameter value as trimmed text (instance first, then type), or None."""
    try:
        v = read_param_text(find_param(elem, name))
        if v:
            return v
    except:
        pass
    try:
        t = doc.GetElement(elem.GetTypeId())
        if t is not None:
            return read_param_text(find_param(t, name))
    except:
        pass
    return None


def similar_param_names(elems, limit=3):
    """Names of parameters containing 'stratus' or 'assembly' on the first few
    selected elements (instance and type), to help spot the real name."""
    found = set()
    for e in elems[:limit]:
        holders = [e]
        try:
            t = doc.GetElement(e.GetTypeId())
            if t is not None:
                holders.append(t)
        except:
            pass
        for h in holders:
            try:
                for q in h.Parameters:
                    n = q.Definition.Name
                    low = n.lower()
                    if "stratus" in low or "assembly" in low:
                        found.add(n)
            except:
                continue
    return sorted(found)


def get_selected_assemblies(elems):
    """Stratus Assembly values from ALL selected elements (parts and families).
    Returns (set of values, number of selected elements that had a value)."""
    vals = set()
    with_value = 0
    for e in elems:
        v = param_text(e, STRATUS_PARAM)
        if v:
            vals.add(v)
            with_value += 1
    return vals, with_value


def collect_ids_to_hide(view, cls, keep_ids, assemblies, model_only=False):
    ids = []
    col = DB.FilteredElementCollector(doc, view.Id).OfClass(cls).WhereElementIsNotElementType()
    for e in col:
        try:
            if str(e.Id) in keep_ids:
                continue
            if model_only:
                cat = e.Category
                if cat is None or cat.CategoryType != DB.CategoryType.Model:
                    continue
                if e.SuperComponent is not None:
                    continue
            if param_text(e, STRATUS_PARAM) in assemblies:
                continue
            if e.CanBeHidden(view):
                ids.append(e.Id)
        except:
            continue
    return ids


def hide_ids(view, ids):
    if not ids:
        return 0
    try:
        view.HideElements(List[DB.ElementId](ids))
        return len(ids)
    except:
        count = 0
        for eid in ids:
            try:
                view.HideElements(List[DB.ElementId]([eid]))
                count += 1
            except:
                pass
        return count


def hide_other_assemblies(view, keep_ids, assemblies):
    """Hide fabrication parts (and optionally model families) that are not
    selected and not in a selected assembly. Returns (parts_hidden, families_hidden)."""
    parts = hide_ids(view, collect_ids_to_hide(view, FabricationPart, keep_ids, assemblies))
    fams = 0
    if FILTER_FAMILIES:
        fams = hide_ids(view, collect_ids_to_hide(view, FamilyInstance, keep_ids, assemblies, True))
    return parts, fams


def hide_unwanted_categories(view):
    """Hide scope boxes, section marks, and Revit links so only the selected
    fabrication parts and families (plus host model context) show."""
    for name in HIDE_CATEGORIES:
        try:
            bic = getattr(DB.BuiltInCategory, name)
            cat_id = DB.ElementId(bic)
            if view.CanCategoryBeHidden(cat_id):
                view.SetCategoryHidden(cat_id, True)
        except:
            pass


def set_scale(view, key):
    denom = VIEW_SCALES.get(key)
    if denom:
        try:
            view.Scale = int(denom)
        except:
            pass


# ---------------- Sheet / title block data ----------------
NO_TITLE_BLOCK = "<No title block>"


def get_title_block_types():
    d = {}
    col = DB.FilteredElementCollector(doc).OfCategory(
        DB.BuiltInCategory.OST_TitleBlocks).WhereElementIsElementType()
    for s in col:
        d["{0} : {1}".format(s.FamilyName, elem_name(s))] = s.Id
    return d


def get_sheets():
    d = {}
    for s in DB.FilteredElementCollector(doc).OfClass(DB.ViewSheet):
        try:
            if s.IsPlaceholder:
                continue
        except:
            pass
        d["{0} - {1}".format(s.SheetNumber, elem_name(s))] = s
    return d


# ---------------- WPF form (modeless) ----------------
# Black on white, Murray Yellow #FFE716 as accent/fill only, Arial.
FORM_XAML = """
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Views From Selection" Width="480" SizeToContent="Height"
        WindowStartupLocation="Manual" Left="40" Top="80" ResizeMode="NoResize"
        FontFamily="Arial" FontSize="13" Background="White" Foreground="Black">
  <StackPanel Margin="18">
    <TextBlock x:Name="info_text" FontSize="11" TextWrapping="Wrap"/>
    <TextBlock FontSize="11" TextWrapping="Wrap" Margin="0,4,0,0"
               Text="You can select elements, use Properties, and change views in Revit while this window is open."/>

    <Border Height="3" Background="#FFE716" Margin="0,10,0,10"/>

    <TextBlock Text="View name" FontWeight="Bold"/>
    <TextBox x:Name="name_box" Margin="0,4,0,4" Padding="4" BorderBrush="Black"/>
    <TextBlock TextWrapping="Wrap" FontSize="11"
               Text="Example: L1-A1-WV_WTR-01. Project Browser names get - Floor Plan, - Elevation View, or - 3D Isometric View added. Sheet titles show the name plus the view type."/>

    <TextBlock x:Name="front_note" TextWrapping="Wrap" FontSize="11" Margin="0,10,0,0"/>

    <Border Height="3" Background="#FFE716" Margin="0,14,0,10"/>

    <TextBlock Text="Sheet" FontWeight="Bold"/>
    <RadioButton x:Name="rb_none" GroupName="mode" Margin="0,6,0,0"
                 Content="Do not place on a sheet" Checked="on_mode_changed"/>
    <RadioButton x:Name="rb_existing" GroupName="mode" Margin="0,6,0,0"
                 Content="Place on an existing sheet" Checked="on_mode_changed"/>
    <StackPanel x:Name="existing_panel" Margin="22,4,0,0">
      <ComboBox x:Name="sheet_combo"/>
    </StackPanel>
    <RadioButton x:Name="rb_new" GroupName="mode" Margin="0,8,0,0"
                 Content="Create a new sheet" Checked="on_mode_changed"/>
    <StackPanel x:Name="new_panel" Margin="22,4,0,0">
      <TextBlock Text="Title block"/>
      <ComboBox x:Name="tb_combo" Margin="0,2,0,6"/>
      <TextBlock Text="Sheet number"/>
      <TextBox x:Name="sheet_number_box" Margin="0,2,0,6" Padding="4" BorderBrush="Black"/>
      <TextBlock Text="Sheet name"/>
      <TextBox x:Name="sheet_name_box" Margin="0,2,0,0" Padding="4" BorderBrush="Black"/>
    </StackPanel>

    <TextBlock x:Name="error_text" FontWeight="Bold" TextWrapping="Wrap" Margin="0,12,0,0"/>

    <TextBlock x:Name="result_text" TextWrapping="Wrap" Margin="0,12,0,0"/>

    <TextBox x:Name="status_box" Visibility="Collapsed" Height="150" IsReadOnly="True" TextWrapping="Wrap"
             VerticalScrollBarVisibility="Auto" BorderBrush="Black" Padding="4"
             FontSize="11" Margin="0,10,0,0"/>

    <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,14,0,0">
      <Button x:Name="info_button" Content="Update selection info" Width="150" Height="28" Margin="0,0,8,0"
              Background="White" BorderBrush="Black" Click="on_info"/>
      <Button Content="Close" Width="80" Height="28" Margin="0,0,8,0"
              Background="White" BorderBrush="Black" IsCancel="True" Click="on_close"/>
      <Button x:Name="create_button" Content="Create" Width="90" Height="28" FontWeight="Bold"
              Background="#FFE716" BorderBrush="Black" IsDefault="True" Click="on_create"/>
    </StackPanel>
  </StackPanel>
</Window>
"""


class ViewsForm(forms.WPFWindow):
    def __init__(self, handler, ext_event):
        forms.WPFWindow.__init__(self, FORM_XAML, literal_string=True)
        self._handler = handler
        self._event = ext_event
        self._sheets = {}
        self._title_blocks = {}
        self._taken_numbers = set()

        self.refresh_lists()
        self.rb_none.IsChecked = True
        self._update_panels()
        self.update_info()
        self.show_result("Select elements in Revit, enter a name, then click Create.")
        self.name_box.Focus()

    # ----- data refresh (call only from a Revit API context) -----
    def refresh_lists(self):
        keep_sheet = self.sheet_combo.SelectedItem
        keep_tb = self.tb_combo.SelectedItem

        self._sheets = get_sheets()
        self._title_blocks = get_title_block_types()
        self._taken_numbers = set(
            sh.SheetNumber for sh in DB.FilteredElementCollector(doc).OfClass(DB.ViewSheet))

        self.sheet_combo.Items.Clear()
        for label in sorted(self._sheets.keys()):
            self.sheet_combo.Items.Add(label)
        self.sheet_combo.SelectedIndex = 0 if self._sheets else -1
        if keep_sheet is not None and str(keep_sheet) in self._sheets:
            self.sheet_combo.SelectedItem = keep_sheet
        self.rb_existing.IsEnabled = bool(self._sheets)

        self.tb_combo.Items.Clear()
        for label in sorted(self._title_blocks.keys()):
            self.tb_combo.Items.Add(label)
        self.tb_combo.Items.Add(NO_TITLE_BLOCK)
        self.tb_combo.SelectedIndex = 0
        if keep_tb is not None and (str(keep_tb) in self._title_blocks or str(keep_tb) == NO_TITLE_BLOCK):
            self.tb_combo.SelectedItem = keep_tb

    def update_info(self):
        try:
            count = len(get_selected_elements())
            self.info_text.Text = "Selected: {0} element(s)   |   Active view: {1}".format(
                count, elem_name(active_view))
        except:
            self.info_text.Text = "Selection info unavailable."

        if PLACE_INTERACTIVE:
            self.front_note.Text = ("After you click Create, you will see a preview of each side of the "
                                    "selection. Pick the face of the carrier for the Elevation View, then "
                                    "click where each view goes on the sheet.")
        else:
            self.front_note.Text = ("After you click Create, you will see a preview of each side of the "
                                    "selection. Pick the face of the carrier for the Elevation View. "
                                    "The views are then placed and fitted on the sheet automatically.")

    # ----- UI helpers -----
    def show_result(self, summary, details=None):
        """One short line for the user. The full details box only shows when
        SHOW_DETAILS is True."""
        self.result_text.Text = summary
        if SHOW_DETAILS and details:
            self.status_box.Visibility = Visibility.Visible
            self.status_box.Text = details
            self.status_box.ScrollToHome()

    def set_busy(self, busy):
        self.create_button.IsEnabled = not busy
        self.info_button.IsEnabled = not busy

    def _update_panels(self):
        self.existing_panel.IsEnabled = bool(self.rb_existing.IsChecked)
        self.new_panel.IsEnabled = bool(self.rb_new.IsChecked)

    def on_mode_changed(self, sender, args):
        try:
            self._update_panels()
        except:
            pass

    def _raise(self, request):
        self.error_text.Text = ""
        self.set_busy(True)
        self._handler.request = request
        self._event.Raise()

    # ----- button handlers (run on the UI thread; Revit work goes through the event) -----
    def on_close(self, sender, args):
        self.Close()

    def on_info(self, sender, args):
        self._raise({"action": "info"})

    def on_create(self, sender, args):
        try:
            name = clean_text(self.name_box.Text, "View name")
            plan = None

            if self.rb_existing.IsChecked:
                pick = self.sheet_combo.SelectedItem
                if pick is None:
                    raise Exception("Select a sheet.")
                plan = {"mode": "existing", "sheet": self._sheets[str(pick)]}

            elif self.rb_new.IsChecked:
                pick = self.tb_combo.SelectedItem
                if pick is None:
                    raise Exception("Select a title block.")
                pick = str(pick)
                tb_id = (DB.ElementId.InvalidElementId if pick == NO_TITLE_BLOCK
                         else self._title_blocks[pick])
                number = clean_text(self.sheet_number_box.Text, "Sheet number")
                if number in self._taken_numbers:
                    raise Exception("Sheet number {0} already exists.".format(number))
                sheet_name = clean_text(self.sheet_name_box.Text, "Sheet name")
                plan = {"mode": "new", "tb_id": tb_id,
                        "number": number, "name": sheet_name}
        except Exception as ex:
            self.error_text.Text = str(ex)
            return

        self.show_result("Working...")
        self._raise({"action": "create", "name": name, "sheet_plan": plan})


class ViewsHandler(IExternalEventHandler):
    """Runs the Revit work. Revit only allows API calls and model changes from an
    external event when a modeless window is open."""

    def __init__(self):
        self.form = None
        self.request = None

    def Execute(self, uiapp):
        form = self.form
        req = self.request or {}
        try:
            refresh_context()
            if req.get("action") == "create":
                summary, details = run_create(req, form)
                form.show_result(summary, details)
        except UserCancel:
            form.show_result("Cancelled. Nothing was created.")
        except Exception as ex:
            form.show_result("Error: {0}".format(ex))
        finally:
            try:
                form.refresh_lists()
                form.update_info()
            except:
                pass
            form.set_busy(False)

    def GetName(self):
        return "Views From Selection"


# ---------------- View creation ----------------
def make_section(section_type, origin, x_axis, y_axis, z_axis, points):
    """Create a section. Callers pass z_axis as the direction TOWARD the viewer.

    Revit's section transform uses BasisZ as the LOOK direction (away from the
    viewer; depth runs along +Z from the section plane). So the transform is
    built with BasisZ = -z_axis and BasisX = Y x Z. x_axis is accepted for call
    compatibility only. The section plane is placed at the near face of the
    selection so everything sits in front of the far clip."""
    look = z_axis.Negate()
    x_dir = y_axis.CrossProduct(look).Normalize()

    tf = DB.Transform.Identity
    tf.Origin = origin
    tf.BasisX = x_dir
    tf.BasisY = y_axis
    tf.BasisZ = look

    lo, hi = get_local_bounds(points, tf)

    near = lo.Z - PAD_Z
    depth = (hi.Z - lo.Z) + 2.0 * PAD_Z
    if depth < 0.25:
        depth = 0.25
    tf.Origin = origin.Add(look.Multiply(near))

    box = DB.BoundingBoxXYZ()
    box.Transform = tf
    box.Min = DB.XYZ(lo.X - PAD_X, lo.Y - PAD_Y, 0.0)
    box.Max = DB.XYZ(hi.X + PAD_X, hi.Y + PAD_Y, depth)

    view = DB.ViewSection.CreateSection(doc, section_type.Id, box)
    set_fine_and_scale(view)
    return view


def api_view_sign(view):
    """Calibrate View.ViewDirection from the active plan view.
    Returns +1 if ViewDirection points toward the viewer, -1 if it points the
    way the viewer looks, or None when the active view is not a plan."""
    if not isinstance(view, DB.ViewPlan):
        return None
    try:
        toward_viewer_z = -1.0 if view.ViewType == DB.ViewType.CeilingPlan else 1.0
        return 1.0 if view.ViewDirection.Z * toward_viewer_z > 0 else -1.0
    except:
        return None


def section_faces(view, z_axis, sign):
    """True if the section's real view direction matches the intended z_axis
    (z_axis points toward the viewer)."""
    return view.ViewDirection.DotProduct(z_axis) * sign > 0


def make_section_checked(section_type, origin, x_axis, y_axis, z_axis, points, sign):
    """Create a section and verify Revit's actual view direction against the
    intended one. If it came out reversed, rebuild it from the opposite side.
    Returns (view, status) with status 'ok', 'corrected', 'not corrected', or 'not checked'."""
    view = make_section(section_type, origin, x_axis, y_axis, z_axis, points)
    if sign is None:
        return view, "not checked"
    try:
        if section_faces(view, z_axis, sign):
            return view, "ok"
    except:
        return view, "not checked"

    try:
        doc.Delete(view.Id)
        view = make_section(section_type, origin, x_axis.Negate(), y_axis,
                            z_axis.Negate(), points)
        if section_faces(view, z_axis, sign):
            return view, "corrected"
        return view, "not corrected"
    except:
        return view, "not checked"


def make_iso(vft_3d, center, right, up, points):
    view = DB.View3D.CreateIsometric(doc, vft_3d.Id)

    world_z = DB.XYZ.BasisZ
    front = up.Negate()

    to_eye = right.Add(front).Add(world_z).Normalize()
    forward = to_eye.Negate()
    up_vec = world_z.Subtract(forward.Multiply(world_z.DotProduct(forward))).Normalize()
    eye = center.Add(to_eye.Multiply(EYE_DISTANCE))

    view.SetOrientation(DB.ViewOrientation3D(eye, up_vec, forward))

    xs = [p.X for p in points]
    ys = [p.Y for p in points]
    zs = [p.Z for p in points]
    box = DB.BoundingBoxXYZ()
    box.Min = DB.XYZ(min(xs) - PAD_X, min(ys) - PAD_Y, min(zs) - PAD_Z)
    box.Max = DB.XYZ(max(xs) + PAD_X, max(ys) + PAD_Y, max(zs) + PAD_Z)
    view.SetSectionBox(box)
    view.IsSectionBoxActive = True

    set_fine_and_scale(view)
    return view


# ---------------- Sheet placement ----------------
def get_sheet_area(sheet):
    col = DB.FilteredElementCollector(doc, sheet.Id).OfCategory(
        DB.BuiltInCategory.OST_TitleBlocks).WhereElementIsNotElementType()
    for tb in col:
        bb = tb.get_BoundingBox(sheet)
        if bb:
            return bb.Min.X, bb.Min.Y, bb.Max.X, bb.Max.Y
    return 0.0, 0.0, FALLBACK_SHEET_W, FALLBACK_SHEET_H


def place_on_sheet(sheet, top, front, iso):
    """Layout: Floor Plan over Elevation View in the left column, 3D Isometric View in the right column.
    Returns a warning string if the layout exceeds the sheet area, else None."""
    xmin, ymin, xmax, ymax = get_sheet_area(sheet)
    xmin += SHEET_MARGIN
    ymin += SHEET_MARGIN
    xmax -= SHEET_MARGIN
    ymax -= SHEET_MARGIN
    cx = (xmin + xmax) / 2.0
    cy = (ymin + ymax) / 2.0

    vps = {}
    for key, v in (("top", top), ("front", front), ("iso", iso)):
        vps[key] = DB.Viewport.Create(doc, sheet.Id, v.Id, DB.XYZ(cx, cy, 0))
    doc.Regenerate()

    size = {}
    for key, vp in vps.items():
        o = vp.GetBoxOutline()
        size[key] = (o.MaximumPoint.X - o.MinimumPoint.X,
                     o.MaximumPoint.Y - o.MinimumPoint.Y)

    left_w = max(size["top"][0], size["front"][0])
    left_h = size["top"][1] + VIEWPORT_GAP + size["front"][1]
    total_w = left_w + VIEWPORT_GAP + size["iso"][0]
    total_h = max(left_h, size["iso"][1])

    top_x = xmin + left_w / 2.0
    top_y = ymax - size["top"][1] / 2.0
    front_x = top_x
    front_y = ymax - size["top"][1] - VIEWPORT_GAP - size["front"][1] / 2.0
    iso_x = xmin + left_w + VIEWPORT_GAP + size["iso"][0] / 2.0
    iso_y = ymax - total_h / 2.0

    vps["top"].SetBoxCenter(DB.XYZ(top_x, top_y, 0))
    vps["front"].SetBoxCenter(DB.XYZ(front_x, front_y, 0))
    vps["iso"].SetBoxCenter(DB.XYZ(iso_x, iso_y, 0))

    if total_w > (xmax - xmin) or total_h > (ymax - ymin):
        return ("Views are larger than the sheet area. Adjust view scale or "
                "crop padding, then reposition the viewports.")
    return None


# ---------------- Main work ----------------
# ---------------- Face preview ----------------
FACE_XAML = """
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Choose the face" Width="780" SizeToContent="Height"
        WindowStartupLocation="CenterScreen" ResizeMode="NoResize"
        FontFamily="Arial" FontSize="13" Background="White" Foreground="Black">
  <StackPanel Margin="18">
    <TextBlock Text="Choose the face for the Elevation View" FontWeight="Bold" FontSize="15"/>
    <TextBlock TextWrapping="Wrap" FontSize="11" Margin="0,4,0,0"
               Text="Pick the side that shows the face of the carrier and its runs, then click Use this face."/>
    <Border Height="3" Background="#FFE716" Margin="0,10,0,10"/>
    <UniformGrid Columns="2">
      <RadioButton x:Name="rb0" GroupName="face" Margin="4">
        <StackPanel>
          <Image x:Name="img0" Height="190" Stretch="Uniform"/>
          <TextBlock x:Name="lbl0" FontWeight="Bold" HorizontalAlignment="Center" Margin="0,4,0,0"/>
        </StackPanel>
      </RadioButton>
      <RadioButton x:Name="rb1" GroupName="face" Margin="4">
        <StackPanel>
          <Image x:Name="img1" Height="190" Stretch="Uniform"/>
          <TextBlock x:Name="lbl1" FontWeight="Bold" HorizontalAlignment="Center" Margin="0,4,0,0"/>
        </StackPanel>
      </RadioButton>
      <RadioButton x:Name="rb2" GroupName="face" Margin="4">
        <StackPanel>
          <Image x:Name="img2" Height="190" Stretch="Uniform"/>
          <TextBlock x:Name="lbl2" FontWeight="Bold" HorizontalAlignment="Center" Margin="0,4,0,0"/>
        </StackPanel>
      </RadioButton>
      <RadioButton x:Name="rb3" GroupName="face" Margin="4">
        <StackPanel>
          <Image x:Name="img3" Height="190" Stretch="Uniform"/>
          <TextBlock x:Name="lbl3" FontWeight="Bold" HorizontalAlignment="Center" Margin="0,4,0,0"/>
        </StackPanel>
      </RadioButton>
    </UniformGrid>
    <TextBlock x:Name="note" TextWrapping="Wrap" FontSize="11" Margin="0,8,0,0"/>
    <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,14,0,0">
      <Button Content="Cancel" Width="90" Height="28" Margin="0,0,8,0"
              Background="White" BorderBrush="Black" IsCancel="True" Click="on_cancel"/>
      <Button Content="Use this face" Width="120" Height="28" FontWeight="Bold"
              Background="#FFE716" BorderBrush="Black" IsDefault="True" Click="on_ok"/>
    </StackPanel>
  </StackPanel>
</Window>
"""


class FaceWindow(forms.WPFWindow):
    def __init__(self, labels, paths):
        forms.WPFWindow.__init__(self, FACE_XAML, literal_string=True)
        self.index = None
        missing = False
        for i in range(4):
            getattr(self, "lbl{0}".format(i)).Text = labels[i]
            path = paths[i]
            if not path:
                missing = True
                continue
            try:
                bmp = BitmapImage()
                bmp.BeginInit()
                bmp.UriSource = Uri(path)
                bmp.CacheOption = BitmapCacheOption.OnLoad   # lets us delete the file afterwards
                bmp.EndInit()
                getattr(self, "img{0}".format(i)).Source = bmp
            except:
                missing = True
        if missing:
            self.note.Text = "A preview could not be drawn for one or more sides. You can still choose by name."
        self.rb0.IsChecked = True

    def on_cancel(self, sender, args):
        self.index = None
        self.Close()

    def on_ok(self, sender, args):
        for i in range(4):
            if getattr(self, "rb{0}".format(i)).IsChecked:
                self.index = i
                break
        self.Close()


def export_face_previews(section_type, center, right, faces, points, dir_sign, assemblies, keep_ids):
    """Build temporary Elevation candidates, export each as a PNG, then roll every
    change back so the model is untouched. Returns (image paths, temp folder)."""
    tmpdir = tempfile.mkdtemp(prefix="vfs_faces_")
    paths = [None] * len(faces)
    tag = str(int(time.time()))
    world_z = DB.XYZ.BasisZ

    tg = DB.TransactionGroup(doc, "Preview faces")
    tg.Start()
    try:
        t = DB.Transaction(doc, "Create preview views")
        try:
            t.Start()
            vids = List[DB.ElementId]()
            tokens = []
            temp_views = []
            for i, (label, z_toward) in enumerate(faces):
                v, status = make_section_checked(section_type, center, right, world_z,
                                                 z_toward, points, dir_sign)
                token = "VFS_FACE_{0}_{1}".format(tag, i)
                v.Name = token
                hide_unwanted_categories(v)
                vids.Add(v.Id)
                tokens.append(token)
                temp_views.append(v)
            if assemblies:
                doc.Regenerate()
                for v in temp_views:
                    hide_other_assemblies(v, keep_ids, assemblies)
            t.Commit()
        except:
            if t.HasStarted():
                t.RollBack()
            raise

        opts = DB.ImageExportOptions()
        opts.ExportRange = DB.ExportRange.SetOfViews
        opts.SetViewsAndSheets(vids)
        opts.FilePath = os.path.join(tmpdir, "face")
        opts.HLRandWFViewsFileType = DB.ImageFileType.PNG
        opts.ShadowViewsFileType = DB.ImageFileType.PNG
        opts.ImageResolution = DB.ImageResolution.DPI_150
        opts.ZoomType = DB.ZoomFitType.FitToPage
        opts.PixelSize = PREVIEW_PIXELS
        doc.ExportImage(opts)

        for f in os.listdir(tmpdir):
            if not f.lower().endswith(".png"):
                continue
            for i, token in enumerate(tokens):
                if token in f:
                    paths[i] = os.path.join(tmpdir, f)
    except:
        pass            # previews are a convenience; fall back to choosing by name
    finally:
        try:
            tg.RollBack()   # removes the temporary views
        except:
            pass
    return paths, tmpdir


def choose_elevation_face(section_type, center, right, faces, points, dir_sign, assemblies, keep_ids):
    """Show the four sides and return the index the user picks."""
    tmpdir = None
    try:
        paths, tmpdir = export_face_previews(section_type, center, right, faces, points,
                                             dir_sign, assemblies, keep_ids)
        window = FaceWindow([f[0] for f in faces], paths)
        window.show_dialog()
        if window.index is None:
            raise UserCancel()
        return window.index
    finally:
        if tmpdir:
            shutil.rmtree(tmpdir, True)


# ---------------- Interactive sheet placement ----------------
def place_views_interactively(sheet, ordered):
    """Switch to the sheet and let the user click where each view goes.
    ordered = [(label, view), ...]. Returns the number of views placed."""
    uidoc.ActiveView = sheet
    native = hasattr(uidoc, "PromptToPlaceViewOnSheet")
    placed = 0
    for label, v in ordered:
        try:
            if native:
                uidoc.PromptToPlaceViewOnSheet(v, False)    # Revit's own placement with live preview
            else:
                pt = uidoc.Selection.PickPoint("Click where to place the {0}".format(label))
                tx = DB.Transaction(doc, "Place view on sheet")
                try:
                    tx.Start()
                    DB.Viewport.Create(doc, sheet.Id, v.Id, pt)
                    tx.Commit()
                except:
                    if tx.HasStarted():
                        tx.RollBack()
                    raise
            placed += 1
        except OperationCanceledException:
            break
    return placed


def auto_place_on_sheet(sheet, views):
    """Automatic layout (used when interactive placement is off or unavailable)."""
    t = DB.Transaction(doc, "Place Views On Sheet")
    try:
        t.Start()
        warn = place_on_sheet(sheet, views["top"], views["front"], views["iso"])
        t.Commit()
        return warn
    except:
        if t.HasStarted():
            t.RollBack()
        raise


# ---------------- Main work ----------------
def run_create(request, form=None):
    """Hide the window while Revit is busy, then run the work."""
    if form is not None:
        form.Hide()
    try:
        return _run_create(request, form)
    finally:
        if form is not None:
            form.Show()


def _run_create(request, form=None):
    """Create the three views from the current selection, optionally place them on
    a sheet, and return (short summary, full details)."""
    start_view = active_view

    elems = get_selected_elements()
    if not elems:
        raise Exception("Select fabrication parts and/or family instances in Revit first, then click Create.")

    points, skipped = get_selection_points(elems)
    if not points:
        raise Exception("No selected element returned a bounding box.")

    section_type = get_view_family_type(DB.ViewFamily.Section)
    if not section_type:
        raise Exception("No Section ViewFamilyType found.")

    vft_3d = get_view_family_type(DB.ViewFamily.ThreeDimensional)
    if not vft_3d:
        raise Exception("No 3D ViewFamilyType found.")

    taken = existing_view_names()
    base_name = request["name"]
    sheet_plan = request["sheet_plan"]

    names = {}
    for key in ("top", "front", "iso"):
        names[key] = unique_name(base_name + VIEW_SUFFIX[key], taken)

    if FILTER_BY_ASSEMBLY:
        assemblies, with_value = get_selected_assemblies(elems)
    else:
        assemblies, with_value = set(), 0

    right, up, ref_pipe = get_pipe_frame(elems, active_view, assemblies)
    world_z = DB.XYZ.BasisZ

    xs = [p.X for p in points]
    ys = [p.Y for p in points]
    zs = [p.Z for p in points]
    center = DB.XYZ((min(xs) + max(xs)) / 2.0,
                    (min(ys) + max(ys)) / 2.0,
                    (min(zs) + max(zs)) / 2.0)

    dir_sign = api_view_sign(active_view)
    checks = {}
    keep_ids = set(str(e.Id) for e in elems)

    # The four sides of the selection. Each vector points from the selection toward
    # the viewer. The Floor Plan and 3D views always use the default frame.
    faces = [("Front", up.Negate()),
             ("Back", up),
             ("Right end", right),
             ("Left end", right.Negate())]
    face_index = 0
    if CHOOSE_FACE_PREVIEW:
        face_index = choose_elevation_face(section_type, center, right, faces, points,
                                           dir_sign, assemblies, keep_ids)
    face_label, front_z = faces[face_index]

    ids = List[DB.ElementId]([e.Id for e in elems])
    views = {}
    hidden_counts = {}

    # Transaction 1: views
    t = DB.Transaction(doc, "Create Views From Selection")
    try:
        t.Start()

        # Floor Plan: looking down. Y = plan up, Z = +world Z
        top, checks["top"] = make_section_checked(section_type, center, right, up, world_z, points, dir_sign)
        top.Name = names["top"]
        apply_template(top, "top")
        views["top"] = top

        # Elevation: Y = world Z, Z toward the viewer (the face the user chose)
        front, checks["front"] = make_section_checked(section_type, center, right, world_z, front_z, points, dir_sign)
        front.Name = names["front"]
        apply_template(front, "front")
        views["front"] = front

        iso = make_iso(vft_3d, center, right, up, points)
        iso.Name = names["iso"]
        apply_template(iso, "iso")
        views["iso"] = iso

        # Title on Sheet: typed name plus the view type, set per view
        for k, v in views.items():
            set_title_on_sheet(v, SHEET_TITLE_FORMAT.format(name=base_name, label=VIEW_LABELS[k]))
            set_scale(v, k)
            hide_unwanted_categories(v)

        if assemblies:
            doc.Regenerate()
            for k, v in views.items():
                hidden_counts[k] = hide_other_assemblies(v, keep_ids, assemblies)

        if ISOLATE_SELECTION:
            for v in views.values():
                try:
                    v.IsolateElementsTemporary(ids)
                except:
                    pass

        t.Commit()
    except:
        if t.HasStarted():
            t.RollBack()
        raise

    # ----- details (only shown when SHOW_DETAILS is True) -----
    msg = "Created views:\n" + "\n".join(
        names[k] for k in ("top", "front", "iso"))
    if FILTER_BY_ASSEMBLY:
        if assemblies:
            msg += "\n\nKept {0}: {1}".format(STRATUS_PARAM, ", ".join(sorted(assemblies)))
            msg += "\n(value read from {0} of {1} selected elements)".format(with_value, len(elems))

            def HC(key, idx=0):
                return hidden_counts.get(key, (0, 0))[idx]
            msg += "\nFabrication parts hidden: Floor Plan {0}, Elevation {1}, 3D {2}".format(
                HC("top"), HC("front"), HC("iso"))
            if FILTER_FAMILIES:
                msg += "\nFamilies hidden: Floor Plan {0}, Elevation {1}, 3D {2}".format(
                    HC("top", 1), HC("front", 1), HC("iso", 1))
        else:
            msg += "\n\nNo '{0}' value found on any of the {1} selected elements, so nothing was hidden by assembly.".format(STRATUS_PARAM, len(elems))
            names_found = similar_param_names(elems)
            if names_found:
                msg += "\nParameters with 'stratus' or 'assembly' in the name: " + "; ".join(names_found)
                msg += "\n(If one is the right one, set STRATUS_PARAM to that exact name. It may also be blank on these elements.)"
            else:
                msg += "\nNo parameter with 'stratus' or 'assembly' in its name was found on the selected elements."
    msg += "\n\nElevation face chosen: {0}.".format(face_label)
    msg += "\nView direction check: Floor Plan {0}, Elevation {1}.".format(
        checks.get("top", "not checked"), checks.get("front", "not checked"))
    if ref_pipe is not None:
        msg += "\n\nAligned to fabrication part Id {0}.".format(ref_pipe.Id)
    else:
        msg += "\n\nNo straight fabrication pipe found. Aligned to the plan orientation."
    if skipped:
        msg += "\n\n{0} selected element(s) had no bounding box and were ignored.".format(skipped)
    msg += "\n\nScript {0}".format(SCRIPT_VERSION)

    # ----- short, user-facing summary -----
    summary = "Created 3 views for {0}. Elevation face: {1}.".format(base_name, face_label)
    if FILTER_BY_ASSEMBLY and not assemblies:
        summary += " No {0} value was found on the selection, so nothing was filtered.".format(STRATUS_PARAM)
    if ref_pipe is None:
        summary += " No straight run was found, so the views follow the plan orientation."

    # ----- sheet -----
    sheet = None
    if sheet_plan:
        try:
            if sheet_plan["mode"] == "existing":
                sheet = sheet_plan["sheet"]
            else:
                t2 = DB.Transaction(doc, "Create Sheet")
                try:
                    t2.Start()
                    sheet = DB.ViewSheet.Create(doc, sheet_plan["tb_id"])
                    sheet.SheetNumber = sheet_plan["number"]
                    sheet.Name = sheet_plan["name"]
                    t2.Commit()
                except:
                    if t2.HasStarted():
                        t2.RollBack()
                    sheet = None
                    raise
        except Exception as sheet_ex:
            summary += " Views were created, but the sheet could not be made: {0}".format(sheet_ex)
            msg += "\n\nSheet error: {0}".format(sheet_ex)

    if sheet is not None:
        sheet_label = "{0} - {1}".format(sheet.SheetNumber, elem_name(sheet))
        ordered = [(VIEW_LABELS["top"], views["top"]),
                   (VIEW_LABELS["front"], views["front"]),
                   (VIEW_LABELS["iso"], views["iso"])]
        done = False
        if PLACE_INTERACTIVE:
            try:
                placed = place_views_interactively(sheet, ordered)
                done = True
                if placed == len(ordered):
                    summary += " Placed on sheet {0}.".format(sheet_label)
                else:
                    summary += (" Placed {0} of {1} views on sheet {2}. Drag the rest from the "
                                "Project Browser.").format(placed, len(ordered), sheet_label)
            except Exception as place_ex:
                msg += "\n\nInteractive placement unavailable: {0}".format(place_ex)
        if not done:
            try:
                warn = auto_place_on_sheet(sheet, views)
                summary += " Placed on sheet {0}.".format(sheet_label)
                if warn:
                    summary += " " + warn
            except Exception as place_ex:
                summary += " Views were created, but placement failed: {0}".format(place_ex)
                msg += "\n\nPlacement error: {0}".format(place_ex)

        # Open the finished sheet, or go back to the starting view
        if OPEN_SHEET_AFTER:
            try:
                uidoc.ActiveView = sheet
            except:
                try:
                    uidoc.RequestViewChange(sheet)
                except:
                    pass
        elif PLACE_INTERACTIVE and start_view is not None:
            try:
                uidoc.ActiveView = start_view
            except:
                pass

    return summary, msg


# ---------------- Start the modeless window ----------------
_handler = None
_event = None
_form = None


def start(uiapp):
    """Called by script.py. Everything stays in this module so it remains loaded
    while the modeless window is open."""
    global _uiapp, _handler, _event, _form
    _uiapp = uiapp
    try:
        refresh_context()
        _handler = ViewsHandler()
        _event = ExternalEvent.Create(_handler)     # must be created in this API context
        _form = ViewsForm(_handler, _event)
        _handler.form = _form
        _form.show(modal=False)
    except Exception as ex:
        TaskDialog.Show("Error", str(ex))


# Run
start(__revit__)