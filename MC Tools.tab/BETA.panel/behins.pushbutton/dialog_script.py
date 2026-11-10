# -*- coding: utf-8 -*-

import math
import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.Windows import (
    Window, Thickness, WindowStartupLocation, ResizeMode, HorizontalAlignment
)
from System.Windows.Controls import (
    StackPanel, TextBox, ListBox, Label, Button, DockPanel, Dock
)
from System.Windows.Input import Keyboard, Key

from Autodesk.Revit.DB import (
    Transaction, TransactionGroup, FabricationPart, ConnectorType, XYZ, Line,
    Plane, SketchPlane, ViewType, FilteredElementCollector, ReferencePlane,
    FamilyInstance, FamilyPlacementType, ElementTransformUtils, FamilySymbol, Element
)
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.Exceptions import OperationCanceledException

# False: plane is vertical (rod stays vertical on sloped pipe, ring is off by the slope angle).
# True : plane is exactly perpendicular to the pipe (ring is square to the pipe, rod tilts with slope).
MATCH_SLOPE = False

TOL = 1e-6
UP = XYZ.BasisZ
PLAN_VIEWS = (ViewType.FloorPlan, ViewType.CeilingPlan, ViewType.EngineeringPlan, ViewType.AreaPlan)
ALLOWED_VIEWS = PLAN_VIEWS + (ViewType.ThreeD, ViewType.Section, ViewType.Detail)

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument


def show(msg, title="Place Clamp At Point"):
    TaskDialog.Show(title, msg)


# ==============================================================================
# WPF FAMILY PICKER DIALOG (NO PYREVIT FORMS)
# ==============================================================================
class ClampFamilyPicker(Window):
    def __init__(self, items_dict):
        self.items_dict = items_dict
        self.all_labels = sorted(items_dict.keys())
        self.filtered_labels = list(self.all_labels)
        self.selected_symbol = None

        self.Title = "Select Clamp Family"
        self.Width = 450
        self.Height = 550
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.ResizeMode = ResizeMode.CanResize

        root_panel = DockPanel()
        root_panel.Margin = Thickness(10)

        # OK / Use Button
        self.ok_button = Button()
        self.ok_button.Content = "Use Family"
        self.ok_button.Height = 30
        self.ok_button.Width = 110
        self.ok_button.HorizontalAlignment = HorizontalAlignment.Right
        self.ok_button.Margin = Thickness(0, 10, 0, 0)
        self.ok_button.Click += self.on_ok_click
        DockPanel.SetDock(self.ok_button, Dock.Bottom)
        root_panel.Children.Add(self.ok_button)

        # Top panel
        top_stack = StackPanel()
        DockPanel.SetDock(top_stack, Dock.Top)

        lbl_search = Label()
        lbl_search.Content = "Search Family Type:"
        top_stack.Children.Add(lbl_search)

        self.search_box = TextBox()
        self.search_box.Margin = Thickness(0, 0, 0, 10)
        self.search_box.TextChanged += self.apply_filter
        self.search_box.KeyDown += self.on_search_key_down
        top_stack.Children.Add(self.search_box)

        lbl_instr = Label()
        lbl_instr.Content = "Double-click family to select or press Enter"
        lbl_instr.Margin = Thickness(0, 0, 0, 5)
        top_stack.Children.Add(lbl_instr)

        root_panel.Children.Add(top_stack)

        # ListBox
        self.list_box = ListBox()
        self.list_box.Margin = Thickness(0, 0, 0, 0)
        self.list_box.MouseDoubleClick += self.on_double_click
        self.list_box.KeyDown += self.on_list_key_down
        root_panel.Children.Add(self.list_box)

        self.Content = root_panel
        self.refresh_list()
        self.search_box.Focus()
        Keyboard.Focus(self.search_box)

    def refresh_list(self):
        self.list_box.ItemsSource = self.filtered_labels
        if self.filtered_labels:
            self.list_box.SelectedIndex = 0

    def apply_filter(self, sender, args):
        query = self.search_box.Text.lower().strip()
        if not query:
            self.filtered_labels = list(self.all_labels)
        else:
            self.filtered_labels = [label for label in self.all_labels if query in label.lower()]
        self.refresh_list()

    def confirm_selection(self):
        idx = self.list_box.SelectedIndex
        if idx < 0 or idx >= len(self.filtered_labels):
            show("Please select a valid clamp family from the list.")
            return False
        selected_label = self.filtered_labels[idx]
        self.selected_symbol = self.items_dict[selected_label]
        return True

    def on_double_click(self, sender, args):
        if self.confirm_selection():
            self.DialogResult = True
            self.Close()

    def on_ok_click(self, sender, args):
        if self.confirm_selection():
            self.DialogResult = True
            self.Close()

    def on_search_key_down(self, sender, args):
        if args.Key == Key.Down:
            self.list_box.Focus()
        elif args.Key == Key.Enter:
            if self.confirm_selection():
                self.DialogResult = True
                self.Close()

    def on_list_key_down(self, sender, args):
        if args.Key == Key.Enter:
            if self.confirm_selection():
                self.DialogResult = True
                self.Close()


class StraightFilter(ISelectionFilter):
    def AllowElement(self, e):
        return isinstance(e, FabricationPart) and get_axis(e) is not None

    def AllowReference(self, ref, pt):
        return True


def pick_symbol():
    """List loaded Work Plane-based family types using WPF dialog."""
    items = {}
    for sym in FilteredElementCollector(doc).OfClass(FamilySymbol):
        try:
            if sym.Family.FamilyPlacementType != FamilyPlacementType.WorkPlaneBased:
                continue
        except Exception:
            continue
        items["{} : {}".format(sym.Family.Name, Element.Name.GetValue(sym))] = sym
    if not items:
        show("No Work Plane-based families are loaded in this project.")
        return None

    dlg = ClampFamilyPicker(items)
    if dlg.ShowDialog() and dlg.DialogResult:
        return dlg.selected_symbol
    return None


def get_axis(part):
    """Centerline end points of a straight, from its two farthest End connectors."""
    try:
        if not part.IsAStraight():
            return None
    except Exception:
        return None
    conns = [c for c in part.ConnectorManager.Connectors if c.ConnectorType == ConnectorType.End]
    if len(conns) < 2:
        return None
    best, pair = -1.0, None
    for i in range(len(conns)):
        for j in range(i + 1, len(conns)):
            dist = conns[i].Origin.DistanceTo(conns[j].Origin)
            if dist > best:
                best, pair = dist, (conns[i].Origin, conns[j].Origin)
    return pair if best > TOL else None


def delete_script_reference_planes(doc):
    """Deletes existing TEMPORARY and leftover CLAMP_PLANE_ reference planes."""
    refs_to_delete = []
    collector = FilteredElementCollector(doc).OfClass(ReferencePlane).WhereElementIsNotElementType()

    for rp in collector:
        if rp.Name == "TEMPORARY" or (rp.Name and rp.Name.startswith("CLAMP_PLANE_")):
            refs_to_delete.append(rp.Id)

    for rid in refs_to_delete:
        try:
            doc.Delete(rid)
        except Exception:
            pass


def try_name_sketch_plane_temporary(sketch_plane):
    try:
        sketch_plane.Name = "TEMPORARY"
    except Exception:
        pass


def create_temporary_reference_plane(doc, curview, plane):
    """Creates a single TEMPORARY reference plane for visual placement in the view."""
    bubble_end = plane.Origin
    free_end = plane.Origin + plane.XVec
    cut_vec = plane.YVec

    ref_plane = doc.Create.NewReferencePlane(bubble_end, free_end, cut_vec, curview)
    ref_plane.Name = "TEMPORARY"
    return ref_plane


def turn_off_workplane(curview):
    """Hides active work plane grid display when tool finishes."""
    t = Transaction(doc, "Hide Work Plane")
    t.Start()
    try:
        curview.HideActiveWorkPlane()
        delete_script_reference_planes(doc)
        t.Commit()
    except Exception:
        t.RollBack()


def place_on_plane(curview, symbol, origin, plane, up_vec):
    """Places clamp on active work plane and updates single TEMPORARY reference plane."""
    note = ""
    tg = TransactionGroup(doc, "Place clamp on work plane")
    tg.Start()
    try:
        # Step 1: Set Active View Sketch Plane
        t1 = Transaction(doc, "Set Work Plane")
        t1.Start()
        sp = SketchPlane.Create(doc, plane)
        try_name_sketch_plane_temporary(sp)
        curview.SketchPlane = sp
        curview.ShowActiveWorkPlane()
        t1.Commit()

        # Step 2: Reset TEMPORARY Reference Plane & Place Clamp
        t2 = Transaction(doc, "Reset TEMPORARY Reference Plane & Place Clamp")
        t2.Start()
        
        delete_script_reference_planes(doc)
        rp = create_temporary_reference_plane(doc, curview, plane)

        if not symbol.IsActive:
            symbol.Activate()
            doc.Regenerate()

        # Place family instance on active plane reference
        inst = None
        try:
            inst = doc.Create.NewFamilyInstance(sp.GetPlaneReference(), origin, plane.XVec, symbol)
        except Exception:
            inst = doc.Create.NewFamilyInstance(rp.GetReference(), origin, plane.XVec, symbol)

        doc.Regenerate()

        # Ensure family Y (rod) points up
        tr = inst.GetTransform()
        dot = tr.BasisY.DotProduct(up_vec)
        if dot < -0.5:
            ElementTransformUtils.RotateElement(
                doc, inst.Id, Line.CreateUnbound(origin, tr.BasisZ), math.pi)
            doc.Regenerate()
            dot = inst.GetTransform().BasisY.DotProduct(up_vec)
        if tr.BasisZ.DotProduct(plane.XVec.CrossProduct(plane.YVec)) < 0:
            note = ("Clamp {} faces the opposite way along the pipe (family Z is "
                    "reversed from the plane normal).".format(inst.Id.IntegerValue))
        if abs(dot) < 0.9:
            note = ("Clamp {} placed, but family Y is not up (dot {:.2f}). "
                    "The rod is probably not on family Y.".format(inst.Id.IntegerValue, dot))

        t2.Commit()
        tg.Assimilate()

    except Exception as ex:
        tg.RollBack()
        raise ex

    return note


def main():
    curview = doc.ActiveView
    if curview.ViewType not in ALLOWED_VIEWS:
        show("Run this tool from a 3D, plan, Section or Detail Section view.")
        return
    is_plan = curview.ViewType in PLAN_VIEWS

    symbol = pick_symbol()
    if symbol is None:
        return

    placed, notes = 0, []
    try:
        while True:
            try:
                pref = uidoc.Selection.PickObject(
                    ObjectType.PointOnElement, StraightFilter(),
                    "Pick a point on a straight (Esc to finish)")
            except OperationCanceledException:
                break
            part = doc.GetElement(pref.ElementId)
            axis = get_axis(part)
            if axis is None:
                notes.append("Skipped {}: not a straight.".format(pref.ElementId.IntegerValue))
                continue
            p0, p1 = axis
            d = (p1 - p0).Normalize()
            horiz = XYZ(d.X, d.Y, 0.0)
            if horiz.GetLength() < 1e-3:
                notes.append("Skipped {}: vertical run.".format(pref.ElementId.IntegerValue))
                continue
            horiz = horiz.Normalize()
            t = max(0.0, min(p0.DistanceTo(p1), (pref.GlobalPoint - p0).DotProduct(d)))
            origin = p0 + d.Multiply(t)

            if horiz.X < -1e-6 or (abs(horiz.X) <= 1e-6 and horiz.Y < 0):
                horiz = horiz.Negate()
            if d.DotProduct(horiz) < 0:
                d = d.Negate()

            perp = UP.CrossProduct(horiz).Normalize()
            up_vec = UP
            if MATCH_SLOPE and not is_plan:
                up_vec = d.CrossProduct(perp).Normalize()
                if up_vec.Z < 0:
                    up_vec = up_vec.Negate()
            plane = Plane.CreateByOriginAndBasis(origin, perp, up_vec)
            try:
                note = place_on_plane(curview, symbol, origin, plane, up_vec)
                placed += 1
                if note:
                    notes.append(note)
            except Exception as ex:
                notes.append("Failed on {}: {}".format(pref.ElementId.IntegerValue, ex))
    finally:
        turn_off_workplane(curview)


if __name__ == "__main__":
    main()