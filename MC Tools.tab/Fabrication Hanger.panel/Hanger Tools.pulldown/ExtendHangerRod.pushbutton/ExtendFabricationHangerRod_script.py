# Imports
import Autodesk
from Autodesk.Revit import DB
from Autodesk.Revit.UI import Selection
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from Autodesk.Revit.DB import FilteredElementCollector, Transaction, ReferencePlane
from Autodesk.Revit.Exceptions import OperationCanceledException
from Parameters.Get_Set_Params import get_parameter_value_by_name_AsValueString
import clr
import sys

clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

from System.Windows import (
    Window, Thickness, WindowStyle,
    ResizeMode, WindowStartupLocation, HorizontalAlignment
)
from System.Windows.Controls import Label, Button, Grid, RowDefinition, ColumnDefinition
from System.Windows.Input import Key

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView


def active_view_has_visible_reference_plane(doc, view):
    """Return True if at least one Reference Plane is visible in the active view."""
    ref_planes = FilteredElementCollector(doc, view.Id).OfClass(ReferencePlane)
    return ref_planes.GetElementCount() > 0


class ProceedCancelDialog(Window):
    def __init__(self):
        self.Title = "Reference Plane Not Visible"
        self.Width = 360
        self.Height = 170
        self.WindowStyle = WindowStyle.SingleBorderWindow
        self.ResizeMode = ResizeMode.NoResize
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.result = False

        self.PreviewKeyDown += self.on_key_down

        grid = Grid()
        grid.Margin = Thickness(10)
        grid.RowDefinitions.Add(RowDefinition())
        grid.RowDefinitions.Add(RowDefinition())
        grid.ColumnDefinitions.Add(ColumnDefinition())
        grid.ColumnDefinitions.Add(ColumnDefinition())
        self.Content = grid

        label = Label()
        label.Content = "No visible Reference Plane was found in the active view.\n\nProceed anyway?"
        label.Margin = Thickness(0, 0, 0, 10)
        Grid.SetRow(label, 0)
        Grid.SetColumnSpan(label, 2)
        grid.Children.Add(label)

        self.proceed_button = Button()
        self.proceed_button.Content = "Proceed"
        self.proceed_button.Width = 80
        self.proceed_button.Height = 25
        self.proceed_button.Margin = Thickness(0, 0, 10, 0)
        self.proceed_button.HorizontalAlignment = HorizontalAlignment.Right
        self.proceed_button.Click += self.on_proceed
        self.proceed_button.IsDefault = True
        Grid.SetRow(self.proceed_button, 1)
        Grid.SetColumn(self.proceed_button, 0)
        grid.Children.Add(self.proceed_button)

        self.cancel_button = Button()
        self.cancel_button.Content = "Cancel"
        self.cancel_button.Width = 80
        self.cancel_button.Height = 25
        self.cancel_button.HorizontalAlignment = HorizontalAlignment.Left
        self.cancel_button.Click += self.on_cancel
        self.cancel_button.IsCancel = True
        Grid.SetRow(self.cancel_button, 1)
        Grid.SetColumn(self.cancel_button, 1)
        grid.Children.Add(self.cancel_button)

    def on_key_down(self, sender, e):
        if e.Key == Key.Enter or e.Key == Key.Space:
            self.on_proceed(None, None)
            e.Handled = True
        elif e.Key == Key.Escape:
            self.on_cancel(None, None)
            e.Handled = True

    def on_proceed(self, sender, args):
        self.result = True
        self.DialogResult = True
        self.Close()

    def on_cancel(self, sender, args):
        self.result = False
        self.DialogResult = False
        self.Close()


# Only prompt if no visible Reference Plane exists in active view
if not active_view_has_visible_reference_plane(doc, curview):
    form = ProceedCancelDialog()
    if not form.ShowDialog():
        sys.exit()


# Selection Filter for Fabrication Hangers
class CustomISelectionFilter(ISelectionFilter):
    def __init__(self, category_name):
        self.category_name = category_name

    def AllowElement(self, e):
        return e.Category and e.Category.Name == self.category_name

    def AllowReference(self, ref, point):
        return False


# Select Fabrication Hangers
try:
    pipesel = uidoc.Selection.PickObjects(
        ObjectType.Element,
        CustomISelectionFilter("MEP Fabrication Hangers"),
        "Select Fabrication Hangers to Extend"
    )
except OperationCanceledException:
    sys.exit()

Hanger = [doc.GetElement(elId) for elId in pipesel]


# Select a Reference Plane
class ReferencePlaneSelectionFilter(ISelectionFilter):
    def AllowElement(self, elem):
        return isinstance(elem, ReferencePlane)

    def AllowReference(self, ref, point):
        return False


try:
    ref_plane_ref = uidoc.Selection.PickObject(
        ObjectType.Element,
        ReferencePlaneSelectionFilter(),
        "Select a reference plane"
    )
except OperationCanceledException:
    sys.exit()

ref_plane = doc.GetElement(ref_plane_ref.ElementId)

# Get the plane geometry
plane = ref_plane.GetPlane()
plane_normal = plane.Normal.Normalize()
plane_origin = plane.Origin

# Ensure plane normal points upward
normal = plane_normal
if normal.Z < 0:
    normal = -normal

# Start transaction
t = Transaction(doc, 'Extend Hanger Rods')
t.Start()

for e in Hanger:
    rod_info = e.GetRodInfo()
    rod_count = rod_info.RodCount
    rod_info.CanRodsBeHosted = False  # Detach rods from structure
    HangerType = get_parameter_value_by_name_AsValueString(e, 'Family')

    if 'Strap' in HangerType and rod_count > 1:
        rod_len = rod_info.GetRodLength(0)
        rod_pos = rod_info.GetRodEndPosition(0)

        vec = rod_pos - plane_origin
        dist = normal.DotProduct(vec)
        new_length = rod_len - dist

        rod_info.SetRodLength(0, new_length)

    else:
        for n in range(rod_count):
            rod_len = rod_info.GetRodLength(n)
            rod_pos = rod_info.GetRodEndPosition(n)

            vec = rod_pos - plane_origin
            dist = normal.DotProduct(vec)
            new_length = rod_len - dist

            rod_info.SetRodLength(n, new_length)

t.Commit()