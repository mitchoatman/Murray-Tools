# Imports
import Autodesk
from Autodesk.Revit import DB
from Autodesk.Revit.UI import Selection, TaskDialog
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, Transaction, ReferencePlane
import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
from System.Windows import Application, Window, Thickness, WindowStyle, ResizeMode, WindowStartupLocation, HorizontalAlignment
from System.Windows.Controls import Label, Button, Grid, RowDefinition, ColumnDefinition

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView

# Yes/No Dialog for Reference Plane
class ConfirmationDialog(Window):
    def __init__(self):
        self.Title = "Confirmation"
        self.Width = 300
        self.Height = 150
        self.WindowStyle = WindowStyle.SingleBorderWindow
        self.ResizeMode = ResizeMode.NoResize
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.result = None

        grid = Grid()
        grid.Margin = Thickness(10)
        grid.RowDefinitions.Add(RowDefinition())
        grid.RowDefinitions.Add(RowDefinition())
        grid.ColumnDefinitions.Add(ColumnDefinition())
        grid.ColumnDefinitions.Add(ColumnDefinition())
        self.Content = grid

        label = Label()
        label.Content = "Have you created a Reference Plane?"
        label.Margin = Thickness(0, 0, 0, 10)
        Grid.SetRow(label, 0)
        Grid.SetColumnSpan(label, 2)
        grid.Children.Add(label)

        yes_button = Button()
        yes_button.Content = "Yes"
        yes_button.Width = 60
        yes_button.Height = 25
        yes_button.Margin = Thickness(0, 0, 10, 0)
        yes_button.HorizontalAlignment = HorizontalAlignment.Right
        yes_button.Click += self.on_yes
        Grid.SetRow(yes_button, 1)
        Grid.SetColumn(yes_button, 0)
        grid.Children.Add(yes_button)

        no_button = Button()
        no_button.Content = "No"
        no_button.Width = 60
        no_button.Height = 25
        no_button.HorizontalAlignment = HorizontalAlignment.Left
        no_button.Click += self.on_no
        Grid.SetRow(no_button, 1)
        Grid.SetColumn(no_button, 1)
        grid.Children.Add(no_button)

    def on_yes(self, sender, args):
        self.result = True
        self.DialogResult = True
        self.Close()

    def on_no(self, sender, args):
        self.result = False
        self.DialogResult = True
        self.Close()

# Show confirmation dialog
form = ConfirmationDialog()
if not (form.ShowDialog() and form.result):
    # Exit quietly if "No" is selected or dialog is cancelled
    import sys
    sys.exit()

# Selection Filter for MEP Fabrication Pipework + All Thread Rod family
class CustomISelectionFilter(ISelectionFilter):
    def AllowElement(self, e):
        if not e.Category:
            return False
        if e.Category.Name != "MEP Fabrication Pipework":
            return False
        fam_param = e.LookupParameter("Family")
        return fam_param and fam_param.AsValueString() == "All Thread Rod"

    def AllowReference(self, ref, point):
        return False  # Only allow element selection

# Select All Thread Rods
try:
    pipesel = uidoc.Selection.PickObjects(ObjectType.Element,
                                          CustomISelectionFilter(),
                                          "Select 'All Thread Rod' Pipework to Extend")
    rods = [doc.GetElement(elId) for elId in pipesel]
except Autodesk.Revit.Exceptions.OperationCanceledException:
    TaskDialog.Show("Selection Cancelled", "Selection Cancelled by User.")
    import sys
    sys.exit()

# Select a Reference Plane
class ReferencePlaneSelectionFilter(ISelectionFilter):
    def AllowElement(self, elem):
        return isinstance(elem, ReferencePlane)

    def AllowReference(self, ref, point):
        return False

try:
    ref_plane_ref = uidoc.Selection.PickObject(ObjectType.Element, ReferencePlaneSelectionFilter(),
                                               "Select a reference plane")
    ref_plane = doc.GetElement(ref_plane_ref.ElementId)
except Autodesk.Revit.Exceptions.OperationCanceledException:
    TaskDialog.Show("Selection Cancelled", "Selection Cancelled by User.")
    import sys
    sys.exit()

# Get the plane geometry
plane = ref_plane.GetPlane()
plane_normal = plane.Normal.Normalize()
plane_origin = plane.Origin

# Start transaction
t = Transaction(doc, 'Extend Rods to Reference Plane')
t.Start()

for e in rods:
    # Use bounding box top point as rod endpoint
    bbox = e.get_BoundingBox(None)
    if not bbox:
        continue

    rod_top = bbox.Max
    rod_bot = bbox.Min
    current_length = rod_top.Z - rod_bot.Z

    # Vector from plane to rod top
    vec_to_plane = rod_top - plane_origin
    distance_to_plane = plane_normal.DotProduct(vec_to_plane)
    intersection_point = rod_top - (distance_to_plane * plane_normal)
    delta_length = intersection_point.DistanceTo(rod_top)

    # Extend or shorten based on side of plane
    if distance_to_plane > 0:
        new_length = current_length - delta_length
    else:
        new_length = current_length + delta_length

    # Set new Length param
    length_param = e.LookupParameter("Length")
    if length_param and not length_param.IsReadOnly:
        length_param.Set(new_length)
    else:
        print("Can't set Length on", e.Id)

t.Commit()