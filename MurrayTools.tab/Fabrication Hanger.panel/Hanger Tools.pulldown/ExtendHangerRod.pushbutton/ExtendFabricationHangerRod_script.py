# Imports
import Autodesk
from Autodesk.Revit import DB
from Autodesk.Revit.UI import Selection
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, Transaction, ReferencePlane
from Parameters.Get_Set_Params import get_parameter_value_by_name_AsValueString
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

# Selection Filter for Fabrication Hangers
class CustomISelectionFilter(ISelectionFilter):
    def __init__(self, category_name):
        self.category_name = category_name

    def AllowElement(self, e):
        return e.Category and e.Category.Name == self.category_name

    def AllowReference(self, ref, point):
        return False  # Only allow element selection

# Select Fabrication Hangers
pipesel = uidoc.Selection.PickObjects(ObjectType.Element,
                                      CustomISelectionFilter("MEP Fabrication Hangers"),
                                      "Select Fabrication Hangers to Extend")
Hanger = [doc.GetElement(elId) for elId in pipesel]

# Select a Reference Plane
class ReferencePlaneSelectionFilter(ISelectionFilter):
    def AllowElement(self, elem):
        return isinstance(elem, ReferencePlane)  # Only allow Reference Planes

    def AllowReference(self, ref, point):
        return False

ref_plane_ref = uidoc.Selection.PickObject(ObjectType.Element, ReferencePlaneSelectionFilter(),
                                           "Select a reference plane")
ref_plane = doc.GetElement(ref_plane_ref.ElementId)

# Get the plane geometry
plane = ref_plane.GetPlane()
plane_normal = plane.Normal.Normalize()  # Ensure the normal is a unit vector
plane_origin = plane.Origin

# Start transaction
t = Transaction(doc, 'Extend Hanger Rods')
t.Start()

for e in Hanger:
    rod_info = e.GetRodInfo()
    rod_count = rod_info.RodCount
    rod_info.CanRodsBeHosted = False # Detach rods from structure
    HangerType = get_parameter_value_by_name_AsValueString(e, 'Family')

    # Ensure the plane normal always points upward (positive Z) regardless of drawing direction
    normal = plane_normal
    if normal.Z < 0:
        normal = -normal

    # Signed distance from rod end to plane (positive = rod end above plane, negative = below)
    if 'Strap' in HangerType and rod_count > 1:
        rod_len = rod_info.GetRodLength(0)
        rod_pos = rod_info.GetRodEndPosition(0)

        vec = rod_pos - plane_origin
        dist = normal.DotProduct(vec)          # signed distance
        new_length = rod_len - dist            # correct adjustment

        rod_info.SetRodLength(0, new_length)

    else:
        for n in range(rod_count):
            rod_len = rod_info.GetRodLength(n)
            rod_pos = rod_info.GetRodEndPosition(n)

            vec = rod_pos - plane_origin
            dist = normal.DotProduct(vec)
            new_length = rod_len - dist        # correct adjustment

            rod_info.SetRodLength(n, new_length)

t.Commit()