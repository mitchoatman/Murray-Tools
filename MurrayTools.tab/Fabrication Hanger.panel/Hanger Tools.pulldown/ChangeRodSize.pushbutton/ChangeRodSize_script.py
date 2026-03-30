import clr
from Autodesk.Revit import DB
from Autodesk.Revit.DB import FabricationPart, FabricationAncillaryUsage, Transaction, FilteredElementCollector, BuiltInCategory, TransactionGroup
from Autodesk.Revit.UI.Selection import *
from Autodesk.Revit.UI import TaskDialog
from Parameters.Add_SharedParameters import Shared_Params
from System import Array
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
from System.Windows import Application, Window, Thickness, WindowStyle, ResizeMode, WindowStartupLocation, HorizontalAlignment
from System.Windows.Controls import Label, ComboBox, Button, Grid, RowDefinition

Shared_Params()

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView

def set_parameter_by_name(element, parameterName, value):
    element.LookupParameter(parameterName).Set(value)

class CustomISelectionFilter(ISelectionFilter):
    def __init__(self, nom_categorie):
        self.nom_categorie = nom_categorie
    def AllowElement(self, e):
        if e.Category.Name == self.nom_categorie:
            return True
        else:
            return False
    def AllowReference(self, ref, point):
        return True

try:
    pipesel = uidoc.Selection.PickObjects(ObjectType.Element,
    CustomISelectionFilter("MEP Fabrication Hangers"), "Select Fabrication Hangers")            
    hangers = [doc.GetElement( elId ) for elId in pipesel]

    if len(hangers) > 0:
        class RodSizeForm(Window):
            def __init__(self):
                self.Title = "Select Rod Size"
                self.Width = 200
                self.Height = 150
                self.WindowStyle = WindowStyle.SingleBorderWindow
                self.ResizeMode = ResizeMode.NoResize
                self.WindowStartupLocation = WindowStartupLocation.CenterScreen
                self.result = None

                grid = Grid()
                grid.Margin = Thickness(10)

                for _ in range(3):
                    grid.RowDefinitions.Add(RowDefinition())

                self.Content = grid

                label = Label()
                label.Content = "Select Rod Size:"
                label.Margin = Thickness(0, -5, 0, 10)
                Grid.SetRow(label, 0)
                grid.Children.Add(label)

                self.combo = ComboBox()
                self.combo.ItemsSource = Array[object](['3/8', '1/2', '5/8', '3/4', '7/8', '1', '1-1/4', '1-1/2'])
                self.combo.SelectedIndex = 0  # Default to '3/8'
                self.combo.Width = 150
                self.combo.Margin = Thickness(0, 0, 0, 10)
                Grid.SetRow(self.combo, 1)
                grid.Children.Add(self.combo)

                ok_button = Button()
                ok_button.Content = "OK"
                ok_button.Width = 60
                ok_button.Height = 25
                ok_button.HorizontalAlignment = HorizontalAlignment.Center
                ok_button.Click += self.on_ok
                Grid.SetRow(ok_button, 2)
                grid.Children.Add(ok_button)
                self.combo.Focus()

            def on_ok(self, sender, args):
                self.result = self.combo.SelectedItem
                self.DialogResult = True
                self.Close()
        
        form = RodSizeForm()
        if form.ShowDialog() and form.DialogResult:
            value = form.result

            if value == '3/8':
                newrodkit = 58
            elif value == '1/2':
                newrodkit = 42
            elif value == '5/8':
                newrodkit = 31
            elif value == '3/4':
                newrodkit = 62
            elif value == '7/8':
                newrodkit = 64
            elif value == '1':
                newrodkit = 67
            elif value == '1-1/4':
                newrodkit = 70
            elif value == '1-1/2':
                newrodkit = 143

            tg = TransactionGroup(doc, "Change Hanger Rod")
            tg.Start()

            t = Transaction(doc, "Set Hanger Rod")
            t.Start()
            for hanger in hangers:
                hanger.HangerRodKit = newrodkit
            t.Commit()

            t = Transaction(doc, "Update FP Parameter")
            t.Start()
            for x in hangers:
                [set_parameter_by_name(x, 'FP_Rod Size', n.AncillaryWidthOrDiameter) for n in x.GetPartAncillaryUsage() if n.AncillaryWidthOrDiameter > 0]
            t.Commit()
            
            #End Transaction Group
            tg.Assimilate()
        else:
            TaskDialog.Show("Selection Cancelled", "No rod size selected.")
    else:
        TaskDialog.Show("Selection Error", "At least one fabrication hanger must be selected.")
except:
    TaskDialog.Show("Selection Cancelled", "Operation cancelled by user.")