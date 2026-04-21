from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.DB import BoundingBoxXYZ, FilteredElementCollector, Transaction, BuiltInCategory, FabricationPart
from pyrevit import forms
import math

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView

# .NET Imports
import clr
clr.AddReference('System')
clr.AddReference('System.Drawing')
clr.AddReference('System.Windows.Forms')
import System
from System.Windows.Forms import *
from System.Drawing import Point, Size, Font, FontStyle


class TXT_Form(Form):
    def __init__(self):
        self.Text          = 'Fuzz Distance'
        self.Size          = Size(250,150)
        self.StartPosition = FormStartPosition.CenterScreen

        self.TopMost     = True  # Keeps the form on top of all other windows
        self.ShowIcon    = False  # Removes the icon from the title bar
        self.MaximizeBox = False
        self.MinimizeBox = False
        self.FormBorderStyle = FormBorderStyle.FixedDialog  # Disallows resizing

        # Initialize valuenum to None
        self.valuenum = None

        #Design of dialog

        #Label for TextBox
        self.label_textbox           = Label()
        self.label_textbox.Text      = 'Fuzz:'
        self.label_textbox.ForeColor = System.Drawing.Color.Black
        self.label_textbox.Font      = Font("Arial", 12, FontStyle.Bold)
        self.label_textbox.Location  = Point(20,20)
        self.label_textbox.Size      = Size(90,40)
        self.Controls.Add(self.label_textbox)

        #TextBox
        self.textBox          = TextBox()
        self.textBox.Text     = '0.0625'
        self.textBox.Location = Point(110, 20)
        self.textBox.Size     = Size(100, 40)
        self.Controls.Add(self.textBox)

        #Button
        self.button          = Button()
        self.button.Text     = 'Set Distance'
        self.button.Location = Point(68, 60)
        self.button.Size     = Size(100, 30)
        self.Controls.Add(self.button)

        self.button.Click      += self.on_click
        # self.button.MouseEnter += self.btn_hover

    def on_click(self, sender, event):
        try:
            self.value = self.textBox.Text
            self.valuenum = float(self.value) / 12
        except ValueError:
            MessageBox.Show("Enter amount of fuzz distance in decimal inches.")
        self.Close()

#Show the Form
form = TXT_Form()
# form.Show()
Application.Run(form)

def GetCenterPoint(ele):
    bBox = doc.GetElement(ele).get_BoundingBox(None)
    center = (bBox.Max + bBox.Min) / 2
    return (center.X, center.Y, center.Z)

def calculate_distance(point1, point2):
    return math.sqrt((point1[0] - point2[0])**2 +
                     (point1[1] - point2[1])**2 +
                     (point1[2] - point2[2])**2)

# Fuzz distance in Revit units (0.25 inches = 0.020833333 feet)
fuzz_distance = form.valuenum

# Create a FilteredElementCollector to get all FabricationPart elements
AllElements = FilteredElementCollector(doc).OfClass(FabricationPart) \
                   .WhereElementIsNotElementType() \
                   .ToElements()

# Get the center point of each selected element
element_ids = []
center_points = []

for reference in AllElements:
    center_point = GetCenterPoint(reference.Id)
    center_points.append(center_point)
    element_ids.append(reference.Id)

# Find the duplicates in the list of center points
duplicates = []
duplicate_element_ids = []
unique_center_points = []

for i, cp in enumerate(center_points):
    found_duplicate = False
    for ucp in unique_center_points:
        if calculate_distance(cp, ucp) <= fuzz_distance:
            found_duplicate = True
            break
    if not found_duplicate:
        unique_center_points.append(cp)
    else:
        duplicates.append(cp)
        duplicate_element_ids.append(element_ids[i])

# Delete the elements that belong to duplicate center points
try:
    if duplicates:
        forms.alert_ifnot(len(duplicates) < 0,
                          ("Delete Duplicate(s): {}".format(len(duplicates))),
                          yes=True, no=True, exitscript=True)
        
        with Transaction(doc, "Delete Elements") as transaction:
            transaction.Start()
            for element_id in duplicate_element_ids:
                doc.Delete(element_id)
            transaction.Commit()
    else:
        forms.show_balloon('Duplicates', 'No Duplicates Found')

except:
    pass
