import Autodesk
from Autodesk.Revit.DB import FilteredElementCollector, BoundingBoxXYZ, XYZ, BuiltInCategory, Transaction, ViewSheet, ViewDuplicateOption, Viewport, ElementId
from Autodesk.Revit.UI.Selection import PickBoxStyle
import sys
import os

# .NET Imports
import clr
clr.AddReference('System')
clr.AddReference('System.Drawing')
clr.AddReference('System.Windows.Forms')
import System
from System.Windows.Forms import *
from System.Drawing import Point, Size, Font, FontStyle

# define the active Revit application and document
DB = Autodesk.Revit.DB
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView
fec = FilteredElementCollector

folder_name = "c:\\Temp"
filepath = os.path.join(folder_name, 'Ribbon_CreateSketch.txt')

if not os.path.exists(folder_name):
    os.makedirs(folder_name)
if not os.path.exists(filepath):
    with open(filepath, 'w') as the_file:
        line1 = ('SK-001' + '\n')
        line2 = 'SKETCH-001'
        the_file.writelines([line1, line2])

with open(filepath, 'r') as file:
    lines = file.readlines()
    lines = [line.rstrip() for line in lines]

class TXT_Form(Form):
    def __init__(self):
        self.Text = 'Sketch Data'
        self.Size = Size(275, 260)
        self.StartPosition = FormStartPosition.CenterScreen
        self.TopMost = True
        self.ShowIcon = False
        self.MaximizeBox = False
        self.MinimizeBox = False
        self.FormBorderStyle = FormBorderStyle.FixedDialog

        self.snumber = None

        # Label for Sheet Number
        self.label_textbox = Label()
        self.label_textbox.Text = 'Sheet Number:'
        self.label_textbox.ForeColor = System.Drawing.Color.Black
        self.label_textbox.Font = Font("Arial", 12, FontStyle.Bold)
        self.label_textbox.Location = Point(15, 20)
        self.label_textbox.AutoSize = True
        self.Controls.Add(self.label_textbox)

        # TextBox for Sheet Number
        self.textBox1 = TextBox()
        self.textBox1.Text = lines[0]
        self.textBox1.Location = Point(15, 50)
        self.textBox1.Size = Size(150, 20)
        self.Controls.Add(self.textBox1)

        # Label for Sheet Name
        self.label_textbox2 = Label()
        self.label_textbox2.Text = 'Sheet Name:'
        self.label_textbox2.ForeColor = System.Drawing.Color.Black
        self.label_textbox2.Font = Font("Arial", 12, FontStyle.Bold)
        self.label_textbox2.Location = Point(15, 85)
        self.label_textbox2.AutoSize = True
        self.Controls.Add(self.label_textbox2)

        # TextBox for Sheet Name
        self.textBox2 = TextBox()
        self.textBox2.Text = lines[1]
        self.textBox2.Location = Point(15, 115)
        self.textBox2.Size = Size(150, 20)
        self.Controls.Add(self.textBox2)

        # Button
        self.button = Button()
        self.button.Text = 'Set Sketch Data'
        self.button.AutoSize = True
        self.Controls.Add(self.button)
        # Center the button after adding to controls
        self.button.Location = Point((self.ClientSize.Width - self.button.Width) // 2, 160)

        self.button.Click += self.on_click

    def on_click(self, sender, event):
        self.snumber = self.textBox1.Text
        self.sname = self.textBox2.Text
        self.Close()

# Show the Form
form = TXT_Form()
Application.Run(form)

if form.snumber is not None:
    snumber = form.snumber
    sname = form.sname

    # Check if sheet already exists
    sheet_exists = any(sheet.SheetNumber == snumber for sheet in fec(doc).OfClass(ViewSheet).ToElements())
    if sheet_exists:
        print("Error", "Sheet Name or Number {} already exists.".format(snumber))
        sys.exit()

    # Check if view already exists
    view_exists = any(view.Name == sname for view in fec(doc).OfClass(DB.View).ToElements())
    if view_exists:
        print("Error", "View with name {} already exists.".format(sname))
        sys.exit()

    from pyrevit import forms
    selected_titleblocks = forms.select_titleblocks(title='Select Titleblock', button_name='Select', no_tb_option='No Title Block', width=500, multiple=False, filterfunc=None)

    if not selected_titleblocks:
        print("Error", "No title block selected.")
        sys.exit()

    if str(curview.ViewType) == 'FloorPlan':
        # Prompt user for box and make sure mins are mins and maxs are maxs
        pickedBox = uidoc.Selection.PickBox(PickBoxStyle.Directional, "Select area for sketch")
        Maxx = pickedBox.Max.X
        Maxy = pickedBox.Max.Y
        Minx = pickedBox.Min.X
        Miny = pickedBox.Min.Y

        newmaxx = max(Maxx, Minx)
        newmaxy = max(Maxy, Miny)
        newminx = min(Maxx, Minx)
        newminy = min(Maxy, Miny)

        # Make bounding box of the points selected
        bbox = BoundingBoxXYZ()
        bbox.Max = XYZ(newmaxx, newmaxy, 0)
        bbox.Min = XYZ(newminx, newminy, 0)

        # Define a transaction variable and describe the transaction
        t = Transaction(doc, 'Create Sketch')

        # Begin new transaction
        t.Start()
        SHEET = ViewSheet.Create(doc, selected_titleblocks)
        SHEET.Name = sname
        SHEET.SheetNumber = snumber
        newView = curview.Duplicate(ViewDuplicateOption.WithDetailing)
        getnewview = doc.GetElement(newView)
        getnewview.Name = sname
        getnewview.get_Parameter(DB.BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP).Set(ElementId.InvalidElementId)
        getnewview.CropBoxActive = True
        getnewview.CropBoxVisible = True
        getnewview.CropBox = bbox
        x = SHEET.Outline.Max.Add(SHEET.Outline.Min).Divide(2.0)[0]
        y = SHEET.Outline.Max.Add(SHEET.Outline.Min).Divide(2.0)[1]
        ViewLocation = XYZ(x, y, 0.0)
        NEWSHEET = Viewport.Create(doc, SHEET.Id, newView, ViewLocation)
        t.Commit()
        uidoc.RequestViewChange(SHEET)

    else:
        # Define a transaction variable and describe the transaction
        t = Transaction(doc, 'Create Sketch')

        # Begin new transaction
        t.Start()
        SHEET = ViewSheet.Create(doc, selected_titleblocks)
        SHEET.Name = sname
        SHEET.SheetNumber = snumber
        newView = curview.Duplicate(ViewDuplicateOption.WithDetailing)
        getnewview = doc.GetElement(newView)
        getnewview.Name = sname
        getnewview.get_Parameter(DB.BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP).Set(ElementId.InvalidElementId)
        getnewview.CropBoxActive = True
        getnewview.CropBoxVisible = True
        x = SHEET.Outline.Max.Add(SHEET.Outline.Min).Divide(2.0)[0]
        y = SHEET.Outline.Max.Add(SHEET.Outline.Min).Divide(2.0)[1]
        ViewLocation = XYZ(x, y, 0.0)
        NEWSHEET = Viewport.Create(doc, SHEET.Id, newView, ViewLocation)
        t.Commit()
        uidoc.RequestViewChange(SHEET)

    # Update the text file with the new values
    with open(filepath, 'w') as the_file:
        line1 = snumber + '\n'
        line2 = sname + '\n'
        the_file.writelines([line1, line2])