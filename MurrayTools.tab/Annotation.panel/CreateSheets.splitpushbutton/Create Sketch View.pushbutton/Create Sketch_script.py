import Autodesk
from Autodesk.Revit.DB import FilteredElementCollector, BoundingBoxXYZ, XYZ, BuiltInCategory, Transaction, ViewSheet, ViewDuplicateOption, Viewport
from Autodesk.Revit.UI.Selection import PickBoxStyle
from rpw.ui.forms import (FlexForm, Label, ComboBox, TextBox, TextBox,
                          Separator, Button, CheckBox)
from pyrevit import forms
from Autodesk.Revit.UI import TaskDialog
import sys

#define the active Revit application and document
DB = Autodesk.Revit.DB
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView
fec = FilteredElementCollector
app = doc.Application
RevitVersion = app.VersionNumber
RevitINT = float(RevitVersion)

# Define dialog options and show it
components = [Label('Sheet Number:'),
              TextBox('sheetnumber', 'SK-001'),
              Label('Sheet Name:'),
              TextBox('sheetname', 'SKETCH 001'),
              Button('Ok')]
form = FlexForm('Title', components)
form.show()

if not form.values:
    sys.exit()  # Exit if form is canceled

snumber = form.values['sheetnumber']
sname = form.values['sheetname']

# Check if sheet already exists
sheet_exists = any(sheet.SheetNumber == snumber for sheet in fec(doc).OfClass(ViewSheet).ToElements())
if sheet_exists:
    TaskDialog.Show("Error", "Sheet with number {} already exists.".format(snumber))
    sys.exit()

# Check if view already exists
view_exists = any(view.Name == sname for view in fec(doc).OfClass(DB.View).ToElements())
if view_exists:
    TaskDialog.Show("Error", "View with name {} already exists.".format(sname))
    sys.exit()

selected_titleblocks = forms.select_titleblocks(title='Select Titleblock', button_name='Select', no_tb_option='No Title Block', width=500, multiple=False, filterfunc=None)

if not selected_titleblocks:
    TaskDialog.Show("Error", "No title block selected.")
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
    t = Transaction(doc, 'This is my new transaction')

    # Begin new transaction
    t.Start()
    SHEET = ViewSheet.Create(doc, selected_titleblocks)
    SHEET.Name = sname
    SHEET.SheetNumber = snumber
    newView = curview.Duplicate(ViewDuplicateOption.WithDetailing)
    getnewview = doc.GetElement(newView)
    getnewview.Name = sname
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
    t = Transaction(doc, 'This is my new transaction')

    # Begin new transaction
    t.Start()
    SHEET = ViewSheet.Create(doc, selected_titleblocks)
    SHEET.Name = sname
    SHEET.SheetNumber = snumber
    newView = curview.Duplicate(ViewDuplicateOption.WithDetailing)
    getnewview = doc.GetElement(newView)
    getnewview.Name = sname
    getnewview.CropBoxActive = True
    getnewview.CropBoxVisible = True
    x = SHEET.Outline.Max.Add(SHEET.Outline.Min).Divide(2.0)[0]
    y = SHEET.Outline.Max.Add(SHEET.Outline.Min).Divide(2.0)[1]
    ViewLocation = XYZ(x, y, 0.0)
    NEWSHEET = Viewport.Create(doc, SHEET.Id, newView, ViewLocation)
    t.Commit()
    uidoc.RequestViewChange(SHEET)
