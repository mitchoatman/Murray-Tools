import Autodesk
from Autodesk.Revit.DB import FilteredElementCollector, BoundingBoxXYZ, XYZ, BuiltInCategory, Transaction, ViewSheet, ViewDuplicateOption, Viewport, BuiltInParameter, ElementId
from Autodesk.Revit.UI.Selection import PickBoxStyle
from rpw.ui.forms import (FlexForm, Label, ComboBox, TextBox, TextBox,
                          Separator, Button, CheckBox)
from pyrevit import forms

#define the active Revit application and document
DB = Autodesk.Revit.DB
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView
fec = FilteredElementCollector
app = doc.Application
RevitVersion = app.VersionNumber
RevitINT = float (RevitVersion)

def clear_scope_box_from_view(myview):
    # Get the crop region parameter for the view
    param = myview.get_Parameter(BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP)
    # Set the parameter value to an invalid ElementId (None)
    param.Set(ElementId.InvalidElementId)

#Define dialog options and show it
components = [Label('Sheet Number:'),
    TextBox('sheetnumber', 'SK-001'),
    Label('Sheet Name:'),
    TextBox('sheetname', 'SKETCH 001'),
    Button('Ok')]
form = FlexForm('Title', components)
form.show()

snumber = (form.values['sheetnumber'])
sname = (form.values['sheetname'])

selected_titleblocks = forms.select_titleblocks(title='Select Titleblock', button_name='Select', no_tb_option='No Title Block', width=500, multiple=False, filterfunc=None)

borders = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_TitleBlocks) \
                                                    .WhereElementIsElementType()

titleblockid = borders.FirstElementId()

if str(curview.ViewType) == 'FloorPlan' :
    #prompt user for box and make sure mins are mins and maxs are maxs
    pickedBox = uidoc.Selection.PickBox(PickBoxStyle.Directional, "Draw rectangle area for sketch boundary")
    Maxx = pickedBox.Max.X
    Maxy = pickedBox.Max.Y
    Minx = pickedBox.Min.X
    Miny = pickedBox.Min.Y

    newmaxx = max(Maxx, Minx)
    newmaxy = max(Maxy, Miny)
    newminx = min(Maxx, Minx)
    newminy = min(Maxy, Miny)

    #make bounding box of the points selected
    bbox = BoundingBoxXYZ()
    bbox.Max = XYZ(newmaxx, newmaxy, 0)
    bbox.Min = XYZ(newminx, newminy, 0)

    borders = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_TitleBlocks) \
                                                        .WhereElementIsElementType()

    titleblockid = borders.FirstElementId()

    #define a transaction variable and describe the transaction
    t = Transaction(doc, 'Create Sketch View')

    # Begin new transaction
    t.Start()
    SHEET = ViewSheet.Create(doc, selected_titleblocks);
    SHEET.Name = sname;
    SHEET.SheetNumber = snumber;
    newView = curview.Duplicate(ViewDuplicateOption.WithDetailing)
    getnewview = doc.GetElement(newView)
    clear_scope_box_from_view(getnewview)
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


if str(curview.ViewType) != 'FloorPlan' :
    #define a transaction variable and describe the transaction
    t = Transaction(doc, 'Create Sketch Sheet')

    # Begin new transaction
    t.Start()
    SHEET = ViewSheet.Create(doc, selected_titleblocks);
    SHEET.Name = sname;
    SHEET.SheetNumber = snumber;
    newView = curview.Duplicate(ViewDuplicateOption.WithDetailing)
    getnewview = doc.GetElement(newView)
    clear_scope_box_from_view(getnewview)
    getnewview.Name = sname
    getnewview.CropBoxActive = True
    getnewview.CropBoxVisible = True
    x = SHEET.Outline.Max.Add(SHEET.Outline.Min).Divide(2.0)[0]
    y = SHEET.Outline.Max.Add(SHEET.Outline.Min).Divide(2.0)[1]
    ViewLocation = XYZ(x, y, 0.0)
    NEWSHEET = Viewport.Create(doc, SHEET.Id, newView, ViewLocation)
    t.Commit()
    uidoc.RequestViewChange(SHEET)