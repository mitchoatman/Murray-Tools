import Autodesk
from Autodesk.Revit.DB import FilteredElementCollector, BoundingBoxXYZ, XYZ, BuiltInCategory, Transaction, ViewSheet, ViewDuplicateOption, Viewport, BuiltInParameter
from Autodesk.Revit.UI.Selection import PickBoxStyle, ISelectionFilter, ObjectType
from rpw.ui.forms import (FlexForm, Label, ComboBox, TextBox, TextBox, Separator, Button, CheckBox)
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

#>>>>>>>>>> GET SELECTED VIEWS
selected_views = forms.select_views(use_selection=True)
if not selected_views:
    forms.alert("No views selected. Please try again.", exitscript = True)

#>>>>>> SELECT TBLOCK
selected_titleblocks = forms.select_titleblocks(title='Select Titleblock', button_name='Select', no_tb_option='No Title Block', width=500, multiple=False, filterfunc=None)

#>>>>>> DEFINE SHEET NUMBER AND NAME
components = [Label('Sheet Number:'),
    TextBox('sheetnumber', 'Sheet Number'),
    Label('Sheet Name:'),
    TextBox('sheetname', 'Sheet Name'),
    Button('Ok')]
form = FlexForm('Title', components)
form.show()

snumber = (form.values['sheetnumber'])
sname = (form.values['sheetname'])

borders = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_TitleBlocks) \
                                                    .WhereElementIsElementType()

titleblockid = borders.FirstElementId()

#define a transaction variable and describe the transaction
t = Transaction(doc, 'Sheet From View')
# Begin new transaction
t.Start()
for view in selected_views:
    SHEET = ViewSheet.Create(doc, selected_titleblocks);
    SHEET.Name = sname;
    SHEET.SheetNumber = snumber;
    x = SHEET.Outline.Max.Add(SHEET.Outline.Min).Divide(2.0)[0]
    y = SHEET.Outline.Max.Add(SHEET.Outline.Min).Divide(2.0)[1]
    ViewLocation = XYZ(x, y, 0.0)
    NEWSHEET = Viewport.Create(doc, SHEET.Id, view.Id, ViewLocation)
t.Commit()
uidoc.RequestViewChange(SHEET)
