from Autodesk.Revit import DB
from Autodesk.Revit.DB import FilteredElementCollector, Transaction, BuiltInCategory, FamilySymbol, Family
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
import os
import math
import clr
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
from System.Windows.Forms import Form, Label, TextBox, Button, DialogResult, FormBorderStyle, FormStartPosition
from System.Drawing import Point

path, filename = os.path.split(__file__)
NewFilename = '\Fabrication Pipe - Size Tag.rfa'
ElevationTagFilename = '\Fabrication Pipe - Elevation Tag.rfa'  # Assuming this is the correct file name

app = __revit__.Application
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

class FamilyLoaderOptionsHandler(DB.IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues.Value = False
        return True
    def OnSharedFamilyFound(self, familyInUse, source, overwriteParameterValues):
        source.Value = DB.FamilySource.Family
        overwriteParameterValues.Value = False
        return True

SizeFamilyName = 'Fabrication Pipe - Size Tag'
SizeFamilyType = 'Size Tag'
ElevationFamilyName = 'Fabrication Pipe - Elevation Tag'
ElevationFamilyType = 'BOP'

family_pathCC2 = path + NewFilename
family_pathCC1 = path + ElevationTagFilename

families = FilteredElementCollector(doc).OfClass(Family)
needs_size_loading = not any(f.Name == SizeFamilyName for f in families)
needs_elevation_loading = not any(f.Name == ElevationFamilyName for f in families)

if needs_size_loading:
    t = Transaction(doc, 'Load Pipe Size Family')
    t.Start()
    fload_handler = FamilyLoaderOptionsHandler()
    doc.LoadFamily(family_pathCC2, fload_handler)
    t.Commit()

if needs_elevation_loading:
    t = Transaction(doc, 'Load Elevation Tag Family')
    t.Start()
    fload_handler = FamilyLoaderOptionsHandler()
    doc.LoadFamily(family_pathCC1, fload_handler)
    t.Commit()

# Get size tag symbol
collector = FilteredElementCollector(doc)
collector.OfCategory(BuiltInCategory.OST_FabricationPipeworkTags)
collector.OfClass(FamilySymbol)

size_symbol = None
elevation_symbol = None
for symbol in collector:
    if symbol.Family.Name == SizeFamilyName and symbol.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString() == SizeFamilyType:
        size_symbol = symbol
    if symbol.Family.Name == ElevationFamilyName and symbol.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString() == ElevationFamilyType:
        elevation_symbol = symbol

if not size_symbol:
    raise Exception("Required size tag family type not found")
if not elevation_symbol:
    raise Exception("Required elevation tag family type not found")

if not size_symbol.IsActive:
    t = Transaction(doc, 'Activate Size Family Symbol')
    t.Start()
    size_symbol.Activate()
    t.Commit()

if not elevation_symbol.IsActive:
    t = Transaction(doc, 'Activate Elevation Family Symbol')
    t.Start()
    elevation_symbol.Activate()
    t.Commit()

class SpacingDialog(Form):
    def __init__(self):
        self.Text = "Tag Spacing"
        self.Width = 250
        self.Height = 150
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False
        self.StartPosition = FormStartPosition.CenterScreen
        self.label = Label()
        self.label.Text = "Enter tag spacing in inches\n6.0 for 1/4, 12.0 for 1/8"
        self.label.Location = Point(20, 10)
        self.label.Width = 210
        self.label.Height = 40
        self.Controls.Add(self.label)
        self.textbox = TextBox()
        self.textbox.Text = "6.0"
        self.textbox.Location = Point(20, 45)
        self.textbox.Width = 100
        self.Controls.Add(self.textbox)
        self.ok_button = Button()
        self.ok_button.Text = "OK"
        self.ok_button.Location = Point(45, 80)
        self.ok_button.Click += self.ok_clicked
        self.Controls.Add(self.ok_button)
        self.cancel_button = Button()
        self.cancel_button.Text = "Cancel"
        self.cancel_button.Location = Point(130, 80)
        self.cancel_button.Click += self.cancel_clicked
        self.Controls.Add(self.cancel_button)
    def ok_clicked(self, sender, args):
        self.DialogResult = DialogResult.OK
        self.Close()
    def cancel_clicked(self, sender, args):
        self.DialogResult = DialogResult.Cancel
        self.Close()

try:
    class FabricationPipeFilter(ISelectionFilter):
        def AllowElement(self, element):
            return element.Category.Id.IntegerValue == int(BuiltInCategory.OST_FabricationPipework)
        def AllowReference(self, reference, point):
            return False

    selected_refs = uidoc.Selection.PickObjects(ObjectType.Element, FabricationPipeFilter(), "Select fabrication pipes")
    selected_pipes = [doc.GetElement(ref.ElementId) for ref in selected_refs]
    
    if not selected_pipes:
        raise Exception("No fabrication pipes selected")

    is_x_direction = False
    if selected_pipes:
        connectors = selected_pipes[0].ConnectorManager.Connectors
        if connectors.Size >= 2:
            conn_list = list(connectors)
            direction = (conn_list[1].Origin - conn_list[0].Origin).Normalize()
            is_x_direction = abs(direction.X) > abs(direction.Y)

    dialog = SpacingDialog()
    if dialog.ShowDialog() == DialogResult.OK:
        try:
            spacing_inches = float(dialog.textbox.Text)
        except ValueError:
            spacing_inches = 6.0
    else:
        raise Exception("Operation cancelled by user")

    spacing = DB.UnitUtils.ConvertToInternalUnits(spacing_inches, DB.UnitTypeId.Inches)
    base_point = uidoc.Selection.PickPoint("Select base point for stacked tags")
    
    t = Transaction(doc, 'Place Stacked Tags')
    t.Start()
    
    # Keep your original sorting logic
    if is_x_direction:
        sorted_pipes = sorted(selected_pipes, key=lambda p: list(p.ConnectorManager.Connectors)[0].Origin.Y, reverse=True)
    else:
        sorted_pipes = sorted(selected_pipes, key=lambda p: list(p.ConnectorManager.Connectors)[0].Origin.X)
    
    # Keep your original offsets
    x_offset = 0.0
    y_offset = -spacing

    # Place size tags as in your original
    for i, pipe in enumerate(sorted_pipes):
        tag_position = DB.XYZ(base_point.X + (i * x_offset), base_point.Y + (i * y_offset), base_point.Z)
        tag = DB.IndependentTag.Create(doc, size_symbol.Id, doc.ActiveView.Id, DB.Reference(pipe), False, DB.TagOrientation.Horizontal, tag_position)

    # Place elevation tag following same stacking rules
    num_pipes = len(sorted_pipes)
    elevation_position = DB.XYZ(
        base_point.X + (num_pipes * x_offset),           # Follows x_offset (0 for vertical stacking)
        base_point.Y + (num_pipes * y_offset),           # One more step in the stacking direction
        base_point.Z
    )
    elevation_tag = DB.IndependentTag.Create(
        doc, 
        elevation_symbol.Id, 
        doc.ActiveView.Id, 
        DB.Reference(sorted_pipes[-1]), 
        False, 
        DB.TagOrientation.Horizontal, 
        elevation_position
    )

    t.Commit()

except Exception as e:
    print "Operation cancelled or failed: %s" % str(e)