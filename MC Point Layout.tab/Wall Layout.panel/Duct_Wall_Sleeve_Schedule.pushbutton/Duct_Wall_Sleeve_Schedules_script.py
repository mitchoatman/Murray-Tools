import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('RevitServices')

from Autodesk.Revit.DB import (
    BuiltInCategory, Transaction, ElementId, ViewSchedule,
    FilteredElementCollector, ParameterElement, ScheduleFieldType,
    BuiltInParameter, ScheduleFilter, ScheduleFilterType,
    FormatOptions, UnitTypeId
)
from Autodesk.Revit.UI import TaskDialog
import System
import os

# Import WPF libraries for custom image display
import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
from System.Windows import Window, Thickness, HorizontalAlignment, VerticalAlignment
from System.Windows.Controls import StackPanel, Image, Button, TextBlock, ScrollViewer
from System.Windows.Media.Imaging import BitmapImage
from System.IO import Path

doc = __revit__.ActiveUIDocument.Document
app = __revit__.Application

def schedule_exists(schedule_name, category_id):
    schedules_collector = FilteredElementCollector(doc).OfClass(ViewSchedule)
    for schedule in schedules_collector:
        if schedule.Name == schedule_name and schedule.Definition.CategoryId == category_id:
            return True
    return False

roundFieldNames = [
    ("TS_Point_Number", "ITEM NO"),
    ("FP_Product Entry", "DIAMETER"),
    ("Diameter", "SIZE (With Annular Space)"),
    ("Elevation from Level", "CL Elevation"),
    ("Family", "SLEEVE TYPE DR-WS=DROP WS=THRU"),
    ("FP_Service Abbreviation", "SYSTEM ABBR."),
    ("FP_Service Name", "SERVICE NAME"),
    ("Comments", "COMMENTS")
]

rectFieldNames = [
    ("TS_Point_Number", "ITEM NO"),
    ("Width", "WIDTH"),
    ("Height", "HEIGHT"),
    ("Elevation from Level", "CL Elevation"),
    ("Family", "SLEEVE TYPE DR-WS=DROP WS=THRU"),
    ("FP_Service Abbreviation", "SYSTEM ABBR."),
    ("FP_Service Name", "SERVICE NAME"),
    ("Comments", "COMMENTS")
]

revit_version = int(app.VersionNumber)
is_revit_2022_or_newer = revit_version >= 2022

categoryId = ElementId(BuiltInCategory.OST_DuctAccessory)

schedules = [
    {"name": "ROUND WALL SLEEVE SCHEDULE", "fields": roundFieldNames, "filter": "RDS", "category": categoryId},
    {"name": "RECTANGLE WALL SLEEVE SCHEDULE", "fields": rectFieldNames, "filter": "RWS", "category": categoryId}
]

def set_fractional_inches_1_8(field):
    fmt = field.GetFormatOptions()
    fmt.UseDefault = False
    fmt.SetUnitTypeId(UnitTypeId.FractionalInches)
    fmt.Accuracy = 0.125  # 1/8"
    field.SetFormatOptions(fmt)

def add_field_by_name(definition, paramName, userColumnName, parameters):
    field = None

    if paramName == "Family":
        paramId = ElementId(BuiltInParameter.ELEM_FAMILY_PARAM)
        field = definition.AddField(ScheduleFieldType.Instance, paramId)
    elif paramName == "Elevation from Level":
        paramId = ElementId(BuiltInParameter.INSTANCE_ELEVATION_PARAM)
        field = definition.AddField(ScheduleFieldType.Instance, paramId)
    elif paramName == "Comments":
        paramId = ElementId(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
        field = definition.AddField(ScheduleFieldType.Instance, paramId)
    else:
        parameter = next((p for p in parameters if p.Name == paramName), None)
        if parameter is not None:
            paramId = parameter.Id
            field = definition.AddField(ScheduleFieldType.Instance, paramId)
        else:
            TaskDialog.Show("Warning", "Parameter '{}' not found for schedule.".format(paramName))
            return None

    field.ColumnHeading = userColumnName

    if userColumnName in ["WIDTH", "HEIGHT", "SIZE (With Annular Space)"]:
        set_fractional_inches_1_8(field)

    return field

t = Transaction(doc, "Create Schedules")
t.Start()

parameters = FilteredElementCollector(doc).OfClass(ParameterElement).ToElements()

for schedule_info in schedules:
    schedule_name = schedule_info["name"]
    fieldNames = schedule_info["fields"]
    family_filter = schedule_info["filter"]
    categoryId = schedule_info["category"]

    if not schedule_exists(schedule_name, categoryId):
        schedule = ViewSchedule.CreateSchedule(doc, categoryId)
        schedule.Name = schedule_name
        definition = schedule.Definition

        family_field = None
        for paramName, userColumnName in fieldNames:
            field = add_field_by_name(definition, paramName, userColumnName, parameters)
            if paramName == "Family" and field is not None:
                family_field = field

        if is_revit_2022_or_newer and family_field is not None:
            schedule_filter = ScheduleFilter(
                family_field.FieldId,
                ScheduleFilterType.EndsWith,
                family_filter
            )
            definition.AddFilter(schedule_filter)
    else:
        TaskDialog.Show("Schedule Exists", "'{}' already exists.".format(schedule_name))

t.Commit()

# --- WPF Dialog with Multiple Images Support ---
class MultiImageDialog(Window):
    def __init__(self, img0_path, img1_path, message):
        super(MultiImageDialog, self).__init__()
        self.Title = "Schedules Created"
        self.Width = 500
        self.Height = 650
        self.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen
        
        # Outer container with scrolling enabled in case images are tall
        scroll = ScrollViewer()
        scroll.VerticalScrollBarVisibility = System.Windows.Controls.ScrollBarVisibility.Auto
        
        panel = StackPanel()
        panel.Margin = Thickness(15)
        
        # Text Message
        text_block = TextBlock()
        text_block.Text = message
        text_block.FontSize = 14
        text_block.Margin = Thickness(0, 0, 0, 15)
        text_block.TextWrapping = System.Windows.TextWrapping.Wrap
        panel.Children.Add(text_block)
        
        # Helper function to add images safely
        def add_image(path):
            if os.path.exists(path):
                img = Image()
                bitmap = BitmapImage()
                bitmap.BeginInit()
                bitmap.UriSource = System.Uri(path)
                bitmap.EndInit()
                img.Source = bitmap
                img.MaxHeight = 220
                img.Margin = Thickness(0, 0, 0, 10)
                img.HorizontalAlignment = HorizontalAlignment.Center
                panel.Children.Add(img)
            else:
                err_block = TextBlock()
                err_block.Text = "[Missing: {}]".format(os.path.basename(path))
                err_block.Foreground = System.Windows.Media.Brushes.Red
                err_block.Margin = Thickness(0, 0, 0, 10)
                panel.Children.Add(err_block)

        # Add both images
        add_image(img0_path)
        add_image(img1_path)
            
        # Close Button
        btn = Button()
        btn.Content = "OK"
        btn.Width = 100
        btn.Height = 30
        btn.Margin = Thickness(0, 10, 0, 0)
        btn.HorizontalAlignment = HorizontalAlignment.Center
        btn.Click += self.close_click
        panel.Children.Add(btn)
        
        scroll.Content = panel
        self.Content = scroll

    def close_click(self, sender, e):
        self.Close()

# Determine script folder dynamically
try:
    script_dir = os.path.dirname(__file__)
except NameError:
    script_dir = os.getcwd()

img0_file = os.path.join(script_dir, "image0.png")
img1_file = os.path.join(script_dir, "image1.png")

# Launch the WPF window with multiple images
msg = "Schedules created. Reference stickers below for necessary fields."
dialog = MultiImageDialog(img0_file, img1_file, msg)
dialog.ShowDialog()