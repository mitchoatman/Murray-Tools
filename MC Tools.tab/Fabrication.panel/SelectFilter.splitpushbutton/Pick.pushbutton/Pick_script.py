import Autodesk
from collections import namedtuple
from System.Collections.Generic import List
from Autodesk.Revit import DB
from Autodesk.Revit.DB import FilteredElementCollector, ElementId
from Autodesk.Revit.UI import Selection
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.UI import TaskDialog
import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference('System.Xaml')
import System
from System.Windows.Controls.Primitives import UniformGrid
from System.Windows import Window, Thickness
from System.Windows.Controls import Label, TextBox, Button, ScrollViewer, StackPanel, Grid, Orientation
from System.Windows.Media import Brushes, FontFamily
from System.Windows import SizeToContent, ResizeMode, HorizontalAlignment, VerticalAlignment, GridLength, GridUnitType

# Categories to exclude
CAT_EXCLUDED = (
    ElementId(DB.BuiltInCategory.OST_RoomSeparationLines),
    ElementId(DB.BuiltInCategory.OST_Cameras),
    ElementId(DB.BuiltInCategory.OST_CurtainGrids),
    ElementId(DB.BuiltInCategory.OST_Elev),
    ElementId(DB.BuiltInCategory.OST_IOSModelGroups),
    ElementId(DB.BuiltInCategory.OST_SitePropertyLineSegment),
    ElementId(DB.BuiltInCategory.OST_SectionBox),
    ElementId(DB.BuiltInCategory.OST_ShaftOpening),
    ElementId(DB.BuiltInCategory.OST_BeamAnalytical),
    ElementId(DB.BuiltInCategory.OST_StructuralFramingOpening),
    ElementId(DB.BuiltInCategory.OST_MEPSpaceSeparationLines),
    ElementId(DB.BuiltInCategory.OST_DuctSystem),
    ElementId(DB.BuiltInCategory.OST_PipingSystem),
    ElementId(DB.BuiltInCategory.OST_CenterLines),
    ElementId(DB.BuiltInCategory.OST_PipeCurvesCenterLine),
    ElementId(DB.BuiltInCategory.OST_FabricationPipeworkCenterLine),
    ElementId(DB.BuiltInCategory.OST_FabricationPipeworkDrop),
    ElementId(DB.BuiltInCategory.OST_FabricationPipeworkRise),
    ElementId(DB.BuiltInCategory.OST_FabricationPipeworkInsulation),
    ElementId(DB.BuiltInCategory.OST_CurtainGridsRoof),
    ElementId(DB.BuiltInCategory.OST_SWallRectOpening),
    ElementId(System.Int64(-2000278)),  # Explicitly cast to Int64
    ElementId(System.Int64(-1)),        # Explicitly cast to Int64
)

CategoryOption = namedtuple('CategoryOption', ['name', 'revit_cat'])

class PickByCategorySelectionFilter(ISelectionFilter):
    def __init__(self, category_opts):
        self.category_ids = [cat_opt.revit_cat.Id for cat_opt in category_opts]

    def AllowElement(self, element):
        if element.Category and element.Category.Id in self.category_ids:
            return True
        return False

    def AllowReference(self, refer, point):
        return False

def get_categories_in_active_view(doc, view):
    collector = FilteredElementCollector(doc, view.Id)
    elements = collector.WhereElementIsNotElementType().ToElements()
    category_ids = set()
    category_dict = {}
    for element in elements:
        if (element.Category and element.Category.Name and 
            element.Category.Id not in CAT_EXCLUDED and 
            element.Category.Id not in category_ids):
            category_ids.add(element.Category.Id)
            category_dict[element.Category.Id] = element.Category
    return sorted(category_dict.values(), key=lambda x: x.Name)

def pick_by_categories(category_opts, uidoc):
    if not category_opts:
        return

    msfilter = PickByCategorySelectionFilter(category_opts)
    
    try:
        selection_list = uidoc.Selection.PickElementsByRectangle(msfilter)
    except Autodesk.Revit.Exceptions.OperationCanceledException:
        return
    except Exception as ex:
        return

    if selection_list:
        filtered_ids = List[ElementId]([e.Id for e in selection_list])
        uidoc.Selection.SetElementIds(filtered_ids)

class FilterSelectionByCategory(Window):
    def __init__(self, category_list):
        self.selected_categories = []
        self.category_list = sorted(category_list, key=lambda x: x.Name)
        self.checkboxes = []
        self.check_all_state = False
        self.InitializeComponents()

    def InitializeComponents(self):
        self.Title = "Filter Selection by Category"
        self.Width = 400
        self.Height = 400
        self.MinWidth = self.Width
        self.MinHeight = self.Height
        self.ResizeMode = ResizeMode.NoResize
        self.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen

        grid = Grid()
        grid.Margin = Thickness(5)
        for i in range(4):  # rows for: label, search box, scroll, buttons
            row = GridLength(1, GridUnitType.Star) if i == 2 else GridLength.Auto
            grid.RowDefinitions.Add(System.Windows.Controls.RowDefinition(Height=row))
        grid.ColumnDefinitions.Add(System.Windows.Controls.ColumnDefinition())

        # Row 0 - Label
        self.label = Label(Content="Search and select categories:")
        self.label.FontFamily = FontFamily("Arial")
        self.label.FontSize = 16
        self.label.Margin = Thickness(0)
        Grid.SetRow(self.label, 0)
        grid.Children.Add(self.label)

        # Row 1 - Search Box
        self.search_box = TextBox(Height=20, FontFamily=FontFamily("Arial"), FontSize=12)
        self.search_box.TextChanged += self.search_changed
        Grid.SetRow(self.search_box, 1)
        grid.Children.Add(self.search_box)

        # Row 2 - Scrollable Checkbox Panel
        self.checkbox_panel = StackPanel(Orientation=System.Windows.Controls.Orientation.Vertical)
        scroll_viewer = ScrollViewer(Content=self.checkbox_panel, VerticalScrollBarVisibility=System.Windows.Controls.ScrollBarVisibility.Auto)
        scroll_viewer.Margin = Thickness(0, 1, 0, 1)
        Grid.SetRow(scroll_viewer, 2)
        grid.Children.Add(scroll_viewer)

        self.update_checkboxes(self.category_list)

        # Row 3 - Button Panel
        button_panel = StackPanel(Orientation=System.Windows.Controls.Orientation.Horizontal, HorizontalAlignment=HorizontalAlignment.Center, Margin=Thickness(0, 10, 0, 10))

        self.select_button = Button(Content="Select", FontFamily=FontFamily("Arial"), FontSize=12, Height=25, Margin=Thickness(0, 0, 20, 0), Width=50)
        self.select_button.Click += self.select_clicked
        button_panel.Children.Add(self.select_button)

        self.check_all_button = Button(Content="Check All", FontFamily=FontFamily("Arial"), FontSize=12, Height=25, Width=70)
        self.check_all_button.Click += self.check_all_clicked
        button_panel.Children.Add(self.check_all_button)

        Grid.SetRow(button_panel, 3)
        grid.Children.Add(button_panel)

        # Set window content
        self.Content = grid
        self.SizeChanged += self.on_resize

    def update_checkboxes(self, categories):
        self.checkbox_panel.Children.Clear()
        self.checkboxes = []
        for category in categories:
            checkbox = System.Windows.Controls.CheckBox(Content=category.Name.upper())
            checkbox.Tag = category   # keep actual category object here
            checkbox.Click += self.checkbox_clicked
            if category.Name in self.selected_categories:
                checkbox.IsChecked = True
            self.checkbox_panel.Children.Add(checkbox)
            self.checkboxes.append(checkbox)

    def check_all_clicked(self, sender, args):
        self.check_all_state = not self.check_all_state
        for cb in self.checkboxes:
            cb.IsChecked = self.check_all_state
        self.selected_categories = [cb.Tag.Name for cb in self.checkboxes if cb.IsChecked]

    def checkbox_clicked(self, sender, args):
        self.selected_categories = [cb.Tag.Name for cb in self.checkboxes if cb.IsChecked]

    def select_clicked(self, sender, args):
        self.selected_categories = [cb.Tag.Name for cb in self.checkboxes if cb.IsChecked]
        self.DialogResult = True
        self.Close()

    def search_changed(self, sender, args):
        search_text = self.search_box.Text.lower()
        filtered = [c for c in self.category_list if search_text in c.Name.lower()]
        self.update_checkboxes(filtered)

    def on_resize(self, sender, args):
        pass


uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document
source_categories = get_categories_in_active_view(doc, doc.ActiveView)

if source_categories:
    form = FilterSelectionByCategory(source_categories)
    if form.ShowDialog():
        category_opts = [CategoryOption(name=x.Name, revit_cat=x) 
                         for x in source_categories if x.Name in form.selected_categories]
        pick_by_categories(category_opts, uidoc)
else:
    TaskDialog.Show("Error", "No categories found in the current view.")