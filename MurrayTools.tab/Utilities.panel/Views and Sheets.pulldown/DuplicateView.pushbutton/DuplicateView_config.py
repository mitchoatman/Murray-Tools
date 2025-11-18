# coding: utf8
from System.Collections.Generic import List
from Autodesk.Revit import DB
from Autodesk.Revit.DB import FilteredElementCollector, ElementId, Transaction
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
from collections import namedtuple

app = __revit__.Application
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
active_view = doc.ActiveView

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

def get_scope_boxes_in_document(doc):
    collector = FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_VolumeOfInterest).WhereElementIsNotElementType().ToElements()
    return sorted(collector, key=lambda x: x.Name)

def get_views_in_document(doc):
    collector = FilteredElementCollector(doc).OfClass(DB.View).WhereElementIsNotElementType().ToElements()
    return sorted([v for v in collector if v.CanBePrinted], key=lambda x: x.Name)

class FilterSelectionByCategory(Window):
    def __init__(self, item_list, title, label_text, is_view_selection=False):
        self.selected_items = []
        self.item_list = sorted(item_list, key=lambda x: x.Name)
        self.checkboxes = []
        self.check_all_state = False
        self.title = title
        self.label_text = label_text
        self.is_view_selection = is_view_selection
        self.InitializeComponents()

    def InitializeComponents(self):
        self.Title = self.title
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
        self.label = Label(Content=self.label_text)
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

        self.update_checkboxes(self.item_list)

        # Row 3 - Button Panel
        button_panel = StackPanel(Orientation=System.Windows.Controls.Orientation.Horizontal, HorizontalAlignment=HorizontalAlignment.Center, Margin=Thickness(0, 10, 0, 10))

        self.select_button = Button(Content="Select", FontFamily=FontFamily("Arial"), FontSize=12, Height=25, Margin=Thickness(10, 0, 10, 0))
        self.select_button.Click += self.select_clicked
        button_panel.Children.Add(self.select_button)

        self.check_all_button = Button(Content="Check All", FontFamily=FontFamily("Arial"), FontSize=12, Height=25, Margin=Thickness(10, 0, 10, 0))
        self.check_all_button.Click += self.check_all_clicked
        button_panel.Children.Add(self.check_all_button)

        Grid.SetRow(button_panel, 3)
        grid.Children.Add(button_panel)

        # Set window content
        self.Content = grid
        self.SizeChanged += self.on_resize

    def update_checkboxes(self, items):
        self.checkbox_panel.Children.Clear()
        self.checkboxes = []
        for item in items:
            # Display view name with view type for view selection
            display_name = "{} ({})".format(item.Name, item.ViewType) if self.is_view_selection else item.Name
            checkbox = System.Windows.Controls.CheckBox(Content=display_name)
            checkbox.Tag = item
            checkbox.Click += self.checkbox_clicked
            if display_name in self.selected_items:
                checkbox.IsChecked = True
            self.checkbox_panel.Children.Add(checkbox)
            self.checkboxes.append(checkbox)

    def search_changed(self, sender, args):
        search_text = self.search_box.Text.lower()
        filtered = [c for c in self.item_list if search_text in c.Name.lower()]
        self.update_checkboxes(filtered)

    def check_all_clicked(self, sender, args):
        self.check_all_state = not self.check_all_state
        for cb in self.checkboxes:
            cb.IsChecked = self.check_all_state
        self.selected_items = [cb.Content for cb in self.checkboxes if cb.IsChecked]

    def checkbox_clicked(self, sender, args):
        self.selected_items = [cb.Content for cb in self.checkboxes if cb.IsChecked]

    def select_clicked(self, sender, args):
        self.selected_items = [cb.Content for cb in self.checkboxes if cb.IsChecked]
        self.DialogResult = True
        self.Close()

    def on_resize(self, sender, args):
        pass

# Select scope boxes (Volumes of Interest)
scope_boxes = get_scope_boxes_in_document(doc)
if scope_boxes:
    form = FilterSelectionByCategory(scope_boxes, "Select Scope Boxes for Dependent Views", "Search and select scope boxes:")
    if form.ShowDialog():
        volumes_of_interest = [x for x in scope_boxes if x.Name in form.selected_items]
else:
    TaskDialog.Show("Error", "No scope boxes found in the document.")
    volumes_of_interest = []

# Select views to duplicate as dependent views
# Check for pre-selected views
selected_view_ids = uidoc.Selection.GetElementIds()
views_to_duplicate = []
if selected_view_ids:
    for view_id in selected_view_ids:
        element = doc.GetElement(view_id)
        if isinstance(element, DB.View) and element.CanBePrinted:
            views_to_duplicate.append(element)

# If no views are selected, prompt with dialog
if not views_to_duplicate:
    views = get_views_in_document(doc)
    if views:
        form = FilterSelectionByCategory(views, "Select Views to Create Dependent Views", "Search and select views:", is_view_selection=True)
        if form.ShowDialog():
            views_to_duplicate = [x for x in views if "{} ({})".format(x.Name, x.ViewType) in form.selected_items]
    else:
        TaskDialog.Show("Error", "No printable views found in the document.")
        views_to_duplicate = []

def create_dependent_views(views_to_duplicate, volumes_of_interest):
    for view in views_to_duplicate:
        for voi in volumes_of_interest:
            try:
                # Duplicate the view as a dependent view
                new_view_id = view.Duplicate(DB.ViewDuplicateOption.AsDependent)
                new_view = doc.GetElement(new_view_id)
                
                # Set the volume of interest to the new dependent view
                parameter = new_view.get_Parameter(
                    DB.BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP
                )
                if parameter:
                    parameter.Set(voi.Id)
                else:
                    TaskDialog.Show("Warning", "Volume of interest parameter not found for view {}.".format(view.Name))
                    doc.Delete(new_view_id)  # Delete the view if parameter setting fails
                    continue
                
                # Set the dependent view name to the combination of the original view name and scope box name
                try:
                    new_view.Name = "{} - {}".format(view.Name, voi.Name)
                except Exception as name_err:
                    if "Name must be unique" in str(name_err):
                        TaskDialog.Show("Error", "Cannot create view '{} - {}': A view with this name already exists.".format(view.Name, voi.Name))
                        doc.Delete(new_view_id)  # Delete the view if name is duplicate
                        continue
                    else:
                        raise
            except AttributeError as e:
                TaskDialog.Show("Error", "Failed to duplicate view {}: {}".format(view.Name, str(e)))

# Create dependent views if selections are made
if volumes_of_interest and views_to_duplicate:
    trans = Transaction(doc, "BatchCreateDependentViews")
    try:
        trans.Start()
        create_dependent_views(views_to_duplicate, volumes_of_interest)
        trans.Commit()
    except Exception as e:
        trans.RollBack()
        TaskDialog.Show("Error", "Transaction failed: {}".format(str(e)))